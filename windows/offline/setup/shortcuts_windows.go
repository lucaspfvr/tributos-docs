package main

import (
	"fmt"
	"path/filepath"
	"runtime"
	"syscall"
	"unsafe"
)

type guid struct {
	data1        uint32
	data2, data3 uint16
	data4        [8]byte
}

var (
	shellLinkClass = guid{0x00021401, 0, 0, [8]byte{0xc0, 0, 0, 0, 0, 0, 0, 0x46}}
	shellLinkW     = guid{0x000214f9, 0, 0, [8]byte{0xc0, 0, 0, 0, 0, 0, 0, 0x46}}
	persistFile    = guid{0x0000010b, 0, 0, [8]byte{0xc0, 0, 0, 0, 0, 0, 0, 0x46}}
)

type comObject struct {
	methods *[21]uintptr
}

//go:uintptrescapes
func (obj *comObject) call(method int, args ...uintptr) uintptr {
	result, _, _ := syscall.SyscallN(obj.methods[method], append([]uintptr{uintptr(unsafe.Pointer(obj))}, args...)...)
	return result
}

func failed(result uintptr) bool {
	return int32(result) < 0
}

func createShortcuts(executable string) error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	ole := syscall.NewLazyDLL("ole32.dll")
	result, _, _ := ole.NewProc("CoInitializeEx").Call(0, 2)
	if failed(result) {
		return fmt.Errorf("não foi possível iniciar a criação de atalhos (%#x)", uint32(result))
	}
	defer ole.NewProc("CoUninitialize").Call()
	var link *comObject
	result, _, _ = ole.NewProc("CoCreateInstance").Call(
		uintptr(unsafe.Pointer(&shellLinkClass)), 0, 1,
		uintptr(unsafe.Pointer(&shellLinkW)), uintptr(unsafe.Pointer(&link)),
	)
	if failed(result) {
		return fmt.Errorf("não foi possível criar o atalho (%#x)", uint32(result))
	}
	defer link.call(2)                                                                      // IUnknown::Release
	if result := link.call(20, uintptr(unsafe.Pointer(wide(executable)))); failed(result) { // SetPath
		return fmt.Errorf("não foi possível definir o destino do atalho (%#x)", uint32(result))
	}
	link.call(9, uintptr(unsafe.Pointer(wide(filepath.Dir(executable))))) // SetWorkingDirectory
	var file *comObject
	if result := link.call(0, uintptr(unsafe.Pointer(&persistFile)), uintptr(unsafe.Pointer(&file))); failed(result) {
		return fmt.Errorf("não foi possível preparar o arquivo de atalho (%#x)", uint32(result))
	}
	defer file.call(2)
	getFolder := syscall.NewLazyDLL("shell32.dll").NewProc("SHGetFolderPathW")
	var lastErr error
	for _, location := range []uintptr{0x10, 0x02} { // CSIDL_DESKTOPDIRECTORY, CSIDL_PROGRAMS
		var buffer [260]uint16
		result, _, _ := getFolder.Call(0, location, 0, 0, uintptr(unsafe.Pointer(&buffer[0])))
		if failed(result) {
			lastErr = fmt.Errorf("pasta de atalho indisponível (%#x)", uint32(result))
			continue
		}
		path := filepath.Join(syscall.UTF16ToString(buffer[:]), "Tributos Docs.lnk")
		if result := file.call(6, uintptr(unsafe.Pointer(wide(path))), 1); failed(result) { // IPersistFile::Save
			lastErr = fmt.Errorf("não foi possível salvar o atalho (%#x)", uint32(result))
		}
	}
	return lastErr
}
