package main

import (
	"os"
	"os/exec"
	"path/filepath"
	"syscall"
	"unsafe"
)

func main() {
	executable, err := os.Executable()
	if err != nil {
		showError("Não foi possível localizar o Tributos Docs.")
		return
	}
	base := filepath.Dir(executable)
	python := filepath.Join(base, "runtime", "pythonw.exe")
	command := exec.Command(python, "-E", "-s", filepath.Join(base, "app.py"))
	command.Dir = base
	command.Env = append(os.Environ(),
		"TCL_LIBRARY="+filepath.Join(base, "runtime", "tcl", "tcl8.6"),
		"TK_LIBRARY="+filepath.Join(base, "runtime", "tcl", "tk8.6"),
	)
	if err := command.Run(); err != nil {
		showError("Não foi possível abrir o Tributos Docs. Execute novamente o instalador para reparar os arquivos do aplicativo.")
	}
}

func showError(body string) {
	dialog := syscall.NewLazyDLL("user32.dll").NewProc("MessageBoxW")
	text, _ := syscall.UTF16PtrFromString(body)
	title, _ := syscall.UTF16PtrFromString("Tributos Docs")
	dialog.Call(0, uintptr(unsafe.Pointer(text)), uintptr(unsafe.Pointer(title)), 0x10)
}
