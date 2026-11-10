package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"syscall"
	"time"
	"unsafe"
)

const (
	aptosURL    = "https://download.microsoft.com/download/8/6/0/860a94fa-7feb-44ef-ac79-c072d9113d69/Microsoft%20Aptos%20Fonts.zip"
	aptosHash   = "6528fd120e719a9f985e94214eca6887d1653b88456916a792a630b02e95b025"
	aptosZIP    = "Microsoft Aptos Fonts.zip"
	fontKey     = `SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`
	maxAptosZIP = 5 << 20
)

type aptosFont struct {
	file, name string
}

var aptosFonts = []aptosFont{
	{"Aptos.ttf", "Aptos (TrueType)"},
	{"Aptos-Bold.ttf", "Aptos Bold (TrueType)"},
	{"Aptos-Italic.ttf", "Aptos Italic (TrueType)"},
	{"Aptos-Bold-Italic.ttf", "Aptos Bold Italic (TrueType)"},
}

func aptosInstalled() bool {
	for _, font := range aptosFonts {
		if !fontRegistered(font.name) {
			return false
		}
	}
	return true
}

func fontRegistered(name string) bool {
	for _, hive := range []syscall.Handle{syscall.HKEY_CURRENT_USER, syscall.HKEY_LOCAL_MACHINE} {
		var key syscall.Handle
		if syscall.RegOpenKeyEx(hive, wide(fontKey), 0, syscall.KEY_READ, &key) != nil {
			continue
		}
		var kind, size uint32
		err := syscall.RegQueryValueEx(key, wide(name), nil, &kind, nil, &size)
		syscall.RegCloseKey(key)
		if err == nil && kind == syscall.REG_SZ && size > 2 {
			return true
		}
	}
	return false
}

func installAptos(executable, local string) error {
	file, err := os.Open(filepath.Join(filepath.Dir(executable), aptosZIP))
	var source io.ReadCloser
	switch {
	case err == nil:
		source = file
	case os.IsNotExist(err):
		client := http.Client{Timeout: 40 * time.Second}
		response, err := client.Get(aptosURL)
		if err != nil {
			return fmt.Errorf("download oficial da Aptos indisponível: %w", err)
		}
		if response.StatusCode != http.StatusOK {
			response.Body.Close()
			return fmt.Errorf("download oficial da Aptos indisponível (HTTP %d)", response.StatusCode)
		}
		source = response.Body
	default:
		return err
	}
	defer source.Close()
	data, err := io.ReadAll(io.LimitReader(source, maxAptosZIP+1))
	if err != nil {
		return err
	}
	if len(data) > maxAptosZIP || fmt.Sprintf("%x", sha256.Sum256(data)) != aptosHash {
		return fmt.Errorf("o ZIP não corresponde à edição oficial da Aptos 4.40")
	}
	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	fonts := filepath.Join(local, "Microsoft", "Windows", "Fonts")
	if err := os.MkdirAll(fonts, 0755); err != nil {
		return err
	}
	advapi := syscall.NewLazyDLL("advapi32.dll")
	createKey := advapi.NewProc("RegCreateKeyExW")
	setValue := advapi.NewProc("RegSetValueExW")
	var key syscall.Handle
	result, _, _ := createKey.Call(uintptr(syscall.HKEY_CURRENT_USER), uintptr(unsafe.Pointer(wide(fontKey))),
		0, 0, 0, syscall.KEY_SET_VALUE, 0, uintptr(unsafe.Pointer(&key)), 0)
	if result != 0 {
		return syscall.Errno(result)
	}
	defer syscall.RegCloseKey(key)
	addFont := syscall.NewLazyDLL("gdi32.dll").NewProc("AddFontResourceW")
	for _, font := range aptosFonts {
		if fontRegistered(font.name) {
			continue
		}
		entry, err := archive.Open(font.file)
		if err != nil {
			return err
		}
		content, err := io.ReadAll(entry)
		entry.Close()
		if err != nil {
			return err
		}
		target := filepath.Join(fonts, font.file)
		if err := os.WriteFile(target, content, 0644); err != nil {
			return err
		}
		value, err := syscall.UTF16FromString(target)
		if err != nil {
			return err
		}
		result, _, _ := setValue.Call(uintptr(key), uintptr(unsafe.Pointer(wide(font.name))), 0, syscall.REG_SZ,
			uintptr(unsafe.Pointer(&value[0])), uintptr(len(value)*2))
		if result != 0 {
			return syscall.Errno(result)
		}
		addFont.Call(uintptr(unsafe.Pointer(wide(target))))
	}
	syscall.NewLazyDLL("user32.dll").NewProc("SendNotifyMessageW").Call(0xffff, 0x001d, 0, 0)
	return nil
}
