package main

import (
	"crypto/sha256"
	"encoding/hex"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func TestReleasePayload(t *testing.T) {
	installer := os.Getenv("TRIBUTOS_INSTALLER_TEST_EXE")
	if installer == "" {
		t.Skip("defina TRIBUTOS_INSTALLER_TEST_EXE para conferir um instalador compilado")
	}
	target := t.TempDir()
	if err := unpack(installer, "app", target); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{
		"Tributos-Docs.exe", "app.py", "tributos/assets/desdobro.docx",
		"runtime/pythonw.exe", "runtime/Lib/os.py", "runtime/DLLs/_sqlite3.pyd",
		"runtime/DLLs/_tkinter.pyd", "runtime/tcl/tcl8.6/init.tcl", "runtime/tcl/tk8.6/tk.tcl",
		"runtime/Lib/site-packages/lxml/__init__.py",
		"runtime/Lib/site-packages/PIL/__init__.py",
		"runtime/Lib/site-packages/pypdfium2_raw/pdfium.dll",
	} {
		if info, err := os.Stat(filepath.Join(target, name)); err != nil || info.Size() == 0 {
			t.Fatalf("payload ausente ou vazio: %s (%v)", name, err)
		}
	}
	if err := unpack(installer, "office", target); err != nil {
		t.Fatal(err)
	}
	file, err := os.Open(filepath.Join(target, officeFile))
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()
	hash := sha256.New()
	if _, err := io.Copy(hash, file); err != nil {
		t.Fatal(err)
	}
	if actual := hex.EncodeToString(hash.Sum(nil)); actual != "480fecf759fc6393bd614fbf43753eb95a8198fc1a065d41b727e3b664cd1e59" {
		t.Fatalf("SHA-256 inesperado no MSI: %s", actual)
	}
}
