package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"strings"
)

const officeFile = "LibreOffice-Win-x64.msi"

func unpack(executable, kind, target string) error {
	archive, err := zip.OpenReader(executable)
	if err != nil {
		return err
	}
	defer archive.Close()

	count := 0
	for _, entry := range archive.File {
		name := entry.Name
		if kind == "app" {
			if !strings.HasPrefix(name, "app/") {
				continue
			}
			name = strings.TrimPrefix(name, "app/")
		} else if name != officeFile {
			continue
		}

		if !entry.Mode().IsRegular() || !filepath.IsLocal(filepath.FromSlash(name)) ||
			name != path.Clean(name) || strings.ContainsAny(name, "\\:") {
			return fmt.Errorf("arquivo inválido no instalador: %s", name)
		}
		if err := extract(entry, filepath.Join(target, filepath.FromSlash(name))); err != nil {
			return err
		}
		count++
	}
	if count == 0 {
		return fmt.Errorf("conteúdo de instalação ausente: %s", kind)
	}
	return nil
}

func extract(entry *zip.File, target string) error {
	if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
		return err
	}
	input, err := entry.Open()
	if err != nil {
		return err
	}
	defer input.Close()
	output, err := os.OpenFile(target, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0644)
	if err != nil {
		return err
	}
	_, copyErr := io.Copy(output, input)
	closeErr := output.Close()
	if copyErr != nil {
		return copyErr
	}
	return closeErr
}
