package main

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"syscall"
	"unsafe"
)

const (
	okCancel   = 0x01
	yesNo      = 0x04
	iconError  = 0x10
	iconInfo   = 0x40
	defaultNo  = 0x100
	foreground = 0x10000
	topmost    = 0x40000
	confirmed  = 1
	yes        = 6
)

func main() {
	body := "Instalar o Tributos Docs neste computador?\n\nO pacote já contém o Python e as bibliotecas do aplicativo."
	body += "\n\nO pacote inclui o LibreOffice para quem precisar de um conversor de PDF e oferece a instalação da fonte Aptos oficial."
	if dialog(body, okCancel|iconInfo) != confirmed {
		return
	}
	if err := install(wordInstalled(), libreOfficeInstalled()); err != nil {
		dialog("A instalação foi interrompida.\n\n"+err.Error(), iconError)
	}
}

func install(word, libreOffice bool) error {
	executable, err := os.Executable()
	if err != nil {
		return fmt.Errorf("não foi possível localizar o arquivo do instalador: %w", err)
	}
	local := os.Getenv("LOCALAPPDATA")
	if local == "" {
		return errors.New("a pasta do usuário não foi localizada")
	}
	programs := filepath.Join(local, "Programs")
	target := filepath.Join(programs, "Tributos Docs")
	if err := os.MkdirAll(programs, 0755); err != nil {
		return err
	}
	stage, err := os.MkdirTemp(programs, "TributosDocs-setup-")
	if err != nil {
		return err
	}
	defer os.RemoveAll(stage)
	if err := unpack(executable, "app", stage); err != nil {
		return fmt.Errorf("não foi possível extrair o aplicativo. Verifique o espaço livre e o download: %w", err)
	}

	installOffice := !word && !libreOffice
	if word && !libreOffice {
		installOffice = dialog("O Microsoft Word foi detectado e pode gerar os PDFs. Deseja instalar também o LibreOffice incluído no pacote?", yesNo|defaultNo|iconInfo) == yes
	}
	restart := false
	if installOffice {
		restart, err = installLibreOffice(executable)
		if err != nil {
			return err
		}
	}

	if err := replaceInstallation(stage, target); err != nil {
		return err
	}
	launcher := filepath.Join(target, "Tributos-Docs.exe")
	result := "O Tributos Docs foi instalado."
	if !aptosInstalled() {
		if dialog("A Avaliação Municipal usa a fonte Aptos. Deseja instalar a fonte agora?\n\n"+
			"O instalador usa o ZIP oficial ao lado deste .exe, se estiver presente, ou baixa a fonte da Microsoft.",
			yesNo|iconInfo) == yes {
			if err := installAptos(executable, local); err != nil {
				result += "\n\nA fonte Aptos ainda precisa ser instalada. " + err.Error()
			} else {
				result += "\n\nA fonte Aptos foi instalada. Se o Word ou o Writer estiver aberto, feche e abra novamente."
			}
		} else {
			result += "\n\nA fonte Aptos poderá ser instalada executando este instalador novamente."
		}
	}
	if err := createShortcuts(launcher); err != nil {
		result += "\n\nUm atalho não pôde ser criado. O aplicativo está em " + launcher
	} else {
		result += "\n\nOs atalhos estão na área de trabalho e no menu Iniciar."
	}
	if restart {
		result += "\n\nO LibreOffice solicitou a reinicialização do Windows."
	}
	if dialog(result+"\n\nDeseja abrir o Tributos Docs agora?", yesNo|iconInfo) == yes {
		if err := exec.Command(launcher).Start(); err != nil {
			return fmt.Errorf("o aplicativo foi instalado, mas não pôde ser aberto: %w", err)
		}
	}
	return nil
}

func replaceInstallation(stage, target string) error {
	backup := stage + "-anterior"
	_, err := os.Stat(target)
	replacing := err == nil
	if err != nil && !os.IsNotExist(err) {
		return err
	}
	if replacing {
		if err := os.Rename(target, backup); err != nil {
			return errors.New("feche o Tributos Docs antes de iniciar a instalação novamente")
		}
	}
	if err := os.Rename(stage, target); err != nil {
		if replacing {
			if restoreErr := os.Rename(backup, target); restoreErr != nil {
				return fmt.Errorf("a instalação anterior foi mantida em %s. Erro: %w", backup, restoreErr)
			}
		}
		return err
	}
	if replacing {
		_ = os.RemoveAll(backup)
	}
	return nil
}

func wordInstalled() bool {
	name := wide("Word.Application\\CLSID")
	for _, view := range []uint32{0, 0x0200} {
		var key syscall.Handle
		if syscall.RegOpenKeyEx(syscall.HKEY_CLASSES_ROOT, name, 0, syscall.KEY_READ|view, &key) == nil {
			syscall.RegCloseKey(key)
			return true
		}
	}
	return false
}

func libreOfficeInstalled() bool {
	for _, root := range []string{os.Getenv("ProgramFiles"), os.Getenv("ProgramFiles(x86)")} {
		if root != "" {
			if info, err := os.Stat(filepath.Join(root, "LibreOffice", "program", "soffice.exe")); err == nil && !info.IsDir() {
				return true
			}
		}
	}
	return false
}

type shellExecuteInfo struct {
	size, mask                        uint32
	window                            uintptr
	verb, file, parameters, directory *uint16
	show                              int32
	instance, idList                  uintptr
	class                             *uint16
	classKey                          uintptr
	hotKey                            uint32
	icon                              uintptr
	process                           syscall.Handle
}

func installLibreOffice(executable string) (bool, error) {
	temp, err := os.MkdirTemp("", "TributosDocs-office-")
	if err != nil {
		return false, err
	}
	defer os.RemoveAll(temp)
	if err := unpack(executable, "office", temp); err != nil {
		return false, fmt.Errorf("não foi possível extrair o instalador do LibreOffice: %w", err)
	}
	msiexec := filepath.Join(os.Getenv("SystemRoot"), "System32", "msiexec.exe")
	info := shellExecuteInfo{
		mask:       0x40 | 0x100, // SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC
		verb:       wide("runas"),
		file:       wide(msiexec),
		parameters: wide("/i " + syscall.EscapeArg(filepath.Join(temp, officeFile)) + " /passive /norestart"),
		show:       1,
	}
	info.size = uint32(unsafe.Sizeof(info))
	result, _, callErr := syscall.NewLazyDLL("shell32.dll").NewProc("ShellExecuteExW").Call(uintptr(unsafe.Pointer(&info)))
	if result == 0 {
		return false, fmt.Errorf("o Windows não iniciou a instalação do LibreOffice: %w", callErr)
	}
	if info.process == 0 {
		return false, errors.New("o Windows não retornou o processo de instalação do LibreOffice")
	}
	defer syscall.CloseHandle(info.process)
	if _, err := syscall.WaitForSingleObject(info.process, syscall.INFINITE); err != nil {
		return false, err
	}
	var code uint32
	if err := syscall.GetExitCodeProcess(info.process, &code); err != nil {
		return false, err
	}
	switch code {
	case 0:
		return false, nil
	case 3010, 1641:
		return true, nil
	default:
		return false, fmt.Errorf("o LibreOffice não foi instalado. Código do Windows %d", code)
	}
}

func dialog(body string, flags uintptr) uintptr {
	messageBox := syscall.NewLazyDLL("user32.dll").NewProc("MessageBoxW")
	result, _, _ := messageBox.Call(0, uintptr(unsafe.Pointer(wide(body))),
		uintptr(unsafe.Pointer(wide("Tributos Docs"))), flags|foreground|topmost)
	return result
}

func wide(text string) *uint16 {
	value, _ := syscall.UTF16PtrFromString(text)
	return value
}
