// UULA Restyle installer for Windows: puts the add-in into PowerPoint (current user, no admin needed).
// It copies the manifest to %LOCALAPPDATA%\UULA Restyle and registers it under
// HKCU\Software\Microsoft\Office\16.0\WEF\Developer, so PowerPoint loads it on its next start.
package main

import (
	_ "embed"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
	"unsafe"
)

//go:embed UULA-Restyle.xml
var manifest []byte

const (
	regKey   = `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
	regValue = "UULA Restyle"
	title    = "UULA Restyle"
	mbOK     = 0x00000040 // information icon
	mbErr    = 0x00000010
	mbYesNoC = 0x00000003 | 0x00000020
	idYes    = 6
	idNo     = 7
)

var (
	user32     = syscall.NewLazyDLL("user32.dll")
	messageBox = user32.NewProc("MessageBoxW")
)

func box(text string, flags uintptr) int {
	t, _ := syscall.UTF16PtrFromString(text)
	c, _ := syscall.UTF16PtrFromString(title)
	r, _, _ := messageBox.Call(0, uintptr(unsafe.Pointer(t)), uintptr(unsafe.Pointer(c)), flags|0x00100000|0x00040000) // RTL reading + topmost
	return int(r)
}

func reg(args ...string) (string, error) {
	cmd := exec.Command("reg", args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	out, err := cmd.CombinedOutput()
	return string(out), err
}

func powerPointOpen() bool {
	cmd := exec.Command("tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH")
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	out, _ := cmd.Output()
	return strings.Contains(strings.ToUpper(string(out)), "POWERPNT.EXE")
}

func main() {
	dir := filepath.Join(os.Getenv("LOCALAPPDATA"), "UULA Restyle")
	path := filepath.Join(dir, "UULA-Restyle.xml")

	if out, err := reg("query", regKey, "/v", regValue); err == nil && strings.Contains(out, regValue) {
		switch box("أداة UULA Restyle مثبّتة على هذا الجهاز.\n\nنعم = إعادة التثبيت (تحديث)\nلا = إزالة الأداة\n\nUULA Restyle is already installed.\nYes = reinstall   No = remove", mbYesNoC) {
		case idNo:
			reg("delete", regKey, "/v", regValue, "/f")
			os.RemoveAll(dir)
			box("أُزيلت الأداة. أغلق PowerPoint وافتحه من جديد.\n\nThe add-in was removed. Restart PowerPoint.", mbOK)
			return
		case idYes:
		default:
			return
		}
	}

	if err := os.MkdirAll(dir, 0o755); err != nil {
		box("تعذّر إنشاء المجلد:\n"+dir+"\n\n"+err.Error(), mbErr)
		return
	}
	if err := os.WriteFile(path, manifest, 0o644); err != nil {
		box("تعذّر نسخ ملف الأداة:\n"+err.Error(), mbErr)
		return
	}
	if out, err := reg("add", regKey, "/v", regValue, "/t", "REG_SZ", "/d", path, "/f"); err != nil {
		box("تعذّر تسجيل الأداة في PowerPoint:\n"+out, mbErr)
		return
	}
	msg := "تم تثبيت أداة UULA Restyle ✓\n\n" +
		"1) أغلق PowerPoint تمامًا ثم افتحه من جديد.\n" +
		"2) ستجد زر UULA Restyle في تبويب Home (الصفحة الرئيسية).\n" +
		"   إذا لم يظهر: Insert ← Add-ins ← My Add-ins ← Developer Add-ins.\n\n" +
		"يلزم Microsoft 365 محدّث، وخط UULA Sans مثبّت على الجهاز.\n\n" +
		"UULA Restyle is installed. Restart PowerPoint; the button is on the Home tab."
	if powerPointOpen() {
		msg += "\n\n⚠ PowerPoint مفتوح الآن: أغلقه وافتحه لتظهر الأداة."
	}
	box(msg, mbOK)
}
