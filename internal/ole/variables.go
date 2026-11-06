//go:build windows
// +build windows

package ole

import "syscall"

// (angepasst: ohne golang.org/x/sys – System-DLLs direkt über syscall)
var (
	modcombase  = syscall.NewLazyDLL("combase.dll")
	modkernel32 = syscall.NewLazyDLL("kernel32.dll")
	modole32    = syscall.NewLazyDLL("ole32.dll")
	modoleaut32 = syscall.NewLazyDLL("oleaut32.dll")
	moduser32   = syscall.NewLazyDLL("user32.dll")
)
