package main

import (
	"bytes"
	"errors"
	"fmt"
	"runtime"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/sys/windows"
)

type HANDLE uintptr

var (
	kernel32 = syscall.NewLazyDLL("kernel32")
	// ヒープから指定したバイト数を割り当てます。
	procGlobalAlloc = kernel32.NewProc("GlobalAlloc")
	// 指定したグローバル メモリ オブジェクトを解放し、ハンドルを無効にします。
	procGlobalFree = kernel32.NewProc("GlobalFree")
	// メモリ オブジェクトをロックし、メモリ ブロックの先頭ポインタを取得します。移動可能メモリにのみ有効です。
	procGlobalLock = kernel32.NewProc("GlobalLock")
	// メモリ オブジェクトのロックを解除します。移動可能メモリにのみ有効です。
	procGlobalUnlock = kernel32.NewProc("GlobalUnlock")
	// 指定した文字列の長さを決定します (終端の null 文字は含まれません)。
	procLstrlen    = kernel32.NewProc("lstrlenA")
	procMoveMemory = kernel32.NewProc("RtlMoveMemory")
)

var (
	moduser32 = syscall.NewLazyDLL("user32.dll")
	// クリップボードを閉じます。
	procCloseClipboard = moduser32.NewProc("CloseClipboard")
	// クリップボードを空にし、クリップボード内のデータへのハンドルを解放します。
	// この関数は、クリップボードが現在開いているウィンドウにクリップボードの所有権を割り当てます。
	procEmptyClipboard = moduser32.NewProc("EmptyClipboard")
	// クリップボードで現在使用できるデータ形式を列挙します。
	// クリップボードのデータ形式は、順序付けされたリストに格納されます、
	// クリップボードのデータ形式の列挙を実行するには、EnumClipboardFormats関数を一連の呼び出しを行います。
	// 呼び出しごとに、format パラメーターは使用可能なクリップボード形式を指定し、関数は次に使用可能なクリップボード形式を返します。
	procEnumClipboardFormats = moduser32.NewProc("EnumClipboardFormats")
	// 特定の形式でクリップボードからデータを取得する。クリップボードは以前に開かれている必要があります。
	procGetClipboardData = moduser32.NewProc("GetClipboardData")
	// 指定した登録済み形式の名前をクリップボードから取得します。この関数は、指定されたバッファーに名前をコピーします。
	procGetClipboardFormatName = moduser32.NewProc("GetClipboardFormatNameW")
	// 検索のためにクリップボードを開き、他のアプリケーションがクリップボードの内容を変更できないようにします。
	procOpenClipboard = moduser32.NewProc("OpenClipboard")
	// 新しいクリップボード形式を登録します。この形式は、有効なクリップボード形式として使用できます。
	procRegisterClipboardFormat = moduser32.NewProc("RegisterClipboardFormatW")
	// 指定したクリップボード形式でクリップボードにデータを配置します。
	// ウィンドウは現在のクリップボード所有者である必要があり、アプリケーションは、OpenClipboard関数を呼び出している必要があります。
	// (WM_RENDERFORMATメッセージに応答する際、クリップボードの所有者はSetClipboardDataを呼び出す前にOpenClipboardを呼び出さないでください)。
	procSetClipboardData = moduser32.NewProc("SetClipboardData")
)

func main() {
	fmt.Println("GetClipboardHtml")
	v, err := GetClipboardHtml()
	if err != nil {
		fmt.Println("ERR:" + err.Error())
	}
	// fmt.Println(v)

	fmt.Println("SetClipboardHtml")
	newV := strings.ReplaceAll(v, "make", "XXXX")
	if err := SetClipboardHTML(newV); err != nil {
		fmt.Println("ERR:" + err.Error())
	}
}

// クリップボードからHTML形式のテキストを取得します
func GetClipboardHtml() (string, error) {
	// LockOSThread は、メソッド全体が開始から終了まで同じスレッドで実行され続けることを保証します
	// （実際にはゴルーチンのスレッド割り当てをロックします）。
	// そうでない場合、ゴルーチンが実行中にスレッドを切り替えると（これは一般的な動作です）、
	// OpenClipboard と CloseClipboard が異なるスレッドで実行され、クリップボードのデッドロックが発生します。
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := waitOpenClipboard(); err != nil {
		return "", fmt.Errorf("ERR: waitOpenClipboard(%w)", err)
	}
	defer procCloseClipboard.Call()

	// HTML Formatを取得します。
	cf, err := GetHtmlClipboardFormat()
	if err != nil {
		return "", err
	}
	fmt.Printf("HTML Format = %d\n", cf)

	// ハンドルを取得します
	h := GetClipboardData(cf)
	if h == 0 {
		return "", errors.New("クリップボードのハンドルの取得に失敗しました。")
	}
	fmt.Printf("HANDLE = %v\n", h)

	// メモリアドレスを取得します
	pMem, err := GlobalLock(h)
	if err != nil {
		return "", err
	}
	defer GlobalUnlock(h)
	fmt.Printf("pMem = %v\n", pMem)

	// 文字列のバイト数を取得します
	slength := Lstrlen(pMem)
	fmt.Printf("sLength=%d\n", slength)

	data := make([]byte, slength)
	MoveMemory((unsafe.Pointer(&data[0])), unsafe.Pointer(pMem), uint32(len(data)))
	// fmt.Printf("data = %s\n", string(data))

	return string(data), nil
}

// クリップボードにテキストとHTMLテキストをコピーします、成否を返します
func SetClipboardHTML(html string) error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	// 新しいクリップボード形式を登録します
	uFormat, err := RegisterClipboardFormat("HTML Format")
	if err != nil {
		return fmt.Errorf("ERR: UTF16PtrFromString(%w)", err)
	}
	fmt.Printf("cfHtml = %v\n", uFormat)

	// クリップボードを開きます
	if err := waitOpenClipboard(); err != nil {
		return fmt.Errorf("ERR: waitOpenClipboard(%w)", err)
	}
	defer procCloseClipboard.Call()

	// クリップボードを空にします
	if err := EmptyClipboard(); err != nil {
		return fmt.Errorf("ERR: EmptyClipboard(%w)", err)
	}

	// HTMLテキストをUTF-8エンコードの文字列でクリップボードにコピーします
	var buffer bytes.Buffer
	// Version
	buffer.WriteString("Version:0.9\r\n")
	// StartHTML
	starthtml, err := GetStartHTMLPos(html)
	if err != nil {
		return err
	}
	buffer.WriteString(fmt.Sprintf("StartHTML:%06d\r\n", starthtml))
	// EndHTML
	buffer.WriteString(fmt.Sprintf("EndHTML:%06d\r\n", GetEndHTMLPos(html)))
	// StartFragment
	startFragment, err := GetStartFragment(html)
	if err != nil {
		return err
	}
	buffer.WriteString(fmt.Sprintf("StartFragment:%06d\r\n", startFragment))
	// EndFragment
	endFragment := GetEndFragment(html)
	buffer.WriteString(fmt.Sprintf("EndFragment:%06d\r\n", endFragment))
	// <html>
	buffer.WriteString("<html>\r\n")
	// <body>
	buffer.WriteString("<body>\r\n")
	// <!--StartFragment-->
	buffer.WriteString("<!--StartFragment-->\r\n")
	// Fragment
	buffer.WriteString(html[startFragment:endFragment+1] + "\r\n")
	// <!--EndFragment-->
	buffer.WriteString("<!--EndFragment-->\r\n")
	// </body>
	buffer.WriteString("</body>\r\n")
	// </html>
	buffer.WriteString("</html>\r\n")
	// NULL末端
	buffer.WriteByte(0)

	// fmt.Printf("buffer=%s", buffer.String())

	hMem, err := globalAlloc(0x0002|0x2000, uint32(len(buffer.Bytes()))) // GMEM_MOVEABLE | GMEM_SHARE
	if err != nil {
		return fmt.Errorf("ERR: globalAlloc(%w)", err)
	}
	defer func() {
		if hMem != 0 {
			globalFree(hMem)
		}
	}()
	fmt.Printf("hMem = %v\n", hMem)
	// 文字列のバイト数を取得します
	slength := Lstrlen(hMem)
	fmt.Printf("sLength=%d\n", slength)

	ptr, err := GlobalLock(HANDLE(hMem))
	if err != nil {
		return fmt.Errorf("ERR: globalLock(%w)", err)
	}
	defer GlobalUnlock(HANDLE(hMem))
	fmt.Printf("ptr = %v\n", ptr)

	MoveMemory(unsafe.Pointer(ptr), unsafe.Pointer(&buffer.Bytes()[0]), uint32(len(buffer.Bytes())))

	if _, err = setClipboardData(uFormat, hMem); err != nil {
		return err
	}
	// hMem = 0

	return nil
}

// Private Function SetClipboardHTML(sText As String, sHtml As String) As Boolean
//     Const GHND As Long = &H42
//     Const CF_UNICODETEXT As Long = 13

//     Dim hHtml As LongPtr: hHtml = GlobalAlloc(GHND, byteLength)
//     If hHtml = 0 Then GoTo Finally ' 関数に失敗
//     Dim pHtml As LongPtr: pHtml = GlobalLock(hHtml)
//     If pHtml = 0 Then GlobalFree hHtml: GoTo Finally ' 関数に失敗
//     MoveMemory pHtml, VarPtr(bytes(0)), byteLength
//     GlobalUnlock hHtml
//     If SetClipboardData(CF_HTML, hHtml) = 0 Then
//         ' コピー失敗
//         GlobalFree hHtml
//         GoTo Finally
//     End If
//     SetClipboardHTML = True

// End Function

// waitOpenClipboard はクリップボードを開き、最大1秒間待機します。
func waitOpenClipboard() error {
	started := time.Now()
	limit := started.Add(time.Second)
	var r uintptr
	var err error
	for time.Now().Before(limit) {
		r, _, err = procOpenClipboard.Call(0)
		if r != 0 {
			return nil
		}
		time.Sleep(time.Millisecond)
	}
	return err
}

func GetClipboardFormatName(format uint) (string, bool) {
	cchMaxCount := 255
	buf := make([]uint16, cchMaxCount)
	ret, _, _ := procGetClipboardFormatName.Call(
		uintptr(format),
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(cchMaxCount))

	if ret > 0 {
		return syscall.UTF16ToString(buf), true
	}

	return "Requested format does not exist or is predefined", false
}

func Lstrlen(lpString uintptr) int {
	ret, _, _ := procLstrlen.Call(lpString)

	return int(ret)
}

func EnumClipboardFormats(format uint) (uint, bool) {
	ret, _, _ := procEnumClipboardFormats.Call(uintptr(format))
	cf := uint(ret)
	if cf == 0 {
		return cf, false
	} else {
		return cf, true
	}
}

func GetHtmlClipboardFormat() (uint, error) {
	var cf uint
	loop := true
	for loop {
		cf, loop = EnumClipboardFormats(cf)
		if loop {
			if name, b := GetClipboardFormatName(cf); b && name == "HTML Format" {
				return cf, nil
			}
		}
	}

	return 0, errors.New("HTML Formatを取得できません。")
}

func GetClipboardData(uFormat uint) HANDLE {
	ret, _, _ := procGetClipboardData.Call(uintptr(uFormat))
	return HANDLE(ret)
}

// func GlobalLock(hMem HANDLE) (unsafe.Pointer, error) {
func GlobalLock(hMem HANDLE) (uintptr, error) {
	ret, _, _ := procGlobalLock.Call(uintptr(hMem))

	if ret == 0 {
		return 0, errors.New("ERROR: GlobalLock")
	}

	return ret, nil
}

func GlobalUnlock(hMem HANDLE) error {
	ret, _, _ := procGlobalUnlock.Call(uintptr(hMem))

	if ret != 0 {
		return nil
	} else {
		return errors.New("ERROR: GlobalUnlock")
	}
}

func MoveMemory(destination, source unsafe.Pointer, length uint32) {
	procMoveMemory.Call(
		uintptr(unsafe.Pointer(destination)),
		uintptr(source),
		uintptr(length))
}

func RegisterClipboardFormat(format string) (uintptr, error) {
	p, err := windows.UTF16PtrFromString(format)
	if err != nil {
		return 0, err
	}
	r1, _, err := procRegisterClipboardFormat.Call(uintptr(unsafe.Pointer(p)))
	if r1 == 0 {
		return 0, err
	}
	return r1, nil
}

func EmptyClipboard() error {
	r1, _, err := procEmptyClipboard.Call()
	if r1 == 0 {
		return err
	}
	return nil
}

func GetStartHTMLPos(html string) (int, error) {
	p := strings.Index(html, "<html>")
	if p == -1 {
		return p, fmt.Errorf("ERR: GetStartHTMLPos(<html>の検索に失敗しました。index=%d", p)
	}

	return p, nil
}

func GetEndHTMLPos(html string) int {
	return len(html) - 1
}

func GetStartFragment(html string) (int, error) {
	p := strings.Index(html, "<!--StartFragment-->")
	if p == -1 {
		return p, fmt.Errorf("ERR: GetStartFragment(<html>の検索に失敗しました。index=%d", p)
	}

	return p + len("<!--StartFragment-->"), nil

}

func GetEndFragment(html string) int {
	return len(html) - 37
}

func globalAlloc(uFlags uintptr, dwBytes uint32) (uintptr, error) {
	r1, _, err := procGlobalAlloc.Call(uFlags, uintptr(dwBytes))
	if r1 == 0 {
		return 0, err
	}
	return r1, nil
}

func globalFree(hMem uintptr) error {
	r1, _, err := procGlobalFree.Call(hMem)
	if r1 != 0 {
		return err
	}
	return nil
}

func setClipboardData(uFormat uintptr, hMem uintptr) (uintptr, error) {
	r1, _, err := procSetClipboardData.Call(uFormat, hMem)
	if r1 == 0 {
		return 0, err
	}
	return r1, nil
}
