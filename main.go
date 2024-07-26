package main

import (
	"bufio"
	"bytes"
	"errors"
	"flag"
	"fmt"
	"os"
	"reflect"
	"regexp"
	"runtime"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/sys/windows"
)

type (
	HANDLE uintptr
	HWND   HANDLE
)

const (
	// 初期バッファサイズ
	initialBufSize = 10000
	// バッファサイズの最大値。Scannerは必要に応じこのサイズまでバッファを大きくして各行をスキャンする。
	// この値がinitialBufSize以下の場合、Scannerはバッファの拡張を一切行わず与えられた初期バッファのみを使う。
	maxBufSize = 1000000
)

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

var (
	s    bool
	o    string
	html bool
)

func main() {
	flag.BoolVar(&s, "s", false, "フォーマット一覧の表示")
	flag.StringVar(&o, "o", "", "ファイル出力")
	flag.BoolVar(&html, "html", false, "HTMLで出力")
	flag.Parse()

	switch {
	case s:
		if err := ShowClipboardFormat(); err != nil {
			fmt.Println("ERR:" + err.Error())
		}
	case o != "":
		switch {
		case html:
			if err := SaveClipboardHTML(o); err != nil {
				fmt.Println("ERR:" + err.Error())
			}
		default:
			if err := SaveClipboard(o); err != nil {
				fmt.Println("ERR:" + err.Error())
			}
		}
	default:
		replacer1 := strings.NewReplacer(
			// `<span><span style="font-family: -apple-system,`, `<tr><td><span><span style="font-family: -apple-system,`,
			`dir="ltr">`, `dir="ltr"><table><tr><td>`,
			`&nbsp;</span></span><span>`, `</span></span></td></tr><tr><td><span>`,
			// `&nbsp;</span></span></span></span>`, `</span></span></td></tr></table></span></span>`,
			`<blockquote`, `<table><tr><td><blockquote`,
			`</blockquote>`, "</blockquote></td></tr></table>",
		)
		replacer2 := strings.NewReplacer(
			"</span><span", "</span><br><span",
		// `</tr></span></span>`, `</tr></table></span></span>`,
		)
		re1 := regexp.MustCompile(`(?U)<span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="\d">(.*)</span>`)
		re2 := regexp.MustCompile(`(?U)<p style="margin: 0px;">(\[\d{4}/\d{2}/\d{2} \d+:\d{2}\]) (.*)</p>`)
		re3 := regexp.MustCompile(`(?U)&nbsp;((</span>)+)</span></span><!--EndFragment-->`)
		re_like := regexp.MustCompile(`like (\d+)`)
		re_heart := regexp.MustCompile(`heart (\d+)`)
		re_sad := regexp.MustCompile(`sad (\d+)`)
		re_wave1 := regexp.MustCompile(`thewave1 (\d+)`)
		re_surprised := regexp.MustCompile(`surprised (\d+)`)
		re_bowing := regexp.MustCompile(`bowing (\d+)`)
		re_doh := regexp.MustCompile(`doh (\d+)`)
		re_thanks := regexp.MustCompile(`thanks (\d+)`)
		re_bow := regexp.MustCompile(`bow (\d+)`)
		// fmt.Println("GetClipboardHtml")
		v, err := GetClipboardHtml()
		if err != nil {
			fmt.Println("ERR:" + err.Error())
			return
		}
		if strings.Contains(v, "<li>") {
			fmt.Println("クリップボードの内容が誤っています。再取得してください。")
			bufio.NewScanner(os.Stdin).Scan()
			return
		}
		// fmt.Println(v)

		// fmt.Println("SetClipboardHtml")
		// newV := strings.ReplaceAll(v, "make", "XXXX")
		// newV := strings.ReplaceAll(v, "</span><span>", "</span><hr><span>")
		// newV = strings.ReplaceAll(newV, `<p style="margin: 0px;"> like`, `<p style="margin: 0px;"><img src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/yes/default/20_f.png?v=v70">`)
		newV := replacer1.Replace(v)
		newV = replacer2.Replace(newV)
		// メンション
		// newV = re.ReplaceAllString(newV, `<span style="font-weight:bold; color:#FF0000;">$1</span><br>`)
		newV = re1.ReplaceAllString(newV, `<p style="font-weight:bold; color:rgb(98, 100, 167);">$1</p>`)
		newV = re2.ReplaceAllString(newV, `<p style="font-weight:bold; font-size: 12px;">$1 $2</p>`)
		newV = re3.ReplaceAllString(newV, `$1</td></tr></table></span></span><!--EndFragment-->`)
		newV = re_like.ReplaceAllString(newV, `&#x1F44D; $1`)
		newV = re_heart.ReplaceAllString(newV, `&#x1f9e1; $1`)
		newV = re_sad.ReplaceAllString(newV, `&#x1f622; $1`)
		newV = re_wave1.ReplaceAllString(newV, `&#x1f30a; $1`)
		newV = re_surprised.ReplaceAllString(newV, `&#x1f632; $1`)
		newV = re_bowing.ReplaceAllString(newV, `&#x1f647; $1`)
		newV = re_doh.ReplaceAllString(newV, `😣 $1`)
		newV = re_thanks.ReplaceAllString(newV, `🙇🏼‍♀️ $1`)
		newV = re_bow.ReplaceAllString(newV, `🙇🏼‍♀️ $1`)
		if err := SetClipboardHTML(newV); err != nil {
			fmt.Println("ERR:" + err.Error())
		}
	}
}

func SaveClipboard(path string) error {
	v, err := GetClipboardHtml()
	if err != nil {
		return err
	}
	if err := os.WriteFile(o, []byte(v), 0644); err != nil {
		return err
	}

	return nil
}

func SaveClipboardHTML(path string) error {
	v, err := GetClipboardHtml()
	if err != nil {
		return err
	}
	outf, err := os.Create(path)
	if err != nil {
		return err
	}
	defer outf.Close()
	scanner := bufio.NewScanner(strings.NewReader(v))
	buf := make([]byte, initialBufSize)
	scanner.Buffer(buf, maxBufSize)
	writer := bufio.NewWriter(outf)
	defer writer.Flush()

	lineNumber := 0
	for scanner.Scan() {
		lineNumber++
		if lineNumber >= 7 {
			if _, err := writer.WriteString(scanner.Text() + "\n"); err != nil {
				return err
			}
		}
	}

	if err := scanner.Err(); err != nil {
		return err
	}

	return nil
}

func ShowClipboardFormat() error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := waitOpenClipboard(); err != nil {
		return fmt.Errorf("waitOpenClipboard(%w)", err)
	}
	defer W32CloseClipboard()

	var cf uint
	for {
		cf = W32EnumClipboardFormats(cf)
		if cf == 0 {
			return nil
		}
		if name, b := W32GetClipboardFormatName(cf); b {
			fmt.Printf("CF=%d : %s\n", cf, name)
		} else {
			fmt.Printf("CF=%d\n", cf)
		}
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
	// fmt.Printf("HTML Format = %d\n", cf)

	// ハンドルを取得します
	hMem := GetClipboardData(cf)
	if hMem == 0 {
		return "", errors.New("クリップボードのハンドルの取得に失敗しました。")
	}
	// fmt.Printf("HANDLE = %v\n", hMem)

	// メモリアドレスを取得します
	pMem, err := GlobalLock(hMem)
	if err != nil {
		return "", err
	}
	defer GlobalUnlock(hMem)
	// fmt.Printf("pMem = %v\n", pMem)

	// 文字列のバイト数を取得します
	slength := Lstrlen(pMem)
	// fmt.Printf("sLength=%d\n", slength)

	// data := make([]byte, slength)
	// MoveMemory((unsafe.Pointer(&data[0])), unsafe.Pointer(pMem), uint32(len(data)))
	// fmt.Printf("data = %s\n", string(data))

	var data []byte
	h := (*reflect.SliceHeader)(unsafe.Pointer(&data))
	h.Data = pMem
	h.Len = slength
	h.Cap = slength

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
	// fmt.Printf("cfHtml = %v\n", uFormat)

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
	// fmt.Printf("hMem = %v\n", hMem)
	// 文字列のバイト数を取得します
	// slength := Lstrlen(hMem)
	// fmt.Printf("sLength=%d\n", slength)

	ptr, err := GlobalLock(HANDLE(hMem))
	if err != nil {
		return fmt.Errorf("ERR: globalLock(%w)", err)
	}
	defer GlobalUnlock(HANDLE(hMem))
	// fmt.Printf("ptr = %v\n", ptr)

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

func W32OpenClipboard(hWndNewOwner HWND) bool {
	ret, _, _ := procOpenClipboard.Call(
		uintptr(hWndNewOwner))
	return ret != 0
}

func W32CloseClipboard() bool {
	ret, _, _ := procCloseClipboard.Call()
	return ret != 0
}

func W32EnumClipboardFormats(format uint) uint {
	ret, _, _ := procEnumClipboardFormats.Call(
		uintptr(format))
	return uint(ret)
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

func W32GetClipboardFormatName(format uint) (string, bool) {
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
