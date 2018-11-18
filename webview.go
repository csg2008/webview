// Package webview implements Go bindings to https://github.com/zserge/webview C library.
//
// Bindings closely repeat the C APIs and include both, a simplified
// single-function API to just open a full-screen webview window, and a more
// advanced and featureful set of APIs, including Go-to-JavaScript bindings.
//
// The library uses gtk-webkit, Cocoa/Webkit and MSHTML (IE8..11) as a browser
// engine and supports Linux, MacOS and Windows 7..10 respectively.
//
package webview

/*
#cgo windows CFLAGS: -DWEBVIEW_WINAPI=1
#cgo windows LDFLAGS: -lwininet -lole32 -lcomctl32 -loleaut32 -luuid -lgdi32

#include <stdlib.h>
#include <stdint.h>
#define WEBVIEW_STATIC
#define WEBVIEW_IMPLEMENTATION
#include "webview.h"

extern void _webviewExternalInvokeCallback(void *, void *);

static inline void CgoWebViewFree(void *w) {
	free((void *)((struct webview *)w)->title);
	free((void *)((struct webview *)w)->url);
	free(w);
}

static inline void *CgoWebViewCreate(int width, int height, char *title, char *url, int resizable, int debug) {
	struct webview *w = (struct webview *) calloc(1, sizeof(*w));
	w->width = width;
	w->height = height;
	w->title = title;
	w->url = url;
	w->resizable = resizable;
	w->debug = debug;
	w->external_invoke_cb = (webview_external_invoke_cb_t) _webviewExternalInvokeCallback;
	if (webview_init(w) != 0) {
		CgoWebViewFree(w);
		return NULL;
	}
	return (void *)w;
}

static inline int CgoWebViewLoop(void *w, int blocking) {
	return webview_loop((struct webview *)w, blocking);
}
 
static inline void CgoWebViewTerminate(void *w) {
	webview_terminate((struct webview *)w);
}

static inline void CgoWebViewExit(void *w) {
	webview_exit((struct webview *)w);
}

static inline int CgoWebViewVersion(void *w, const char* key, char* ptrChar) {
	return webview_version((struct webview *)w, key, ptrChar);
}

static inline int CgoWebViewGetURL(void *w, void *ptrBstr) {
	return webview_get_url((struct webview *)w, ptrBstr);
}

static inline int CgoWebViewGetHTML(void *w, void *ptrBstr) {
	return webview_get_html((struct webview *)w, ptrBstr);
}

static inline int CgoWebViewGetTitle(void *w, void *ptrBstr) {
	return webview_get_title((struct webview *)w, ptrBstr);
}

static inline int CgoWebViewCookies(void *w, void *ptrBstr) {
	return webview_get_cookies((struct webview *)w, ptrBstr);
}

static inline int CgoWebViewSessionCookie(void *w, const char* url, const char* key, char* ptrChar) {
	return webview_get_session_cookie((struct webview *)w, url, key, ptrChar);
}

static inline int CgoWebViewGoTo(void *w, char *url, int pop) {
	return webview_set_url((struct webview *)w, url, pop);
}

static inline int CgoWebViewRefresh(void *w) {
	return webview_refresh((struct webview *)w);
}

static inline void CgoWebViewSetTitle(void *w, char *title) {
	webview_set_title((struct webview *)w, title);
}

static inline void CgoWebViewSetFullscreen(void *w, int fullscreen) {
	webview_set_fullscreen((struct webview *)w, fullscreen);
}

static inline void CgoWebViewSetColor(void *w, uint8_t r, uint8_t g, uint8_t b, uint8_t a) {
	webview_set_color((struct webview *)w, r, g, b, a);
}

static inline void CgoDialog(void *w, int dlgtype, int flags,
		char *title, char *arg, char *res, size_t ressz) {
	webview_dialog(w, dlgtype, flags,
		(const char*)title, (const char*) arg, res, ressz);
}

static inline int CgoWebViewEval(void *w, char *js) {
	return webview_eval((struct webview *)w, js);
}

static inline void CgoWebViewInjectCSS(void *w, char *css) {
	webview_inject_css((struct webview *)w, css);
}

extern void _webviewDispatchGoCallback(void *);
static inline void _webview_dispatch_cb(struct webview *w, void *arg) {
	_webviewDispatchGoCallback(arg);
}
static inline void CgoWebViewDispatch(void *w, uintptr_t arg) {
	webview_dispatch((struct webview *)w, _webview_dispatch_cb, (void *)arg);
}
*/
import "C"
import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"html/template"
	"log"
	"reflect"
	"runtime"
	"sync"
	"unicode"
	"unsafe"

	"github.com/lxn/win"
)

func init() {
	// Ensure that main.main is called from the main thread
	runtime.LockOSThread()
}

// Open is a simplified API to open a single native window with a full-size webview in
// it. It can be helpful if you want to communicate with the core app using XHR
// or WebSockets (as opposed to using JavaScript bindings).
//
// Window appearance can be customized using title, width, height and resizable parameters.
// URL must be provided and can user either a http or https protocol, or be a
// local file:// URL. On some platforms "data:" URLs are also supported
// (Linux/MacOS).
func Open(title, url string, w, h int, resizable bool) error {
	titleStr := C.CString(title)
	defer C.free(unsafe.Pointer(titleStr))
	urlStr := C.CString(url)
	defer C.free(unsafe.Pointer(urlStr))
	resize := C.int(0)
	if resizable {
		resize = C.int(1)
	}

	r := C.webview(titleStr, urlStr, C.int(w), C.int(h), resize)
	if r != 0 {
		return errors.New("failed to create webview")
	}
	return nil
}

// Debug prints a debug string using stderr on Linux/BSD, NSLog on MacOS and
// OutputDebugString on Windows.
func Debug(a ...interface{}) {
	s := C.CString(fmt.Sprint(a...))
	defer C.free(unsafe.Pointer(s))
	C.webview_print_log(s)
}

// Debugf prints a formatted debug string using stderr on Linux/BSD, NSLog on
// MacOS and OutputDebugString on Windows.
func Debugf(format string, a ...interface{}) {
	s := C.CString(fmt.Sprintf(format, a...))
	defer C.free(unsafe.Pointer(s))
	C.webview_print_log(s)
}

// ExternalInvokeCallbackFunc is a function type that is called every time
// "window.external.invoke()" is called from JavaScript. Data is the only
// obligatory string parameter passed into the "invoke(data)" function from
// JavaScript. To pass more complex data serialized JSON or base64 encoded
// string can be used.
type ExternalInvokeCallbackFunc func(w WebView, data string)

// MouseClickEventFunc mouse click event handle
type MouseClickEventFunc func(key int)

// Settings is a set of parameters to customize the initial WebView appearance
// and behavior. It is passed into the webview.New() constructor.
type Settings struct {
	// WebView main window title
	Title string
	// URL to open in a webview
	URL string
	// Window width in pixels
	Width int
	// Window height in pixels
	Height int
	// Allows/disallows window resizing
	Resizable bool
	// Enable debugging tools (Linux/BSD/MacOS, on Windows use Firebug)
	Debug bool
	// ClickEvent windows click event handle
	ClickEvent MouseClickEventFunc
	// A callback that is executed when JavaScript calls "window.external.invoke()"
	ExternalInvokeCallback ExternalInvokeCallbackFunc
}

// WebView is an interface that wraps the basic methods for controlling the UI
// loop, handling multithreading and providing JavaScript bindings.
type WebView interface {
	// Run() starts the main UI loop until the user closes the webview window or
	// Terminate() is called.
	Run()
	// Loop() runs a single iteration of the main UI.
	Loop(blocking bool) bool
	// SetTitle() changes window title. This method must be called from the main
	// thread only. See Dispatch() for more details.
	SetTitle(title string)
	// SetFullscreen() controls window full-screen mode. This method must be
	// called from the main thread only. See Dispatch() for more details.
	SetFullscreen(fullscreen bool)
	// SetColor() changes window background color. This method must be called from
	// the main thread only. See Dispatch() for more details.
	SetColor(r, g, b, a uint8)
	// Eval() evaluates an arbitrary JS code inside the webview. This method must
	// be called from the main thread only. See Dispatch() for more details.
	Eval(js string) error
	// InjectJS() injects an arbitrary block of CSS code using the JS API. This
	// method must be called from the main thread only. See Dispatch() for more
	// details.
	InjectCSS(css string)
	// Dialog() opens a system dialog of the given type and title. String
	// argument can be provided for certain dialogs, such as alert boxes. For
	// alert boxes argument is a message inside the dialog box.
	Dialog(dlgType DialogType, flags int, title string, arg string) string
	// Terminate() breaks the main UI loop. This method must be called from the main thread
	// only. See Dispatch() for more details.
	Terminate()
	// Dispatch() schedules some arbitrary function to be executed on the main UI
	// thread. This may be helpful if you want to run some JavaScript from
	// background threads/goroutines, or to terminate the app.
	Dispatch(func())
	// Exit() closes the window and cleans up the resources. Use Terminate() to
	// forcefully break out of the main UI loop.
	Exit()
	// Refresh current page
	Refresh() bool
	// Version return ie version
	Version(key string) string
	// GetURL return current browse url
	GetURL() string
	// GetHTML return current page source
	GetHTML() string
	// GetTitle return current page title
	GetTitle() string
	// GetCookies return current cookies
	GetCookies() string
	// GetSessionCookie return spec http only session cookie value
	GetSessionCookie(url string, key string) string
	// GoTo location to spec url
	GoTo(url string, pop bool) int
	// Bind() registers a binding between a given value and a JavaScript object with the
	// given name.  A value must be a struct or a struct pointer. All methods are
	// available under their camel-case names, starting with a lower-case letter,
	// e.g. "FooBar" becomes "fooBar" in JavaScript.
	// Bind() returns a function that updates JavaScript object with the current
	// Go value. You only need to call it if you change Go value asynchronously.
	Bind(name string, v interface{}) (sync func(), err error)
}

// DialogType is an enumeration of all supported system dialog types
type DialogType int

const (
	// DialogTypeOpen is a system file open dialog
	DialogTypeOpen DialogType = iota
	// DialogTypeSave is a system file save dialog
	DialogTypeSave
	// DialogTypeAlert is a system alert dialog (message box)
	DialogTypeAlert
)

const (
	// DialogFlagFile is a normal file picker dialog
	DialogFlagFile = C.WEBVIEW_DIALOG_FLAG_FILE
	// DialogFlagDirectory is an open directory dialog
	DialogFlagDirectory = C.WEBVIEW_DIALOG_FLAG_DIRECTORY
	// DialogFlagInfo is an info alert dialog
	DialogFlagInfo = C.WEBVIEW_DIALOG_FLAG_INFO
	// DialogFlagWarning is a warning alert dialog
	DialogFlagWarning = C.WEBVIEW_DIALOG_FLAG_WARNING
	// DialogFlagError is an error dialog
	DialogFlagError = C.WEBVIEW_DIALOG_FLAG_ERROR
)

var (
	m     sync.Mutex
	index uintptr
	cef   MouseClickEventFunc
	fns   = map[uintptr]func(){}
	cbs   = map[WebView]ExternalInvokeCallbackFunc{}
)

type webview struct {
	w unsafe.Pointer
	s *Settings
}

var _ WebView = &webview{}

func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

// New creates and opens a new webview window using the given settings. The
// returned object implements the WebView interface. This function returns nil
// if a window can not be created.
func New(settings Settings) WebView {
	if settings.Width == 0 {
		settings.Width = 640
	}
	if settings.Height == 0 {
		settings.Height = 480
	}
	if settings.Title == "" {
		settings.Title = "WebView"
	}
	w := &webview{s: &settings}
	w.w = C.CgoWebViewCreate(C.int(settings.Width), C.int(settings.Height),
		C.CString(settings.Title), C.CString(settings.URL),
		C.int(boolToInt(settings.Resizable)), C.int(boolToInt(settings.Debug)))
	m.Lock()
	cef = settings.ClickEvent
	if settings.ExternalInvokeCallback != nil {
		cbs[w] = settings.ExternalInvokeCallback
	} else {
		cbs[w] = func(w WebView, data string) {}
	}
	m.Unlock()
	return w
}

func (w *webview) Loop(blocking bool) bool {
	block := C.int(0)
	if blocking {
		block = 1
	}
	return C.CgoWebViewLoop(w.w, block) == 0
}

func (w *webview) Run() {
	for w.Loop(true) {
	}
}

func (w *webview) Exit() {
	C.CgoWebViewExit(w.w)
}

func (w *webview) Refresh() bool {
	if ret := C.CgoWebViewRefresh(w.w); 0 == ret {
		return true
	}

	return false
}

// Version 读取 IE 版本
// key 可选值 svcVersion， Version, 详细参考：HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer
func (w *webview) Version(key string) string {
	var data = (*C.char)(C.malloc(C.size_t(1024)))
	var ptrKey = C.CString(key)
	defer C.free(unsafe.Pointer(ptrKey))
	if ret := C.CgoWebViewVersion(w.w, ptrKey, data); 0 == ret && nil != data {
		return C.GoString(data)
	}
	
	return ""
}

func (w *webview) GetSessionCookie(url string, key string) string {
	var data = (*C.char)(C.malloc(C.size_t(1024)))
	var ptrUrl = C.CString(url)
	var ptrKey = C.CString(key)
	defer C.free(unsafe.Pointer(ptrUrl))
	defer C.free(unsafe.Pointer(ptrKey))
	if ret := C.CgoWebViewSessionCookie(w.w, ptrUrl, ptrKey, data); 0 == ret && nil != data {
		return C.GoString(data)
	}
	
	return ""
}

func (w *webview) GetCookies() string {
	var ptrBstr *uint16 /*BSTR*/
	if ret := C.CgoWebViewCookies(w.w, unsafe.Pointer(&ptrBstr)); 0 == ret && nil != ptrBstr {
		return win.BSTRToString(ptrBstr);
	}
	
	return ""
}

func (w *webview) GetURL() string {
	var ptrBstr *uint16 /*BSTR*/
	if ret := C.CgoWebViewGetURL(w.w, unsafe.Pointer(&ptrBstr)); 0 == ret && nil != ptrBstr {
		return win.BSTRToString(ptrBstr);
	}
	
	return ""
}

func (w *webview) GetHTML() string {
	var ptrBstr *uint16 /*BSTR*/
	if ret := C.CgoWebViewGetHTML(w.w, unsafe.Pointer(&ptrBstr)); 0 == ret && nil != ptrBstr {
		return win.BSTRToString(ptrBstr);
	}
	
	return ""
}

func (w *webview) GetTitle() string {
	var ptrBstr *uint16 /*BSTR*/
	if ret := C.CgoWebViewGetTitle(w.w, unsafe.Pointer(&ptrBstr)); 0 == ret && nil != ptrBstr {
		return win.BSTRToString(ptrBstr);
	}
	
	return ""
}

func (w *webview) GoTo(url string, pop bool) int {
	var n int
	if pop {
		n = 1
	}

	p := C.CString(url)
	t := C.int(n)
	defer C.free(unsafe.Pointer(p))
	return int(C.CgoWebViewGoTo(w.w, p, t))
}

func (w *webview) Dispatch(f func()) {
	m.Lock()
	for ; fns[index] != nil; index++ {
	}
	fns[index] = f
	m.Unlock()
	C.CgoWebViewDispatch(w.w, C.uintptr_t(index))
}

func (w *webview) SetTitle(title string) {
	p := C.CString(title)
	defer C.free(unsafe.Pointer(p))
	C.CgoWebViewSetTitle(w.w, p)
}

func (w *webview) SetColor(r, g, b, a uint8) {
	C.CgoWebViewSetColor(w.w, C.uint8_t(r), C.uint8_t(g), C.uint8_t(b), C.uint8_t(a))
}

func (w *webview) SetFullscreen(fullscreen bool) {
	C.CgoWebViewSetFullscreen(w.w, C.int(boolToInt(fullscreen)))
}

func (w *webview) Dialog(dlgType DialogType, flags int, title string, arg string) string {
	const maxPath = 4096
	titlePtr := C.CString(title)
	defer C.free(unsafe.Pointer(titlePtr))
	argPtr := C.CString(arg)
	defer C.free(unsafe.Pointer(argPtr))
	resultPtr := (*C.char)(C.calloc((C.size_t)(unsafe.Sizeof((*C.char)(nil))), (C.size_t)(maxPath)))
	defer C.free(unsafe.Pointer(resultPtr))
	C.CgoDialog(w.w, C.int(dlgType), C.int(flags), titlePtr,
		argPtr, resultPtr, C.size_t(maxPath))
	return C.GoString(resultPtr)
}

func (w *webview) Eval(js string) error {
	p := C.CString(js)
	defer C.free(unsafe.Pointer(p))
	switch C.CgoWebViewEval(w.w, p) {
	case -1:
		return errors.New("evaluation failed")
	}
	return nil
}

func (w *webview) InjectCSS(css string) {
	p := C.CString(css)
	defer C.free(unsafe.Pointer(p))
	C.CgoWebViewInjectCSS(w.w, p)
}

func (w *webview) Terminate() {
	C.CgoWebViewTerminate(w.w)
}

//export _webviewDispatchGoCallback
func _webviewDispatchGoCallback(index unsafe.Pointer) {
	var f func()
	m.Lock()
	f = fns[uintptr(index)]
	delete(fns, uintptr(index))
	m.Unlock()
	f()
}

//export _webviewExternalInvokeCallback
func _webviewExternalInvokeCallback(w unsafe.Pointer, data unsafe.Pointer) {
	m.Lock()
	var (
		cb ExternalInvokeCallbackFunc
		wv WebView
	)
	for wv, cb = range cbs {
		if wv.(*webview).w == w {
			break
		}
	}
	m.Unlock()
	cb(wv, C.GoString((*C.char)(data)))
}

//export _webviewMouseNotify
func _webviewMouseNotify(args unsafe.Pointer) {
	if nil != cef {
		cef(0)
	}
}

var bindTmpl = template.Must(template.New("").Parse(`
if (typeof {{.Name}} === 'undefined') {
	{{.Name}} = {};
}
{{ range .Methods }}
{{$.Name}}.{{.JSName}} = function({{.JSArgs}}) {
	window.external.invoke(JSON.stringify({scope: "{{$.Name}}", method: "{{.Name}}", params: [{{.JSArgs}}]}));
};
{{ end }}
`))

type binding struct {
	Value   interface{}
	Name    string
	Methods []methodInfo
}

func newBinding(name string, v interface{}) (*binding, error) {
	methods, err := getMethods(v)
	if err != nil {
		return nil, err
	}
	return &binding{Name: name, Value: v, Methods: methods}, nil
}

func (b *binding) JS() (string, error) {
	js := &bytes.Buffer{}
	err := bindTmpl.Execute(js, b)
	return js.String(), err
}

func (b *binding) Sync() (string, error) {
	js, err := json.Marshal(b.Value)
	if err == nil {
		return fmt.Sprintf("%[1]s.data=%[2]s;if(%[1]s.render){%[1]s.render(%[2]s);}", b.Name, string(js)), nil
	}
	return "", err
}

func (b *binding) Call(js string) bool {
	type rpcCall struct {
		Scope  string        `json:"scope"`
		Method string        `json:"method"`
		Params []interface{} `json:"params"`
	}

	rpc := rpcCall{}
	if err := json.Unmarshal([]byte(js), &rpc); err != nil {
		return false
	}
	if rpc.Scope != b.Name {
		return false
	}
	var mi *methodInfo
	for i := 0; i < len(b.Methods); i++ {
		if b.Methods[i].Name == rpc.Method {
			mi = &b.Methods[i]
			break
		}
	}
	if mi == nil {
		return false
	}
	args := make([]reflect.Value, mi.Arity(), mi.Arity())
	for i := 0; i < mi.Arity(); i++ {
		val := reflect.ValueOf(rpc.Params[i])
		arg := mi.Value.Type().In(i)
		u := reflect.New(arg)
		if b, err := json.Marshal(val.Interface()); err == nil {
			if err = json.Unmarshal(b, u.Interface()); err == nil {
				args[i] = reflect.Indirect(u)
			}
		}
		if !args[i].IsValid() {
			return false
		}
	}
	mi.Value.Call(args)
	return true
}

type methodInfo struct {
	Name  string
	Value reflect.Value
}

func (mi methodInfo) Arity() int { return mi.Value.Type().NumIn() }

func (mi methodInfo) JSName() string {
	r := []rune(mi.Name)
	if len(r) > 0 {
		r[0] = unicode.ToLower(r[0])
	}
	return string(r)
}

func (mi methodInfo) JSArgs() (js string) {
	for i := 0; i < mi.Arity(); i++ {
		if i > 0 {
			js = js + ","
		}
		js = js + fmt.Sprintf("a%d", i)
	}
	return js
}

func getMethods(obj interface{}) ([]methodInfo, error) {
	p := reflect.ValueOf(obj)
	v := reflect.Indirect(p)
	t := reflect.TypeOf(obj)
	if t == nil {
		return nil, errors.New("object can not be nil")
	}
	k := t.Kind()
	if k == reflect.Ptr {
		k = v.Type().Kind()
	}
	if k != reflect.Struct {
		return nil, errors.New("must be a struct or a pointer to a struct")
	}

	methods := []methodInfo{}
	for i := 0; i < t.NumMethod(); i++ {
		method := t.Method(i)
		if !unicode.IsUpper([]rune(method.Name)[0]) {
			continue
		}
		mi := methodInfo{
			Name:  method.Name,
			Value: p.MethodByName(method.Name),
		}
		methods = append(methods, mi)
	}

	return methods, nil
}

func (w *webview) Bind(name string, v interface{}) (sync func(), err error) {
	b, err := newBinding(name, v)
	if err != nil {
		return nil, err
	}
	js, err := b.JS()
	if err != nil {
		return nil, err
	}
	sync = func() {
		if js, err := b.Sync(); err != nil {
			log.Println(err)
		} else {
			w.Eval(js)
		}
	}

	m.Lock()
	cb := cbs[w]
	cbs[w] = func(w WebView, data string) {
		if ok := b.Call(data); ok {
			sync()
		} else {
			cb(w, data)
		}
	}
	m.Unlock()

	w.Eval(js)
	sync()
	return sync, nil
}