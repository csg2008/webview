package main

import (
	"fmt"
	"log"
	"net"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/csg2008/webview"
)

const (
	windowWidth  = 480
	windowHeight = 320
)

var indexHTML = `
<!doctype html>
<html>
	<head>
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<title>window-go</title>
	</head>
	<body>
		<button onclick="external.invoke('github')">Github</button>
		<button onclick="external.invoke('close')">Close</button>
		<button onclick="external.invoke('version')">Version</button>
		<button onclick="external.invoke('refresh')">Refresh</button>
		<button onclick="external.invoke('element')">Element</button>
		<button onclick="external.invoke('fullscreen')">Fullscreen</button>
		<button onclick="external.invoke('unfullscreen')">Unfullscreen</button>
		<button onclick="external.invoke('open')">Open</button>
		<button onclick="external.invoke('opendir')">Open directory</button>
		<button onclick="external.invoke('save')">Save</button>
		<button onclick="external.invoke('message')">Message</button>
		<button onclick="external.invoke('info')">Info</button>
		<button onclick="external.invoke('warning')">Warning</button>
		<button onclick="external.invoke('error')">Error</button>
		<button onclick="external.invoke('changeTitle:'+document.getElementById('new-title').value)">
			Change title
		</button>
		<input id="new-title" type="text" />
		<button onclick="external.invoke('changeColor:'+document.getElementById('new-color').value)">
			Change color
		</button>
		<input id="new-color" value="#e91e63" type="color" />
		<a href="http://www.baidu.com" target="_blank">baidu</a>
	</body>
</html>
`

func startServer() string {
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		log.Fatal(err)
	}
	go func() {
		defer ln.Close()
		http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
			var uid = &http.Cookie{Name: "uid", Value: "test", Path: "/", HttpOnly: false, MaxAge: int(time.Hour * 24 / time.Second)}
			var time = &http.Cookie{Name: "time", Value: strconv.FormatInt(time.Now().Unix(), 10), Path: "/", HttpOnly: false, MaxAge: 0}

			http.SetCookie(w, uid)
			http.SetCookie(w, time)
			w.Write([]byte(indexHTML))
		})
		log.Fatal(http.Serve(ln, nil))
	}()
	return "http://" + ln.Addr().String()
}

func handleRPC(w webview.WebView, data string) {
	switch {
	case data == "version":
		w.Dialog(webview.DialogTypeAlert, 0, "About", "Internet Explorer Version: "+w.Version("svcVersion"))
	case data == "element":
		fmt.Println("url:", w.GetURL())
		fmt.Println("cookies:", w.GetCookies())
		fmt.Println("title:", w.GetTitle())
		fmt.Println("html:", w.GetHTML())
	case data == "github":
		w.GoTo("https://github.com/csg2008/webview", true)
	case data == "refresh":
		w.Refresh()
	case data == "close":
		w.Terminate()
	case data == "fullscreen":
		w.SetFullscreen(true)
	case data == "unfullscreen":
		w.SetFullscreen(false)
	case data == "open":
		log.Println("open", w.Dialog(webview.DialogTypeOpen, 0, "Open file", ""))
	case data == "opendir":
		log.Println("open", w.Dialog(webview.DialogTypeOpen, webview.DialogFlagDirectory, "Open directory", ""))
	case data == "save":
		log.Println("save", w.Dialog(webview.DialogTypeSave, 0, "Save file", ""))
	case data == "message":
		w.Dialog(webview.DialogTypeAlert, 0, "Hello", "Hello, world!")
	case data == "info":
		w.Dialog(webview.DialogTypeAlert, webview.DialogFlagInfo, "Hello", "Hello, info!")
	case data == "warning":
		w.Dialog(webview.DialogTypeAlert, webview.DialogFlagWarning, "Hello", "Hello, warning!")
	case data == "error":
		w.Dialog(webview.DialogTypeAlert, webview.DialogFlagError, "Hello", "Hello, error!")
	case strings.HasPrefix(data, "changeTitle:"):
		w.SetTitle(strings.TrimPrefix(data, "changeTitle:"))
	case strings.HasPrefix(data, "changeColor:"):
		hex := strings.TrimPrefix(strings.TrimPrefix(data, "changeColor:"), "#")
		num := len(hex) / 2
		if !(num == 3 || num == 4) {
			log.Println("Color must be RRGGBB or RRGGBBAA")
			return
		}
		i, err := strconv.ParseUint(hex, 16, 64)
		if err != nil {
			log.Println(err)
			return
		}
		if num == 3 {
			r := uint8((i >> 16) & 0xFF)
			g := uint8((i >> 8) & 0xFF)
			b := uint8(i & 0xFF)
			w.SetColor(r, g, b, 255)
			return
		}
		if num == 4 {
			r := uint8((i >> 24) & 0xFF)
			g := uint8((i >> 16) & 0xFF)
			b := uint8((i >> 8) & 0xFF)
			a := uint8(i & 0xFF)
			w.SetColor(r, g, b, a)
			return
		}
	}
}

func main() {
	var w webview.WebView
	var url = startServer()
	w = webview.New(webview.Settings{
		Width:                  windowWidth,
		Height:                 windowHeight,
		Title:                  "Simple window demo",
		Resizable:              true,
		URL:                    url,
		ExternalInvokeCallback: handleRPC,
		ClickEvent: func(key int) {
			fmt.Println("url:", w.GetURL())
		},
	})
	w.SetColor(255, 255, 255, 255)
	defer w.Exit()
	w.Run()
}
