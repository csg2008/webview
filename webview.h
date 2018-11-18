/*
 * MIT License
 *
 * Copyright (c) 2017 Serge Zaitsev
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
#ifndef WEBVIEW_H
#define WEBVIEW_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stdio.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#define CINTERFACE
#include <windows.h>
#include <wininet.h>
#include <commctrl.h>
#include <exdisp.h>
#include <mshtml.h>
#include <mshtmhst.h>
#include <shobjidl.h>
#include <crtdbg.h>    // for _ASSERT()
#include "ExDispid.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "Wininet.lib")
#pragma comment(lib, "oleaut32.lib")

#define WEBVIEW_API static
#define WEBVIEW_DIALOG_FLAG_FILE (0 << 0)
#define WEBVIEW_DIALOG_FLAG_DIRECTORY (1 << 0)
#define WEBVIEW_DIALOG_FLAG_INFO (1 << 1)
#define WEBVIEW_DIALOG_FLAG_WARNING (2 << 1)
#define WEBVIEW_DIALOG_FLAG_ERROR (3 << 1)
#define WEBVIEW_DIALOG_FLAG_ALERT_MASK (3 << 1)

//消息类
#define WM_NEW_IEVIEW      WM_USER+101
#define WM_WEBVIEW_DISPATCH (WM_APP + 1)

#define DEFAULT_URL                                                            \
  "data:text/"                                                                 \
  "html,%3C%21DOCTYPE%20html%3E%0A%3Chtml%20lang=%22en%22%3E%0A%3Chead%3E%"    \
  "3Cmeta%20charset=%22utf-8%22%3E%3Cmeta%20http-equiv=%22X-UA-Compatible%22%" \
  "20content=%22IE=edge%22%3E%3C%2Fhead%3E%0A%3Cbody%3E%3Cdiv%20id=%22app%22%" \
  "3E%3C%2Fdiv%3E%3Cscript%20type=%22text%2Fjavascript%22%3E%3C%2Fscript%3E%"  \
  "3C%2Fbody%3E%0A%3C%2Fhtml%3E"

#define CSS_INJECT_FUNCTION                                                    \
  "(function(e){var "                                                          \
  "t=document.createElement('style'),d=document.head||document."               \
  "getElementsByTagName('head')[0];t.setAttribute('type','text/"               \
  "css'),t.styleSheet?t.styleSheet.cssText=e:t.appendChild(document."          \
  "createTextNode(e)),d.appendChild(t)})"

// The following macro can be found in IE8 SDK
#ifndef INTERNET_COOKIE_NON_SCRIPT
#define INTERNET_COOKIE_NON_SCRIPT      0x00001000
#endif

#ifndef INTERNET_COOKIE_HTTPONLY
#define INTERNET_COOKIE_HTTPONLY        0x00002000
#endif

#ifdef __cplusplus
#define iid_ref(x) &(x)
#define iid_unref(x) *(x)
#else
#define iid_ref(x) (x)
#define iid_unref(x) (x)
#endif

// A running count of how many windows we have open that contain a browser object
static unsigned char WindowCount = 0;

// This is used by DisplayHTMLStr(). It can be global because we never change it.
static const SAFEARRAYBOUND ArrayBound = {1, 0};

// The class name of our Window to host the browser. It can be anything of your choosing.
static const TCHAR *classname = "WebView";

static unsigned char AppUrl[] = {'a', 0, 'p', 0, 'p', 0, ':', 0};
static wchar_t Blank[] = {L"about:blank"};

static unsigned char _IID_IHTMLWindow3[] = {0xae, 0xf4, 0x50, 0x30, 0xb5, 0x98, 0xcf, 0x11, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b};

// Some misc stuff used by our IDispatch
static const BSTR    OnBeforeOnLoad = L"onbeforeunload";
static const WCHAR  BeforeUnload[] = L"beforeunload";

extern void _webviewMouseNotify(void *arg);

struct webview;

struct webview_priv {
  HWND hwnd;
  IOleObject **browser;
  BOOL is_fullscreen;
  DWORD saved_style;
  DWORD saved_ex_style;
  RECT saved_rect;
};

typedef void (*webview_external_invoke_cb_t)(struct webview *w, const char *arg);

struct webview {
  const char *url;
  const char *title;
  int width;
  int height;
  int resizable;
  int debug;
  webview_external_invoke_cb_t external_invoke_cb;
  struct webview_priv priv;
  void *userdata;
};

enum webview_dialog_type {
  WEBVIEW_DIALOG_TYPE_OPEN  = 0,
  WEBVIEW_DIALOG_TYPE_SAVE  = 1,
  WEBVIEW_DIALOG_TYPE_ALERT = 2
};

typedef void (*webview_dispatch_fn)(struct webview *w, void *arg);

struct webview_dispatch_arg {
  webview_dispatch_fn fn;
  struct webview *w;
  void *arg;
};

typedef struct {
  NMHDR            nmhdr;
  IHTMLEventObj *  htmlEvent;
  LPCTSTR          eventStr;
} WEBPARAMS; 


// Our _IDispatchEx struct. This is just an IDispatch with some
// extra fields appended to it for our use in storing extra
// info we need for the purpose of reacting to events that happen
// to some element on a web page.
typedef struct {
  IDispatch*       dispatchObj;  // The mandatory IDispatch.
  DWORD            refCount;     // Our reference count.
  IHTMLWindow2 *   htmlWindow2;  // Where we store the IHTMLWindow2 so that our IDispatch's Invoke() can get it.
  HWND             hwnd;         // The window hosting the browser page. Our IDispatch's Invoke() sends messages when an event of interest occurs.
  short            id;           // Any numeric value of your choosing that you wish to associate with this IDispatch.
  unsigned short   extraSize;    // Byte size of any extra fields prepended to this struct.
  IUnknown         *object;      // Some object associated with the web page element this IDispatch is for.
  void             *userdata;    // An extra pointer.
} _IDispatchEx;

// We need to return an IOleInPlaceFrame struct to the browser object. And one of our IOleInPlaceFrame
// functions (Frame_GetWindow) is going to need to access our window handle. So let's create our own
// struct that starts off with an IOleInPlaceFrame struct (and that's important -- the IOleInPlaceFrame
// struct *must* be first), and then has an extra data field where we can store our own window's HWND.
//
// And because we may want to create multiple windows, each hosting its own browser object (to
// display its own web page), then we need to create a IOleInPlaceFrame struct for each window. So,
// we're not going to declare our IOleInPlaceFrame struct globally. We'll allocate it later using
// GlobalAlloc, and then stuff the appropriate HWND in it then, and also stuff a pointer to
// MyIOleInPlaceFrameTable in it. But let's just define it here.
typedef struct {
  IOleInPlaceFrame  frame;    // The IOleInPlaceFrame must be first!

  ///////////////////////////////////////////////////
  // Here you add any extra variables that you need
  // to access in your IOleInPlaceFrame functions.
  // You don't want those functions to access global
  // variables, because then you couldn't use more
  // than one browser object. (ie, You couldn't have
  // multiple windows, each with its own embedded
  // browser object to display a different web page).
  //
  // So here is where I added my extra HWND that my
  // IOleInPlaceFrame function Frame_GetWindow() needs
  // to access.
  ///////////////////////////////////////////////////
  HWND        window;
} _IOleInPlaceFrameEx;

// We need to pass our IOleClientSite structure to OleCreate (which in turn gives it to the browser).
// But the browser is also going to ask our IOleClientSite's QueryInterface() to return a pointer to
// our IOleInPlaceSite and/or IDocHostUIHandler structs. So we'll need to have those pointers handy.
// Plus, some of our IOleClientSite and IOleInPlaceSite functions will need to have the HWND to our
// window, and also a pointer to our IOleInPlaceFrame struct. So let's create a single struct that
// has the IOleClientSite, IOleInPlaceSite, IDocHostUIHandler, and IOleInPlaceFrame structs all inside
// it (so we can easily get a pointer to any one from any of those structs' functions). As long as the
// IOleClientSite struct is the very first thing in this custom struct, it's all ok. We can still pass
// it to OleCreate() and pretend that it's an ordinary IOleClientSite. We'll call this new struct a
// _IOleClientSiteEx.
//
// And because we may want to create multiple windows, each hosting its own browser object (to
// display its own web page), then we need to create a unique _IOleClientSiteEx struct for
// each window. So, we're not going to declare this struct globally. We'll allocate it later
// using GlobalAlloc, and then initialize the structs within it.
typedef struct {
  IOleInPlaceSite       inplace;     // My IOleInPlaceSite object. Must be first with in _IOleInPlaceSiteEx.

  ///////////////////////////////////////////////////
  // Here you add any extra variables that you need
  // to access in your IOleInPlaceSite functions.
  //
  // So here is where I added my IOleInPlaceFrame
  // struct. If you need extra variables, add them
  // at the end.
  ///////////////////////////////////////////////////
  _IOleInPlaceFrameEx    frame;    // My IOleInPlaceFrame object. Must be first within my _IOleInPlaceFrameEx
} _IOleInPlaceSiteEx;

typedef struct {
  IDocHostUIHandler    ui;      // My IDocHostUIHandler object. Must be first.

  ///////////////////////////////////////////////////
  // Here you add any extra variables that you need
  // to access in your IDocHostUIHandler functions.
  ///////////////////////////////////////////////////
} _IDocHostUIHandlerEx;

typedef struct {
  IInternetSecurityManager mgr;
} _IInternetSecurityManagerEx;

typedef struct {
  IServiceProvider provider;
  _IInternetSecurityManagerEx mgr;
} _IServiceProviderEx;

typedef struct {
  IOleClientSite        client;    // My IOleClientSite object. Must be first.
  _IOleInPlaceSiteEx    inplace;  // My IOleInPlaceSite object. A convenient place to put it.
  _IDocHostUIHandlerEx  ui;      // My IDocHostUIHandler object. Must be first within my _IDocHostUIHandlerEx.

  ///////////////////////////////////////////////////
  // Here you add any extra variables that you need
  // to access in your IOleClientSite functions.
  ///////////////////////////////////////////////////
  IDispatch external;
  _IServiceProviderEx provider;
} _IOleClientSiteEx;

/* **************************** asciiToNum() ***************************
 * Converts the string of digits (expressed in base 10) to a 32-bit
 * DWORD.
 *
 * val =  Pointer to the nul-terminated string of digits to convert.
 *
 * RETURNS: The integer value as a DWORD.
 *
 * NOTE: Skips leading spaces before the first digit.
*/
static DWORD asciiToNum(OLECHAR *val) {
  OLECHAR      chr;
  DWORD        len;

  // Result is initially 0
  len = 0;

  // Skip leading spaces
  while (*val == ' ' || *val == 0x09) val++;

  // Convert next digit
  while (*val) {
    chr = *(val)++ - '0';
    if ((DWORD)chr > 9) break;
    len += (len + (len << 3) + chr);
  }

  return(len);
}

static int iid_eq(REFIID a, const IID *b) {
  return memcmp((const void *)iid_ref(a), (const void *)b, sizeof(GUID)) == 0;
}

static const char *webview_check_url(const char *url) {
  if (url == NULL || strlen(url) == 0) {
    return DEFAULT_URL;
  }

  return url;
}

WEBVIEW_API int webview(const char *title, const char *url, int width, int height, int resizable);
WEBVIEW_API int webview_init(struct webview *w);
WEBVIEW_API int webview_loop(struct webview *w, int blocking);
WEBVIEW_API int webview_eval(struct webview *w, const char *js);
WEBVIEW_API int webview_inject_css(struct webview *w, const char *css);
WEBVIEW_API void webview_set_title(struct webview *w, const char *title);
WEBVIEW_API void webview_set_fullscreen(struct webview *w, int fullscreen);
WEBVIEW_API void webview_set_color(struct webview *w, uint8_t r, uint8_t g, uint8_t b, uint8_t a);
WEBVIEW_API void webview_dialog(struct webview *w, enum webview_dialog_type dlgtype, int flags, const char *title, const char *arg, char *result, size_t resultsz);
WEBVIEW_API void webview_dispatch(struct webview *w, webview_dispatch_fn fn, void *arg);
WEBVIEW_API void webview_terminate(struct webview *w);
WEBVIEW_API void webview_exit(struct webview *w);
WEBVIEW_API int webview_get_url(struct webview *w, void *ptrBstr);
WEBVIEW_API int webview_get_title(struct webview *w, void *ptrBstr);
WEBVIEW_API int webview_get_html(struct webview *w, void *ptrBstr);
WEBVIEW_API int webview_get_cookies(struct webview *w, void *ptrBstr);
WEBVIEW_API int webview_get_session_cookie(struct webview *w, const char* url, const char* key, char* ptr);
WEBVIEW_API int webview_set_url(struct webview *w, const char *url, int pop);
WEBVIEW_API int webview_refresh(struct webview *w);
WEBVIEW_API void webview_debug(const char *format, ...);
WEBVIEW_API void webview_print_log(const char *s);
WEBVIEW_API int webview_version(struct webview *w, const char* key, char* ptr);

#ifdef WEBVIEW_IMPLEMENTATION
#undef WEBVIEW_IMPLEMENTATION

WEBVIEW_API int webview(const char *title, const char *url, int width, int height, int resizable) {
  struct webview webview;
  memset(&webview, 0, sizeof(webview));
  webview.title = title;
  webview.url = url;
  webview.width = width;
  webview.height = height;
  webview.resizable = resizable;
  int r = webview_init(&webview);
  if (r != 0) {
    return r;
  }
  while (webview_loop(&webview, 1) == 0) {
  }
  webview_exit(&webview);
  return 0;
}

WEBVIEW_API void webview_debug(const char *format, ...) {
  char buf[4096];
  va_list ap;
  va_start(ap, format);
  vsnprintf(buf, sizeof(buf), format, ap);
  webview_print_log(buf);
  va_end(ap);
}

WEBVIEW_API int webview_js_encode(const char *s, char *esc, size_t n) {
  int r = 1; /* At least one byte for trailing zero */
  for (; *s; s++) {
    const unsigned char c = *s;
    if (c >= 0x20 && c < 0x80 && strchr("<>\\'\"", c) == NULL) {
      if (n > 0) {
        *esc++ = c;
        n--;
      }
      r++;
    } else {
      if (n > 0) {
        snprintf(esc, n, "\\x%02x", (int)c);
        esc += 4;
        n -= 4;
      }
      r += 4;
    }
  }
  return r;
}

WEBVIEW_API int webview_inject_css(struct webview *w, const char *css) {
  int n = webview_js_encode(css, NULL, 0);
  char *esc = (char *)calloc(1, sizeof(CSS_INJECT_FUNCTION) + n + 4);
  if (esc == NULL) {
    return -1;
  }
  char *js = (char *)calloc(1, n);
  webview_js_encode(css, js, n);
  snprintf(esc, sizeof(CSS_INJECT_FUNCTION) + n + 4, "%s(\"%s\")", CSS_INJECT_FUNCTION, js);
  int r = webview_eval(w, esc);
  free(js);
  free(esc);
  return r;
}

static inline WCHAR *webview_to_utf16(const char *s) {
  DWORD size = MultiByteToWideChar(CP_UTF8, 0, s, -1, 0, 0);
  WCHAR *ws = (WCHAR *)GlobalAlloc(GMEM_FIXED, sizeof(WCHAR) * size);
  if (ws == NULL) {
    return NULL;
  }
  MultiByteToWideChar(CP_UTF8, 0, s, -1, ws, size);
  return ws;
}

static inline char *webview_from_utf16(WCHAR *ws) {
  int n = WideCharToMultiByte(CP_UTF8, 0, ws, -1, NULL, 0, NULL, NULL);
  char *s = (char *)GlobalAlloc(GMEM_FIXED, n);
  if (s == NULL) {
    return NULL;
  }
  WideCharToMultiByte(CP_UTF8, 0, ws, -1, s, n, NULL, NULL);
  return s;
}

static VARIANT StrToVARIANT(LPCSTR str) {
  VARIANT ret;
  VariantInit(&ret);
  ret.vt = VT_BSTR;
#ifndef UNICODE
    {
      wchar_t *buffer = webview_to_utf16(str);
      if (buffer != NULL) {
        ret.bstrVal = SysAllocString(buffer);
        GlobalFree(buffer);
      }
    }
#else
    ret.bstrVal = SysAllocString(str);
#endif

  return ret;
}




////////////////////////////////////// My IStorage functions  /////////////////////////////////////////
// NOTE: The browser object doesn't use the IStorage functions, so most of these are us just returning
// E_NOTIMPL so that anyone who *does* call these functions knows nothing is being done here.

static HRESULT STDMETHODCALLTYPE Storage_QueryInterface(IStorage FAR* This, REFIID riid, LPVOID FAR* ppvObj) {
  return E_NOTIMPL;
}

static ULONG STDMETHODCALLTYPE Storage_AddRef(IStorage FAR* This) {
  return S_FALSE;
}

static ULONG STDMETHODCALLTYPE Storage_Release(IStorage FAR* This) {
  return S_FALSE;
}

static HRESULT STDMETHODCALLTYPE Storage_CreateStream(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStream **ppstm) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_OpenStream(IStorage FAR* This, const WCHAR * pwcsName, void *reserved1, DWORD grfMode, DWORD reserved2, IStream **ppstm) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_CreateStorage(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStorage **ppstg) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_OpenStorage(IStorage FAR* This, const WCHAR * pwcsName, IStorage * pstgPriority, DWORD grfMode, SNB snbExclude, DWORD reserved, IStorage **ppstg) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_CopyTo(IStorage FAR* This, DWORD ciidExclude, IID const *rgiidExclude, SNB snbExclude,IStorage *pstgDest) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_MoveElementTo(IStorage FAR* This, const OLECHAR *pwcsName,IStorage * pstgDest, const OLECHAR *pwcsNewName, DWORD grfFlags) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_Commit(IStorage FAR* This, DWORD grfCommitFlags) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_Revert(IStorage FAR* This) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_EnumElements(IStorage FAR* This, DWORD reserved1, void * reserved2, DWORD reserved3, IEnumSTATSTG ** ppenum) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_DestroyElement(IStorage FAR* This, const OLECHAR *pwcsName) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_RenameElement(IStorage FAR* This, const WCHAR *pwcsOldName, const WCHAR *pwcsNewName) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_SetElementTimes(IStorage FAR* This, const WCHAR *pwcsName, FILETIME const *pctime, FILETIME const *patime, FILETIME const *pmtime) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_SetClass(IStorage FAR* This, REFCLSID clsid) {
  return(S_OK);
}

static HRESULT STDMETHODCALLTYPE Storage_SetStateBits(IStorage FAR* This, DWORD grfStateBits, DWORD grfMask) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Storage_Stat(IStorage FAR* This, STATSTG * pstatstg, DWORD grfStatFlag) {
  return E_NOTIMPL;
}

// Our IStorage VTable. This is the array of pointers to the above functions in our C
// program that someone may call in order to store some data to disk. We must define a
// particular set of functions that comprise the IStorage set of functions (see above),
// and then stuff pointers to those functions in their respective 'slots' in this table.
// We want the browser to use this VTable with our IStorage structure (object).
static IStorageVtbl MyIStorageTable = {
  Storage_QueryInterface,
  Storage_AddRef,
  Storage_Release,
  Storage_CreateStream,
  Storage_OpenStream,
  Storage_CreateStorage,
  Storage_OpenStorage,
  Storage_CopyTo,
  Storage_MoveElementTo,
  Storage_Commit,
  Storage_Revert,
  Storage_EnumElements,
  Storage_DestroyElement,
  Storage_RenameElement,
  Storage_SetElementTimes,
  Storage_SetClass,
  Storage_SetStateBits,
  Storage_Stat
};

// Our IStorage structure. NOTE: All it contains is a pointer to our IStorageVtbl, so we can easily initialize it
// here instead of doing that programmably.
static IStorage      MyIStorage = { &MyIStorageTable };



#ifndef _ATL_MAX_VARTYPES
#define _ATL_MAX_VARTYPES 8
#endif

static void __stdcall CWebBrowser2_SetSecureLockIcon(long nSecureLockIcon) {
//    T* pT = static_cast<T*>(this);
//    pT->OnSetSecureLockIcon(nSecureLockIcon);
}
  
static  void __stdcall CWebBrowser2_NavigateError(IDispatch* pDisp, VARIANT* pvURL, VARIANT* pvTargetFrameName, VARIANT* pvStatusCode, VARIANT_BOOL* pbCancel) {
//     T* pT = static_cast<T*>(this);
//     ATLASSERT(V_VT(pvURL) == VT_BSTR);
//     ATLASSERT(V_VT(pvTargetFrameName) == VT_BSTR);
//     ATLASSERT(V_VT(pvStatusCode) == (VT_I4));
//     ATLASSERT(pbCancel != NULL);
//     *pbCancel=pT->OnNavigateError(pDisp,V_BSTR(pvURL),V_BSTR(pvTargetFrameName),V_I4(pvStatusCode));
}
  
static  void __stdcall CWebBrowser2_PrivacyImpactedStateChange(VARIANT_BOOL bImpacted) {
//     T* pT = static_cast<T*>(this);
//     pT->OnPrivacyImpactedStateChange(bImpacted);
}
  
static  void __stdcall CWebBrowser2_StatusTextChange(BSTR szText) {
//     T* pT = static_cast<T*>(this);
//     pT->OnStatusTextChange(szText);
}
  
static  void __stdcall CWebBrowser2_ProgressChange(long nProgress, long nProgressMax) {
//     T* pT = static_cast<T*>(this);
//     pT->OnProgressChange(nProgress, nProgressMax);
}
  
static  void __stdcall CWebBrowser2_CommandStateChange(long nCommand, VARIANT_BOOL bEnable) {
//     T* pT = static_cast<T*>(this);
//     pT->OnCommandStateChange(nCommand, (bEnable==VARIANT_TRUE) ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_DownloadBegin(IDispatch* pDisp) {
//     T* pT = static_cast<T*>(this);
//     pT->OnDownloadBegin();
}
  
static  void __stdcall CWebBrowser2_DownloadComplete(IDispatch* pDisp) {
//     T* pT = static_cast<T*>(this);
//     pT->OnDownloadComplete();
}
  
static  void __stdcall CWebBrowser2_TitleChange(IDispatch* pDisp, BSTR szText) {
//     T* pT = static_cast<T*>(this);
//     pT->OnTitleChange(szText);
//    SetWindowTextW( g_frmMain , (LPCWSTR)szText);
      printf("title change %s\n", "cweb");
}
  
static  void __stdcall CWebBrowser2_NavigateComplete2(IDispatch* pDisp, VARIANT* pvURL) {
//    FormMain_OnChangeUrlAndTitle();
//    FormMain_OnChangeTabLabel();
//     T* pT = static_cast<T*>(this);
//     ATLASSERT(V_VT(pvURL) == VT_BSTR);
//     pT->OnNavigateComplete2(pDisp, V_BSTR(pvURL));
}
  
static  void __stdcall CWebBrowser2_BeforeNavigate2(IDispatch* pDisp, VARIANT* pvURL, VARIANT* pvFlags, VARIANT* pvTargetFrameName, VARIANT* pvPostData, VARIANT* pvHeaders, VARIANT_BOOL* pbCancel) {
//     T* pT = static_cast<T*>(this);
//     ATLASSERT(V_VT(pvURL) == VT_BSTR);
//     ATLASSERT(V_VT(pvTargetFrameName) == VT_BSTR);
//     ATLASSERT(V_VT(pvPostData) == (VT_VARIANT | VT_BYREF));
//     ATLASSERT(V_VT(pvHeaders) == VT_BSTR);
//     ATLASSERT(pbCancel != NULL);
//     
//     VARIANT* vtPostedData = V_VARIANTREF(pvPostData);
//     CSimpleArray<BYTE> pArray;
//     if (V_VT(vtPostedData) & VT_ARRAY)
//     {
//       ATLASSERT(V_ARRAY(vtPostedData)->cDims == 1 && V_ARRAY(vtPostedData)->cbElements == 1);
//       long nLowBound=0,nUpperBound=0;
//       SafeArrayGetLBound(V_ARRAY(vtPostedData), 1, &nLowBound);
//       SafeArrayGetUBound(V_ARRAY(vtPostedData), 1, &nUpperBound);
//       DWORD dwSize=nUpperBound+1-nLowBound;
//       LPVOID pData=NULL;
//       SafeArrayAccessData(V_ARRAY(vtPostedData),&pData);
//       ATLASSERT(pData);
//       
//       pArray.m_nSize=pArray.m_nAllocSize=dwSize;
//       pArray.m_aT=(BYTE*)malloc(dwSize);
//       ATLASSERT(pArray.m_aT);
//       CopyMemory(pArray.GetData(), pData, dwSize);
//       SafeArrayUnaccessData(V_ARRAY(vtPostedData));
//     }
//     *pbCancel=pT->OnBeforeNavigate2(pDisp, V_BSTR(pvURL), V_I4(pvFlags), V_BSTR(pvTargetFrameName), pArray, V_BSTR(pvHeaders)) ? VARIANT_TRUE : VARIANT_FALSE;
}
  
static  void __stdcall CWebBrowser2_PropertyChange(BSTR szProperty) {
//     T* pT = static_cast<T*>(this);
//     pT->OnPropertyChange(szProperty);
}
  
static  void __stdcall CWebBrowser2_NewWindow2(IUnknown* EventSin, IDispatch** ppDisp, VARIANT_BOOL* pbCancel) {
/*
    DWORD dwThreadID = 0;
    HANDLE hThread =NULL;
    MSG        msg;
    LPDISPATCH result = NULL;
    IWebBrowser2  *pBrowserApp;
    IOleObject    *browserObject;
    WNDCLASSEX  wc  = {0};
    HRESULT hr  = -1;
    LPDISPATCH DbgppDisp = NULL;

    *pbCancel=FALSE;

    SendMessageW(
      g_frmMain,
      WM_NEW_IEVIEW,
      (WPARAM)&(msg.hwnd),
      0);

    browserObject = *((IOleObject **)GetWindowLong(msg.hwnd, GWL_USERDATA));
    hr = browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**) &pBrowserApp);
    if ( SUCCEEDED(hr) )
    {
      hr = pBrowserApp->lpVtbl->get_Application(pBrowserApp,&DbgppDisp);
      if ( SUCCEEDED(hr) )
      {
        *ppDisp  = DbgppDisp;
      }  
    }
*/
}
  
static  void __stdcall CWebBrowser2_DocumentComplete(IDispatch* pDisp, VARIANT* pvURL) {
//     T* pT = static_cast<T*>(this);
//     ATLASSERT(V_VT(pvURL) == VT_BSTR);
//     pT->OnDocumentComplete(pDisp, V_BSTR(pvURL));
}
  
static  void __stdcall CWebBrowser2_OnQuit() {
//     T* pT = static_cast<T*>(this);
//     pT->OnQuit();
}
  
static  void __stdcall CWebBrowser2_OnVisible(VARIANT_BOOL bVisible) {
//     T* pT = static_cast<T*>(this);
//     pT->OnVisible(bVisible == VARIANT_TRUE ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_OnToolBar(VARIANT_BOOL bToolBar) {
//     T* pT = static_cast<T*>(this);
//     pT->OnToolBar(bToolBar == VARIANT_TRUE ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_OnMenuBar(VARIANT_BOOL bMenuBar) {
//     T* pT = static_cast<T*>(this);
//     pT->OnMenuBar(bMenuBar == VARIANT_TRUE ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_OnStatusBar(VARIANT_BOOL bStatusBar) {
//     T* pT = static_cast<T*>(this);
//     pT->OnStatusBar(bStatusBar == VARIANT_TRUE ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_OnFullScreen(VARIANT_BOOL bFullScreen) {
//     T* pT = static_cast<T*>(this);
//     pT->OnFullScreen(bFullScreen == VARIANT_TRUE ? TRUE : FALSE);
}
  
static  void __stdcall CWebBrowser2_OnTheaterMode(VARIANT_BOOL bTheaterMode) {
  // T* pT = static_cast<T*>(this);
  // pT->OnTheaterMode(bTheaterMode == VARIANT_TRUE ? TRUE : FALSE);
}

typedef struct _ATL_FUNC_INFO_ {
  CALLCONV   cc;
  VARTYPE    vtReturn;
  SHORT      nParams;
  VARTYPE    pVarTypes[_ATL_MAX_VARTYPES];
} _ATL_FUNC_INFO;

typedef struct _ATL_EVENT_ENTRY_ {
  UINT nControlID;         //ID identifying object instance
  const IID* piid;         //dispinterface IID
  int nOffset;             //offset of dispinterface from this pointer
  DISPID dispid;           //DISPID of method/property
  void *pfn;               //method to invoke
  _ATL_FUNC_INFO* pInfo;
} _ATL_EVENT_ENTRY;


__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_StatusTextChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_TitleChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_PropertyChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DownloadBeginStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DownloadCompleteStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnQuitStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NewWindow2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_BYREF|VT_BOOL,VT_BYREF|VT_DISPATCH}}; 
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_CommandStateChangeStruct = {CC_STDCALL, VT_EMPTY, 2, {VT_I4,VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_BeforeNavigate2Struct = {CC_STDCALL, VT_EMPTY, 7, {VT_DISPATCH,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_ProgressChangeStruct = {CC_STDCALL, VT_EMPTY, 2, {VT_I4,VT_I4}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NavigateComplete2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_BYREF|VT_VARIANT}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DocumentComplete2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_BYREF|VT_VARIANT}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnVisibleStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnToolBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnMenuBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnStatusBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnFullScreenStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnTheaterModeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_SetSecureLockIconStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_I4}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NavigateErrorStruct = {CC_STDCALL, VT_EMPTY, 5, {VT_BYREF|VT_BOOL,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_DISPATCH}};
__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_PrivacyImpactedStateChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};

#define  WEBVIEW_ENTRY_ID 0 // 控件ID

//Sink map is used to set up event handling
#define BEGIN_SINK_MAP(_class)\
  static const _class* _GetSinkMap() {\
  static const _ATL_EVENT_ENTRY map[] = {

#define SINK_ENTRY_INFO(id, iid, dispid, fn, info) {id, &iid, (intptr_t)((void*)((void*)8))-8, dispid, (void*)fn, info},
#define SINK_ENTRY_EX(id, iid, dispid, fn) SINK_ENTRY_INFO(id, iid, dispid, fn, NULL)
#define SINK_ENTRY(id, dispid, fn) SINK_ENTRY_EX(id, IID_NULL, dispid, fn)
#define END_SINK_MAP() {0, NULL, 0, 0, NULL, NULL} }; return map;}

//   static const void* _GetSinkMap()
//   {
//     static const _ATL_EVENT_ENTRY map[] = 
//     {
//       {0, &DIID_DWebBrowserEvents2, (intptr_t)((void*)((void*)8))-8, DISPID_STATUSTEXTCHANGE, 0,&CWebBrowser2EventsBase_StatusTextChangeStruct}//,
//       //{0, NULL, 0, 0, NULL, 0}
//     };
//   };

BEGIN_SINK_MAP(void)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_STATUSTEXTCHANGE, CWebBrowser2_StatusTextChange, &CWebBrowser2EventsBase_StatusTextChangeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_PROGRESSCHANGE, CWebBrowser2_ProgressChange, &CWebBrowser2EventsBase_ProgressChangeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_COMMANDSTATECHANGE, CWebBrowser2_CommandStateChange, &CWebBrowser2EventsBase_CommandStateChangeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_DOWNLOADBEGIN, CWebBrowser2_DownloadBegin, &CWebBrowser2EventsBase_DownloadBeginStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_DOWNLOADCOMPLETE, CWebBrowser2_DownloadComplete, &CWebBrowser2EventsBase_DownloadCompleteStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_TITLECHANGE, CWebBrowser2_TitleChange, &CWebBrowser2EventsBase_TitleChangeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, CWebBrowser2_NavigateComplete2, &CWebBrowser2EventsBase_NavigateComplete2Struct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_BEFORENAVIGATE2, CWebBrowser2_BeforeNavigate2, &CWebBrowser2EventsBase_BeforeNavigate2Struct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_PROPERTYCHANGE, CWebBrowser2_PropertyChange, &CWebBrowser2EventsBase_PropertyChangeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_NEWWINDOW2, CWebBrowser2_NewWindow2, &CWebBrowser2EventsBase_NewWindow2Struct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, CWebBrowser2_DocumentComplete, &CWebBrowser2EventsBase_DocumentComplete2Struct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONQUIT, CWebBrowser2_OnQuit, &CWebBrowser2EventsBase_OnQuitStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONVISIBLE, CWebBrowser2_OnVisible, &CWebBrowser2EventsBase_OnVisibleStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONTOOLBAR, CWebBrowser2_OnToolBar, &CWebBrowser2EventsBase_OnToolBarStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONMENUBAR, CWebBrowser2_OnMenuBar, &CWebBrowser2EventsBase_OnMenuBarStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONSTATUSBAR, CWebBrowser2_OnStatusBar, &CWebBrowser2EventsBase_OnStatusBarStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONFULLSCREEN, CWebBrowser2_OnFullScreen, &CWebBrowser2EventsBase_OnFullScreenStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_ONTHEATERMODE, CWebBrowser2_OnTheaterMode, &CWebBrowser2EventsBase_OnTheaterModeStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_SETSECURELOCKICON, CWebBrowser2_SetSecureLockIcon, &CWebBrowser2EventsBase_SetSecureLockIconStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_NAVIGATEERROR, CWebBrowser2_NavigateError, &CWebBrowser2EventsBase_NavigateErrorStruct)
  SINK_ENTRY_INFO(WEBVIEW_ENTRY_ID, DIID_DWebBrowserEvents2, DISPID_PRIVACYIMPACTEDSTATECHANGE, CWebBrowser2_PrivacyImpactedStateChange, &CWebBrowser2EventsBase_PrivacyImpactedStateChangeStruct)
END_SINK_MAP()




//////////////////////////////////// My IDispatch functions  //////////////////////////////////////
// The browser uses our IDispatch to give feedback when certain actions occur on the web page.
#define WEBVIEW_JS_INVOKE_ID 0x1000
#pragma pack(push,1)
typedef  union _pfn_ {
  intptr_t dwFunc;
  PVOID    pfn;
}  _pfn;
typedef struct _CComStdCallThunk {  
  void*      pVtable;
  void*      pFunc;
  intptr_t   m_mov;          // mov dword ptr [esp+4], pThis
  intptr_t   m_this;         //
  BYTE       m_jmp;          // jmp func
  intptr_t   m_relproc;      // relative jmp
  //PVOID pfnCComStdCallThunk_Init;
}CComStdCallThunk, *PCComStdCallThunk;
#pragma pack(pop)

static void CComStdCallThunk_Init(PCComStdCallThunk This, PVOID dw, void* pThis) {
  _pfn pfn;
  pfn.pfn = dw;
  This->pVtable = &This->pFunc;
  This->pFunc = &This->m_mov;
  This->m_mov = 0x042444C7;
  This->m_this = (intptr_t)pThis;
  This->m_jmp = 0xE9;
  This->m_relproc = (intptr_t)pfn.dwFunc - ((intptr_t)This+sizeof(CComStdCallThunk));
  FlushInstructionCache(GetCurrentProcess(), This, sizeof(CComStdCallThunk));
}

static ULONG STDMETHODCALLTYPE IDispatch_AddRef(IDispatch FAR *This) {
  return(InterlockedIncrement(&((_IDispatchEx *)This)->refCount));
}

static HRESULT STDMETHODCALLTYPE IDispatch_QueryInterface(IDispatch * This, REFIID riid, void **ppvObject) {
  *ppvObject = 0;

  if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IDispatch, sizeof(GUID))) {
    *ppvObject = (void *)This;

    // Increment its usage count. The caller will be expected to call our
    // IDispatch's Release() (ie, Dispatch_Release) when it's done with
    // our IDispatch.
    IDispatch_AddRef(This);

    return S_OK;
  }

  *ppvObject = 0;
  return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE IDispatch_Release(IDispatch FAR *This) {
  if (InterlockedDecrement( &((_IDispatchEx *)This)->refCount ) == 0)  {
    /* If you uncomment the following line you should get one message
     * when the document unloads for each successful call to
     * CreateEventHandler. If not, check you are setting all events
     * (that you set), to null or detaching them.
     */
    // OutputDebugString("One event handler destroyed");

    GlobalFree(((char *)This - ((_IDispatchEx *)This)->extraSize));
    return(0);
  }

  return(((_IDispatchEx *)This)->refCount);
}

static HRESULT STDMETHODCALLTYPE IDispatch_GetTypeInfoCount(IDispatch FAR *This, UINT *pctinfo) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IDispatch_GetTypeInfo(IDispatch FAR *This, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IDispatch_GetIDsOfNames(IDispatch FAR *This, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
  if (1 == cNames && 0 == wcscmp(rgszNames[0], L"invoke")) {
    rgDispId[0] = WEBVIEW_JS_INVOKE_ID;
  }
  
  return S_OK;
}

static void webDetach(_IDispatchEx *lpDispatchEx) {
//   IHTMLWindow3  *htmlWindow3;
// 
//   // Get the IHTMLWindow3 and call its detachEvent() to disconnect our _IDispatchEx
//   // from the element on the web page.
//   if (!lpDispatchEx->htmlWindow2->lpVtbl->QueryInterface(lpDispatchEx->htmlWindow2, (GUID *)&_IID_IHTMLWindow3[0], (void **)&htmlWindow3) && htmlWindow3)
//   {
//     htmlWindow3->lpVtbl->detachEvent(htmlWindow3, OnBeforeOnLoad, (LPDISPATCH)lpDispatchEx);
//     htmlWindow3->lpVtbl->Release(htmlWindow3);
//   }
// 
//   // Release any object that was originally passed to CreateEventHandler().
//   if (lpDispatchEx->object) lpDispatchEx->object->lpVtbl->Release(lpDispatchEx->object);
// 
//   // We don't need the IHTMLWindow2 any more (that we got in CreateEventHandler).
//   lpDispatchEx->htmlWindow2->lpVtbl->Release(lpDispatchEx->htmlWindow2);
}

//Helper for finding the function index for a DISPID
static HRESULT IDispatch_GetFuncInfoFromId(const IID* iid, DISPID dispidMember, LCID lcid, _ATL_FUNC_INFO* info) {
  return E_NOTIMPL;
}

static HRESULT IDispatch_InvokeFromFuncInfo(void* This, void *pEvent, _ATL_FUNC_INFO* pInfo, DISPPARAMS* pdispparams, VARIANT* pvarResult) {
  CComStdCallThunk thunk;
  VARIANT tmpResult;
  int i=0;
  VARIANTARG** pVarArgs = pInfo->nParams ? (VARIANTARG**)_alloca( sizeof(VARIANTARG*)*(pInfo->nParams) ) : 0;

  for ( i=0; i<pInfo->nParams; i++) {
    pVarArgs[i] = &pdispparams->rgvarg[pInfo->nParams - i - 1];
  }

  CComStdCallThunk_Init(&thunk, pEvent, This);
  
  if (pvarResult == NULL) {
    pvarResult = &tmpResult;
  }

  return DispCallFunc(&thunk, 0, pInfo->cc, pInfo->vtReturn, pInfo->nParams, pInfo->pVarTypes, pVarArgs, pvarResult);
}

static HRESULT STDMETHODCALLTYPE IDispatch_Invoke_org(IDispatch *This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, unsigned int *puArgErr) {
  WEBPARAMS    webParams;
  BSTR      strType;

  // Get the IHTMLEventObj from the associated window.
  if (!((_IDispatchEx *)This)->htmlWindow2->lpVtbl->get_event(((_IDispatchEx *)This)->htmlWindow2, &webParams.htmlEvent) && webParams.htmlEvent) {
    // Get the event's type (ie, a BSTR) by calling the IHTMLEventObj's get_type().
    webParams.htmlEvent->lpVtbl->get_type(webParams.htmlEvent, &strType);
    if (strType) {
      // Set the WEBPARAMS.NMHDR struct's hwndFrom to the window hosting the browser,
      // and set idFrom to 0 to let him know that this is a message that was sent
      // as a result of an action occuring on the web page (that the app asked to be
      // informed of).
      webParams.nmhdr.hwndFrom = ((_IDispatchEx *)This)->hwnd;

      // Is the type "beforeunload"? If so, set WEBPARAMS.NMHDR struct's "code" to 0, otherwise 1.
      if (!(webParams.nmhdr.code = lstrcmpW(strType, &BeforeUnload[0]))) {
        // This is a "beforeunload" event. NOTE: This should be the last event before our
        // _IDispatchEx is destroyed.

        // If the id number is negative, then this is the app's way of telling us that it
        // expects us to stuff the "eventStr" arg (passed to CreateEventHandler) into the
        // WEBPARAMS->eventStr.
        if (((_IDispatchEx *)This)->id < 0) {  
user:        webParams.eventStr = (LPCTSTR)((_IDispatchEx *)This)->userdata;
        }

        // We can free the BSTR we got above since we no longer need it.
        goto freestr;
      }

      // It's some other type of event. Set the WEBPARAMS->eventStr so he can do a lstrcmp()
      // to see what event happened.
      else {
        // Let app know that this is not a "beforeunload" event.
        webParams.nmhdr.code = 1;

        // If the app wants us to set the event type string into the WEBPARAMS, then get that
        // string as UNICODE or ANSI -- whichever is appropriate for the app window.
        if (((_IDispatchEx *)This)->id < 0) goto user;
        if (!IsWindowUnicode(webParams.nmhdr.hwndFrom))
        {
          // For ANSI, we need to convert the BSTR to an ANSI string, and then we no longer
          // need the BSTR.
          webParams.nmhdr.idFrom = WideCharToMultiByte(CP_ACP, 0, (WCHAR *)strType, -1, 0, 0, 0, 0);
          if (!(webParams.eventStr = GlobalAlloc(GMEM_FIXED, sizeof(char) * webParams.nmhdr.idFrom))) goto bad;
          WideCharToMultiByte(CP_ACP, 0, (WCHAR *)strType, -1, (char *)webParams.eventStr, webParams.nmhdr.idFrom, 0, 0);
freestr:      SysFreeString(strType);
          strType = 0;
        }

        // For UNICODE, we can use the BSTR as is. We can't free the BSTR yet.
        else
          webParams.eventStr = (LPCTSTR)strType;
      }
  
      // Send a WM_NOTIFY message to the window with the _IDispatchEx as WPARAM, and
      // the WEBPARAMS as LPARAM.
      webParams.nmhdr.idFrom = 0;
      SendMessage(webParams.nmhdr.hwndFrom, WM_NOTIFY, (WPARAM)This, (LPARAM)&webParams);

      // Free anything allocated or gotten above
bad:    if (strType) SysFreeString(strType);
      else if (webParams.eventStr && ((_IDispatchEx *)This)->id >= 0) GlobalFree((void *)webParams.eventStr);

      // If this was the "beforeunload" event, detach this IDispatch from that event for its web page element.
      // This should also cause the IE engine to call this IDispatch's Dispatch_Release().
      if (!webParams.nmhdr.code) webDetach((_IDispatchEx *)This);
    }

    // Release the IHTMLEventObj.
    webParams.htmlEvent->lpVtbl->Release(webParams.htmlEvent);
  }

  return S_OK;
}

static HRESULT STDMETHODCALLTYPE IDispatch_Invoke(IDispatch *This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, unsigned int *puArgErr) {
  size_t offset = (size_t) & ((_IOleClientSiteEx *)NULL)->external;
  _IOleClientSiteEx *ex = (_IOleClientSiteEx *)((char *)(This)-offset);
  struct webview    *wv = (struct webview *)GetWindowLongPtr(ex->inplace.frame.window, GWLP_USERDATA);

  if (pDispParams->cArgs == 1 && pDispParams->rgvarg[0].vt == VT_BSTR) {
    BSTR bstr = pDispParams->rgvarg[0].bstrVal;
    char *s = webview_from_utf16(bstr);
    if (s != NULL) {
      if (dispIdMember == WEBVIEW_JS_INVOKE_ID) {
        if (wv->external_invoke_cb != NULL) {
          wv->external_invoke_cb(wv, s);
        }
      }

      GlobalFree(s);
    }
  }

  const _ATL_EVENT_ENTRY* pMap = _GetSinkMap();
  const _ATL_EVENT_ENTRY* pFound = NULL;
  _ATL_FUNC_INFO info;
  _ATL_FUNC_INFO* pInfo;
  HRESULT hr = 0;

  printf("dispID = %d\n", dispIdMember);

  while (pMap->piid != NULL) {
    //comparing pointers here should be adequate
    if ((pMap->nControlID == WEBVIEW_ENTRY_ID) && (pMap->dispid == dispIdMember) &&  (pMap->piid == &DIID_DWebBrowserEvents2)) {
      pFound = pMap;
      break;
    }

    pMap++;
  }
  if (pFound == NULL) {
    return S_FALSE;
  }

  if (pFound->pInfo != NULL) {
    pInfo = pFound->pInfo;
  } else {
    pInfo = &info;
    hr = IDispatch_GetFuncInfoFromId(&DIID_DWebBrowserEvents2, dispIdMember, lcid, &info);
    if (FAILED(hr)) {
      return S_OK;
    }
  }

  return IDispatch_InvokeFromFuncInfo(This, pFound->pfn, pInfo, pDispParams, pVarResult);
}

static DWORD AdviseEvent(struct webview *w, LPUNKNOWN punkSink) {
  IOleObject    *browserObject;
  LPCONNECTIONPOINT pConnPt = NULL;
  DWORD dwCookie = 0;
  LPCONNECTIONPOINTCONTAINER pConnPtCont;
  
  browserObject = *w->priv.browser;
  
  // We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
  // object, so we can call some of the functions in the former's table.
  //GetEventIID( &IID_IConnectionPointContainer );

  if ((browserObject->lpVtbl != NULL) && !((browserObject->lpVtbl)->QueryInterface( browserObject, &IID_IConnectionPointContainer,(LPVOID*)&pConnPtCont ) )) {
    if (SUCCEEDED(pConnPtCont->lpVtbl->FindConnectionPoint(pConnPtCont, &DIID_DWebBrowserEvents2, &pConnPt))) {
      pConnPt->lpVtbl->Advise(pConnPt, punkSink, &dwCookie);
      
      pConnPt->lpVtbl->Release( pConnPt );
    }
    
    pConnPtCont->lpVtbl->Release( pConnPtCont );

    return dwCookie;
  }

  return 0;
}

// The VTable for our _IDispatchEx object.
static IDispatchVtbl ExternalDispatchTable = {
    IDispatch_QueryInterface,
    IDispatch_AddRef,
    IDispatch_Release,
    IDispatch_GetTypeInfoCount,
    IDispatch_GetTypeInfo,
    IDispatch_GetIDsOfNames,
    IDispatch_Invoke
};

////////////////////////////////////// My IOleClientSite functions  /////////////////////////////////////
// We give the browser object a pointer to our IOleClientSite object when we call OleCreate() or DoVerb().
// Our IOleClientSite functions that the browser may call

/************************* Site_QueryInterface() *************************
 * The browser object calls this when it wants a pointer to one of our
 * IOleClientSite, IDocHostUIHandler, or IOleInPlaceSite structures. They
 * are all accessible via the _IOleClientSiteEx struct we allocated in
 * EmbedBrowserObject() and passed to DoVerb() and OleCreate().
 *
 * This =    A pointer to whatever _IOleClientSiteEx struct we passed to
 *        OleCreate() or DoVerb().
 * riid =    A GUID struct that the browser passes us to clue us as to
 *        which type of struct (object) it would like a pointer
 *        returned for.
 * ppvObject =  Where the browser wants us to return a pointer to the
 *        appropriate struct. (ie, It passes us a handle to fill in).
 *
 * RETURNS: S_OK if we return the struct, or E_NOINTERFACE if we don't have
 * the requested struct.
*/
static HRESULT STDMETHODCALLTYPE Site_QueryInterface(IOleClientSite FAR *This, REFIID riid, void **ppvObject) {
  // It just so happens that the first arg passed to us is our _IOleClientSiteEx struct we allocated
  // and passed to DoVerb() and OleCreate(). Nevermind that 'This' is declared is an IOleClientSite *.
  // Remember that in EmbedBrowserObject(), we allocated our own _IOleClientSiteEx struct, and lied
  // to OleCreate() and DoVerb() -- passing our _IOleClientSiteEx struct and saying it was an
  // IOleClientSite struct. It's ok. An _IOleClientSiteEx starts with an embedded IOleClientSite, so
  // the browser didn't mind. So that's what the browser object is passing us now. The browser doesn't
  // know that it's really an _IOleClientSiteEx struct. But we do. So we can recast it and use it as
  // so here.

  // If the browser is asking us to match IID_IOleClientSite, then it wants us to return a pointer to
  // our IOleClientSite struct. Then the browser will use the VTable in that struct to call our
  // IOleClientSite functions. It will also pass this same pointer to all of our IOleClientSite
  // functions.
  //
  // Actually, we're going to lie to the browser again. We're going to return our own _IOleClientSiteEx
  // struct, and tell the browser that it's a IOleClientSite struct. It's ok. The first thing in our
  // _IOleClientSiteEx is an embedded IOleClientSite, so the browser doesn't mind. We want the browser
  // to continue passing our _IOleClientSiteEx pointer wherever it would normally pass a IOleClientSite
  // pointer.
  // 
  // The IUnknown interface uses the same VTable as the first object in our _IOleClientSiteEx
  // struct (which happens to be an IOleClientSite). So if the browser is asking us to match
  // IID_IUnknown, then we'll also return a pointer to our _IOleClientSiteEx.

  if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IOleClientSite, sizeof(GUID))) {
    *ppvObject = &((_IOleClientSiteEx *)This)->client;

  // If the browser is asking us to match IID_IOleInPlaceSite, then it wants us to return a pointer to
  // our IOleInPlaceSite struct. Then the browser will use the VTable in that struct to call our
  // IOleInPlaceSite functions.  It will also pass this same pointer to all of our IOleInPlaceSite
  // functions (except for Site_QueryInterface, Site_AddRef, and Site_Release. Those will always get
  // the pointer to our _IOleClientSiteEx.
  //
  // Actually, we're going to lie to the browser. We're going to return our own _IOleInPlaceSiteEx
  // struct, and tell the browser that it's a IOleInPlaceSite struct. It's ok. The first thing in
  // our _IOleInPlaceSiteEx is an embedded IOleInPlaceSite, so the browser doesn't mind. We want the
  // browser to continue passing our _IOleInPlaceSiteEx pointer wherever it would normally pass a
  // IOleInPlaceSite pointer.
  } else if (!memcmp(riid, &IID_IOleInPlaceSite, sizeof(GUID))) {
    *ppvObject = &((_IOleClientSiteEx *)This)->inplace;

  // If the browser is asking us to match IID_IDocHostUIHandler, then it wants us to return a pointer to
  // our IDocHostUIHandler struct. Then the browser will use the VTable in that struct to call our
  // IDocHostUIHandler functions.  It will also pass this same pointer to all of our IDocHostUIHandler
  // functions (except for Site_QueryInterface, Site_AddRef, and Site_Release. Those will always get
  // the pointer to our _IOleClientSiteEx.
  //
  // Actually, we're going to lie to the browser. We're going to return our own _IDocHostUIHandlerEx
  // struct, and tell the browser that it's a IDocHostUIHandler struct. It's ok. The first thing in
  // our _IDocHostUIHandlerEx is an embedded IDocHostUIHandler, so the browser doesn't mind. We want the
  // browser to continue passing our _IDocHostUIHandlerEx pointer wherever it would normally pass a
  // IDocHostUIHandler pointer. My, we're really playing dirty tricks on the browser here. heheh.
  } else if (!memcmp(riid, &IID_IDocHostUIHandler, sizeof(GUID))) {
    *ppvObject = &((_IOleClientSiteEx *)This)->ui;

  } else if (iid_eq(riid, &IID_IServiceProvider)) {
    *ppvObject = &((_IOleClientSiteEx *)This)->provider;

  // For other types of objects the browser wants, just report that we don't have any such objects.
  // NOTE: If you want to add additional functionality to your browser hosting, you may need to
  // provide some more objects here. You'll have to investigate what the browser is asking for
  // (ie, what REFIID it is passing).
  } else {
    *ppvObject = 0;
    return E_NOINTERFACE;
  }

  return S_OK;
}

static ULONG STDMETHODCALLTYPE Site_AddRef(IOleClientSite FAR *This) {
  return S_FALSE;
}

static ULONG STDMETHODCALLTYPE Site_Release(IOleClientSite FAR *This) {
  return S_FALSE;
}

static HRESULT STDMETHODCALLTYPE Site_SaveObject(IOleClientSite FAR *This) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Site_GetMoniker(IOleClientSite FAR *This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Site_GetContainer(IOleClientSite FAR *This, LPOLECONTAINER FAR *ppContainer) {
  // Tell the browser that we are a simple object and don't support a container
  *ppContainer = 0;
  return E_NOINTERFACE;
}

static HRESULT STDMETHODCALLTYPE Site_ShowObject(IOleClientSite FAR *This) {
  return NOERROR;
}

static HRESULT STDMETHODCALLTYPE Site_OnShowWindow(IOleClientSite FAR *This, BOOL fShow) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Site_RequestNewObjectLayout(IOleClientSite FAR *This) {
  return E_NOTIMPL;
}

// Our IOleClientSite VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleClientSite set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleClientSite structure.
static IOleClientSiteVtbl MyIOleClientSiteTable = {
    Site_QueryInterface,
    Site_AddRef,
    Site_Release,
    Site_SaveObject,
    Site_GetMoniker,
    Site_GetContainer,
    Site_ShowObject,
    Site_OnShowWindow,
    Site_RequestNewObjectLayout
};



////////////////////////////////////// My IOleInPlaceSite functions  /////////////////////////////////////
// The browser object asks us for the pointer to our IOleInPlaceSite object by calling our IOleClientSite's
// QueryInterface (ie, Site_QueryInterface) and specifying a REFIID of IID_IOleInPlaceSite.


// Our IOleInPlaceSite functions that the browser may call
static HRESULT STDMETHODCALLTYPE InPlace_QueryInterface(IOleInPlaceSite FAR *This, REFIID riid, LPVOID FAR *ppvObj) {
  // The browser assumes that our IOleInPlaceSite object is associated with our IOleClientSite
  // object. So it is possible that the browser may call our IOleInPlaceSite's QueryInterface()
  // to ask us to return a pointer to our IOleClientSite, in the same way that the browser calls
  // our IOleClientSite's QueryInterface() to ask for a pointer to our IOleInPlaceSite.
  //
  // Rather than duplicate much of the code in IOleClientSite's QueryInterface, let's just get
  // a pointer to our _IOleClientSiteEx object, substitute it as the 'This' arg, and call our
  // our IOleClientSite's QueryInterface. Note that since our IOleInPlaceSite is embedded right
  // inside our _IOleClientSiteEx, and comes immediately after the IOleClientSite, we can employ
  // the following trickery to get the pointer to our _IOleClientSiteEx.
  return (Site_QueryInterface((IOleClientSite *)((char *)This - sizeof(IOleClientSite)), riid, ppvObj));
}

static ULONG STDMETHODCALLTYPE InPlace_AddRef(IOleInPlaceSite FAR *This) {
  return S_FALSE;
}

static ULONG STDMETHODCALLTYPE InPlace_Release(IOleInPlaceSite FAR *This) {
  return S_FALSE;
}

static HRESULT STDMETHODCALLTYPE InPlace_GetWindow(IOleInPlaceSite FAR *This, HWND FAR *lphwnd) {
  // Return the HWND of the window that contains this browser object. We stored that
  // HWND in our _IOleInPlaceSiteEx struct. Nevermind that the function declaration for
  // Site_GetWindow says that 'This' is an IOleInPlaceSite *. Remember that in
  // EmbedBrowserObject(), we allocated our own _IOleInPlaceSiteEx struct which
  // contained an embedded IOleInPlaceSite struct within it. And when the browser
  // called Site_QueryInterface() to get a pointer to our IOleInPlaceSite object, we
  // returned a pointer to our _IOleClientSiteEx. The browser doesn't know this. But
  // we do. That's what we're really being passed, so we can recast it and use it as
  // so here.
  *lphwnd = ((_IOleInPlaceSiteEx FAR *)This)->frame.window;
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_ContextSensitiveHelp(IOleInPlaceSite FAR *This, BOOL fEnterMode) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE InPlace_CanInPlaceActivate(IOleInPlaceSite FAR *This) {
  // Tell the browser we can in place activate
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceActivate(IOleInPlaceSite FAR *This) {
  // Tell the browser we did it ok
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_OnUIActivate(IOleInPlaceSite FAR *This) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_GetWindowContext(IOleInPlaceSite FAR *This, LPOLEINPLACEFRAME FAR *lplpFrame, LPOLEINPLACEUIWINDOW FAR *lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) {
  // Give the browser the pointer to our IOleInPlaceFrame struct. We stored that pointer
  // in our _IOleInPlaceSiteEx struct. Nevermind that the function declaration for
  // Site_GetWindowContext says that 'This' is an IOleInPlaceSite *. Remember that in
  // EmbedBrowserObject(), we allocated our own _IOleInPlaceSiteEx struct which
  // contained an embedded IOleInPlaceSite struct within it. And when the browser
  // called Site_QueryInterface() to get a pointer to our IOleInPlaceSite object, we
  // returned a pointer to our _IOleClientSiteEx. The browser doesn't know this. But
  // we do. That's what we're really being passed, so we can recast it and use it as
  // so here.
  //
  // Actually, we're giving the browser a pointer to our own _IOleInPlaceSiteEx struct,
  // but telling the browser that it's a IOleInPlaceSite struct. No problem. Our
  // _IOleInPlaceSiteEx starts with an embedded IOleInPlaceSite, so the browser is
  // cool with it. And we want the browser to pass a pointer to this _IOleInPlaceSiteEx
  // wherever it would pass a IOleInPlaceSite struct to our IOleInPlaceSite functions.
  *lplpFrame = (LPOLEINPLACEFRAME) & ((_IOleInPlaceSiteEx *)This)->frame;

  // We have no OLEINPLACEUIWINDOW
  *lplpDoc = 0;

  // Fill in some other info for the browser
  lpFrameInfo->fMDIApp = FALSE;
  lpFrameInfo->hwndFrame = ((_IOleInPlaceFrameEx *)*lplpFrame)->window;
  lpFrameInfo->haccel = 0;
  lpFrameInfo->cAccelEntries = 0;

  // Give the browser the dimensions of where it can draw. We give it our entire window to fill.
  // We do this in InPlace_OnPosRectChange() which is called right when a window is first
  // created anyway, so no need to duplicate it here.
  //  GetClientRect(lpFrameInfo->hwndFrame, lprcPosRect);
  //  GetClientRect(lpFrameInfo->hwndFrame, lprcClipRect);

  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_Scroll(IOleInPlaceSite FAR *This, SIZE scrollExtent) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE InPlace_OnUIDeactivate(IOleInPlaceSite FAR *This, BOOL fUndoable) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceDeactivate(IOleInPlaceSite FAR *This) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE InPlace_DiscardUndoState(IOleInPlaceSite FAR *This) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE InPlace_DeactivateAndUndo(IOleInPlaceSite FAR *This) {
  return E_NOTIMPL;
}

// Called when the position of the browser object is changed, such as when we call the IWebBrowser2's put_Width(),
// put_Height(), put_Left(), or put_Right().
static HRESULT STDMETHODCALLTYPE InPlace_OnPosRectChange(IOleInPlaceSite FAR *This, LPCRECT lprcPosRect) {
  IOleObject *browserObject;
  IOleInPlaceObject *inplace;

  // We need to get the browser's IOleInPlaceObject object so we can call its SetObjectRects
  // function.
  browserObject = *((IOleObject **)((char *)This - sizeof(IOleObject *) - sizeof(IOleClientSite)));
  if (!browserObject->lpVtbl->QueryInterface(browserObject, iid_unref(&IID_IOleInPlaceObject), (void **)&inplace)) {
    // Give the browser the dimensions of where it can draw.
    inplace->lpVtbl->SetObjectRects(inplace, lprcPosRect, lprcPosRect);
    inplace->lpVtbl->Release(inplace);
  }
  return S_OK;
}

// Our IOleInPlaceSite VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleInPlaceSite set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleInPlaceSite structure.
static IOleInPlaceSiteVtbl MyIOleInPlaceSiteTable = {
    InPlace_QueryInterface,
    InPlace_AddRef,
    InPlace_Release,
    InPlace_GetWindow,
    InPlace_ContextSensitiveHelp,
    InPlace_CanInPlaceActivate,
    InPlace_OnInPlaceActivate,
    InPlace_OnUIActivate,
    InPlace_GetWindowContext,
    InPlace_Scroll,
    InPlace_OnUIDeactivate,
    InPlace_OnInPlaceDeactivate,
    InPlace_DiscardUndoState,
    InPlace_DeactivateAndUndo,
    InPlace_OnPosRectChange
};


////////////////////////////////////// My IOleInPlaceFrame functions  ////////////////////////////////////////
// Our IOleInPlaceFrame functions that the browser may call

static HRESULT STDMETHODCALLTYPE Frame_QueryInterface(IOleInPlaceFrame FAR *This, REFIID riid, LPVOID FAR *ppvObj) {
  return E_NOTIMPL;
}

static ULONG STDMETHODCALLTYPE Frame_AddRef(IOleInPlaceFrame FAR *This) {
  return S_FALSE;
}

static ULONG STDMETHODCALLTYPE Frame_Release(IOleInPlaceFrame FAR *This) {
  return S_FALSE;
}

static HRESULT STDMETHODCALLTYPE Frame_GetWindow(IOleInPlaceFrame FAR *This, HWND FAR *lphwnd) {
  // Give the browser the HWND to our window that contains the browser object. We
  // stored that HWND in our IOleInPlaceFrame struct. Nevermind that the function
  // declaration for Frame_GetWindow says that 'This' is an IOleInPlaceFrame *. Remember
  // that in EmbedBrowserObject(), we allocated our own IOleInPlaceFrameEx struct which
  // contained an embedded IOleInPlaceFrame struct within it. And then we lied when
  // Site_GetWindowContext() returned that IOleInPlaceFrameEx. So that's what the
  // browser is passing us. It doesn't know that. But we do. So we can recast it and
  // use it as so here.
  *lphwnd = ((_IOleInPlaceFrameEx *)This)->window;
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE Frame_ContextSensitiveHelp(IOleInPlaceFrame FAR *This, BOOL fEnterMode) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_GetBorder(IOleInPlaceFrame FAR *This, LPRECT lprectBorder) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_RequestBorderSpace(IOleInPlaceFrame FAR *This, LPCBORDERWIDTHS pborderwidths) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_SetBorderSpace(IOleInPlaceFrame FAR *This, LPCBORDERWIDTHS pborderwidths) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_SetActiveObject(IOleInPlaceFrame FAR *This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE Frame_InsertMenus(IOleInPlaceFrame FAR *This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_SetMenu(IOleInPlaceFrame FAR *This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE Frame_RemoveMenus(IOleInPlaceFrame FAR *This, HMENU hmenuShared) {
  return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Frame_SetStatusText(IOleInPlaceFrame FAR *This, LPCOLESTR pszStatusText) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE Frame_EnableModeless(IOleInPlaceFrame FAR *This, BOOL fEnable) {
  return S_OK;
}

static HRESULT STDMETHODCALLTYPE Frame_TranslateAccelerator(IOleInPlaceFrame FAR *This, LPMSG lpmsg, WORD wID) {
  return E_NOTIMPL;
}

// Our IOleInPlaceFrame VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleInPlaceFrame set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleInPlaceFrame structure.
static IOleInPlaceFrameVtbl MyIOleInPlaceFrameTable = {
    Frame_QueryInterface,
    Frame_AddRef,
    Frame_Release,
    Frame_GetWindow,
    Frame_ContextSensitiveHelp,
    Frame_GetBorder,
    Frame_RequestBorderSpace,
    Frame_SetBorderSpace,
    Frame_SetActiveObject,
    Frame_InsertMenus,
    Frame_SetMenu,
    Frame_RemoveMenus,
    Frame_SetStatusText,
    Frame_EnableModeless,
    Frame_TranslateAccelerator
};



//////////////////////////////////// My IDocHostUIHandler functions  //////////////////////////////////////
// The browser object asks us for the pointer to our IDocHostUIHandler object by calling our IOleClientSite's
// QueryInterface (ie, Site_QueryInterface) and specifying a REFIID of IID_IDocHostUIHandler.
//
// NOTE: You need at least IE 4.0. Previous versions do not ask for, nor utilize, our IDocHostUIHandler functions.
//
// Our IDocHostUIHandler functions that the browser may call
//
// The browser assumes that our IDocHostUIHandler object is associated with our IOleClientSite
// object. So it is possible that the browser may call our IDocHostUIHandler's QueryInterface()
// to ask us to return a pointer to our IOleClientSite, in the same way that the browser calls
// our IOleClientSite's QueryInterface() to ask for a pointer to our IDocHostUIHandler.
//
// Rather than duplicate much of the code in IOleClientSite's QueryInterface, let's just get
// a pointer to our _IOleClientSiteEx object, substitute it as the 'This' arg, and call our
// our IOleClientSite's QueryInterface. Note that since our _IDocHostUIHandlerEx is embedded right
// inside our _IOleClientSiteEx, and comes immediately after the _IOleInPlaceSiteEx, we can employ
// the following trickery to get the pointer to our _IOleClientSiteEx.
static HRESULT STDMETHODCALLTYPE UI_QueryInterface(IDocHostUIHandler FAR *This, REFIID riid, LPVOID FAR *ppvObj) {
  return (Site_QueryInterface((IOleClientSite *)((char *)This - sizeof(IOleClientSite) - sizeof(_IOleInPlaceSiteEx)), riid, ppvObj));
}

static ULONG STDMETHODCALLTYPE UI_AddRef(IDocHostUIHandler FAR *This) {
  return S_FALSE;
}

static ULONG STDMETHODCALLTYPE UI_Release(IDocHostUIHandler FAR *This) {
  return S_FALSE;
}

// Called when the browser object is about to display its context menu.
//
// If desired, we can pop up your own custom context menu here. Of course,
// we would need some way to get our window handle, so what we'd probably
// do is like what we did with our IOleInPlaceFrame object. We'd define and create
// our own IDocHostUIHandlerEx object with an embedded IDocHostUIHandler at the
// start of it. Then we'd add an extra HWND field to store our window handle.
// It could look like this:
//
// typedef struct _IDocHostUIHandlerEx {
//    IDocHostUIHandler  ui;    // The IDocHostUIHandler must be first!
//    HWND        window;
// } IDocHostUIHandlerEx;
//
// Of course, we'd need a unique IDocHostUIHandlerEx object for each window, so
// EmbedBrowserObject() would have to allocate one of those too. And that's
// what we'd pass to our browser object (and it would in turn pass it to us
// here, instead of 'This' being a IDocHostUIHandler *).
//
// We will return S_OK to tell the browser not to display its default context menu,
// or return S_FALSE to let the browser show its default context menu. For this
// example, we wish to disable the browser's default context menu.
static HRESULT STDMETHODCALLTYPE UI_ShowContextMenu(IDocHostUIHandler FAR *This, DWORD dwID, POINT __RPC_FAR *ppt, IUnknown __RPC_FAR *pcmdtReserved, IDispatch __RPC_FAR *pdispReserved) {
  return S_FALSE;
}

// Called at initialization of the browser object UI. We can set various features of the browser object here.
static HRESULT STDMETHODCALLTYPE UI_GetHostInfo(IDocHostUIHandler FAR *This, DOCHOSTUIINFO __RPC_FAR *pInfo) {
  pInfo->cbSize = sizeof(DOCHOSTUIINFO);

  // Set some flags. We don't want any 3D border. You can do other things like hide
  // the scroll bar (DOCHOSTUIFLAG_SCROLL_NO), display picture display (DOCHOSTUIFLAG_NOPICS),
  // disable any script running when the page is loaded (DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE),
  // open a site in a new browser window when the user clicks on some link (DOCHOSTUIFLAG_OPENNEWWIN),
  // and lots of other things. See the MSDN docs on the DOCHOSTUIINFO struct passed to us.
  pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER;
  //pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE ;

  // Set what happens when the user double-clicks on the object. Here we use the default.
  pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
  return S_OK;
}

// Called when the browser object shows its UI. This allows us to replace its menus and toolbars by creating our
// own and displaying them here.
//
// We've already got our own UI in place so just return S_OK to tell the browser
// not to display its menus/toolbars. Otherwise we'd return S_FALSE to let it do
// that.
static HRESULT STDMETHODCALLTYPE UI_ShowUI(IDocHostUIHandler FAR *This, DWORD dwID, IOleInPlaceActiveObject __RPC_FAR *pActiveObject, IOleCommandTarget __RPC_FAR *pCommandTarget, IOleInPlaceFrame __RPC_FAR *pFrame, IOleInPlaceUIWindow __RPC_FAR *pDoc) {
  return S_OK;
}

// Called when browser object hides its UI. This allows us to hide any menus/toolbars we created in ShowUI.
static HRESULT STDMETHODCALLTYPE UI_HideUI(IDocHostUIHandler FAR *This) {
  return S_OK;
}

// Called when the browser object wants to notify us that the command state has changed. We should update any
// controls we have that are dependent upon our embedded object, such as "Back", "Forward", "Stop", or "Home"
// buttons.
//
// We update our UI in our window message loop so we don't do anything here.
static HRESULT STDMETHODCALLTYPE UI_UpdateUI(IDocHostUIHandler FAR *This) {
  return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's EnableModeless() function. Also
// called when the browser displays a modal dialog box.
static HRESULT STDMETHODCALLTYPE UI_EnableModeless(IDocHostUIHandler FAR *This, BOOL fEnable) {
  return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's OnDocWindowActivate() function.
// This informs off of when the object is getting/losing the focus.
static HRESULT STDMETHODCALLTYPE UI_OnDocWindowActivate(IDocHostUIHandler FAR *This, BOOL fActivate) {
  return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's OnFrameWindowActivate() function.
static HRESULT STDMETHODCALLTYPE UI_OnFrameWindowActivate(IDocHostUIHandler FAR *This, BOOL fActivate) {
  return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's ResizeBorder() function.
static HRESULT STDMETHODCALLTYPE UI_ResizeBorder(IDocHostUIHandler FAR *This, LPCRECT prcBorder, IOleInPlaceUIWindow __RPC_FAR *pUIWindow, BOOL fRameWindow) {
  return S_OK;
}

// Called from the browser object's TranslateAccelerator routines to translate key strokes to commands.
// We don't intercept any keystrokes, so we do nothing here. But for example, if we wanted to
// override the TAB key, perhaps do something with it ourselves, and then tell the browser
// not to do anything with this keystroke, we'd do:
//
//  if (pMsg && pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_TAB)
//  {
//    // Here we do something as a result of a TAB key press.
//
//    // Tell the browser not to do anything with it.
//    return(S_FALSE);
//  }
//
//  // Otherwise, let the browser do something with this message.
//  return(S_OK);

// For our example, we want to make sure that the user can invoke some key to popup the context
// menu, so we'll tell it to ignore all messages.
static HRESULT STDMETHODCALLTYPE UI_TranslateAccelerator(IDocHostUIHandler FAR *This, LPMSG lpMsg, const GUID __RPC_FAR *pguidCmdGroup, DWORD nCmdID) {
  return S_FALSE;
}

// Called by the browser object to find where the host wishes the browser to get its options in the registry.
// We can use this to prevent the browser from using its default settings in the registry, by telling it to use
// some other registry key we've setup with the options we want.
// Let the browser use its default registry settings.
static HRESULT STDMETHODCALLTYPE UI_GetOptionKeyPath(IDocHostUIHandler FAR *This, LPOLESTR __RPC_FAR *pchKey, DWORD dw) {
  return S_FALSE;
}

// Called by the browser object when it is used as a drop target. We can supply our own IDropTarget object,
// IDropTarget functions, and IDropTarget VTable if we want to determine what happens when someone drags and
// drops some object on our embedded browser object.
// Return our IDropTarget object associated with this IDocHostUIHandler object. I don't
// know why we don't do this via UI_QueryInterface(), but we don't.

// NOTE: If we want/need an IDropTarget interface, then we would have had to setup our own
// IDropTarget functions, IDropTarget VTable, and create an IDropTarget object. We'd want to put
// a pointer to the IDropTarget object in our own custom IDocHostUIHandlerEx object (like how
// we may add an HWND field for the use of UI_ShowContextMenu). So when we created our
// IDocHostUIHandlerEx object, maybe we'd add a 'idrop' field to the end of it, and
// store a pointer to our IDropTarget object there. Then we could return this pointer as so:
//
// *pDropTarget = ((IDocHostUIHandlerEx *)This)->idrop;
// return(S_OK);

// But for our purposes, we don't need an IDropTarget object, so we'll tell whomever is calling
// us that we don't have one.
static HRESULT STDMETHODCALLTYPE UI_GetDropTarget(IDocHostUIHandler FAR *This, IDropTarget __RPC_FAR *pDropTarget, IDropTarget __RPC_FAR *__RPC_FAR *ppDropTarget) {
  return S_FALSE;
}

// Called by the browser when it wants a pointer to our IDispatch object. This object allows us to expose
// our own automation interface (ie, our own COM objects) to other entities that are running within the
// context of the browser so they can call our functions if they want. An example could be a javascript
// running in the URL we display could call our IDispatch functions. We'd write them so that any args passed
// to them would use the generic datatypes like a BSTR for utmost flexibility.
//
// Return our IDispatch object associated with this IDocHostUIHandler object. I don't
// know why we don't do this via UI_QueryInterface(), but we don't.

// NOTE: If we want/need an IDispatch interface, then we would have had to setup our own
// IDispatch functions, IDispatch VTable, and create an IDispatch object. We'd want to put
// a pointer to the IDispatch object in our custom _IDocHostUIHandlerEx object (like how
// we may add an HWND field for the use of UI_ShowContextMenu). So when we defined our
// _IDocHostUIHandlerEx object, maybe we'd add a 'idispatch' field to the end of it, and
// store a pointer to our IDispatch object there. Then we could return this pointer as so:
//
// *ppDispatch = ((_IDocHostUIHandlerEx *)This)->idispatch;
// return(S_OK);

// But for our purposes, we don't need an IDispatch object, so we'll tell whomever is calling
// us that we don't have one. Note: We must set ppDispatch to 0 if we don't return our own
// IDispatch object.
static HRESULT STDMETHODCALLTYPE UI_GetExternal(IDocHostUIHandler FAR *This, IDispatch __RPC_FAR *__RPC_FAR *ppDispatch) {
  *ppDispatch = (IDispatch *)(This + 1);
  return S_OK;
}

// Called by the browser object to give us an opportunity to modify the URL to be loaded.
static HRESULT STDMETHODCALLTYPE UI_TranslateUrl(IDocHostUIHandler FAR *This, DWORD dwTranslate, OLECHAR __RPC_FAR *pchURLIn, OLECHAR __RPC_FAR *__RPC_FAR *ppchURLOut) {
  unsigned short  *src;
  unsigned short  *dest;
  DWORD      len;

  // Get length of URL
  src = pchURLIn;
  while ((*(src)++));
  --src;
  len = src - pchURLIn; 

  // See if the URL starts with 'app:'
  if (len >= 4 && !_wcsnicmp(pchURLIn, (wchar_t *)&AppUrl[0], 4)) {
    // Allocate a new buffer to return an "about:blank" URL
    if ((dest = (OLECHAR *)CoTaskMemAlloc(12<<1))) {
      HWND  hwnd;

      *ppchURLOut = dest;

      // Return "about:blank"
      CopyMemory(dest, &Blank[0], 12<<1);

      // Convert the number after the "app:"
      len = asciiToNum(pchURLIn + 4);

      // Get our host window. That was stored in our _IOleInPlaceFrameEx
      hwnd = ((_IOleInPlaceSiteEx *)((char *)This - sizeof(_IOleInPlaceSiteEx)))->frame.window;

      // Post a message to this window using WM_APP + 50, and pass the number converted above.
      // Do not SendMessage()!. Post instead, since the browser does not like us changing
      // the URL within this here callback. Why not WM_APP? Because the browser control appears
      // to use messages within the range of WM_APP to WM_APP + 2 (and maybe more). Microsoft
      // is a bit careless here. Supposedly, the operating system components, of which
      // MSHTML.DLL is now a part, are not supposed to use WM_APP and over, but no...
      // What we should probably do is RegisterWindowMessage()
      PostMessage(hwnd, WM_APP + 50, (WPARAM)len, 0);

      // Tell browser that we returned a URL
      return(S_OK);
    }
  }

  // We don't need to modify the URL. Note: We need to set ppchURLOut to 0 if we don't
  // return an OLECHAR (buffer) containing a modified version of pchURLIn.
  *ppchURLOut = 0;

  return S_FALSE;
}

// Called by the browser when it does cut/paste to the clipboard. This allows us to block certain clipboard
// formats or support additional clipboard formats.
static HRESULT STDMETHODCALLTYPE UI_FilterDataObject(IDocHostUIHandler FAR* This, IDataObject __RPC_FAR *pDO, IDataObject __RPC_FAR *__RPC_FAR *ppDORet) {
  // Return our IDataObject object associated with this IDocHostUIHandler object. I don't
  // know why we don't do this via UI_QueryInterface(), but we don't.

  // NOTE: If we want/need an IDataObject interface, then we would have had to setup our own
  // IDataObject functions, IDataObject VTable, and create an IDataObject object. We'd want to put
  // a pointer to the IDataObject object in our custom _IDocHostUIHandlerEx object (like how
  // we may add an HWND field for the use of UI_ShowContextMenu). So when we defined our
  // _IDocHostUIHandlerEx object, maybe we'd add a 'idata' field to the end of it, and
  // store a pointer to our IDataObject object there. Then we could return this pointer as so:
  //
  // *ppDORet = ((_IDocHostUIHandlerEx FAR *)This)->idata;
    // return(S_OK);

  // But for our purposes, we don't need an IDataObject object, so we'll tell whomever is calling
  // us that we don't have one. Note: We must set ppDORet to 0 if we don't return our own
  // IDataObject object.
  *ppDORet = 0;
  return S_FALSE;
}

// Our IDocHostUIHandler VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to replace/set certain user interface considerations
// (such as whether to display a pop-up context menu when the user right-clicks on the embedded
// browser object). We must define a particular set of functions that comprise the
// IDocHostUIHandler set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IDocHostUIHandler structure.
static IDocHostUIHandlerVtbl MyIDocHostUIHandlerTable = {
    UI_QueryInterface,
    UI_AddRef,
    UI_Release,
    UI_ShowContextMenu,
    UI_GetHostInfo,
    UI_ShowUI,
    UI_HideUI,
    UI_UpdateUI,
    UI_EnableModeless,
    UI_OnDocWindowActivate,
    UI_OnFrameWindowActivate,
    UI_ResizeBorder,
    UI_TranslateAccelerator,
    UI_GetOptionKeyPath,
    UI_GetDropTarget,
    UI_GetExternal,
    UI_TranslateUrl,
    UI_FilterDataObject
};




static HRESULT STDMETHODCALLTYPE IS_QueryInterface(IInternetSecurityManager FAR *This, REFIID riid, void **ppvObject) {
  return E_NOTIMPL;
}
static ULONG STDMETHODCALLTYPE IS_AddRef(IInternetSecurityManager FAR *This) { return 1; }
static ULONG STDMETHODCALLTYPE IS_Release(IInternetSecurityManager FAR *This) { return 1; }
static HRESULT STDMETHODCALLTYPE IS_SetSecuritySite(IInternetSecurityManager FAR *This, IInternetSecurityMgrSite *pSited) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_GetSecuritySite(IInternetSecurityManager FAR *This, IInternetSecurityMgrSite **ppSite) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_MapUrlToZone(IInternetSecurityManager FAR *This, LPCWSTR pwszUrl, DWORD *pdwZone, DWORD dwFlags) {
  *pdwZone = URLZONE_LOCAL_MACHINE;
  return S_OK;
}
static HRESULT STDMETHODCALLTYPE IS_GetSecurityId(IInternetSecurityManager FAR *This, LPCWSTR pwszUrl, BYTE *pbSecurityId, DWORD *pcbSecurityId, DWORD_PTR dwReserved) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_ProcessUrlAction(IInternetSecurityManager FAR *This, LPCWSTR pwszUrl, DWORD dwAction, BYTE *pPolicy,  DWORD cbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwFlags, DWORD dwReserved) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_QueryCustomPolicy(IInternetSecurityManager FAR *This, LPCWSTR pwszUrl, REFGUID guidKey, BYTE **ppPolicy, DWORD *pcbPolicy, BYTE *pContext, DWORD cbContext, DWORD dwReserved) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_SetZoneMapping(IInternetSecurityManager FAR *This, DWORD dwZone, LPCWSTR lpszPattern, DWORD dwFlags) {
  return INET_E_DEFAULT_ACTION;
}
static HRESULT STDMETHODCALLTYPE IS_GetZoneMappings(IInternetSecurityManager FAR *This, DWORD dwZone, IEnumString **ppenumString, DWORD dwFlags) {
  return INET_E_DEFAULT_ACTION;
}
static IInternetSecurityManagerVtbl MyInternetSecurityManagerTable = {IS_QueryInterface, IS_AddRef, IS_Release, IS_SetSecuritySite, IS_GetSecuritySite, IS_MapUrlToZone, IS_GetSecurityId, IS_ProcessUrlAction, IS_QueryCustomPolicy, IS_SetZoneMapping, IS_GetZoneMappings};

static HRESULT STDMETHODCALLTYPE SP_QueryInterface(IServiceProvider FAR *This, REFIID riid, void **ppvObject) {
  return (Site_QueryInterface(
      (IOleClientSite *)((char *)This - sizeof(IOleClientSite) - sizeof(_IOleInPlaceSiteEx) - sizeof(_IDocHostUIHandlerEx) - sizeof(IDispatch)), riid, ppvObject));
}
static ULONG STDMETHODCALLTYPE SP_AddRef(IServiceProvider FAR *This) { return 1; }
static ULONG STDMETHODCALLTYPE SP_Release(IServiceProvider FAR *This) { return 1; }
static HRESULT STDMETHODCALLTYPE SP_QueryService(IServiceProvider FAR *This, REFGUID siid, REFIID riid, void **ppvObject) {
  if (iid_eq(siid, &IID_IInternetSecurityManager) && iid_eq(riid, &IID_IInternetSecurityManager)) {
    *ppvObject = &((_IServiceProviderEx *)This)->mgr;
  } else {
    *ppvObject = 0;
    return (E_NOINTERFACE);
  }
  return S_OK;
}
static IServiceProviderVtbl MyServiceProviderTable = {SP_QueryInterface, SP_AddRef, SP_Release, SP_QueryService};

/*************************** UnEmbedBrowserObject() ************************
 * Called to detach the browser object from our host window, and free its
 * resources, right before we destroy our window.
 *
 * hwnd =    Handle to the window hosting the browser object.
 *
 * NOTE: The pointer to the browser object must have been stored in the
 * window's USERDATA field. In other words, don't call UnEmbedBrowserObject().
 * with a HWND that wasn't successfully passed to EmbedBrowserObject().
*/
static void UnEmbedBrowserObject(struct webview *w) {
  if (w->priv.browser != NULL) {
    // Unembed the browser object, and release its resources.
    (*w->priv.browser)->lpVtbl->Close(*w->priv.browser, OLECLOSE_NOSAVE);
    (*w->priv.browser)->lpVtbl->Release(*w->priv.browser);
    GlobalFree(w->priv.browser);
    w->priv.browser = NULL;
  }
}

/***************************** EmbedBrowserObject() **************************
 * Puts the browser object inside our host window, and save a pointer to this
 * window's browser object in the window's GWL_USERDATA field.
 *
 * hwnd =    Handle of our window into which we embed the browser object.
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: We tell the browser object to occupy the entire client area of the
 * window.
 *
 * NOTE: No HTML page will be displayed here. We can do that with a subsequent
 * call to either DisplayHTMLPage() or DisplayHTMLStr(). This is merely once-only
 * initialization for using the browser object. In a nutshell, what we do
 * here is get a pointer to the browser object in our window's GWL_USERDATA
 * so we can access that object's functions whenever we want, and we also pass
 * pointers to our IOleClientSite, IOleInPlaceFrame, and IStorage structs so that
 * the browser can call our functions in our struct's VTables.
*/
static int EmbedBrowserObject(struct webview *w) {
  RECT              rect;
  LPCLASSFACTORY    pClassFactory = NULL;
  IWebBrowser2      *webBrowser2  = NULL;
  _IOleClientSiteEx *_iOleClientSiteEx = NULL;

  IOleObject **browserObject = (IOleObject **)GlobalAlloc(GMEM_FIXED, sizeof(IOleObject *) + sizeof(_IOleClientSiteEx));
  if (browserObject == NULL) {
    return -1;
  }
  w->priv.browser = browserObject;

  // Our IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame functions need to get our window handle. We
  // could store that in some global. But then, that would mean that our functions would work with only that
  // one window. If we want to create multiple windows, each hosting its own browser object (to display its
  // own web page), then we need to create unique IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame
  // structs for each window. And we'll put an extra field at the end of those structs to store our extra
  // data such as a window handle. So, our functions won't have to touch global data, and can therefore be
  // re-entrant and work with multiple objects/windows.
  //
  // Remember that a pointer to our IOleClientSite we create here will be passed as the first arg to every
  // one of our IOleClientSite functions. Ditto with the IOleInPlaceFrame object we create here, and the
  // IOleInPlaceFrame functions. So, our functions are able to retrieve the window handle we'll store here,
  // and then, they'll work with all such windows containing a browser control.
  //
  // Furthermore, since the browser will be calling our IOleClientSite's QueryInterface to get a pointer to
  // our IOleInPlaceSite and IDocHostUIHandler objects, that means that our IOleClientSite QueryInterface
  // must have an easy way to grab those pointers. Probably the easiest thing to do is just embed our
  // IOleInPlaceSite and IDocHostUIHandler objects inside of an extended IOleClientSite which we'll call
  // a _IOleClientSiteEx. As long as they come after the pointer to the IOleClientSite VTable, then we're
  // ok.
  //
  // Of course, we need to GlobalAlloc the above structs now. We'll just get all 4 with a single call to
  // GlobalAlloc, especially since 3 of them are all contained inside of our _IOleClientSiteEx anyway.
  //
  // Um, we're not actually allocating a IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame structs. Because
  // we're appending our own fields to them, we're getting an IOleInPlaceFrameEx and an _IOleClientSiteEx (which
  // contains both the IOleClientSite and IOleInPlaceSite. But as far as the browser is concerned, it thinks that
  // we're giving it the plain old IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame.
  //
  // So, we're not actually allocating separate IOleClientSite, IOleInPlaceSite, IOleInPlaceFrame and
  // IDocHostUIHandler structs.
  //

  // Initialize our IOleClientSite object with a pointer to our IOleClientSite VTable.
  _iOleClientSiteEx = (_IOleClientSiteEx *)(browserObject + 1);
  _iOleClientSiteEx->client.lpVtbl = &MyIOleClientSiteTable;

  // Initialize our IOleInPlaceSite object with a pointer to our IOleInPlaceSite VTable.
  _iOleClientSiteEx->inplace.inplace.lpVtbl = &MyIOleInPlaceSiteTable;

  // Initialize our IOleInPlaceFrame object with a pointer to our IOleInPlaceFrame VTable.
  _iOleClientSiteEx->inplace.frame.frame.lpVtbl = &MyIOleInPlaceFrameTable;

  // Save our HWND (in the IOleInPlaceFrame object) so our IOleInPlaceFrame functions can retrieve it.
  _iOleClientSiteEx->inplace.frame.window = w->priv.hwnd;

  // Initialize our IDocHostUIHandler object with a pointer to our IDocHostUIHandler VTable.
  _iOleClientSiteEx->ui.ui.lpVtbl = &MyIDocHostUIHandlerTable;

  _iOleClientSiteEx->external.lpVtbl = &ExternalDispatchTable;

  _iOleClientSiteEx->provider.provider.lpVtbl = &MyServiceProviderTable;
  _iOleClientSiteEx->provider.mgr.mgr.lpVtbl = &MyInternetSecurityManagerTable;

  // Get a pointer to the browser object and lock it down (so it doesn't "disappear" while we're using
  // it in this program). We do this by calling the OS function CoGetClassObject().
  //  
  // NOTE: We need this pointer to interact with and control the browser. With normal WIN32 controls such as a
  // Static, Edit, Combobox, etc, you obtain an HWND and send messages to it with SendMessage(). Not so with
  // the browser object. You need to get a pointer to its "base structure" (as returned by CoGetClassObject()). This
  // structure contains an array of pointers to functions you can call within the browser object. Actually, the
  // base structure contains a 'lpVtbl' field that is a pointer to that array. We'll call the array a 'VTable'.
  //
  // For example, the browser object happens to have a SetHostNames() function we want to call. So, after we
  // retrieve the pointer to the browser object (in a local we'll name 'browserObject'), then we can call that
  // function, and pass it args, as so:
  //
  // browserObject->lpVtbl->SetHostNames(browserObject, SomeString, SomeString);
  //
  // There's our pointer to the browser object in 'browserObject'. And there's the pointer to the browser object's
  // VTable in 'browserObject->lpVtbl'. And the pointer to the SetHostNames function happens to be stored in an
  // field named 'SetHostNames' within the VTable. So we are actually indirectly calling SetHostNames by using
  // a pointer to it. That's how you use a VTable.
  //
  // NOTE: We pass our _IOleClientSiteEx struct and lie -- saying that it's a IOleClientSite. It's ok. A
  // _IOleClientSiteEx struct starts with an embedded IOleClientSite. So the browser won't care, and we want
  // this extended struct passed to our IOleClientSite functions.

  // Get a pointer to the browser object's IClassFactory object via CoGetClassObject()
  pClassFactory = 0;
  if (!CoGetClassObject(iid_unref(&CLSID_WebBrowser), CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER, NULL, iid_unref(&IID_IClassFactory), (void **)&pClassFactory) && pClassFactory) {
    // Call the IClassFactory's CreateInstance() to create a browser object
    if (!pClassFactory->lpVtbl->CreateInstance(pClassFactory, 0, iid_unref(&IID_IOleObject), (void **)browserObject)) {
      // Free the IClassFactory. We need it only to create a browser object instance
      pClassFactory->lpVtbl->Release(pClassFactory);

      // Ok, we now have the pointer to the browser object in 'browserObject'. Let's save this in the
      // memory block we allocated above, and then save the pointer to that whole thing in our window's
      // USERDATA field. That way, if we need multiple windows each hosting its own browser object, we can
      // call EmbedBrowserObject() for each one, and easily associate the appropriate browser object with
      // its matching window and its own objects containing per-window data.
      //*((IOleObject **)ptr) = browserObject;
      //SetWindowLong(hwnd, GWL_USERDATA, (LONG)ptr);

      // Give the browser a pointer to my IOleClientSite object
      if (!(*browserObject)->lpVtbl->SetClientSite(*browserObject, (IOleClientSite *)_iOleClientSiteEx)) {
        // We can now call the browser object's SetHostNames function. SetHostNames lets the browser object know our
        // application's name and the name of the document in which we're embedding the browser. (Since we have no
        // document name, we'll pass a 0 for the latter). When the browser object is opened for editing, it displays
        // these names in its titlebar.
        //  
        // We are passing 3 args to SetHostNames. You'll note that the first arg to SetHostNames is the base
        // address of our browser control. This is something that you always have to remember when working in C
        // (as opposed to C++). When calling a VTable function, the first arg to that function must always be the
        // structure which contains the VTable. (In this case, that's the browser control itself). Why? That's
        // because that function is always assumed to be written in C++. And the first argument to any C++ function
        // must be its 'this' pointer (ie, the base address of its class, which in this case is our browser object
        // pointer). In C++, you don't have to pass this first arg, because the C++ compiler is smart enough to
        // produce an executable that always adds this first arg. In fact, the C++ compiler is smart enough to
        // know to fetch the function pointer from the VTable, so you don't even need to reference that. In other
        // words, the C++ equivalent code would be:
        //
        // browserObject->SetHostNames(L"My Host Name", 0);
        //
        // So, when you're trying to convert C++ code to C, always remember to add this first arg whenever you're
        // dealing with a VTable (ie, the field is usually named 'lpVtbl') in the standard objects, and also add
        // the reference to the VTable itself.
        //
        // Oh yeah, the L is because we need UNICODE strings. And BTW, the host and document names can be anything
        // you want.
        (*browserObject)->lpVtbl->SetHostNames(*browserObject, L"My Host Name", 0);

        GetClientRect(w->priv.hwnd, &rect);

        // Let browser object know that it is embedded in an OLE container.
        if (!OleSetContainedObject((struct IUnknown *)(*browserObject), TRUE) &&

          // Set the display area of our browser control the same as our window's size
          // and actually put the browser object into our window.
          !(*browserObject)->lpVtbl->DoVerb((*browserObject), OLEIVERB_SHOW, NULL, (IOleClientSite *)_iOleClientSiteEx, -1, w->priv.hwnd, &rect) &&

          // Ok, now things may seem to get even trickier, One of those function pointers in the browser's VTable is
          // to the QueryInterface() function. What does this function do? It lets us grab the base address of any
          // other object that may be embedded within the browser object. And this other object has its own VTable
          // containing pointers to more functions we can call for that object.
          //
          // We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
          // object, so we can call some of the functions in the former's table. For example, one IWebBrowser2 function
          // we intend to call below will be Navigate2(). So we call the browser object's QueryInterface to get our
          // pointer to the IWebBrowser2 object.
          !(*browserObject)->lpVtbl->QueryInterface((*browserObject), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2))
        {
          // Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
          // webBrowser2->lpVtbl.

          // Let's call several functions in the IWebBrowser2 object to position the browser display area
          // in our window. The functions we call are put_Left(), put_Top(), put_Width(), and put_Height().
          // Note that we reference the IWebBrowser2 object's VTable to get pointers to those functions. And
          // also note that the first arg we pass to each is the pointer to the IWebBrowser2 object.

          webBrowser2->lpVtbl->put_Left(webBrowser2, 0);
          webBrowser2->lpVtbl->put_Top(webBrowser2, 0);
          webBrowser2->lpVtbl->put_Width(webBrowser2, rect.right);
          webBrowser2->lpVtbl->put_Height(webBrowser2, rect.bottom);

          // We no longer need the IWebBrowser2 object (ie, we don't plan to call any more functions in it
          // right now, so we can release our hold on it). Note that we'll still maintain our hold on the
          // browser object until we're done with that object.
          webBrowser2->lpVtbl->Release(webBrowser2);

          // Success
          return S_OK;
        }
      }

      // Something went wrong setting up the browser!
      UnEmbedBrowserObject(w);
      return(-4);
    }

    pClassFactory->lpVtbl->Release(pClassFactory);

    // Can't create an instance of the browser!
    return(-3);
  }

  if (browserObject != NULL) {
    GlobalFree(browserObject);
  }

  // Can't get the web browser's IClassFactory!
  return(-2);
}

#define WEBVIEW_DATA_URL_PREFIX "data:text/html,"
/******************************* DisplayHTMLPage() ****************************
 * Displays a URL, or HTML file on disk.
 *
 * hwnd =    Handle to the window hosting the browser object.
 * webPageName =  Pointer to nul-terminated name of the URL/file.
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
*/
static int DisplayHTMLPage(struct webview *w, int pop) {
  IWebBrowser2 *webBrowser2;
  LPDISPATCH lpDispatch;
  IHTMLDocument2 *htmlDoc2;
  IOleObject *browserObject;
  SAFEARRAY *sfArray;
  VARIANT myURL;
  VARIANT *pVar;
  BSTR bstr;

  browserObject = *w->priv.browser;
  int isDataURL = 0;
  const char *webview_url = webview_check_url(w->url);
  if (!browserObject->lpVtbl->QueryInterface(browserObject, iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2)) {
    // ignore js error
    webBrowser2->lpVtbl->put_Silent(webBrowser2,VARIANT_TRUE);

    // Our URL (ie, web address, such as "http://www.microsoft.com" or an HTM filename on disk
    // such as "c:\myfile.htm") must be passed to the IWebBrowser2's Navigate2() function as a BSTR.
    // A BSTR is like a pascal version of a double-byte character string. In other words, the
    // first unsigned short is a count of how many characters are in the string, and then this
    // is followed by those characters, each expressed as an unsigned short (rather than a
    // char). The string is not nul-terminated. The OS function SysAllocString can allocate and
    // copy a UNICODE C string to a BSTR. Of course, we'll need to free that BSTR after we're done
    // with it. If we're not using UNICODE, we first have to convert to a UNICODE string.
    //
    // What's more, our BSTR needs to be stuffed into a VARIENT struct, and that VARIENT struct is
    // then passed to Navigate2(). Why? The VARIENT struct makes it possible to define generic
    // 'datatypes' that can be used with all languages. Not all languages support things like
    // nul-terminated C strings. So, by using a VARIENT, whose first field tells what sort of
    // data (ie, string, float, etc) is in the VARIENT, COM interfaces can be used by just about
    // any language.

    LPCSTR webPageName;
    isDataURL = (strncmp(webview_url, WEBVIEW_DATA_URL_PREFIX, strlen(WEBVIEW_DATA_URL_PREFIX)) == 0);
    if (isDataURL) {
      webPageName = "about:blank";
    } else {
      webPageName = (LPCSTR)webview_url;
    }
    myURL = StrToVARIANT(webPageName);
    if (!myURL.bstrVal) {
      webBrowser2->lpVtbl->Release(webBrowser2);
      return (-6);
    }

    // Call the Navigate2() function to actually display the page.
    //printf("pop: %d\n", pop);
    //if (0 == pop ) {
      webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);
    //} else {
    //  VARIANT vFlags;
    //  V_VT(&vFlags) = VT_I4;
    //  V_I4(&vFlags) = navOpenInNewWindow;
    //  webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, &vFlags, 0, 0, 0);
    //}
    
    // Free any resources (including the BSTR we allocated above).
    VariantClear(&myURL);
    if (!isDataURL) {
      return 0;
    }

    char *url = (char *)calloc(1, strlen(webview_url) + 1);
    char *q = url;
    for (const char *p = webview_url + strlen(WEBVIEW_DATA_URL_PREFIX); *q = *p;
         p++, q++) {
      if (*q == '%' && *(p + 1) && *(p + 2)) {
        sscanf(p + 1, "%02x", q);
        p = p + 2;
      }
    }

    if (webBrowser2->lpVtbl->get_Document(webBrowser2, &lpDispatch) == S_OK) {
      if (lpDispatch->lpVtbl->QueryInterface(lpDispatch, iid_unref(&IID_IHTMLDocument2), (void **)&htmlDoc2) == S_OK) {
        if ((sfArray = SafeArrayCreate(VT_VARIANT, 1, (SAFEARRAYBOUND *)&ArrayBound))) {
          if (!SafeArrayAccessData(sfArray, (void **)&pVar)) {
            pVar->vt = VT_BSTR;
#ifndef UNICODE
            {
              wchar_t *buffer = webview_to_utf16(url);
              if (buffer == NULL) {
                goto release;
              }
              bstr = SysAllocString(buffer);
              GlobalFree(buffer);
            }
#else
            bstr = SysAllocString(string);
#endif
            if ((pVar->bstrVal = bstr)) {
              htmlDoc2->lpVtbl->write(htmlDoc2, sfArray);
              htmlDoc2->lpVtbl->close(htmlDoc2);
            }
          }
          SafeArrayDestroy(sfArray);
        }
      release:
        free(url);
        htmlDoc2->lpVtbl->Release(htmlDoc2);
      }
      lpDispatch->lpVtbl->Release(lpDispatch);
    }
    webBrowser2->lpVtbl->Release(webBrowser2);
    return (0);
  }
  return (-5);
}

/******************************* DisplayHTMLStr() ****************************
 * Takes a string containing some HTML BODY, and displays it in the specified
 * window. For example, perhaps you want to display the HTML text of...
 *
 * <P>This is a picture.<P><IMG src="mypic.jpg">
 *
 * hwnd =    Handle to the window hosting the browser object.
 * string =    Pointer to nul-terminated string containing the HTML BODY.
 *        (NOTE: No <BODY></BODY> tags are required in the string).
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
 */
static int DisplayHTMLStr(struct webview *w, LPCTSTR string) {  
  IWebBrowser2  *webBrowser2;
  LPDISPATCH    lpDispatch;
  IHTMLDocument2  *htmlDoc2;
  IOleObject    *browserObject;
  SAFEARRAY    *sfArray;
  VARIANT      myURL;
  VARIANT      *pVar;
  BSTR      bstr;

  // Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
  // we initially attached the browser object to this window.
  browserObject = *w->priv.browser;

  // Assume an error.
  bstr = 0;

  // We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
  // object, so we can call some of the functions in the former's table.
  if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2)) {
    // Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
    // webBrowser2->lpVtbl.

    // Call the IWebBrowser2 object's get_Document so we can get its DISPATCH object. I don't know why you
    // don't get the DISPATCH object via the browser object's QueryInterface(), but you don't.
    lpDispatch = 0;
    webBrowser2->lpVtbl->get_Document(webBrowser2, &lpDispatch);
    if (!lpDispatch) {
      // Before we can get_Document(), we actually need to have some HTML page loaded in the browser. So,
      // let's load an empty HTML page. Then, once we have that empty page, we can get_Document() and
      // write() to stuff our HTML string into it.
      VariantInit(&myURL);
      myURL.vt = VT_BSTR;
      myURL.bstrVal = SysAllocString(&Blank[0]);

      // Call the Navigate2() function to actually display the page.
      webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);

      // Free any resources (including the BSTR).
      VariantClear(&myURL);
    }

    // Call the IWebBrowser2 object's get_Document so we can get its DISPATCH object. I don't know why you
    // don't get the DISPATCH object via the browser object's QueryInterface(), but you don't.
    if (!webBrowser2->lpVtbl->get_Document(webBrowser2, &lpDispatch) && lpDispatch) {
      // We want to get a pointer to the IHTMLDocument2 object embedded within the DISPATCH
      // object, so we can call some of the functions in the former's table.
      if (!lpDispatch->lpVtbl->QueryInterface(lpDispatch, &IID_IHTMLDocument2, (void**)&htmlDoc2)) {
        // Ok, now the pointer to our IHTMLDocument2 object is in 'htmlDoc2', and so its VTable is
        // htmlDoc2->lpVtbl.

        // Our HTML must be in the form of a BSTR. And it must be passed to write() in an
        // array of "VARIENT" structs. So let's create all that.
        if ((sfArray = SafeArrayCreate(VT_VARIANT, 1, (SAFEARRAYBOUND *)&ArrayBound))) {
          if (!SafeArrayAccessData(sfArray, (void**)&pVar)) {
            pVar->vt = VT_BSTR;
#ifndef UNICODE
            {
            wchar_t    *buffer;
            DWORD    size;

            size = MultiByteToWideChar(CP_ACP, 0, string, -1, 0, 0);
            if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto bad;
            MultiByteToWideChar(CP_ACP, 0, string, -1, buffer, size);
            bstr = SysAllocString(buffer);
            GlobalFree(buffer);
            }
#else
            bstr = SysAllocString(string);
#endif
            // Store our BSTR pointer in the VARIENT.
            if ((pVar->bstrVal = bstr)) {
              // Pass the VARIENT with its BSTR to write() in order to shove our desired HTML string
              // into the body of that empty page we created above.
              htmlDoc2->lpVtbl->write(htmlDoc2, sfArray);

              // Close the document. If we don't do this, subsequent calls to DisplayHTMLStr
              // would append to the current contents of the page
              htmlDoc2->lpVtbl->close(htmlDoc2);

              // Normally, we'd need to free our BSTR, but SafeArrayDestroy() does it for us
              //SysFreeString(bstr);
            }
          }

          // Free the array. This also frees the VARIENT that SafeArrayAccessData created for us,
          // and even frees the BSTR we allocated with SysAllocString
          SafeArrayDestroy(sfArray);
        }

        // Release the IHTMLDocument2 object.
bad:      htmlDoc2->lpVtbl->Release(htmlDoc2);
      }

      // Release the DISPATCH object.
      lpDispatch->lpVtbl->Release(lpDispatch);
    }

    // Release the IWebBrowser2 object.
    webBrowser2->lpVtbl->Release(webBrowser2);
  }

  // No error?
  if (bstr) return(0);

  // An error
  return(-1);
}

/****************************** WindowProc() ***************************
 * Our message handler for our Main window.
 */
static LRESULT CALLBACK windowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
  struct webview *w = (struct webview *)GetWindowLongPtr(hwnd, GWLP_USERDATA);
  switch (uMsg) {
    case WM_CREATE: {
      // Embed the browser object into our host window. We need do this only
      // once. Note that the browser object will start calling some of our
      // IOleInPlaceFrame and IOleClientSite functions as soon as we start
      // calling browser object functions in EmbedBrowserObject().
      w = (struct webview *)((CREATESTRUCT *)lParam)->lpCreateParams;
      w->priv.hwnd = hwnd;

      // Embed the browser object into our host window.
      if (EmbedBrowserObject(w)) {
        return(-1);
      }

      AdviseEvent(w, (IUnknown *)&ExternalDispatchTable);
      
      // Another window created with an embedded browser object
      ++WindowCount;

      // Success
      return(0);
    }

    case WM_DESTROY: {
      // Detach the browser object from this window, and free resources.
       // Post the WM_QUIT message to quit the message loop in WinMain()
      UnEmbedBrowserObject(w);

      // One less window
      --WindowCount;

      // If all the windows are now closed, quit this app
      if (!WindowCount) PostQuitMessage(0);

      return TRUE;
    }

    case WM_SIZE: {
      // Resize the browser object to fit the window

      IWebBrowser2 *webBrowser2;
      IOleObject *browser = *w->priv.browser;

      // We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
      // object, so we can call some of the functions in the former's table.
      if (browser->lpVtbl->QueryInterface(browser, iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) == S_OK) {
        RECT rect;
        GetClientRect(hwnd, &rect);

        // Call are put_Width() and put_Height() to set the new width/height.
        webBrowser2->lpVtbl->put_Width(webBrowser2, rect.right);
        webBrowser2->lpVtbl->put_Height(webBrowser2, rect.bottom);
      }

      return TRUE;
    }
    case WM_WEBVIEW_DISPATCH: {
      webview_dispatch_fn f = (webview_dispatch_fn)wParam;
      void *arg = (void *)lParam;
      (*f)(w, arg);
      return TRUE;
    }
    //case WM_PARENTNOTIFY:
    case WM_MOUSEACTIVATE: {
      void *arg = (void *)lParam;
      _webviewMouseNotify(arg);
    }

    case WM_ERASEBKGND: {
      HBRUSH  hBrush;
      HGDIOBJ  orig;
      RECT  rec;

      // Erase the background behind the buttons. Note: The browser object takes care
      // of completely redrawing its area, so we don't need to erase there.
      hBrush = GetStockObject(WHITE_BRUSH);
      GetClientRect(hwnd, &rec);
      orig = SelectObject((HDC)wParam, hBrush);
      Rectangle((HDC)wParam, 0, 0, rec.right, 40);
      SelectObject((HDC)wParam, orig);
      return(TRUE);
    }

    case WM_GETMINMAXINFO:
    {
      LPMINMAXINFO    lpmmi;
      NONCLIENTMETRICS  nonClientMetrics;

      lpmmi = (LPMINMAXINFO)lParam;
      nonClientMetrics.cbSize = sizeof(NONCLIENTMETRICS);
      SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, &nonClientMetrics, 0);

      // Force height to be at least 80 + size of titlebar (if we had a menubar,
      // we would also add in iMenuHeight)

      // DISABLED!!! BUG ON WINDOWS XP
      //lpmmi->ptMinTrackSize.y = nonClientMetrics.iCaptionHeight + (nonClientMetrics.iBorderWidth << 1) + 80;

      return(0);
    }

    case WM_APP:
    {
      //LPCTSTR    str;

      // Our IDocHostUIHandler's TranslateUrl() must have posted us this message
      // to tell us to display some HTML string. The string number is wParam.

      // Find the desired string
      /*str = &HtmlStrings[0];
      while (--wParam)
      {
        if (!(*str)) return(0);
        while (*(str)++);
      }
      */
      // Display it.
      //DisplayHTMLStr(hwnd, str);

      return(0);
    }

    case WM_CLOSE:
    {
      // Close this window. This will also cause the child window hosting the browser
      // control to receive its WM_DESTROY
      DestroyWindow(hwnd);

      return(1);
    }

  }
  return DefWindowProc(hwnd, uMsg, wParam, lParam);
}


#define WEBVIEW_KEY_FEATURE_BROWSER_EMULATION                                  \
  "Software\\Microsoft\\Internet "                                             \
  "Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION"

static int webview_fix_ie_compat_mode() {
  HKEY hKey;
  DWORD ie_version = 11000;
  TCHAR appname[MAX_PATH + 1];
  TCHAR *p;
  if (GetModuleFileName(NULL, appname, MAX_PATH + 1) == 0) {
    return -1;
  }
  for (p = &appname[strlen(appname) - 1]; p != appname && *p != '\\'; p--) {
  }
  p++;
  if (RegCreateKey(HKEY_CURRENT_USER, WEBVIEW_KEY_FEATURE_BROWSER_EMULATION, &hKey) != ERROR_SUCCESS) {
    return -1;
  }
  if (RegSetValueEx(hKey, p, 0, REG_DWORD, (BYTE *)&ie_version, sizeof(ie_version)) != ERROR_SUCCESS) {
    RegCloseKey(hKey);
    return -1;
  }
  RegCloseKey(hKey);
  return 0;
}


#define WEBPAGE_GOBACK    0
#define WEBPAGE_GOFORWARD  1
#define WEBPAGE_GOHOME    2
#define WEBPAGE_SEARCH    3
#define WEBPAGE_REFRESH    4
#define WEBPAGE_STOP    5

/******************************* DoPageAction() **************************
 * Implements the functionality of a "Back". "Forward", "Home", "Search",
 * "Refresh", or "Stop" button.
 *
 * hwnd =    Handle to the window hosting the browser object.
 * action =    One of the following:
 *        0 = Move back to the previously viewed web page.
 *        1 = Move forward to the previously viewed web page.
 *        2 = Move to the home page.
 *        3 = Search.
 *        4 = Refresh the page.
 *        5 = Stop the currently loading page.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
 */

static void DoPageAction(HWND hwnd, DWORD action) {  
  IWebBrowser2  *webBrowser2;
  IOleObject    *browserObject;

  // Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
  // we initially attached the browser object to this window.
  browserObject = *((IOleObject **) GetWindowLongPtr(hwnd, GWLP_USERDATA));

  // We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
  // object, so we can call some of the functions in the former's table.
  if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2)) {
    // Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
    // webBrowser2->lpVtbl.

    // ignore js error
    webBrowser2->lpVtbl->put_Silent(webBrowser2,VARIANT_TRUE);
    
    // Call the desired function
    switch (action) {
      case WEBPAGE_GOBACK:
      {
        // Call the IWebBrowser2 object's GoBack function.
        webBrowser2->lpVtbl->GoBack(webBrowser2);
        break;
      }

      case WEBPAGE_GOFORWARD:
      {
        // Call the IWebBrowser2 object's GoForward function.
        webBrowser2->lpVtbl->GoForward(webBrowser2);
        break;
      }

      case WEBPAGE_GOHOME:
      {
        // Call the IWebBrowser2 object's GoHome function.
        webBrowser2->lpVtbl->GoHome(webBrowser2);
        break;
      }

      case WEBPAGE_SEARCH:
      {
        // Call the IWebBrowser2 object's GoSearch function.
        webBrowser2->lpVtbl->GoSearch(webBrowser2);
        break;
      }

      case WEBPAGE_REFRESH:
      {
        // Call the IWebBrowser2 object's Refresh function.
        webBrowser2->lpVtbl->Refresh(webBrowser2);
      }

      case WEBPAGE_STOP:
      {
        // Call the IWebBrowser2 object's Stop function.
        webBrowser2->lpVtbl->Stop(webBrowser2);
      }
    }

    // Release the IWebBrowser2 object.
    webBrowser2->lpVtbl->Release(webBrowser2);
  }
}

WEBVIEW_API int webview_init(struct webview *w) {
  WNDCLASSEX wc;
  HINSTANCE hInstance;
  DWORD style;
  RECT clientRect;
  RECT rect;

  if (webview_fix_ie_compat_mode() < 0) {
    return -1;
  }

  hInstance = GetModuleHandle(NULL);
  if (hInstance == NULL) {
    return -1;
  }
  if (OleInitialize(NULL) != S_OK) {
    return -1;
  }

  // Register the class of our Main window. 'windowProc' is our message handler
  // and 'ClassName' is the class name. You can choose any class name you want.
  ZeroMemory(&wc, sizeof(WNDCLASSEX));
  wc.cbSize = sizeof(WNDCLASSEX);
  wc.hInstance = hInstance;
  wc.lpfnWndProc = windowProc;
  wc.lpszClassName = classname;
  RegisterClassEx(&wc);

  style = WS_OVERLAPPEDWINDOW;
  if (!w->resizable) {
    style = WS_OVERLAPPED | WS_CAPTION | WS_MINIMIZEBOX | WS_SYSMENU;
  }

  rect.left = 0;
  rect.top = 0;
  rect.right = w->width;
  rect.bottom = w->height;
  AdjustWindowRect(&rect, WS_OVERLAPPEDWINDOW, 0);

  GetClientRect(GetDesktopWindow(), &clientRect);
  int left = (clientRect.right / 2) - ((rect.right - rect.left) / 2);
  int top = (clientRect.bottom / 2) - ((rect.bottom - rect.top) / 2);
  rect.right = rect.right - rect.left + left;
  rect.left = left;
  rect.bottom = rect.bottom - rect.top + top;
  rect.top = top;

  w->priv.hwnd = CreateWindowEx(0, classname, w->title, style, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, HWND_DESKTOP, NULL, hInstance, (void *)w);
  if (w->priv.hwnd == 0) {
    OleUninitialize();
    return -1;
  }

  SetWindowLongPtr(w->priv.hwnd, GWLP_USERDATA, (LONG_PTR)w);

  DisplayHTMLPage(w, 0);

  SetWindowText(w->priv.hwnd, w->title);

  // Show the window.
  ShowWindow(w->priv.hwnd, SW_SHOWDEFAULT);
  UpdateWindow(w->priv.hwnd);
  SetFocus(w->priv.hwnd);

  return 0;
}

WEBVIEW_API int webview_loop(struct webview *w, int blocking) {
  MSG msg;
  if (blocking) {
    GetMessage(&msg, 0, 0, 0);
  } else {
    PeekMessage(&msg, 0, 0, 0, PM_REMOVE);
  }
  switch (msg.message) {
  case WM_QUIT:
    return -1;
  case WM_COMMAND:
  case WM_KEYDOWN:
  case WM_KEYUP: {
    HRESULT r = S_OK;
    IWebBrowser2 *webBrowser2;
    IOleObject *browser = *w->priv.browser;
    if (browser->lpVtbl->QueryInterface(browser, iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) == S_OK) {
      IOleInPlaceActiveObject *pIOIPAO;
      if (browser->lpVtbl->QueryInterface(browser, iid_unref(&IID_IOleInPlaceActiveObject), (void **)&pIOIPAO) == S_OK) {
        r = pIOIPAO->lpVtbl->TranslateAccelerator(pIOIPAO, &msg);
        pIOIPAO->lpVtbl->Release(pIOIPAO);
      }
      webBrowser2->lpVtbl->Release(webBrowser2);
    }
    if (r != S_FALSE) {
      break;
    }
  }
  default:
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return 0;
}

WEBVIEW_API int webview_eval(struct webview *w, const char *js) {
  IWebBrowser2 *webBrowser2;
  IHTMLDocument2 *htmlDoc2;
  IDispatch *docDispatch;
  IDispatch *scriptDispatch;
  if ((*w->priv.browser) ->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  if (webBrowser2->lpVtbl->get_Document(webBrowser2, &docDispatch) != S_OK) {
    return -1;
  }
  if (docDispatch->lpVtbl->QueryInterface(docDispatch, iid_unref(&IID_IHTMLDocument2), (void **)&htmlDoc2) != S_OK) {
    return -1;
  }
  if (htmlDoc2->lpVtbl->get_Script(htmlDoc2, &scriptDispatch) != S_OK) {
    return -1;
  }
  DISPID dispid;
  BSTR evalStr = SysAllocString(L"eval");
  if (scriptDispatch->lpVtbl->GetIDsOfNames(scriptDispatch, iid_unref(&IID_NULL), &evalStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid) != S_OK) {
    SysFreeString(evalStr);
    return -1;
  }
  SysFreeString(evalStr);

  DISPPARAMS params;
  VARIANT arg;
  VARIANT result;
  EXCEPINFO excepInfo;
  UINT nArgErr = (UINT)-1;
  params.cArgs = 1;
  params.cNamedArgs = 0;
  params.rgvarg = &arg;
  arg.vt = VT_BSTR;
  static const char *prologue = "(function(){";
  static const char *epilogue = ";})();";
  int n = strlen(prologue) + strlen(epilogue) + strlen(js) + 1;
  char *eval = (char *)malloc(n);
  snprintf(eval, n, "%s%s%s", prologue, js, epilogue);
  wchar_t *buf = webview_to_utf16(eval);
  if (buf == NULL) {
    return -1;
  }
  arg.bstrVal = SysAllocString(buf);
  if (scriptDispatch->lpVtbl->Invoke(scriptDispatch, dispid, iid_unref(&IID_NULL), 0, DISPATCH_METHOD, &params, &result, &excepInfo, &nArgErr) != S_OK) {
    return -1;
  }
  SysFreeString(arg.bstrVal);
  free(eval);
  scriptDispatch->lpVtbl->Release(scriptDispatch);
  htmlDoc2->lpVtbl->Release(htmlDoc2);
  docDispatch->lpVtbl->Release(docDispatch);
  return 0;
}

WEBVIEW_API void webview_dispatch(struct webview *w, webview_dispatch_fn fn, void *arg) {
  PostMessageW(w->priv.hwnd, WM_WEBVIEW_DISPATCH, (WPARAM)fn, (LPARAM)arg);
}

WEBVIEW_API int webview_version(struct webview *w, const char* key, char* ptr) {
  long ret;
  HKEY hKey;
  DWORD type=REG_SZ;
  DWORD size=MAX_PATH;
  
  ret = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Internet Explorer", 0, KEY_READ, &hKey);
  if(ret!=ERROR_SUCCESS)   {
    RegCloseKey(hKey);
    return 1;
  }

  ret=RegQueryValueEx(hKey, key, NULL, &type, ptr, &size);
  if(ret!=ERROR_SUCCESS)   {
    RegCloseKey(hKey);
    return 1;
  }

  RegCloseKey(hKey);

  return 0;
}

WEBVIEW_API void webview_set_title(struct webview *w, const char *title) {
  SetWindowText(w->priv.hwnd, title);
}

WEBVIEW_API int webview_refresh(struct webview *w) {
  IWebBrowser2 *webBrowser2;
  if ((*w->priv.browser) ->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  return webBrowser2->lpVtbl->Refresh(webBrowser2);
}

WEBVIEW_API int webview_set_url(struct webview *w, const char *url, int pop) {
  w->url = url;
  return DisplayHTMLPage(w, pop);
}

WEBVIEW_API int webview_get_url(struct webview *w, void *ptr) {
  IWebBrowser2 *webBrowser2;
  if ((*w->priv.browser) ->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  BSTR *ptrBSTR;
  ptrBSTR = (BSTR*)ptr;
  if (webBrowser2->lpVtbl->get_LocationURL(webBrowser2, ptrBSTR) != S_OK) {
    return -1;
  }

  return 0;
}

WEBVIEW_API int webview_get_title(struct webview *w, void *ptr) {
  IWebBrowser2 *webBrowser2;
  IHTMLDocument2 *htmlDoc2;
  IDispatch *docDispatch;
  if ((*w->priv.browser)->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  if (webBrowser2->lpVtbl->get_Document(webBrowser2, &docDispatch) != S_OK) {
    return -1;
  }
  if (docDispatch->lpVtbl->QueryInterface(docDispatch, iid_unref(&IID_IHTMLDocument2), (void **)&htmlDoc2) == S_OK) {
    BSTR *ptrBSTR;
    ptrBSTR = (BSTR*)ptr;
    htmlDoc2->lpVtbl->get_title(htmlDoc2, ptrBSTR);
    htmlDoc2->lpVtbl->Release(htmlDoc2);
  }

  docDispatch->lpVtbl->Release(docDispatch);
  
  return 0;
}

WEBVIEW_API int webview_get_html(struct webview *w, void *ptr) {
  IWebBrowser2 *webBrowser2;
  IHTMLDocument2 *htmlDoc2;
  IDispatch *docDispatch;
  if ((*w->priv.browser)->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  if (webBrowser2->lpVtbl->get_Document(webBrowser2, &docDispatch) != S_OK) {
    return -1;
  }
  if (docDispatch->lpVtbl->QueryInterface(docDispatch, iid_unref(&IID_IHTMLDocument2), (void **)&htmlDoc2) == S_OK) {
    BSTR *ptrBSTR;
    ptrBSTR = (BSTR*)ptr;

    IHTMLElement * pBody;
    IHTMLElement * pHtml;
    htmlDoc2->lpVtbl->get_body(htmlDoc2, &pBody);
    pBody->lpVtbl->get_parentElement(pBody, &pHtml);
    pHtml->lpVtbl->get_outerHTML(pHtml, ptrBSTR);
    pBody->lpVtbl->Release(pBody);
    pHtml->lpVtbl->Release(pHtml);
    htmlDoc2->lpVtbl->Release(htmlDoc2);
  }

  docDispatch->lpVtbl->Release(docDispatch);
  
  return 0;
}

WEBVIEW_API int webview_get_session_cookie(struct webview *w, const char* url, const char* key, char* ptr) {
  DWORD dwSize = 1024;

  LPSTR lpData;
  lpData = (LPSTR) ptr;
  BOOL ret = InternetGetCookieEx(url, key, lpData, &dwSize, INTERNET_COOKIE_HTTPONLY | INTERNET_COOKIE_THIRD_PARTY , NULL);
  if (ret) {
    return 0;
  }
  
  return 1;
}

WEBVIEW_API int webview_get_cookies(struct webview *w, void *ptr) {
  IWebBrowser2 *webBrowser2;
  IHTMLDocument2 *htmlDoc2;
  IDispatch *docDispatch;
  if ((*w->priv.browser)->lpVtbl->QueryInterface((*w->priv.browser), iid_unref(&IID_IWebBrowser2), (void **)&webBrowser2) != S_OK) {
    return -1;
  }

  if (webBrowser2->lpVtbl->get_Document(webBrowser2, &docDispatch) != S_OK) {
    return -1;
  }
  if (docDispatch->lpVtbl->QueryInterface(docDispatch, iid_unref(&IID_IHTMLDocument2), (void **)&htmlDoc2) == S_OK) {
    BSTR *ptrBSTR;
    ptrBSTR = (BSTR*)ptr;
    htmlDoc2->lpVtbl->get_cookie(htmlDoc2, ptrBSTR);
    htmlDoc2->lpVtbl->Release(htmlDoc2);
  }

  docDispatch->lpVtbl->Release(docDispatch);
  
  return 0;
}

WEBVIEW_API void webview_set_fullscreen(struct webview *w, int fullscreen) {
  if (w->priv.is_fullscreen == !!fullscreen) {
    return;
  }
  if (w->priv.is_fullscreen == 0) {
    w->priv.saved_style = GetWindowLong(w->priv.hwnd, GWL_STYLE);
    w->priv.saved_ex_style = GetWindowLong(w->priv.hwnd, GWL_EXSTYLE);
    GetWindowRect(w->priv.hwnd, &w->priv.saved_rect);
  }
  w->priv.is_fullscreen = !!fullscreen;
  if (fullscreen) {
    MONITORINFO monitor_info;
    SetWindowLong(w->priv.hwnd, GWL_STYLE, w->priv.saved_style & ~(WS_CAPTION | WS_THICKFRAME));
    SetWindowLong(w->priv.hwnd, GWL_EXSTYLE, w->priv.saved_ex_style & ~(WS_EX_DLGMODALFRAME | WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE | WS_EX_STATICEDGE));
    monitor_info.cbSize = sizeof(monitor_info);
    GetMonitorInfo(MonitorFromWindow(w->priv.hwnd, MONITOR_DEFAULTTONEAREST), &monitor_info);
    RECT r;
    r.left = monitor_info.rcMonitor.left;
    r.top = monitor_info.rcMonitor.top;
    r.right = monitor_info.rcMonitor.right;
    r.bottom = monitor_info.rcMonitor.bottom;
    SetWindowPos(w->priv.hwnd, NULL, r.left, r.top, r.right - r.left, r.bottom - r.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
  } else {
    SetWindowLong(w->priv.hwnd, GWL_STYLE, w->priv.saved_style);
    SetWindowLong(w->priv.hwnd, GWL_EXSTYLE, w->priv.saved_ex_style);
    SetWindowPos(w->priv.hwnd, NULL, w->priv.saved_rect.left, w->priv.saved_rect.top, w->priv.saved_rect.right - w->priv.saved_rect.left, w->priv.saved_rect.bottom - w->priv.saved_rect.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
  }
}

WEBVIEW_API void webview_set_color(struct webview *w, uint8_t r, uint8_t g, uint8_t b, uint8_t a) {
  HBRUSH brush = CreateSolidBrush(RGB(r, g, b));
  SetClassLongPtr(w->priv.hwnd, GCLP_HBRBACKGROUND, (LONG_PTR)brush);
}

/* These are missing parts from MinGW */
#ifndef __IFileDialog_INTERFACE_DEFINED__
#define __IFileDialog_INTERFACE_DEFINED__
enum _FILEOPENDIALOGOPTIONS {
  FOS_OVERWRITEPROMPT = 0x2,
  FOS_STRICTFILETYPES = 0x4,
  FOS_NOCHANGEDIR = 0x8,
  FOS_PICKFOLDERS = 0x20,
  FOS_FORCEFILESYSTEM = 0x40,
  FOS_ALLNONSTORAGEITEMS = 0x80,
  FOS_NOVALIDATE = 0x100,
  FOS_ALLOWMULTISELECT = 0x200,
  FOS_PATHMUSTEXIST = 0x800,
  FOS_FILEMUSTEXIST = 0x1000,
  FOS_CREATEPROMPT = 0x2000,
  FOS_SHAREAWARE = 0x4000,
  FOS_NOREADONLYRETURN = 0x8000,
  FOS_NOTESTFILECREATE = 0x10000,
  FOS_HIDEMRUPLACES = 0x20000,
  FOS_HIDEPINNEDPLACES = 0x40000,
  FOS_NODEREFERENCELINKS = 0x100000,
  FOS_DONTADDTORECENT = 0x2000000,
  FOS_FORCESHOWHIDDEN = 0x10000000,
  FOS_DEFAULTNOMINIMODE = 0x20000000,
  FOS_FORCEPREVIEWPANEON = 0x40000000
};
typedef DWORD FILEOPENDIALOGOPTIONS;
typedef enum FDAP { FDAP_BOTTOM = 0, FDAP_TOP = 1 } FDAP;
DEFINE_GUID(IID_IFileDialog, 0x42f85136, 0xdb7e, 0x439c, 0x85, 0xf1, 0xe4, 0x07, 0x5d, 0x13, 0x5f, 0xc8);
typedef struct IFileDialogVtbl {
  BEGIN_INTERFACE
  HRESULT(STDMETHODCALLTYPE *QueryInterface)
  (IFileDialog *This, REFIID riid, void **ppvObject);
  ULONG(STDMETHODCALLTYPE *AddRef)(IFileDialog *This);
  ULONG(STDMETHODCALLTYPE *Release)(IFileDialog *This);
  HRESULT(STDMETHODCALLTYPE *Show)(IFileDialog *This, HWND hwndOwner);
  HRESULT(STDMETHODCALLTYPE *SetFileTypes)
  (IFileDialog *This, UINT cFileTypes, const COMDLG_FILTERSPEC *rgFilterSpec);
  HRESULT(STDMETHODCALLTYPE *SetFileTypeIndex)
  (IFileDialog *This, UINT iFileType);
  HRESULT(STDMETHODCALLTYPE *GetFileTypeIndex)
  (IFileDialog *This, UINT *piFileType);
  HRESULT(STDMETHODCALLTYPE *Advise)
  (IFileDialog *This, IFileDialogEvents *pfde, DWORD *pdwCookie);
  HRESULT(STDMETHODCALLTYPE *Unadvise)(IFileDialog *This, DWORD dwCookie);
  HRESULT(STDMETHODCALLTYPE *SetOptions)
  (IFileDialog *This, FILEOPENDIALOGOPTIONS fos);
  HRESULT(STDMETHODCALLTYPE *GetOptions)
  (IFileDialog *This, FILEOPENDIALOGOPTIONS *pfos);
  HRESULT(STDMETHODCALLTYPE *SetDefaultFolder)
  (IFileDialog *This, IShellItem *psi);
  HRESULT(STDMETHODCALLTYPE *SetFolder)(IFileDialog *This, IShellItem *psi);
  HRESULT(STDMETHODCALLTYPE *GetFolder)(IFileDialog *This, IShellItem **ppsi);
  HRESULT(STDMETHODCALLTYPE *GetCurrentSelection)
  (IFileDialog *This, IShellItem **ppsi);
  HRESULT(STDMETHODCALLTYPE *SetFileName)(IFileDialog *This, LPCWSTR pszName);
  HRESULT(STDMETHODCALLTYPE *GetFileName)(IFileDialog *This, LPWSTR *pszName);
  HRESULT(STDMETHODCALLTYPE *SetTitle)(IFileDialog *This, LPCWSTR pszTitle);
  HRESULT(STDMETHODCALLTYPE *SetOkButtonLabel)
  (IFileDialog *This, LPCWSTR pszText);
  HRESULT(STDMETHODCALLTYPE *SetFileNameLabel)
  (IFileDialog *This, LPCWSTR pszLabel);
  HRESULT(STDMETHODCALLTYPE *GetResult)(IFileDialog *This, IShellItem **ppsi);
  HRESULT(STDMETHODCALLTYPE *AddPlace)
  (IFileDialog *This, IShellItem *psi, FDAP fdap);
  HRESULT(STDMETHODCALLTYPE *SetDefaultExtension)
  (IFileDialog *This, LPCWSTR pszDefaultExtension);
  HRESULT(STDMETHODCALLTYPE *Close)(IFileDialog *This, HRESULT hr);
  HRESULT(STDMETHODCALLTYPE *SetClientGuid)(IFileDialog *This, REFGUID guid);
  HRESULT(STDMETHODCALLTYPE *ClearClientData)(IFileDialog *This);
  HRESULT(STDMETHODCALLTYPE *SetFilter)
  (IFileDialog *This, IShellItemFilter *pFilter);
  END_INTERFACE
} IFileDialogVtbl;
interface IFileDialog {
  CONST_VTBL IFileDialogVtbl *lpVtbl;
};
DEFINE_GUID(IID_IFileOpenDialog, 0xd57c7288, 0xd4ad, 0x4768, 0xbe, 0x02, 0x9d, 0x96, 0x95, 0x32, 0xd9, 0x60);
DEFINE_GUID(IID_IFileSaveDialog, 0x84bccd23, 0x5fde, 0x4cdb, 0xae, 0xa4, 0xaf, 0x64, 0xb8, 0x3d, 0x78, 0xab);
#endif

WEBVIEW_API void webview_dialog(struct webview *w, enum webview_dialog_type dlgtype, int flags, const char *title, const char *arg, char *result, size_t resultsz) {
  if (dlgtype == WEBVIEW_DIALOG_TYPE_OPEN || dlgtype == WEBVIEW_DIALOG_TYPE_SAVE) {
    IFileDialog *dlg = NULL;
    IShellItem *res = NULL;
    WCHAR *ws = NULL;
    char *s = NULL;
    FILEOPENDIALOGOPTIONS opts, add_opts;
    if (dlgtype == WEBVIEW_DIALOG_TYPE_OPEN) {
      if (CoCreateInstance(iid_unref(&CLSID_FileOpenDialog), NULL, CLSCTX_INPROC_SERVER, iid_unref(&IID_IFileOpenDialog), (void **)&dlg) != S_OK) {
        goto error_dlg;
      }
      if (flags & WEBVIEW_DIALOG_FLAG_DIRECTORY) {
        add_opts |= FOS_PICKFOLDERS;
      }
      add_opts |= FOS_NOCHANGEDIR | FOS_ALLNONSTORAGEITEMS | FOS_NOVALIDATE |
                  FOS_PATHMUSTEXIST | FOS_FILEMUSTEXIST | FOS_SHAREAWARE |
                  FOS_NOTESTFILECREATE | FOS_NODEREFERENCELINKS |
                  FOS_FORCESHOWHIDDEN | FOS_DEFAULTNOMINIMODE;
    } else {
      if (CoCreateInstance( iid_unref(&CLSID_FileSaveDialog), NULL, CLSCTX_INPROC_SERVER, iid_unref(&IID_IFileSaveDialog), (void **)&dlg) != S_OK) {
        goto error_dlg;
      }
      add_opts |= FOS_OVERWRITEPROMPT | FOS_NOCHANGEDIR |
                  FOS_ALLNONSTORAGEITEMS | FOS_NOVALIDATE | FOS_SHAREAWARE |
                  FOS_NOTESTFILECREATE | FOS_NODEREFERENCELINKS |
                  FOS_FORCESHOWHIDDEN | FOS_DEFAULTNOMINIMODE;
    }
    if (dlg->lpVtbl->GetOptions(dlg, &opts) != S_OK) {
      goto error_dlg;
    }
    opts &= ~FOS_NOREADONLYRETURN;
    opts |= add_opts;
    if (dlg->lpVtbl->SetOptions(dlg, opts) != S_OK) {
      goto error_dlg;
    }
    if (dlg->lpVtbl->Show(dlg, w->priv.hwnd) != S_OK) {
      goto error_dlg;
    }
    if (dlg->lpVtbl->GetResult(dlg, &res) != S_OK) {
      goto error_dlg;
    }
    if (res->lpVtbl->GetDisplayName(res, SIGDN_FILESYSPATH, &ws) != S_OK) {
      goto error_result;
    }
    s = webview_from_utf16(ws);
    strncpy(result, s, resultsz);
    result[resultsz - 1] = '\0';
    CoTaskMemFree(ws);
  error_result:
    res->lpVtbl->Release(res);
  error_dlg:
    dlg->lpVtbl->Release(dlg);
    return;
  } else if (dlgtype == WEBVIEW_DIALOG_TYPE_ALERT) {
    UINT type = MB_OK;
    switch (flags & WEBVIEW_DIALOG_FLAG_ALERT_MASK) {
    case WEBVIEW_DIALOG_FLAG_INFO:
      type |= MB_ICONINFORMATION;
      break;
    case WEBVIEW_DIALOG_FLAG_WARNING:
      type |= MB_ICONWARNING;
      break;
    case WEBVIEW_DIALOG_FLAG_ERROR:
      type |= MB_ICONERROR;
      break;
    }
    MessageBox(w->priv.hwnd, arg, title, type);
  }
}

WEBVIEW_API void webview_terminate(struct webview *w) { PostQuitMessage(0); }
WEBVIEW_API void webview_exit(struct webview *w) { OleUninitialize(); }
WEBVIEW_API void webview_print_log(const char *s) { OutputDebugString(s); }

#endif /* WEBVIEW_IMPLEMENTATION */

#ifdef __cplusplus
}
#endif

#endif /* WEBVIEW_H */
