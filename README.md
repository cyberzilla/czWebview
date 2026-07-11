# czWebview

**Full-featured VB6 WebView2 UserControl** — Embed modern web content (HTML5, CSS3, ES6+) directly in VB6 applications using Microsoft Edge WebView2 Runtime. Zero external dependencies, no OCX registration required.

---

## ✨ Features

### Core
- 🌐 Navigate, GoBack, GoForward, Reload, Stop
- ⚡ ExecuteScript — synchronous & async JavaScript execution with JSON results
- 🔗 JavaScript Bridge — bidirectional messaging (`PostWebMessage` / `chrome.webview.postMessage`)
- 🧩 Host Objects — expose VB6 COM objects to JavaScript
- 📦 Single-project embed — just add `.ctl` + `.cls` + `.tlb` to your project

### Advanced
- 🖨️ Print UI & PrintToPdf — print pages or save as PDF
- 📸 CapturePreview — screenshot pages as PNG/JPG byte arrays
- 🍪 Cookie Management — Get, Add, Update, Delete cookies
- 🧹 Clear Browsing Data — cache, history, cookies, passwords, etc.
- ⏸️ Suspend/Resume — memory optimization for background tabs
- 🔧 DevTools Protocol — call Chrome DevTools Protocol methods directly
- 📥 Download handling with custom file path
- 🔒 Basic Authentication handling
- 🌍 Virtual Host Name → Folder Mapping (serve local files)
- 🔑 Permission requests (camera, microphone, geolocation)
- 💬 Script Dialogs (alert/confirm/prompt)
- 📋 Context Menu events
- 🔍 Web Resource Request filtering & response interception
- 🖥️ Full Screen element detection
- 🔇 Mute/unmute audio
- 🏷️ User Agent customization
- 📊 Status Bar text
- 🔎 Zoom control (programmatic + pinch zoom toggle)

### Architecture
- **44 callback interfaces** — full WebView2 API coverage (matching wqweto's cWebView2)
- **29 events** — complete lifecycle control
- **34 public methods** — comprehensive API surface
- **48 properties** — fine-grained settings control
- **Clean separation** — `.ctl` = public API, `.cls` = internal COM callbacks

---

## 📋 Requirements

| Component | Required | Notes |
|---|---|---|
| VB6 IDE | ✅ | SP6 recommended |
| WebView2 Runtime | ✅ | Pre-installed on Windows 10 (1803+) and Windows 11 |
| WebView2Loader.dll | ✅ | Included in `External\` folder (~115KB, x86) |
| czWebview.tlb | ✅ | Pre-compiled included, or build with `typelib\compile_tlb.bat` |
| Windows SDK (MIDL) | ❌ | Only needed to recompile the TLB |

---

## 🚀 Quick Start

### 1. Add to your VB6 project

```
1. Project → Add User Control → Existing → select src\czWebview.ctl
2. Project → Add Class Module → Existing → select src\czWebviewCallback.cls
3. Project → References → Browse → select typelib\czWebview.tlb
4. Place WebView2Loader.dll next to your .exe (or in app folder)
```

### 2. Place the control on your Form

After adding the `.ctl` to your project, `czWebview` appears in the Toolbox. Drop it on your Form like any other control, or place it at design-time in the `.frm` header:

```
Begin YourProject.czWebview czWebview1
   Height = 5775
   Left   = 0
   Top    = 480
   Width  = 10200
End
```

### 3. Handle events

```vb
Private Sub czWebview1_InitComplete(ByVal Success As Boolean, ByVal ErrorCode As Long)
    If Success Then
        czWebview1.Navigate "https://www.google.com"
    Else
        MsgBox "WebView2 init failed: &H" & Hex$(ErrorCode), vbCritical
    End If
End Sub

Private Sub czWebview1_TitleChanged(ByVal Title As String)
    Me.Caption = Title
End Sub

Private Sub czWebview1_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    lblStatus.Caption = IIf(IsSuccess, "Done", "Error: " & WebErrorStatus)
End Sub

Private Sub czWebview1_WebMessageReceived(ByVal Message As String, ByVal IsJSON As Boolean)
    Debug.Print "From JS: " & Message
End Sub
```

### 4. JavaScript Bridge

**VB6 → JavaScript:**
```vb
' Execute JavaScript synchronously (returns JSON result)
Dim result As String
result = czWebview1.ExecuteScript("document.title")
' result = """My Page Title"""

' Execute async (result via ScriptCompleted event)
czWebview1.ExecuteScriptAsync "fetch('/api/data').then(r => r.json())"

' Send message to page
czWebview1.PostWebMessageAsString "Hello from VB6!"
czWebview1.PostWebMessageAsJSON "{""action"":""update"",""value"":42}"
```

**JavaScript → VB6:**
```javascript
// Send message to VB6
chrome.webview.postMessage("Hello from JavaScript!");

// Listen for messages from VB6
chrome.webview.addEventListener('message', function(e) {
    console.log('From VB6:', e.data);
});
```

---

## 📁 Project Structure

```
czWebview/
├── src/
│   ├── czWebview.ctl               # ★ Main UserControl — public API
│   └── czWebviewCallback.cls       # Internal callback handler (44 interfaces)
├── typelib/
│   ├── czWebview.odl               # Interface definitions (ODL source)
│   ├── czWebview.tlb               # Compiled type library
│   └── compile_tlb.bat             # Auto-compile script (uses MIDL)
├── External/
│   └── WebView2Loader.dll          # Microsoft x86 loader
├── demo/
│   ├── SimpleTest/                 # Simple browser demo
│   │   ├── TestProject.vbp
│   │   └── frmTest.frm
│   └── WhatsappBot/                # WhatsApp Web bot demo
│       ├── WhatsappBot.vbp
│       └── frmWhatsappBot.frm
└── README.md
```

---

## 🔧 Compiling the Type Library

A pre-compiled `.tlb` is included. To recompile after editing the `.odl`:

```batch
cd typelib
compile_tlb.bat
```

This requires the Windows SDK (MIDL compiler). The script auto-detects the SDK path.

---

## 📖 API Reference

### Properties

| Property | Type | Access | Description |
|---|---|---|---|
| `URL` | String | Get/Let | Default property. Design-time URL or navigate at runtime |
| `UserDataFolder` | String | Get/Let | Custom user data folder path |
| `DocumentURL` | String | Get | Current document URL |
| `DocumentTitle` | String | Get | Current page title |
| `BrowserVersion` | String | Get | WebView2 runtime version string |
| `BrowserProcessId` | Long | Get | Browser process ID |
| `IsInitialized` | Boolean | Get | `True` when WebView2 is ready |
| `CanGoBack` | Boolean | Get | Navigation history has back entry |
| `CanGoForward` | Boolean | Get | Navigation history has forward entry |
| `ZoomFactor` | Double | Get/Let | Zoom level (1.0 = 100%) |
| `TimeOutSeconds` | Double | Get/Let | Default timeout for sync operations (default: 8) |
| `ContainsFullScreenElement` | Boolean | Get | Page has full-screen element |
| `StatusBarText` | String | Get | Current status bar text |
| `IsSuspended` | Boolean | Get | WebView is suspended |
| `IsMuted` | Boolean | Get/Let | Mute/unmute audio |

#### Settings Properties

| Property | Type | Access | Description |
|---|---|---|---|
| `IsScriptEnabled` | Boolean | Get/Let | Enable/disable JavaScript |
| `IsWebMessageEnabled` | Boolean | Get/Let | Enable/disable web messaging |
| `AreDevToolsEnabled` | Boolean | Get/Let | Enable/disable DevTools (F12) |
| `AreDefaultContextMenusEnabled` | Boolean | Get/Let | Enable/disable right-click menu |
| `AreDefaultScriptDialogsEnabled` | Boolean | Get/Let | Enable/disable JS dialogs |
| `AreHostObjectsAllowed` | Boolean | Get/Let | Enable/disable host objects |
| `IsStatusBarEnabled` | Boolean | Get/Let | Enable/disable status bar |
| `IsZoomControlEnabled` | Boolean | Get/Let | Enable/disable zoom control |
| `UserAgent` | String | Get/Let | Custom user agent string |
| `AreBrowserAcceleratorKeysEnabled` | Boolean | Get/Let | Enable/disable browser shortcuts |
| `IsPasswordAutosaveEnabled` | Boolean | Get/Let | Enable/disable password save |
| `IsGeneralAutofillEnabled` | Boolean | Get/Let | Enable/disable autofill |
| `IsPinchZoomEnabled` | Boolean | Get/Let | Enable/disable pinch zoom |
| `IsSwipeNavigationEnabled` | Boolean | Get/Let | Enable/disable swipe navigation |

---

### Methods

#### Navigation

| Method | Returns | Description |
|---|---|---|
| `Navigate(URL)` | — | Navigate to URL (async) |
| `NavigateSync(URL, [Timeout]) ` | Boolean | Navigate and wait until complete |
| `NavigateToString(HTML)` | — | Render HTML string directly |
| `GoBack()` | — | Navigate back |
| `GoForward()` | — | Navigate forward |
| `Reload()` | — | Refresh page |
| `StopNavigation()` | — | Stop loading |

#### JavaScript

| Method | Returns | Description |
|---|---|---|
| `ExecuteScript(script, [Timeout])` | String | Execute JS synchronously, returns JSON result |
| `ExecuteScriptAsync(script)` | — | Execute JS asynchronously (result via `ScriptCompleted` event) |
| `AddScriptToExecuteOnDocumentCreated(script)` | — | Inject JS that runs on every page load |
| `PostWebMessageAsString(msg)` | — | Send string message to page |
| `PostWebMessageAsJSON(json)` | — | Send JSON message to page |
| `AddHostObjectToScript(name, obj)` | — | Expose VB6 COM object to JavaScript |
| `RemoveHostObjectFromScript(name)` | — | Remove exposed host object |
| `CallDevToolsProtocolMethod(method, params)` | String | Call Chrome DevTools Protocol method |

#### Capture & Print

| Method | Returns | Description |
|---|---|---|
| `CapturePreview([ImageFormat])` | Byte() | Screenshot as PNG/JPG byte array |
| `PrintToPdf(FilePath, [Timeout])` | Boolean | Save page as PDF file |
| `ShowPrintUI([PrintDialogKind])` | — | Show browser or system print dialog |

#### Cookie Management

| Method | Returns | Description |
|---|---|---|
| `GetCookies([URL], [Timeout])` | Collection | Get cookies (each item is a Collection with Name, Value, Domain, Path, etc.) |
| `AddOrUpdateCookie(Name, Value, Domain, [Path], [Expires], [IsSecure], [IsHttpOnly], [SameSite])` | — | Create or update a cookie |
| `DeleteCookies(Name, URI)` | — | Delete cookies by name and URI |
| `DeleteAllCookies()` | — | Delete all cookies |

#### Resource Management

| Method | Returns | Description |
|---|---|---|
| `AddWebResourceRequestedFilter(filter, [context])` | — | Add URL filter for `WebResourceRequested` event |
| `RemoveWebResourceRequestedFilter(filter, context)` | — | Remove URL filter |
| `SetVirtualHostNameToFolderMapping(host, folder, [access])` | — | Map hostname to local folder |
| `ClearVirtualHostNameToFolderMapping(host)` | — | Remove hostname mapping |

#### Lifecycle

| Method | Returns | Description |
|---|---|---|
| `Init([UserDataFolder], [BrowserArgs], [Timeout])` | Long | Initialize WebView2 (usually auto-called) |
| `TrySuspend([Timeout])` | Boolean | Suspend to save memory |
| `ResumeFromSuspend()` | — | Resume from suspended state |
| `ClearBrowsingData(DataKinds, [Timeout])` | Boolean | Clear browsing data (cache, cookies, history, etc.) |
| `SetFocus([Reason])` | — | Set focus to WebView |
| `OpenDevToolsWindow()` | — | Open DevTools |
| `OpenTaskManagerWindow()` | — | Open browser task manager |
| `Shutdown()` | — | Release all resources |

---

### Events

#### Navigation

| Event | Parameters | Description |
|---|---|---|
| `InitComplete` | `Success`, `ErrorCode` | WebView2 initialization completed |
| `NavigationStarting` | `IsUserInitiated`, `IsRedirected`, `URI`, `Cancel` | Before navigation (set `Cancel=True` to block) |
| `NavigationCompleted` | `IsSuccess`, `WebErrorStatus` | Navigation finished |
| `DocumentComplete` | — | Successful navigation completed |
| `SourceChanged` | — | URL changed |
| `ContentLoading` | `IsErrorPage` | Content loading started |
| `DOMContentLoaded` | — | DOM content loaded |

#### Page Info

| Event | Parameters | Description |
|---|---|---|
| `TitleChanged` | `Title` | Page title changed |
| `HistoryChanged` | — | Navigation history updated |
| `StatusBarTextChanged` | `Text` | Status bar text changed |
| `ContainsFullScreenElementChanged` | `IsFullScreen` | Full screen state changed |
| `ZoomFactorChanged` | — | Zoom level changed |

#### User Interaction

| Event | Parameters | Description |
|---|---|---|
| `WebMessageReceived` | `Message`, `IsJSON` | Message from JavaScript |
| `ScriptCompleted` | `ErrorCode`, `Result`, `Token` | Async script execution result |
| `NewWindowRequested` | `IsUserInitiated`, `Handled`, `URI`, `NewWindowFeatures` | New window/tab requested |
| `ContextMenuRequested` | `PageURI`, `LinkURI`, `SelectionText`, `ScreenX`, `ScreenY`, `Handled` | Right-click context menu |
| `ScriptDialogOpening` | `Kind`, `Accept`, `ResultText`, `URI`, `Message`, `DefaultText` | JS alert/confirm/prompt |
| `PermissionRequested` | `IsUserInitiated`, `State`, `URI`, `PermissionKind` | Permission request |
| `AcceleratorKeyPressed` | `KeyState`, `IsExtendedKey`, `WasKeyDown`, `IsKeyReleased`, `IsMenuKeyDown`, `RepeatCount`, `ScanCode`, `Handled` | Key press in WebView |

#### Downloads & Resources

| Event | Parameters | Description |
|---|---|---|
| `DownloadStarting` | `URI`, `MimeType`, `SuggestedPath`, `CancelDownload`, `NewFilePath` | Download starting |
| `WebResourceRequested` | `URI`, `ResourceContext`, `Permit` | Web resource intercepted |
| `WebResourceResponseReceived` | `ReqURI`, `ReqMethod`, `RespStatus`, `RespReasonPhrase`, `RespHeaders` | Response received |

#### Focus & Window

| Event | Parameters | Description |
|---|---|---|
| `WebViewGotFocus` | — | WebView gained focus |
| `WebViewLostFocus` | — | WebView lost focus |
| `MoveFocusRequested` | `Reason`, `Handled` | Focus move requested |
| `WindowCloseRequested` | — | `window.close()` called |

#### Other

| Event | Parameters | Description |
|---|---|---|
| `ProcessFailed` | — | Browser process crashed |
| `BasicAuthenticationRequested` | `URI`, `Challenge`, `UserName`, `Password`, `Cancel` | HTTP Basic Auth |
| `FrameCreated` | `FrameName` | New iframe created |

---

### Enums

| Enum | Values | Used By |
|---|---|---|
| `czWebView2FocusReason` | `PROGRAMMATIC`, `NEXT`, `PREVIOUS` | `SetFocus`, `MoveFocusRequested` |
| `czWebView2ImageCaptureFormat` | `PNG`, `JPG` | `CapturePreview` |
| `czWebView2PermissionKind` | `UNKNOWN`, `MICROPHONE`, `CAMERA`, `GEOLOCATION`, `NOTIFICATIONS`, `OTHER_SENSORS`, `CLIPBOARD_READ` | `PermissionRequested` |
| `czWebView2PermissionState` | `DEFAULT`, `ALLOW`, `DENY` | `PermissionRequested`, `WebResourceRequested` |
| `czWebView2ScriptDialogKind` | `ALERT`, `CONFIRM`, `PROMPT`, `BEFOREUNLOAD` | `ScriptDialogOpening` |
| `czWebView2ResourceFilter` | `ALL`, `DOCUMENT`, `STYLESHEET`, `IMAGE`, `MEDIA`, `FONT`, `SCRIPT`, `XML_HTTP_REQUEST`, `FETCH`, etc. | `AddWebResourceRequestedFilter` |
| `czWebView2AccKeyState` | `KEY_DOWN`, `KEY_UP`, `SYSKEY_DOWN`, `SYSKEY_UP` | `AcceleratorKeyPressed` |
| `czWebView2HostResourceAccessKind` | `DENY`, `ALLOW`, `DENY_CORS` | `SetVirtualHostNameToFolderMapping` |
| `czWebView2PreferredColorScheme` | `AUTO`, `LIGHT`, `DARK` | Color scheme preference |
| `czWebView2PrintDialogKind` | `BROWSER`, `SYSTEM` | `ShowPrintUI` |
| `czWebView2CookieSameSiteKind` | `NONE`, `LAX`, `STRICT` | `AddOrUpdateCookie` |
| `czWebView2BrowsingDataKinds` | `FILE_SYSTEMS`, `INDEXED_DB`, `LOCAL_STORAGE`, `COOKIES`, `DISK_CACHE`, `BROWSING_HISTORY`, `ALL_PROFILE`, etc. | `ClearBrowsingData` |

---

## 💡 Usage Examples

### Screenshot to file
```vb
Dim baPng() As Byte
baPng = czWebview1.CapturePreview(czCaptureAs_PNG)
Open "screenshot.png" For Binary As #1
Put #1, , baPng
Close #1
```

### Save page as PDF
```vb
If czWebview1.PrintToPdf("C:\output\page.pdf") Then
    MsgBox "PDF saved!"
End If
```

### Cookie management
```vb
' Add a cookie
czWebview1.AddOrUpdateCookie "session", "abc123", ".example.com", "/", , True, True

' Get all cookies for a URL
Dim cCookies As Collection
Set cCookies = czWebview1.GetCookies("https://example.com")
Dim cItem As Collection
For Each cItem In cCookies
    Debug.Print cItem("Name") & " = " & cItem("Value")
Next

' Clear all cookies
czWebview1.DeleteAllCookies
```

### Block specific URLs
```vb
Private Sub Form_Load()
    czWebview1.AddWebResourceRequestedFilter "*.ads.example.com/*"
End Sub

Private Sub czWebview1_WebResourceRequested(ByVal URI As String, ByVal ResourceContext As czWebView2ResourceFilter, Permit As czWebView2PermissionState)
    Permit = czPERMISSION_STATE_DENY
End Sub
```

### Serve local files via virtual hostname
```vb
czWebview1.SetVirtualHostNameToFolderMapping "app.local", App.Path & "\www", czHostResourceAccess_ALLOW
czWebview1.Navigate "https://app.local/index.html"
```

### Suspend to save memory
```vb
' When minimized or hidden
If czWebview1.TrySuspend() Then
    Debug.Print "Suspended! IsSuspended=" & czWebview1.IsSuspended
End If

' When restored
czWebview1.ResumeFromSuspend
```

### Clear browsing data
```vb
czWebview1.ClearBrowsingData czBrowsingData_COOKIES Or czBrowsingData_DISK_CACHE
```

### DevTools Protocol
```vb
' Emulate mobile device
Dim sResult As String
sResult = czWebview1.CallDevToolsProtocolMethod("Emulation.setDeviceMetricsOverride", _
    "{""width"":375,""height"":812,""deviceScaleFactor"":3,""mobile"":true}")
```

---

## 📝 License

MIT

## 🙏 Credits

- Interface patterns inspired by [cWebView2](https://github.com/wqweto/cWebView2) by wqweto
- WebView2 Runtime by Microsoft
- Built with ❤️ for the VB6 community
