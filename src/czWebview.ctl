VERSION 5.00
Begin VB.UserControl czWebview 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "czWebview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' czWebview UserControl
' Zero-dependency WebView2 wrapper for VB6 — full-featured, matching
' wqweto's cWebView2 API surface. Requires czWebview.tlb (type library
' reference) and WebView2Loader.dll at runtime.
'
' Architecture:
'   - czWebviewCallback.cls handles ALL callback/handler interfaces
'   - The UserControl provides the public API (properties, methods, events)
'   - Friend subs receive forwarded events from the callback class
'   - Message pump (SpinThreadMessagePump) enables synchronous blocking
'     calls from VB6 while allowing COM callbacks to arrive
'=========================================================================
Option Explicit

'=========================================================================
' Public enums
'=========================================================================

Public Enum czWebView2FocusReason
    czFocusReason_PROGRAMMATIC = 0
    czFocusReason_NEXT = 1
    czFocusReason_PREVIOUS = 2
End Enum

Public Enum czWebView2ImageCaptureFormat
    czCaptureAs_PNG = 0
    czCaptureAs_JPG = 1
End Enum

Public Enum czWebView2PermissionKind
    czPERMISSION_KIND_UNKNOWN = 0
    czPERMISSION_KIND_MICROPHONE = 1
    czPERMISSION_KIND_CAMERA = 2
    czPERMISSION_KIND_GEOLOCATION = 3
    czPERMISSION_KIND_NOTIFICATIONS = 4
    czPERMISSION_KIND_OTHER_SENSORS = 5
    czPERMISSION_KIND_CLIPBOARD_READ = 6
End Enum

Public Enum czWebView2PermissionState
    czPERMISSION_STATE_DEFAULT = 0
    czPERMISSION_STATE_ALLOW = 1
    czPERMISSION_STATE_DENY = 2
End Enum

Public Enum czWebView2ScriptDialogKind
    czSCRIPT_DIALOG_ALERT = 0
    czSCRIPT_DIALOG_CONFIRM = 1
    czSCRIPT_DIALOG_PROMPT = 2
    czSCRIPT_DIALOG_BEFOREUNLOAD = 3
End Enum

Public Enum czWebView2ResourceFilter
    czFilter_ALL = 0
    czFilter_DOCUMENT = 1
    czFilter_STYLESHEET = 2
    czFilter_IMAGE = 3
    czFilter_MEDIA = 4
    czFilter_FONT = 5
    czFilter_SCRIPT = 6
    czFilter_XML_HTTP_REQUEST = 7
    czFilter_FETCH = 8
    czFilter_TEXT_TRACK = 9
    czFilter_EVENT_SOURCE = 10
    czFilter_WEBSOCKET = 11
    czFilter_MANIFEST = 12
    czFilter_SIGNED_EXCHANGE = 13
    czFilter_PING = 14
    czFilter_CSP_VIOLATION_REPORT = 15
    czFilter_OTHER = 16
End Enum

Public Enum czWebView2AccKeyState
    czKey_DOWN = 0
    czKey_UP = 1
    czSysKey_DOWN = 2
    czSysKey_UP = 3
End Enum

Public Enum czWebView2HostResourceAccessKind
    czHostResourceAccess_DENY = 0
    czHostResourceAccess_ALLOW = 1
    czHostResourceAccess_DENY_CORS = 2
End Enum

Public Enum czWebView2PreferredColorScheme
    czColorScheme_AUTO = 0
    czColorScheme_LIGHT = 1
    czColorScheme_DARK = 2
End Enum

Public Enum czWebView2PrintDialogKind
    czPrintDialog_BROWSER = 0
    czPrintDialog_SYSTEM = 1
End Enum

Public Enum czWebView2CookieSameSiteKind
    czSameSite_NONE = 0
    czSameSite_LAX = 1
    czSameSite_STRICT = 2
End Enum

Public Enum czWebView2BrowsingDataKinds
    czBrowsingData_FILE_SYSTEMS = &H1
    czBrowsingData_INDEXED_DB = &H2
    czBrowsingData_LOCAL_STORAGE = &H4
    czBrowsingData_WEB_SQL = &H8
    czBrowsingData_CACHE_STORAGE = &H10
    czBrowsingData_ALL_DOM_STORAGE = &H20
    czBrowsingData_COOKIES = &H40
    czBrowsingData_ALL_SITE = &H80
    czBrowsingData_DISK_CACHE = &H100
    czBrowsingData_DOWNLOAD_HISTORY = &H200
    czBrowsingData_GENERAL_AUTOFILL = &H400
    czBrowsingData_PASSWORD_AUTOSAVE = &H800
    czBrowsingData_BROWSING_HISTORY = &H1000
    czBrowsingData_SETTINGS = &H2000
    czBrowsingData_ALL_PROFILE = &H3FFF
End Enum

'=========================================================================
' Public events
'=========================================================================
Public Event InitComplete(ByVal Success As Boolean, ByVal errorCode As Long)
Public Event NavigationStarting(ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal URI As String, Cancel As Boolean)
Public Event NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
Public Event DocumentComplete()
Public Event SourceChanged()
Public Event ProcessFailed()
Public Event TitleChanged(ByVal Title As String)
Public Event ContentLoading(ByVal IsErrorPage As Boolean)
Public Event HistoryChanged()
Public Event WindowCloseRequested()
Public Event WebMessageReceived(ByVal Message As String, ByVal IsJSON As Boolean)
Public Event ScriptCompleted(ByVal errorCode As Long, ByVal Result As String, ByVal Token As Currency)
Public Event NewWindowRequested(ByVal IsUserInitiated As Boolean, Handled As Boolean, ByVal URI As String, NewWindowFeatures As Collection)
Public Event PermissionRequested(ByVal IsUserInitiated As Boolean, State As czWebView2PermissionState, ByVal URI As String, ByVal PermissionKind As czWebView2PermissionKind)
Public Event ScriptDialogOpening(ByVal Kind As czWebView2ScriptDialogKind, Accept As Boolean, ResultText As String, ByVal URI As String, ByVal Message As String, ByVal DefaultText As String)
Public Event AcceleratorKeyPressed(ByVal KeyState As czWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, Handled As Boolean)
Public Event DownloadStarting(ByVal URI As String, ByVal MimeType As String, ByVal SuggestedPath As String, CancelDownload As Boolean, NewFilePath As String)
Public Event WebResourceRequested(ByVal URI As String, ByVal ResourceContext As czWebView2ResourceFilter, Handled As Boolean)
Public Event WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, RespHeaders As Collection)
Public Event ContextMenuRequested(ByVal PageURI As String, ByVal LinkURI As String, ByVal SelectionText As String, ByVal ScreenX As Long, ByVal ScreenY As Long, Handled As Boolean)
Public Event WebViewGotFocus()
Public Event WebViewLostFocus()
Public Event ZoomFactorChanged()
Public Event MoveFocusRequested(ByVal Reason As czWebView2FocusReason, Handled As Boolean)
Public Event StatusBarTextChanged(ByVal Text As String)
Public Event ContainsFullScreenElementChanged(ByVal IsFullScreen As Boolean)
Public Event DOMContentLoaded()
Public Event BasicAuthenticationRequested(ByVal URI As String, ByVal Challenge As String, UserName As String, Password As String, Cancel As Boolean)
Public Event FrameCreated(ByVal FrameName As String)

'=========================================================================
' API declarations
'=========================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As APIRECT) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpsz As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, ByVal UserDataFolder As Long, ByVal environmentOptions As Long, ByVal environmentCreatedHandler As Long) As Long
Private Declare Function GetAvailableCoreWebView2BrowserVersionString Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, versionInfo As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function CoWaitForMultipleHandles Lib "ole32" (ByVal dwFlags As Long, ByVal dwTimeout As Long, ByVal cHandles As Long, pHandles As Any, lpdwindex As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByVal lpMsg As Long, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByVal lpMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByVal lpMsg As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IVBStream) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long

Private Type APIMSG
    hWnd                As Long
    lMessage            As Long
    wParam              As Long
    lParam              As Long
    time                As Long
    ptX                 As Long
    ptY                 As Long
End Type

'=========================================================================
' Constants
'=========================================================================
Private Const PM_REMOVE As Long = 1
Private Const WM_NCMOUSEFIRST As Long = &HA0
Private Const WM_NCMOUSELAST As Long = &HA9
Private Const WM_KEYFIRST As Long = &H100
Private Const WM_KEYLAST As Long = &H108
Private Const WM_MOUSEFIRST As Long = &H200
Private Const WM_MOUSELAST As Long = &H209
Private Const E_MOD_NOT_FOUND As Long = &H8007007E
Private Const RPC_E_TIMEOUT As Long = &H8001011F
Private Const SECS_PER_DAY As Long = 86400
Private Const STREAM_SEEK_SET As Long = 0
Private Const STREAM_SEEK_END As Long = 2
Private Const QS_ALLINPUT As Long = &H4FF
Private Const GMEM_MOVEABLE As Long = &H2
'--- Deferred settings bitmask flags
Private Const DEF_SCRIPT As Long = 1
Private Const DEF_WEBMSG As Long = 2
Private Const DEF_DIALOGS As Long = 4
Private Const DEF_STATUS As Long = 8
Private Const DEF_DEVTOOLS As Long = &H10
Private Const DEF_CTXMENU As Long = &H20
Private Const DEF_HOSTOBJ As Long = &H40
Private Const DEF_ZOOM As Long = &H80
Private Const DEF_ERRORPAGE As Long = &H100
Private Const DEF_ACCELKEYS As Long = &H200
Private Const DEF_PWDSAVE As Long = &H400
Private Const DEF_AUTOFILL As Long = &H800
Private Const DEF_PINCH As Long = &H1000
Private Const DEF_SWIPE As Long = &H2000
'--- WebView2 default values (all True except DEF_PWDSAVE)
Private Const DEF_WV2_DEFAULTS As Long = DEF_SCRIPT Or DEF_WEBMSG Or DEF_DIALOGS Or DEF_STATUS Or DEF_DEVTOOLS Or DEF_CTXMENU Or DEF_HOSTOBJ Or DEF_ZOOM Or DEF_ERRORPAGE Or DEF_ACCELKEYS Or DEF_AUTOFILL Or DEF_PINCH Or DEF_SWIPE

'=========================================================================
' Member variables
'=========================================================================
Private m_oEnvironment              As IVBCoreWebView2Environment
Private m_oController               As IVBCoreWebView2Controller
Private m_oWebView                  As IVBCoreWebView2
Private m_oWebView4                 As IVBCoreWebView2_4
Private m_oWebView11                As IVBCoreWebView2_11
Private m_oWebView13                As IVBCoreWebView2_13
Private m_oWebView16                As IVBCoreWebView2_16
Private m_oSettings                 As IVBCoreWebView2Settings
Private m_oSettings6                As IVBCoreWebView2Settings6
Private m_oEventSink                As czWebviewCallback    '--- main event sink
Private m_oFocusGotSink             As czWebviewCallback    '--- discriminated GotFocus
Private m_oFocusLostSink            As czWebviewCallback    '--- discriminated LostFocus
Private m_bInitComplete             As Boolean
Private m_hInitError                As Long
Private m_bNavDone                  As Boolean
Private m_bNavSuccess               As Boolean
Private m_lNavErrorStatus           As Long
Private m_bShuttingDown             As Boolean
Private m_dblTimeOutSeconds         As Double
Private m_cTokenSeq                 As Currency
Private m_cScriptResults            As Collection
Private m_sUserDataFolder           As String
Private m_sInitialURL               As String
Private m_cPendingCallbacks         As Collection
Private m_lDeferredMask             As Long
Private m_lDeferredValues           As Long
Private m_sDeferredUserAgent        As String
Private m_oCurrentResourceArgs      As IVBCoreWebView2WebResourceRequestedEventArgs  '--- held during WebResourceRequested event

'=========================================================================
' Properties
'=========================================================================

Public Property Get IsInitialized() As Boolean
    IsInitialized = (Not m_oWebView Is Nothing)
End Property

Public Property Get CanGoBack() As Boolean
    If Not m_oWebView Is Nothing Then
        CanGoBack = m_oWebView.CanGoBack
    End If
End Property

Public Property Get CanGoForward() As Boolean
    If Not m_oWebView Is Nothing Then
        CanGoForward = m_oWebView.CanGoForward
    End If
End Property

Public Property Get DocumentURL() As String
    If Not m_oWebView Is Nothing Then
        DocumentURL = pvToString(m_oWebView.Source)
    End If
End Property

Public Property Get DocumentTitle() As String
    If Not m_oWebView Is Nothing Then
        DocumentTitle = pvToString(m_oWebView.DocumentTitle)
    End If
End Property

Public Property Get BrowserVersion() As String
    If Not m_oEnvironment Is Nothing Then
        BrowserVersion = pvToString(m_oEnvironment.BrowserVersionString)
    End If
End Property

Public Property Get BrowserProcessId() As Long
    If Not m_oWebView Is Nothing Then
        BrowserProcessId = m_oWebView.BrowserProcessId
    End If
End Property

Public Property Get ZoomFactor() As Double
    If Not m_oController Is Nothing Then
        ZoomFactor = m_oController.ZoomFactor
    End If
End Property

Public Property Let ZoomFactor(ByVal dblValue As Double)
    If Not m_oController Is Nothing Then
        m_oController.ZoomFactor = dblValue
    End If
End Property

Public Property Get IsScriptEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsScriptEnabled = m_oSettings.IsScriptEnabled
    Else
        IsScriptEnabled = (m_lDeferredValues And DEF_SCRIPT) <> 0
    End If
End Property

Public Property Let IsScriptEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsScriptEnabled = -bValue
    Else
        pvSetDeferredBool DEF_SCRIPT, bValue
    End If
End Property

Public Property Get IsWebMessageEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsWebMessageEnabled = m_oSettings.IsWebMessageEnabled
    Else
        IsWebMessageEnabled = (m_lDeferredValues And DEF_WEBMSG) <> 0
    End If
End Property

Public Property Let IsWebMessageEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsWebMessageEnabled = -bValue
    Else
        pvSetDeferredBool DEF_WEBMSG, bValue
    End If
End Property

Public Property Get AreDefaultScriptDialogsEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDefaultScriptDialogsEnabled = m_oSettings.AreDefaultScriptDialogsEnabled
    Else
        AreDefaultScriptDialogsEnabled = (m_lDeferredValues And DEF_DIALOGS) <> 0
    End If
End Property

Public Property Let AreDefaultScriptDialogsEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDefaultScriptDialogsEnabled = -bValue
    Else
        pvSetDeferredBool DEF_DIALOGS, bValue
    End If
End Property

Public Property Get IsStatusBarEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsStatusBarEnabled = m_oSettings.IsStatusBarEnabled
    Else
        IsStatusBarEnabled = (m_lDeferredValues And DEF_STATUS) <> 0
    End If
End Property

Public Property Let IsStatusBarEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsStatusBarEnabled = -bValue
    Else
        pvSetDeferredBool DEF_STATUS, bValue
    End If
End Property

Public Property Get AreDevToolsEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDevToolsEnabled = m_oSettings.AreDevToolsEnabled
    Else
        AreDevToolsEnabled = (m_lDeferredValues And DEF_DEVTOOLS) <> 0
    End If
End Property

Public Property Let AreDevToolsEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDevToolsEnabled = -bValue
    Else
        pvSetDeferredBool DEF_DEVTOOLS, bValue
    End If
End Property

Public Property Get AreDefaultContextMenusEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDefaultContextMenusEnabled = m_oSettings.AreDefaultContextMenusEnabled
    Else
        AreDefaultContextMenusEnabled = (m_lDeferredValues And DEF_CTXMENU) <> 0
    End If
End Property

Public Property Let AreDefaultContextMenusEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDefaultContextMenusEnabled = -bValue
    Else
        pvSetDeferredBool DEF_CTXMENU, bValue
    End If
End Property

Public Property Get AreHostObjectsAllowed() As Boolean
    If Not m_oSettings Is Nothing Then
        AreHostObjectsAllowed = m_oSettings.AreHostObjectsAllowed
    Else
        AreHostObjectsAllowed = (m_lDeferredValues And DEF_HOSTOBJ) <> 0
    End If
End Property

Public Property Let AreHostObjectsAllowed(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreHostObjectsAllowed = -bValue
    Else
        pvSetDeferredBool DEF_HOSTOBJ, bValue
    End If
End Property

Public Property Get IsZoomControlEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsZoomControlEnabled = m_oSettings.IsZoomControlEnabled
    Else
        IsZoomControlEnabled = (m_lDeferredValues And DEF_ZOOM) <> 0
    End If
End Property

Public Property Let IsZoomControlEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsZoomControlEnabled = -bValue
    Else
        pvSetDeferredBool DEF_ZOOM, bValue
    End If
End Property

Public Property Get UserAgent() As String
    If Not m_oSettings6 Is Nothing Then
        UserAgent = pvToString(m_oSettings6.UserAgent)
    ElseIf LenB(m_sDeferredUserAgent) <> 0 Then
        UserAgent = m_sDeferredUserAgent
    End If
End Property

Public Property Let UserAgent(ByVal sValue As String)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.UserAgent = StrPtr(sValue)
    Else
        m_sDeferredUserAgent = sValue
    End If
End Property

Public Property Get AreBrowserAcceleratorKeysEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        AreBrowserAcceleratorKeysEnabled = m_oSettings6.AreBrowserAcceleratorKeysEnabled
    Else
        AreBrowserAcceleratorKeysEnabled = (m_lDeferredValues And DEF_ACCELKEYS) <> 0
    End If
End Property

Public Property Let AreBrowserAcceleratorKeysEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.AreBrowserAcceleratorKeysEnabled = -bValue
    Else
        pvSetDeferredBool DEF_ACCELKEYS, bValue
    End If
End Property

Public Property Get IsPasswordAutosaveEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsPasswordAutosaveEnabled = m_oSettings6.IsPasswordAutosaveEnabled
    Else
        IsPasswordAutosaveEnabled = (m_lDeferredValues And DEF_PWDSAVE) <> 0
    End If
End Property

Public Property Let IsPasswordAutosaveEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsPasswordAutosaveEnabled = -bValue
    Else
        pvSetDeferredBool DEF_PWDSAVE, bValue
    End If
End Property

Public Property Get IsGeneralAutofillEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsGeneralAutofillEnabled = m_oSettings6.IsGeneralAutofillEnabled
    Else
        IsGeneralAutofillEnabled = (m_lDeferredValues And DEF_AUTOFILL) <> 0
    End If
End Property

Public Property Let IsGeneralAutofillEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsGeneralAutofillEnabled = -bValue
    Else
        pvSetDeferredBool DEF_AUTOFILL, bValue
    End If
End Property

Public Property Get IsPinchZoomEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsPinchZoomEnabled = m_oSettings6.IsPinchZoomEnabled
    Else
        IsPinchZoomEnabled = (m_lDeferredValues And DEF_PINCH) <> 0
    End If
End Property

Public Property Let IsPinchZoomEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsPinchZoomEnabled = -bValue
    Else
        pvSetDeferredBool DEF_PINCH, bValue
    End If
End Property

Public Property Get IsSwipeNavigationEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsSwipeNavigationEnabled = m_oSettings6.IsSwipeNavigationEnabled
    Else
        IsSwipeNavigationEnabled = (m_lDeferredValues And DEF_SWIPE) <> 0
    End If
End Property

Public Property Let IsSwipeNavigationEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsSwipeNavigationEnabled = -bValue
    Else
        pvSetDeferredBool DEF_SWIPE, bValue
    End If
End Property

Public Property Get IsMuted() As Boolean
    If Not m_oWebView11 Is Nothing Then
        IsMuted = m_oWebView11.IsMuted
    End If
End Property

Public Property Let IsMuted(ByVal bValue As Boolean)
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.IsMuted = -bValue
    End If
End Property

Public Property Get StatusBarText() As String
    If Not m_oWebView13 Is Nothing Then
        StatusBarText = pvToString(m_oWebView13.StatusBarText)
    End If
End Property

Public Property Get ContainsFullScreenElement() As Boolean
    If Not m_oWebView Is Nothing Then
        ContainsFullScreenElement = m_oWebView.ContainsFullScreenElement
    End If
End Property

Public Property Get IsSuspended() As Boolean
    If Not m_oWebView4 Is Nothing Then
        IsSuspended = m_oWebView4.IsSuspended
    End If
End Property

Public Property Get TimeOutSeconds() As Double
    TimeOutSeconds = m_dblTimeOutSeconds
End Property

Public Property Let TimeOutSeconds(ByVal dblValue As Double)
    m_dblTimeOutSeconds = dblValue
End Property

Public Property Get URL() As String
Attribute URL.VB_UserMemId = 0
    URL = m_sInitialURL
End Property

Public Property Let URL(ByVal sValue As String)
    m_sInitialURL = sValue
    If Not m_oWebView Is Nothing Then
        Navigate sValue
    End If
End Property

Public Property Get UserDataFolder() As String
    UserDataFolder = m_sUserDataFolder
End Property

Public Property Let UserDataFolder(ByVal sValue As String)
    m_sUserDataFolder = sValue
End Property

'=========================================================================
' Public methods
'=========================================================================

Public Function Init( _
        Optional ByVal UserDataFolder As String, _
        Optional ByVal AdditionalBrowserArguments As String, _
        Optional ByVal SecondsToWaitForInitComplete As Double = 8) As Long
    Dim hResult         As Long

    On Error GoTo EH
    If m_bInitComplete And Not m_oWebView Is Nothing Then
        Exit Function
    End If
    m_bShuttingDown = False
    m_bInitComplete = False
    m_hInitError = 0
    If LenB(UserDataFolder) <> 0 Then
        m_sUserDataFolder = UserDataFolder
    End If
    '--- create callback for environment options + handler
    Set m_oEventSink = New czWebviewCallback
    m_oEventSink.InitOwner Me
    If LenB(AdditionalBrowserArguments) <> 0 Then
        m_oEventSink.AdditionalBrowserArguments = AdditionalBrowserArguments
    End If
    hResult = pvCreateWebView2Environment(vbNullString, m_sUserDataFolder, m_oEventSink, m_oEventSink)
    If hResult <> 0 Then
        Init = hResult
        RaiseEvent InitComplete(False, hResult)
        Exit Function
    End If
    If Not pvPumpUntil(m_bInitComplete, SecondsToWaitForInitComplete) Then
        Init = RPC_E_TIMEOUT
        RaiseEvent InitComplete(False, RPC_E_TIMEOUT)
        Exit Function
    End If
    Init = m_hInitError
    Exit Function
EH:
    Init = Err.Number
    Debug.Print "czWebview.Init error: " & Err.Description
End Function

Public Sub Navigate(ByVal URI As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        m_oWebView.Navigate StrPtr(URI)
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.Navigate error: " & Err.Description
End Sub

Public Sub NavigateToString(ByVal sHTMLContent As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        m_oWebView.NavigateToString StrPtr(sHTMLContent)
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.NavigateToString error: " & Err.Description
End Sub

Public Function NavigateSync(ByVal URI As String, Optional ByVal SecondsTimeout As Double = 8) As Boolean
    On Error GoTo EH
    If m_oWebView Is Nothing Then Exit Function
    m_bNavDone = False
    m_bNavSuccess = True
    m_oWebView.Navigate StrPtr(URI)
    If pvPumpUntil(m_bNavDone, SecondsTimeout) Then
        NavigateSync = m_bNavSuccess
    End If
    Exit Function
EH:
    Debug.Print "czWebview.NavigateSync error: " & Err.Description
End Function

Public Sub Reload()
    If Not m_oWebView Is Nothing Then m_oWebView.Reload
End Sub

Public Sub GoBack()
    If Not m_oWebView Is Nothing Then m_oWebView.GoBack
End Sub

Public Sub GoForward()
    If Not m_oWebView Is Nothing Then m_oWebView.GoForward
End Sub

Public Sub StopNavigation()
    If Not m_oWebView Is Nothing Then m_oWebView.Stop
End Sub

Public Sub OpenDevToolsWindow()
    If Not m_oWebView Is Nothing Then m_oWebView.OpenDevToolsWindow
End Sub

Public Sub OpenTaskManagerWindow()
    On Error GoTo EH
    If Not m_oWebView11 Is Nothing Then m_oWebView11.OpenTaskManagerWindow
    Exit Sub
EH:
    Debug.Print "czWebview.OpenTaskManagerWindow error: " & Err.Description
End Sub

Public Function ExecuteScript(ByVal sScript As String, Optional ByVal SecondsTimeout As Double = -1) As String
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView Is Nothing Then Exit Function
    If SecondsTimeout <= 0 Then SecondsTimeout = m_dblTimeOutSeconds
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    m_oWebView.ExecuteScript StrPtr(sScript), oHandler
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) >= 0 Then
        ExecuteScript = vItem(1)
    End If
    Exit Function
EH:
    Debug.Print "czWebview.ExecuteScript error: " & Err.Description
End Function

Public Sub ExecuteScriptAsync(ByVal sScript As String)
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency

    On Error GoTo EH
    If m_oWebView Is Nothing Then Exit Sub
    cToken = pvNextToken()
    Set oHandler = pvNewHandler(cToken)
    m_oWebView.ExecuteScript StrPtr(sScript), oHandler
    m_cScriptResults.Add True, "A" & cToken
    Exit Sub
EH:
    Debug.Print "czWebview.ExecuteScriptAsync error: " & Err.Description
End Sub

Public Sub AddScriptToExecuteOnDocumentCreated(ByVal sScript As String)
    Dim oHandler        As czWebviewCallback
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        Set oHandler = pvNewHandler()
        m_oWebView.AddScriptToExecuteOnDocumentCreated StrPtr(sScript), oHandler
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.AddScriptToExecuteOnDocumentCreated error: " & Err.Description
End Sub

Public Sub PostWebMessageAsString(ByVal sMessage As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then m_oWebView.PostWebMessageAsString StrPtr(sMessage)
    Exit Sub
EH:
    Debug.Print "czWebview.PostWebMessageAsString error: " & Err.Description
End Sub

Public Sub PostWebMessageAsJSON(ByVal sJson As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then m_oWebView.PostWebMessageAsJSON StrPtr(sJson)
    Exit Sub
EH:
    Debug.Print "czWebview.PostWebMessageAsJSON error: " & Err.Description
End Sub

Public Sub AddHostObjectToScript(ByVal sName As String, ByVal obj As Object)
    Dim vObj As Variant
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        Set vObj = obj
        m_oWebView.AddHostObjectToScript StrPtr(sName), vObj
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.AddHostObjectToScript error: " & Err.Description
End Sub

Public Sub RemoveHostObjectFromScript(ByVal sName As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then m_oWebView.RemoveHostObjectFromScript StrPtr(sName)
    Exit Sub
EH:
    Debug.Print "czWebview.RemoveHostObjectFromScript error: " & Err.Description
End Sub

Public Sub AddWebResourceRequestedFilter(ByVal sFilter As String, Optional ByVal FilterContext As czWebView2ResourceFilter = czFilter_ALL)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then m_oWebView.AddWebResourceRequestedFilter StrPtr(sFilter), FilterContext
    Exit Sub
EH:
    Debug.Print "czWebview.AddWebResourceRequestedFilter error: " & Err.Description
End Sub

Public Sub RemoveWebResourceRequestedFilter(ByVal sFilter As String, ByVal FilterContext As czWebView2ResourceFilter)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then m_oWebView.RemoveWebResourceRequestedFilter StrPtr(sFilter), FilterContext
    Exit Sub
EH:
    Debug.Print "czWebview.RemoveWebResourceRequestedFilter error: " & Err.Description
End Sub

'=========================================================================
' SetWebResourceResponse — call from within WebResourceRequested event
' to serve custom content (e.g. from czStorage byte arrays).
'
' Usage in event handler:
'   Sub czWebview1_WebResourceRequested(URI, ResourceContext, Handled)
'       czWebview1.SetWebResourceResponse 200, "text/html", bMyData()
'       Handled = True
'   End Sub
'=========================================================================
Public Sub SetWebResourceResponse(ByVal lStatusCode As Long, ByVal sMIMEType As String, bData() As Byte)
    Dim oStream     As IVBStream
    Dim oResponse   As IVBCoreWebView2WebResourceResponse
    Dim hGlobal     As Long
    Dim pMem        As Long
    Dim lSize       As Long
    Dim sHeaders    As String
    Dim sReason     As String
    On Error GoTo EH
    If m_oCurrentResourceArgs Is Nothing Then
        Err.Raise 5, "czWebview", "SetWebResourceResponse can only be called from within the WebResourceRequested event handler"
    End If
    '--- Build IStream from byte array
    lSize = UBound(bData) - LBound(bData) + 1
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, lSize)
    If hGlobal = 0 Then Err.Raise 7  '--- Out of memory
    pMem = GlobalLock(hGlobal)
    CopyMemory ByVal pMem, bData(LBound(bData)), lSize
    GlobalUnlock hGlobal
    CreateStreamOnHGlobal hGlobal, 1, oStream   '--- fDeleteOnRelease=True, stream owns the HGLOBAL
    '--- Build response
    sHeaders = "Content-Type: " & sMIMEType
    If lStatusCode >= 200 And lStatusCode < 300 Then sReason = "OK" Else sReason = "Error"
    Set oResponse = m_oEnvironment.CreateWebResourceResponse(oStream, lStatusCode, StrPtr(sReason), StrPtr(sHeaders))
    m_oCurrentResourceArgs.Response = oResponse
    Exit Sub
EH:
    Debug.Print "czWebview.SetWebResourceResponse error: " & Err.Description
End Sub

Public Sub SetWebResourceResponseString(ByVal lStatusCode As Long, ByVal sMIMEType As String, ByVal sContent As String)
    Dim bUTF8() As Byte
    bUTF8 = StrConv(sContent, vbFromUnicode)
    SetWebResourceResponse lStatusCode, sMIMEType & "; charset=utf-8", bUTF8
End Sub

Public Sub SetVirtualHostNameToFolderMapping(ByVal HostName As String, ByVal FolderPath As String, Optional ByVal AccessKind As czWebView2HostResourceAccessKind = czHostResourceAccess_DENY)
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then m_oWebView4.SetVirtualHostNameToFolderMapping StrPtr(HostName), StrPtr(FolderPath), AccessKind
    Exit Sub
EH:
    Debug.Print "czWebview.SetVirtualHostNameToFolderMapping error: " & Err.Description
End Sub

Public Sub ClearVirtualHostNameToFolderMapping(ByVal HostName As String)
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then m_oWebView4.ClearVirtualHostNameToFolderMapping StrPtr(HostName)
    Exit Sub
EH:
    Debug.Print "czWebview.ClearVirtualHostNameToFolderMapping error: " & Err.Description
End Sub

Public Sub ShowPrintUI(Optional ByVal PrintDialogKind As czWebView2PrintDialogKind = czPrintDialog_BROWSER)
    On Error GoTo EH
    If Not m_oWebView16 Is Nothing Then m_oWebView16.ShowPrintUI PrintDialogKind
    Exit Sub
EH:
    Debug.Print "czWebview.ShowPrintUI error: " & Err.Description
End Sub

Public Function CapturePreview(Optional ByVal ImageFormat As czWebView2ImageCaptureFormat = czCaptureAs_PNG) As Byte()
    Dim oHandler        As czWebviewCallback
    Dim oStream         As IVBStream
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim cPos            As Currency
    Dim lSize           As Long
    Dim lRead           As Long
    Dim baResult()      As Byte

    On Error GoTo EH
    If m_oWebView Is Nothing Then Exit Function
    If CreateStreamOnHGlobal(0, 1, oStream) < 0 Then Exit Function
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    m_oWebView.CapturePreview ImageFormat, oStream, oHandler
    If m_dblTimeOutSeconds <= 0 Then Exit Function
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dblTimeOutSeconds) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then Exit Function
    oStream.Seek 0, STREAM_SEEK_END, cPos
    If cPos * 10000@ > 2147483647@ Then Exit Function
    lSize = CLng(cPos * 10000@)
    If lSize > 0 Then
        oStream.Seek 0, STREAM_SEEK_SET, cPos
        ReDim baResult(0 To lSize - 1)
        oStream.Read VarPtr(baResult(0)), lSize, lRead
        CapturePreview = baResult
    End If
    Exit Function
EH:
    Debug.Print "czWebview.CapturePreview error: " & Err.Description
End Function

Public Function PrintToPdf(ByVal ResultFilePath As String, Optional ByVal SecondsTimeout As Double = 8) As Boolean
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView11 Is Nothing Then Exit Function
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    m_oWebView11.PrintToPdf StrPtr(ResultFilePath), Nothing, oHandler
    If SecondsTimeout <= 0 Then
        PrintToPdf = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then Exit Function
    PrintToPdf = (vItem(1) = "true")
    Exit Function
EH:
    Debug.Print "czWebview.PrintToPdf error: " & Err.Description
End Function

Public Function CallDevToolsProtocolMethod(ByVal MethodName As String, ByVal ParamsAsJSON As String) As String
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView Is Nothing Then Exit Function
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    m_oWebView.CallDevToolsProtocolMethod StrPtr(MethodName), StrPtr(ParamsAsJSON), oHandler
    If m_dblTimeOutSeconds <= 0 Then Exit Function
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dblTimeOutSeconds) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) >= 0 Then
        CallDevToolsProtocolMethod = vItem(1)
    End If
    Exit Function
EH:
    Debug.Print "czWebview.CallDevToolsProtocolMethod error: " & Err.Description
End Function

Public Function GetCookies(Optional ByVal URI As String, Optional ByVal SecondsTimeout As Double = 8) As Collection
    Dim oManager        As IVBCoreWebView2CookieManager
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If oManager Is Nothing Then Exit Function
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    oManager.GetCookies StrPtr(URI), oHandler
    If SecondsTimeout <= 0 Then Exit Function
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then Exit Function
    Set GetCookies = vItem(1)
    Exit Function
EH:
    Debug.Print "czWebview.GetCookies error: " & Err.Description
End Function

Public Sub AddOrUpdateCookie(ByVal Name As String, ByVal Value As String, ByVal Domain As String, Optional ByVal Path As String = "/", Optional ByVal Expires As Date, Optional ByVal IsSecure As Boolean, Optional ByVal IsHttpOnly As Boolean, Optional ByVal SameSite As czWebView2CookieSameSiteKind = czSameSite_LAX)
    Dim oManager        As IVBCoreWebView2CookieManager
    Dim oCookie         As IVBCoreWebView2Cookie

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If oManager Is Nothing Then Exit Sub
    Set oCookie = oManager.CreateCookie(StrPtr(Name), StrPtr(Value), StrPtr(Domain), StrPtr(Path))
    If Expires <> 0 Then
        oCookie.Expires = (Expires - #1/1/1970#) * SECS_PER_DAY
    End If
    oCookie.IsSecure = -IsSecure
    oCookie.IsHttpOnly = -IsHttpOnly
    oCookie.SameSite = SameSite
    oManager.AddOrUpdateCookie oCookie
    Exit Sub
EH:
    Debug.Print "czWebview.AddOrUpdateCookie error: " & Err.Description
End Sub

Public Sub DeleteCookies(ByVal Name As String, ByVal URI As String)
    Dim oManager        As IVBCoreWebView2CookieManager
    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If Not oManager Is Nothing Then oManager.DeleteCookies StrPtr(Name), StrPtr(URI)
    Exit Sub
EH:
    Debug.Print "czWebview.DeleteCookies error: " & Err.Description
End Sub

Public Sub DeleteAllCookies()
    Dim oManager        As IVBCoreWebView2CookieManager
    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If Not oManager Is Nothing Then oManager.DeleteAllCookies
    Exit Sub
EH:
    Debug.Print "czWebview.DeleteAllCookies error: " & Err.Description
End Sub

Public Function TrySuspend(Optional ByVal SecondsTimeout As Double = 8) As Boolean
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView4 Is Nothing Or m_oController Is Nothing Then Exit Function
    m_oController.IsVisible = 0
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    m_oWebView4.TrySuspend oHandler
    If SecondsTimeout <= 0 Then
        TrySuspend = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then Exit Function
    TrySuspend = (vItem(1) = "true")
    Exit Function
EH:
    Debug.Print "czWebview.TrySuspend error: " & Err.Description
End Function

Public Sub ResumeFromSuspend()
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then m_oWebView4.ResumeFromSuspend
    If Not m_oController Is Nothing Then m_oController.IsVisible = 1
    Exit Sub
EH:
    Debug.Print "czWebview.ResumeFromSuspend error: " & Err.Description
End Sub

Public Function ClearBrowsingData(ByVal DataKinds As czWebView2BrowsingDataKinds, Optional ByVal SecondsTimeout As Double = 8) As Boolean
    Dim oProfile        As IVBCoreWebView2Profile7
    Dim oHandler        As czWebviewCallback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    Set oProfile = pvGetProfile4()
    If oProfile Is Nothing Then Exit Function
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = pvNewHandler(cToken)
    oProfile.ClearBrowsingData DataKinds, oHandler
    If SecondsTimeout <= 0 Then
        ClearBrowsingData = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    ClearBrowsingData = (vItem(0) >= 0)
    Exit Function
EH:
    Debug.Print "czWebview.ClearBrowsingData error: " & Err.Description
End Function

Public Sub SetFocus(Optional ByVal Reason As czWebView2FocusReason = czFocusReason_PROGRAMMATIC)
    On Error GoTo EH
    If Not m_oController Is Nothing Then m_oController.MoveFocus Reason
    Exit Sub
EH:
    Debug.Print "czWebview.SetFocus error: " & Err.Description
End Sub

Public Sub Shutdown()
    On Error Resume Next
    m_bShuttingDown = True
    '--- Terminate all pending async callbacks
    Dim i As Long
    Dim oCb As czWebviewCallback
    For i = 1 To m_cPendingCallbacks.Count
        Set oCb = m_cPendingCallbacks(i)
        oCb.Terminate
    Next i
    Set m_cPendingCallbacks = New Collection
    Set m_cScriptResults = New Collection
    '--- Release WebView2 objects
    Set m_oWebView = Nothing
    Set m_oWebView4 = Nothing
    Set m_oWebView11 = Nothing
    Set m_oWebView13 = Nothing
    Set m_oWebView16 = Nothing
    If Not m_oController Is Nothing Then
        m_oController.Close
        Set m_oController = Nothing
    End If
    Set m_oEnvironment = Nothing
    If Not m_oEventSink Is Nothing Then
        m_oEventSink.Terminate
        Set m_oEventSink = Nothing
    End If
    If Not m_oFocusGotSink Is Nothing Then
        m_oFocusGotSink.Terminate
        Set m_oFocusGotSink = Nothing
    End If
    If Not m_oFocusLostSink Is Nothing Then
        m_oFocusLostSink.Terminate
        Set m_oFocusLostSink = Nothing
    End If
    Set m_oSettings = Nothing
    Set m_oSettings6 = Nothing
End Sub

'=========================================================================
' UserControl events
'=========================================================================

Private Sub UserControl_Initialize()
    m_dblTimeOutSeconds = 8
    Set m_cScriptResults = New Collection
    Set m_cPendingCallbacks = New Collection
    m_sInitialURL = "about:blank"
    m_lDeferredValues = DEF_WV2_DEFAULTS
End Sub

Private Sub UserControl_InitProperties()
    m_sInitialURL = "about:blank"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_sInitialURL = PropBag.ReadProperty("URL", "about:blank")
    m_sUserDataFolder = PropBag.ReadProperty("UserDataFolder", "")
    m_lDeferredMask = PropBag.ReadProperty("DeferredMask", 0)
    m_lDeferredValues = PropBag.ReadProperty("DeferredValues", DEF_WV2_DEFAULTS)
    m_sDeferredUserAgent = PropBag.ReadProperty("DeferredUserAgent", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "URL", m_sInitialURL, "about:blank"
    PropBag.WriteProperty "UserDataFolder", m_sUserDataFolder, ""
    PropBag.WriteProperty "DeferredMask", m_lDeferredMask, 0
    PropBag.WriteProperty "DeferredValues", m_lDeferredValues, DEF_WV2_DEFAULTS
    PropBag.WriteProperty "DeferredUserAgent", m_sDeferredUserAgent, ""
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        If m_oWebView Is Nothing And Not m_bShuttingDown Then
            Dim hResult As Long
            hResult = Init(m_sUserDataFolder)
            If hResult = 0 And m_sInitialURL <> "about:blank" And LenB(m_sInitialURL) <> 0 Then
                Navigate m_sInitialURL
            End If
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    pvSyncSizeToHost
End Sub

Private Sub UserControl_Terminate()
    Shutdown
End Sub

'=========================================================================
' Internal methods
'=========================================================================

Private Sub pvSyncSizeToHost()
    Dim uRect As APIRECT
    On Error GoTo EH
    If Not m_oController Is Nothing Then
        Call GetClientRect(UserControl.hWnd, uRect)
        m_oController.SetBounds 0, 0, uRect.Right, uRect.Bottom
        m_oController.NotifyParentWindowPositionChanged
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.pvSyncSizeToHost error: " & Err.Description
End Sub

Private Function pvCreateWebView2Environment( _
            sBrowserExecutableFolder As String, _
            ByVal sUserDataFolder As String, _
            oOptions As IVBWebView2EnvironmentOptions, _
            oHandler As IVBWebView2CreateEnvironmentCompletedHandler) As Long
    On Error GoTo EH
    If Not pvEnsureLoader() Then
        pvCreateWebView2Environment = E_MOD_NOT_FOUND
        Exit Function
    End If
    If LenB(sUserDataFolder) = 0 Then
        sUserDataFolder = Environ$("LOCALAPPDATA") & "\" & UCase$(App.EXEName)
    End If
    pvCreateWebView2Environment = CreateCoreWebView2EnvironmentWithOptions(StrPtr(sBrowserExecutableFolder), StrPtr(sUserDataFolder), ObjPtr(oOptions), ObjPtr(oHandler))
    Exit Function
EH:
    pvCreateWebView2Environment = Err.Number
    Debug.Print "czWebview.pvCreateWebView2Environment error: " & Err.Description
End Function

Private Function pvPumpUntil(bDone As Boolean, ByVal dblTimeout As Double) As Boolean
    Dim dblTimer As Double
    On Error GoTo EH
    If dblTimeout <= 0 Then
        pvPumpUntil = True
        Exit Function
    End If
    dblTimer = pvTimerEx + dblTimeout
    Do While pvTimerEx <= dblTimer
        If m_bShuttingDown Then Exit Do
        pvSpinMessagePump Flush:=True
        If bDone Then
            pvPumpUntil = True
            Exit Do
        End If
    Loop
    Exit Function
EH:
    Debug.Print "czWebview.pvPumpUntil error: " & Err.Description
End Function

Private Function pvPumpUntilKey(cCol As Collection, sKey As String, ByVal dblTimeout As Double) As Boolean
    Dim dblTimer As Double
    On Error GoTo EH
    dblTimer = pvTimerEx + dblTimeout
    Do While pvTimerEx <= dblTimer
        If m_bShuttingDown Then Exit Do
        pvSpinMessagePump Flush:=True
        If pvSearchCollection(cCol, sKey) Then
            pvPumpUntilKey = True
            Exit Do
        End If
    Loop
    Exit Function
EH:
    Debug.Print "czWebview.pvPumpUntilKey error: " & Err.Description
End Function

Private Sub pvSpinMessagePump(Optional ByVal Flush As Boolean, Optional ByVal hWndFilter As Long, Optional ByVal FromMsg As Long, Optional ByVal ToMsg As Long)
    Dim uMsg As APIMSG
    If Flush Then Call MsgWaitForMultipleObjects(0, ByVal 0&, 0, 10, QS_ALLINPUT)
    Call CoWaitForMultipleHandles(0, 0, 0, 0, 0)
    Do While PeekMessage(VarPtr(uMsg), hWndFilter, FromMsg, ToMsg, PM_REMOVE) <> 0
        Call TranslateMessage(VarPtr(uMsg))
        If Flush And (uMsg.lMessage >= WM_KEYFIRST And uMsg.lMessage <= WM_KEYLAST _
                        Or uMsg.lMessage >= WM_MOUSEFIRST And uMsg.lMessage <= WM_MOUSELAST _
                        Or uMsg.lMessage >= WM_NCMOUSEFIRST And uMsg.lMessage <= WM_NCMOUSELAST) Then
            '--- flush input during blocking wait
        Else
            Call DispatchMessage(VarPtr(uMsg))
        End If
        Call CoWaitForMultipleHandles(0, 0, 0, 0, 0)
    Loop
End Sub

Private Property Get pvTimerEx() As Double
    Dim cFreq As Currency
    Dim cValue As Currency
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    pvTimerEx = cValue / cFreq
End Property

Private Function pvToString(ByVal lPtr As Long) As String
    Call SysReAllocString(VarPtr(pvToString), lPtr)
    Call CoTaskMemFree(lPtr)
End Function

Private Function pvNextToken() As Currency
    m_cTokenSeq = m_cTokenSeq + 1
    pvNextToken = m_cTokenSeq
End Function

Private Function pvEnsureLoader() As Boolean
    Dim sLoader As String
    If GetModuleHandle(StrPtr("WebView2Loader.dll")) <> 0 Then
        pvEnsureLoader = True
        Exit Function
    End If
    '--- Search: App.Path, App.Path\External, then walk up parent dirs
    If pvFileExists(App.Path & "\WebView2Loader.dll") Then
        sLoader = App.Path & "\WebView2Loader.dll"
    Else
        sLoader = pvLocateFile(App.Path & "\External\WebView2Loader.dll")
    End If
    If LenB(sLoader) <> 0 Then
        pvEnsureLoader = (LoadLibrary(StrPtr(sLoader)) <> 0)
    End If
End Function

Private Function pvLocateFile(sFile As String) As String
    Dim sDir As String
    Dim sName As String
    Dim lPos As Long
    If InStrRev(sFile, "\") > 0 Then
        sDir = Left$(sFile, InStrRev(sFile, "\"))
        sName = Mid$(sFile, InStrRev(sFile, "\") + 1)
        Do While Not pvFileExists(sDir & sName)
            If Len(sDir) > 1 Then
                lPos = InStrRev(sDir, "\", Len(sDir) - 1)
                If lPos > 0 Then
                    sDir = Left$(sDir, lPos)
                    If Left$(sDir, 2) = "\\" And InStrRev(sDir, "\", Len(sDir) - 1) <= 2 Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Loop
        pvLocateFile = sDir & sName
    ElseIf pvFileExists(sFile) Then
        pvLocateFile = sFile
    End If
End Function

Private Function pvFileExists(sFile As String) As Boolean
    If GetFileAttributes(StrPtr(sFile)) = -1 Then
        pvFileExists = (Err.LastDllError = 32)
    Else
        pvFileExists = True
    End If
End Function

Private Function pvSearchCollection(cCol As Collection, sKey As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = cCol.Item(sKey)
    pvSearchCollection = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function pvNewHandler(Optional ByVal cToken As Currency) As czWebviewCallback
    Set pvNewHandler = New czWebviewCallback
    pvNewHandler.InitOwner Me, cToken
    m_cPendingCallbacks.Add pvNewHandler
End Function

Private Sub pvSetDeferredBool(ByVal lBit As Long, ByVal bValue As Boolean)
    If bValue Then
        m_lDeferredValues = m_lDeferredValues Or lBit
    Else
        m_lDeferredValues = m_lDeferredValues And Not lBit
    End If
    m_lDeferredMask = m_lDeferredMask Or lBit
End Sub

Private Sub pvApplyDeferredSettings()
    '--- Apply deferred boolean settings (bitmask)
    If m_lDeferredMask <> 0 And Not m_oSettings Is Nothing Then
        If (m_lDeferredMask And DEF_SCRIPT) <> 0 Then m_oSettings.IsScriptEnabled = -(CBool(m_lDeferredValues And DEF_SCRIPT))
        If (m_lDeferredMask And DEF_WEBMSG) <> 0 Then m_oSettings.IsWebMessageEnabled = -(CBool(m_lDeferredValues And DEF_WEBMSG))
        If (m_lDeferredMask And DEF_DIALOGS) <> 0 Then m_oSettings.AreDefaultScriptDialogsEnabled = -(CBool(m_lDeferredValues And DEF_DIALOGS))
        If (m_lDeferredMask And DEF_STATUS) <> 0 Then m_oSettings.IsStatusBarEnabled = -(CBool(m_lDeferredValues And DEF_STATUS))
        If (m_lDeferredMask And DEF_DEVTOOLS) <> 0 Then m_oSettings.AreDevToolsEnabled = -(CBool(m_lDeferredValues And DEF_DEVTOOLS))
        If (m_lDeferredMask And DEF_CTXMENU) <> 0 Then m_oSettings.AreDefaultContextMenusEnabled = -(CBool(m_lDeferredValues And DEF_CTXMENU))
        If (m_lDeferredMask And DEF_HOSTOBJ) <> 0 Then m_oSettings.AreHostObjectsAllowed = -(CBool(m_lDeferredValues And DEF_HOSTOBJ))
        If (m_lDeferredMask And DEF_ZOOM) <> 0 Then m_oSettings.IsZoomControlEnabled = -(CBool(m_lDeferredValues And DEF_ZOOM))
        If (m_lDeferredMask And DEF_ERRORPAGE) <> 0 Then m_oSettings.IsBuiltInErrorPageEnabled = -(CBool(m_lDeferredValues And DEF_ERRORPAGE))
    End If
    '--- Apply deferred Settings6 booleans
    If m_lDeferredMask <> 0 And Not m_oSettings6 Is Nothing Then
        If (m_lDeferredMask And DEF_ACCELKEYS) <> 0 Then m_oSettings6.AreBrowserAcceleratorKeysEnabled = -(CBool(m_lDeferredValues And DEF_ACCELKEYS))
        If (m_lDeferredMask And DEF_PWDSAVE) <> 0 Then m_oSettings6.IsPasswordAutosaveEnabled = -(CBool(m_lDeferredValues And DEF_PWDSAVE))
        If (m_lDeferredMask And DEF_AUTOFILL) <> 0 Then m_oSettings6.IsGeneralAutofillEnabled = -(CBool(m_lDeferredValues And DEF_AUTOFILL))
        If (m_lDeferredMask And DEF_PINCH) <> 0 Then m_oSettings6.IsPinchZoomEnabled = -(CBool(m_lDeferredValues And DEF_PINCH))
        If (m_lDeferredMask And DEF_SWIPE) <> 0 Then m_oSettings6.IsSwipeNavigationEnabled = -(CBool(m_lDeferredValues And DEF_SWIPE))
    End If
    '--- Apply deferred UserAgent
    If LenB(m_sDeferredUserAgent) <> 0 And Not m_oSettings6 Is Nothing Then
        m_oSettings6.UserAgent = StrPtr(m_sDeferredUserAgent)
        m_sDeferredUserAgent = vbNullString
    End If
    m_lDeferredMask = 0
End Sub

Private Function pvGetCookieManager() As IVBCoreWebView2CookieManager
    On Error Resume Next
    If Not m_oWebView4 Is Nothing Then
        Set pvGetCookieManager = m_oWebView4.CookieManager
    End If
End Function

Private Function pvGetProfile4() As IVBCoreWebView2Profile7
    On Error Resume Next
    If Not m_oWebView13 Is Nothing Then
        Set pvGetProfile4 = m_oWebView13.Profile
    End If
End Function

'=========================================================================
' Friend sinks — called by czWebviewCallback
'=========================================================================

Friend Sub OnEnvironmentCreated(ByVal hResult As Long, ByVal oEnv As IVBCoreWebView2Environment)
    On Error GoTo EH
    If hResult <> 0 Then
        m_hInitError = hResult
        m_bInitComplete = True
        Exit Sub
    End If
    Set m_oEnvironment = oEnv
    '--- Create controller, using the event sink as handler
    m_oEnvironment.CreateCoreWebView2Controller UserControl.hWnd, m_oEventSink
    Exit Sub
EH:
    m_hInitError = Err.Number
    m_bInitComplete = True
    Debug.Print "czWebview.OnEnvironmentCreated error: " & Err.Description
End Sub

Friend Sub OnControllerCreated(ByVal hResult As Long, ByVal oCtrl As IVBCoreWebView2Controller)
    On Error GoTo EH
    If hResult <> 0 Then
        m_hInitError = hResult
        m_bInitComplete = True
        RaiseEvent InitComplete(False, hResult)
        Exit Sub
    End If
    Set m_oController = oCtrl
    Set m_oWebView = m_oController.CoreWebView2
    Set m_oSettings = m_oWebView.Settings
    '--- Newer interfaces may be unavailable on older runtimes
    On Error Resume Next
    Set m_oSettings6 = m_oSettings
    Set m_oWebView4 = m_oWebView
    Set m_oWebView11 = m_oWebView
    Set m_oWebView13 = m_oWebView
    Set m_oWebView16 = m_oWebView
    On Error GoTo EH
    '--- Register ALL events via the event sink callback
    m_oWebView.add_NavigationStarting m_oEventSink
    m_oWebView.add_NavigationCompleted m_oEventSink
    m_oWebView.add_SourceChanged m_oEventSink
    m_oWebView.add_ProcessFailed m_oEventSink
    m_oWebView.add_WebMessageReceived m_oEventSink
    m_oWebView.add_ScriptDialogOpening m_oEventSink
    m_oWebView.add_PermissionRequested m_oEventSink
    m_oWebView.add_NewWindowRequested m_oEventSink
    m_oWebView.add_WebResourceRequested m_oEventSink
    m_oWebView.add_DocumentTitleChanged m_oEventSink
    m_oWebView.add_ContentLoading m_oEventSink
    m_oWebView.add_HistoryChanged m_oEventSink
    m_oWebView.add_WindowCloseRequested m_oEventSink
    m_oWebView.add_ContainsFullScreenElementChanged m_oEventSink
    m_oController.add_AcceleratorKeyPressed m_oEventSink
    m_oController.add_ZoomFactorChanged m_oEventSink
    m_oController.add_MoveFocusRequested m_oEventSink
    '--- GotFocus/LostFocus need separate callback instances for discrimination
    Set m_oFocusGotSink = New czWebviewCallback
    m_oFocusGotSink.InitOwner Me, 1
    m_oController.add_GotFocus m_oFocusGotSink
    Set m_oFocusLostSink = New czWebviewCallback
    m_oFocusLostSink.InitOwner Me, 2
    m_oController.add_LostFocus m_oFocusLostSink
    '--- Extended event subscriptions (may fail on older runtimes)
    On Error Resume Next
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.add_DownloadStarting m_oEventSink
        m_oWebView4.add_FrameCreated m_oEventSink
        m_oWebView4.add_DOMContentLoaded m_oEventSink
    End If
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.add_ContextMenuRequested m_oEventSink
        m_oWebView11.add_BasicAuthenticationRequested m_oEventSink
    End If
    If Not m_oWebView13 Is Nothing Then
        m_oWebView13.add_StatusBarTextChanged m_oEventSink
    End If
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.add_WebResourceResponseReceived m_oEventSink
    End If
    On Error GoTo EH
    '--- Apply deferred settings
    pvApplyDeferredSettings
    '--- Size and show
    pvSyncSizeToHost
    m_oController.IsVisible = 1
    m_hInitError = 0
    m_bInitComplete = True
    RaiseEvent InitComplete(True, 0)
    Exit Sub
EH:
    m_hInitError = Err.Number
    m_bInitComplete = True
    RaiseEvent InitComplete(False, Err.Number)
    Debug.Print "czWebview.OnControllerCreated error: " & Err.Description
End Sub

Friend Sub OnExecuteScriptCompleted(ByVal hResult As Long, sResultJson As String, ByVal cToken As Currency)
    On Error GoTo EH
    If cToken = 0 Then
        '--- fire-and-forget
    ElseIf pvSearchCollection(m_cScriptResults, "A" & cToken) Then
        '--- async result: raise event
        m_cScriptResults.Remove "A" & cToken
        RaiseEvent ScriptCompleted(hResult, sResultJson, cToken)
    ElseIf pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        '--- sync waiter: store result
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(hResult, sResultJson), "T" & cToken
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.OnExecuteScriptCompleted error: " & Err.Description
End Sub

Friend Sub OnGetCookies(ByVal hResult As Long, ByVal oResult As IVBCoreWebView2CookieList, ByVal cToken As Currency)
    Dim cCookies        As Collection
    Dim i               As Long
    Dim oCookie         As IVBCoreWebView2Cookie
    Dim cItem           As Collection

    On Error GoTo EH
    Set cCookies = New Collection
    If hResult >= 0 And Not oResult Is Nothing Then
        For i = 0 To oResult.Count - 1
            Set oCookie = oResult.GetValueAtIndex(i)
            Set cItem = New Collection
            cItem.Add pvToString(oCookie.Name), "Name"
            cItem.Add pvToString(oCookie.Value), "Value"
            cItem.Add pvToString(oCookie.Domain), "Domain"
            cItem.Add pvToString(oCookie.Path), "Path"
            cItem.Add oCookie.IsSecure <> 0, "IsSecure"
            cItem.Add oCookie.IsHttpOnly <> 0, "IsHttpOnly"
            cItem.Add oCookie.SameSite, "SameSite"
            cCookies.Add cItem
        Next
    End If
    If pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(hResult, cCookies), "T" & cToken
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.OnGetCookies error: " & Err.Description
End Sub

Friend Sub OnPrintToPdfStream(ByVal hResult As Long, ByVal oPdfStream As IVBStream, ByVal cToken As Currency)
    Dim cPos            As Currency
    Dim lSize           As Long
    Dim lRead           As Long
    Dim baResult()      As Byte

    On Error GoTo EH
    If hResult >= 0 And Not oPdfStream Is Nothing Then
        oPdfStream.Seek 0, STREAM_SEEK_END, cPos
        If cPos * 10000@ > 2147483647@ Then GoTo EH
        lSize = CLng(cPos * 10000@)
        If lSize > 0 Then
            oPdfStream.Seek 0, STREAM_SEEK_SET, cPos
            ReDim baResult(0 To lSize - 1)
            oPdfStream.Read VarPtr(baResult(0)), lSize, lRead
        End If
    End If
    If pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(hResult, baResult), "T" & cToken
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.OnPrintToPdfStream error: " & Err.Description
End Sub

Friend Sub OnGetBrowserExtensions(ByVal hResult As Long, ByVal oResult As IVBCoreWebView2BrowserExtensionList, ByVal cToken As Currency)
    On Error GoTo EH
    If pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(hResult, vbNullString), "T" & cToken
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.OnGetBrowserExtensions error: " & Err.Description
End Sub

Friend Sub SinkNavigationStarting(ByVal oArgs As IVBCoreWebView2NavigationStartingEventArgs)
    Dim bCancel As Boolean
    On Error GoTo EH
    bCancel = oArgs.Cancel
    RaiseEvent NavigationStarting(CBool(oArgs.IsUserInitiated), CBool(oArgs.IsRedirected), pvToString(oArgs.URI), bCancel)
    oArgs.Cancel = -bCancel
    Exit Sub
EH:
    Debug.Print "czWebview.SinkNavigationStarting error: " & Err.Description
End Sub

Friend Sub SinkNavigationCompleted(ByVal oArgs As IVBCoreWebView2NavigationCompletedEventArgs)
    On Error GoTo EH
    m_bNavDone = True
    m_bNavSuccess = CBool(oArgs.IsSuccess)
    m_lNavErrorStatus = oArgs.WebErrorStatus
    RaiseEvent NavigationCompleted(m_bNavSuccess, m_lNavErrorStatus)
    If m_bNavSuccess Then RaiseEvent DocumentComplete
    Exit Sub
EH:
    Debug.Print "czWebview.SinkNavigationCompleted error: " & Err.Description
End Sub

Friend Sub SinkSourceChanged()
    RaiseEvent SourceChanged
End Sub

Friend Sub SinkProcessFailed()
    RaiseEvent ProcessFailed
End Sub

Friend Sub SinkFocusChanged(ByVal bGotFocus As Boolean)
    If bGotFocus Then
        RaiseEvent WebViewGotFocus
    Else
        RaiseEvent WebViewLostFocus
    End If
End Sub

Friend Sub SinkWebMessageReceived(ByVal bIsString As Boolean, sMsg As String, sJson As String)
    On Error GoTo EH
    If bIsString Then
        RaiseEvent WebMessageReceived(sMsg, False)
    Else
        RaiseEvent WebMessageReceived(sJson, True)
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.SinkWebMessageReceived error: " & Err.Description
End Sub

Friend Sub SinkScriptDialogOpening(ByVal lKind As Long, bAccept As Boolean, sResultText As String, ByVal sURI As String, ByVal sMessage As String, ByVal sDefaultText As String)
    RaiseEvent ScriptDialogOpening(lKind, bAccept, sResultText, sURI, sMessage, sDefaultText)
End Sub

Friend Sub SinkPermissionRequested(ByVal bIsUser As Long, eState As Long, ByVal sURI As String, ByVal lKind As Long)
    Dim eTyped As czWebView2PermissionState
    eTyped = eState
    RaiseEvent PermissionRequested(bIsUser <> 0, eTyped, sURI, lKind)
    eState = eTyped
End Sub

Friend Sub SinkNewWindowRequested(ByVal bIsUser As Long, bHandled As Boolean, ByVal sURI As String, cFeatures As Collection)
    RaiseEvent NewWindowRequested(bIsUser <> 0, bHandled, sURI, cFeatures)
End Sub

Friend Sub SinkAcceleratorKeyPressed(ByVal lKind As Long, ByVal bExtended As Boolean, ByVal bWasDown As Boolean, ByVal bReleased As Boolean, ByVal bMenuDown As Boolean, ByVal lRepeat As Long, ByVal lScan As Long, bHandled As Boolean)
    RaiseEvent AcceleratorKeyPressed(lKind, bExtended, bWasDown, bReleased, bMenuDown, lRepeat, lScan, bHandled)
End Sub

Friend Sub SinkWebResourceRequested(ByVal oArgs As IVBCoreWebView2WebResourceRequestedEventArgs)
    Dim oRequest As IVBCoreWebView2WebResourceRequest
    Dim bHandled As Boolean
    On Error GoTo EH
    Set oRequest = oArgs.Request
    '--- Store args so SetWebResourceResponse can use them
    Set m_oCurrentResourceArgs = oArgs
    RaiseEvent WebResourceRequested(pvToString(oRequest.URI), oArgs.ResourceContext, bHandled)
    Set m_oCurrentResourceArgs = Nothing
    '--- If not handled by SetWebResourceResponse, block with 403
    If Not bHandled Then
        '--- Legacy permit/deny: if user didn't handle, allow by default
    End If
    Exit Sub
EH:
    Set m_oCurrentResourceArgs = Nothing
    Debug.Print "czWebview.SinkWebResourceRequested error: " & Err.Description
End Sub

Friend Sub SinkWebResourceResponseReceived(ByVal oArgs As IVBCoreWebView2WebResourceResponseReceivedEventArgs)
    Dim oReq As IVBCoreWebView2WebResourceRequest
    Dim oResp As IVBCoreWebView2WebResourceResponseView
    Dim cHeaders As Collection
    Dim oHdrs As IVBCoreWebView2HttpResponseHeaders
    Dim oIter As IVBCoreWebView2HttpHeadersCollectionIterator
    Dim lName As Long
    Dim lValue As Long
    Dim sHdrName As String
    Dim sHdrValue As String
    On Error GoTo EH
    Set oReq = oArgs.Request
    Set oResp = oArgs.Response
    Set cHeaders = New Collection
    '--- Parse response headers
    Set oHdrs = oResp.Headers
    If Not oHdrs Is Nothing Then
        Set oIter = oHdrs.GetIterator()
        If Not oIter Is Nothing Then
            Do While oIter.HasCurrentHeader <> 0
                oIter.GetCurrentHeader lName, lValue
                sHdrName = pvToString(lName)
                sHdrValue = pvToString(lValue)
                If LenB(sHdrName) <> 0 Then
                    On Error Resume Next
                    cHeaders.Add sHdrValue, sHdrName
                    On Error GoTo EH
                End If
                If oIter.MoveNext() = 0 Then Exit Do
            Loop
        End If
    End If
    RaiseEvent WebResourceResponseReceived(pvToString(oReq.URI), pvToString(oReq.Method), oResp.StatusCode, pvToString(oResp.ReasonPhrase), cHeaders)
    Exit Sub
EH:
    Debug.Print "czWebview.SinkWebResourceResponseReceived error: " & Err.Description
End Sub

Friend Sub SinkFrameCreated(ByVal sFrameName As String)
    RaiseEvent FrameCreated(sFrameName)
End Sub

Friend Sub SinkDownloadStarting(ByVal oArgs As IVBCoreWebView2DownloadStartingEventArgs)
    Dim oOper As IVBCoreWebView2DownloadOperation
    Dim sPath As String
    Dim bCancel As Boolean
    Dim sNewPath As String
    On Error GoTo EH
    Set oOper = oArgs.DownloadOperation
    sPath = pvToString(oArgs.ResultFilePath)
    bCancel = CBool(oArgs.Cancel)
    RaiseEvent DownloadStarting(pvToString(oOper.URI), pvToString(oOper.MimeType), sPath, bCancel, sNewPath)
    oArgs.Cancel = -bCancel
    If LenB(sNewPath) <> 0 Then oArgs.ResultFilePath = StrPtr(sNewPath)
    Exit Sub
EH:
    Debug.Print "czWebview.SinkDownloadStarting error: " & Err.Description
End Sub

Friend Sub SinkDocumentTitleChanged()
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        RaiseEvent TitleChanged(pvToString(m_oWebView.DocumentTitle))
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.SinkDocumentTitleChanged error: " & Err.Description
End Sub

Friend Sub SinkContentLoading(ByVal bIsErrorPage As Boolean)
    RaiseEvent ContentLoading(bIsErrorPage)
End Sub

Friend Sub SinkHistoryChanged()
    RaiseEvent HistoryChanged
End Sub

Friend Sub SinkWindowCloseRequested()
    RaiseEvent WindowCloseRequested
End Sub

Friend Sub SinkZoomFactorChanged()
    RaiseEvent ZoomFactorChanged
End Sub

Friend Sub SinkMoveFocusRequested(ByVal lReason As Long, bHandled As Boolean)
    RaiseEvent MoveFocusRequested(lReason, bHandled)
End Sub

Friend Sub SinkContextMenuRequested(ByVal sPageURI As String, ByVal sLinkURI As String, ByVal sSelectionText As String, ByVal lX As Long, ByVal lY As Long, bHandled As Boolean)
    RaiseEvent ContextMenuRequested(sPageURI, sLinkURI, sSelectionText, lX, lY, bHandled)
End Sub

Friend Sub SinkDOMContentLoaded()
    RaiseEvent DOMContentLoaded
End Sub

Friend Sub SinkStatusBarTextChanged()
    On Error GoTo EH
    RaiseEvent StatusBarTextChanged(StatusBarText)
    Exit Sub
EH:
    Debug.Print "czWebview.SinkStatusBarTextChanged error: " & Err.Description
End Sub

Friend Sub SinkContainsFullScreenElementChanged()
    On Error GoTo EH
    RaiseEvent ContainsFullScreenElementChanged(ContainsFullScreenElement)
    Exit Sub
EH:
    Debug.Print "czWebview.SinkContainsFullScreenElementChanged error: " & Err.Description
End Sub

Friend Sub SinkBasicAuthenticationRequested(ByVal oArgs As IVBCoreWebView2BasicAuthenticationRequestedEventArgs)
    Dim oResponse As IVBCoreWebView2BasicAuthenticationResponse
    Dim sUserName As String
    Dim sPassword As String
    Dim bCancel As Boolean
    On Error GoTo EH
    bCancel = CBool(oArgs.Cancel)
    RaiseEvent BasicAuthenticationRequested(pvToString(oArgs.URI), pvToString(oArgs.Challenge), sUserName, sPassword, bCancel)
    oArgs.Cancel = -bCancel
    If LenB(sUserName) <> 0 Or LenB(sPassword) <> 0 Then
        Set oResponse = oArgs.Response
        oResponse.UserName = StrPtr(sUserName)
        oResponse.Password = StrPtr(sPassword)
    End If
    Exit Sub
EH:
    Debug.Print "czWebview.SinkBasicAuthenticationRequested error: " & Err.Description
End Sub
