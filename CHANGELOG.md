# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-07-21

### Added
- **Core Navigation**: Navigate, GoBack, GoForward, Reload, Stop, NavigateSync, NavigateToString
- **JavaScript Bridge**: ExecuteScript (sync & async), PostWebMessage (String & JSON), AddHostObjectToScript
- **29 Events**: Full lifecycle control including NavigationStarting, NavigationCompleted, WebMessageReceived, etc.
- **48 Properties**: Fine-grained settings control (IsScriptEnabled, AreDevToolsEnabled, UserAgent, ZoomFactor, etc.)
- **34 Public Methods**: Comprehensive API surface covering navigation, JS execution, capture, cookies, and more
- **44 Callback Interfaces**: Full WebView2 API coverage matching wqweto's cWebView2
- **Print & Capture**: PrintToPdf, ShowPrintUI, CapturePreview (PNG/JPG)
- **Cookie Management**: GetCookies, AddOrUpdateCookie, DeleteCookies, DeleteAllCookies
- **Download Handling**: DownloadStarting event with custom file path support
- **Basic Authentication**: BasicAuthenticationRequested event
- **Virtual Host Mapping**: SetVirtualHostNameToFolderMapping for serving local files
- **Permission Requests**: Camera, microphone, geolocation, notifications
- **Script Dialogs**: Alert, confirm, prompt handling
- **Context Menu Events**: Right-click menu interception
- **Web Resource Filtering**: Request filtering & response interception
- **Full Screen Detection**: ContainsFullScreenElementChanged event
- **Audio Control**: Mute/unmute via IsMuted property
- **User Agent Customization**: Custom UserAgent string
- **Zoom Control**: Programmatic zoom + pinch zoom toggle
- **Suspend/Resume**: Memory optimization for background tabs
- **DevTools Protocol**: Direct Chrome DevTools Protocol method calls
- **Clear Browsing Data**: Cache, history, cookies, passwords, etc.
- **Deferred Settings**: Properties can be set before WebView2 initializes
- **Demo Projects**: SimpleTest browser, WhatsApp Bot, Calculator Web App
- **Type Library**: Pre-compiled czWebview.tlb with compile_tlb.bat auto-build script
- **MIT License**

[1.0.0]: https://github.com/cyberzilla/czWebview/releases/tag/v1.0.0
