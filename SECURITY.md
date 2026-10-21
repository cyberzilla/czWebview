# Security Policy

## Supported Versions

| Version | Supported          |
|---------|--------------------|
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

If you discover a security vulnerability within czWebview, please report it responsibly:

1. **Do NOT open a public issue** for security vulnerabilities
2. **Email**: Send details to the maintainer directly
3. **Include**: Steps to reproduce, potential impact, and suggested fix (if any)

We will acknowledge receipt within 48 hours and provide a timeline for resolution.

## Security Considerations

### WebView2 Runtime
- czWebview relies on the Microsoft Edge WebView2 Runtime for all web content rendering
- WebView2 Runtime is automatically updated by Microsoft with security patches
- Ensure users have the latest WebView2 Runtime installed

### JavaScript Bridge
- The `PostWebMessage` / `WebMessageReceived` bridge allows bidirectional communication between VB6 and web content
- **Always validate messages** received from JavaScript before processing them in VB6
- **Never execute untrusted JavaScript** via `ExecuteScript` without sanitization
- Consider disabling `IsWebMessageEnabled` if your application doesn't need the JS bridge

### Host Objects
- `AddHostObjectToScript` exposes VB6 COM objects directly to JavaScript
- **Only expose objects with minimal, well-defined interfaces**
- **Never expose objects** that can access the file system, registry, or network without restrictions
- Consider disabling `AreHostObjectsAllowed` if not needed

### Web Resource Filtering
- Use `AddWebResourceRequestedFilter` to control which URLs the WebView2 can access
- Block access to sensitive internal resources
- Validate and sanitize all URLs before navigation

### User Data Folder
- WebView2 stores browsing data (cookies, cache, etc.) in the UserDataFolder
- Use a dedicated, application-specific folder path
- Call `ClearBrowsingData` when sensitive data should not persist
- Consider using `ExclusiveUserDataFolderAccess` to prevent other applications from accessing the data

### Recommendations
1. Set `AreDevToolsEnabled = False` in production builds
2. Set `AreDefaultContextMenusEnabled = False` unless needed
3. Use `NavigationStarting` event to validate/block unexpected URLs
4. Implement proper error handling for all WebView2 events
5. Keep WebView2Loader.dll updated to the latest version
