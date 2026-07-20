VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "czWebview Test Browser"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin TestProject.czWebview czWebview1 
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10398
      DeferredMask    =   48
      DeferredValues  =   15311
   End
   Begin VB.CommandButton cmdDevTools 
      Caption         =   "F12"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3300
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   9660
      TabIndex        =   3
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Text            =   "https://www.github.com/cyberzilla/czWebview"
      Top             =   60
      Width           =   4995
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6480
      Width           =   10200
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' frmTest — czWebview Test Browser
' The czWebview UserControl is placed directly on the form at design-time
' since it's part of the same project. No Controls.Add needed.
'=========================================================================
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    czWebview1.Move 0, 480, ScaleWidth, ScaleHeight - 480 - lblStatus.Height
    lblStatus.Top = ScaleHeight - lblStatus.Height
    lblStatus.Width = ScaleWidth
    txtURL.Width = ScaleWidth - txtURL.Left - cmdGo.Width - cmdDevTools.Width - 180
    cmdGo.Left = txtURL.Left + txtURL.Width + 60
    cmdDevTools.Left = cmdGo.Left + cmdGo.Width + 60
End Sub

Private Sub cmdBack_Click()
    czWebview1.GoBack
End Sub

Private Sub cmdForward_Click()
    czWebview1.GoForward
End Sub

Private Sub cmdRefresh_Click()
    czWebview1.Reload
End Sub

Private Sub cmdStop_Click()
    czWebview1.StopNavigation
End Sub

Private Sub cmdGo_Click()
    Dim sURL As String
    sURL = Trim$(txtURL.Text)
    If LenB(sURL) <> 0 Then
        If InStr(1, sURL, "://") = 0 Then
            sURL = "https://" & sURL
        End If
        czWebview1.Navigate sURL
    End If
End Sub

Private Sub cmdDevTools_Click()
    czWebview1.OpenDevToolsWindow
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdGo_Click
    End If
End Sub

'=========================================================================
' WebView2 events — handled directly on the design-time control
'=========================================================================

Private Sub czWebview1_InitComplete(ByVal Success As Boolean, ByVal ErrorCode As Long)
    If Success Then
        lblStatus.Caption = "WebView2 initialized. Version: " & czWebview1.BrowserVersion
        cmdGo_Click
    Else
        lblStatus.Caption = "WebView2 init FAILED! Error: &H" & Hex$(ErrorCode)
        MsgBox "WebView2 initialization failed." & vbCrLf & _
               "Error code: &H" & Hex$(ErrorCode) & vbCrLf & vbCrLf & _
               "Make sure WebView2Loader.dll is in the app folder" & vbCrLf & _
               "and WebView2 Runtime is installed.", vbCritical
    End If
End Sub

Private Sub czWebview1_NavigationStarting(ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal URI As String, Cancel As Boolean)
    lblStatus.Caption = "Loading: " & URI
    txtURL.Text = URI
End Sub

Private Sub czWebview1_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    If IsSuccess Then
        lblStatus.Caption = "Done"
    Else
        lblStatus.Caption = "Navigation failed. Error: " & WebErrorStatus
    End If
    cmdBack.Enabled = czWebview1.CanGoBack
    cmdForward.Enabled = czWebview1.CanGoForward
End Sub

Private Sub czWebview1_TitleChanged(ByVal Title As String)
    Me.Caption = Title & " - czWebview Test"
End Sub

Private Sub czWebview1_SourceChanged()
    txtURL.Text = czWebview1.DocumentURL
End Sub

Private Sub czWebview1_StatusBarTextChanged(ByVal Text As String)
    If LenB(Text) <> 0 Then lblStatus.Caption = Text
End Sub

Private Sub czWebview1_NewWindowRequested(ByVal IsUserInitiated As Boolean, Handled As Boolean, ByVal URI As String, NewWindowFeatures As Collection)
    czWebview1.Navigate URI
    Handled = True
End Sub

Private Sub czWebview1_WebMessageReceived(ByVal Message As String, ByVal IsJSON As Boolean)
    Debug.Print "WebMessage: " & Message & " (IsJSON=" & IsJSON & ")"
End Sub

Private Sub czWebview1_DownloadStarting(ByVal URI As String, ByVal MimeType As String, ByVal SuggestedPath As String, CancelDownload As Boolean, NewFilePath As String)
    lblStatus.Caption = "Downloading: " & URI
End Sub
