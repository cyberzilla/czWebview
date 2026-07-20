VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czCalc"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin CalculatorWebApp.czWebview czWebview1 
      Height          =   7380
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' frmCalculator — Calculator Web App Demo
' Loads HTML/CSS/JS from the www\ folder and renders inline via
' NavigateToString. Demonstrates czWebview as a native-looking app shell.
'=========================================================================
Option Explicit

Private Sub Form_Load()
    czWebview1.AreDefaultContextMenusEnabled = False
    czWebview1.AreDevToolsEnabled = False
    czWebview1.IsZoomControlEnabled = False
    czWebview1.IsStatusBarEnabled = False
    czWebview1.AreBrowserAcceleratorKeysEnabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    czWebview1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub czWebview1_InitComplete(ByVal Success As Boolean, ByVal ErrorCode As Long)
    If Success Then
        Dim sHTML As String, sCSS As String, sJS As String
        '--- Load from www folder
        sHTML = ReadTextFile(App.Path & "\www\index.html")
        sCSS = ReadTextFile(App.Path & "\www\style.css")
        sJS = ReadTextFile(App.Path & "\www\app.js")
        '--- Inline CSS and JS
        sHTML = Replace(sHTML, "<link rel=""stylesheet"" href=""{{STYLE}}"">", "<style>" & sCSS & "</style>")
        sHTML = Replace(sHTML, "<script src=""{{SCRIPT}}""></script>", "<script>" & sJS & "</script>")
        czWebview1.NavigateToString sHTML
    Else
        MsgBox "WebView2 init failed: &H" & Hex$(ErrorCode), vbCritical
    End If
End Sub

Private Function ReadTextFile(ByVal sPath As String) As String
    Dim f As Integer, bData() As Byte
    If Dir$(sPath) = "" Then Exit Function
    f = FreeFile
    Open sPath For Binary As #f
    ReDim bData(LOF(f) - 1)
    Get #f, , bData
    Close #f
    ReadTextFile = StrConv(bData, vbUnicode)
End Function
