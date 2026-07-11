VERSION 5.00
Begin VB.Form frmWhatsappBot 
   Caption         =   "czWebview WhatsApp Bot"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin WhatsappBot.czWebview czWebview1 
      Height          =   8175
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   10200
   End
   Begin VB.Frame fraBot 
      Caption         =   " Bot Control Panel "
      Height          =   8175
      Left            =   10320
      TabIndex        =   1
      Top             =   60
      Width           =   4020
      Begin VB.Timer tmrAutoReply 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   3360
         Top             =   7560
      End
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "Clear Log"
         Height          =   375
         Left            =   2100
         TabIndex        =   13
         Top             =   7680
         Width           =   1815
      End
      Begin VB.CheckBox chkAutoReply 
         Caption         =   "Enable Auto-Reply Bot"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   3795
      End
      Begin VB.TextBox txtAutoReply 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "Hello! I'm a bot built with VB6 + czWebview. I received your message."
         Top             =   6240
         Width           =   3795
      End
      Begin VB.ListBox lstLog 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   3795
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send Message"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   3060
         Width           =   3795
      End
      Begin VB.TextBox txtMessage 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2100
         Width           =   3795
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   3795
      End
      Begin VB.CommandButton cmdInjectBot 
         Caption         =   "Inject Bot Script"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3795
      End
      Begin VB.CommandButton cmdGetChats 
         Caption         =   "Get Active Chats"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   3795
      End
      Begin VB.Label lblAutoReply 
         Caption         =   "Auto-Reply Template:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   5940
         Width           =   3795
      End
      Begin VB.Label lblLog 
         Caption         =   "Event Log:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   3795
      End
      Begin VB.Label lblMessage 
         Caption         =   "Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1860
         Width           =   3795
      End
      Begin VB.Label lblContact 
         Caption         =   "Contact / Group Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3795
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status: Waiting for WhatsApp Web..."
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.Label lblFooter 
      Alignment       =   2  'Center
      Caption         =   "Ready"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   14400
   End
End
Attribute VB_Name = "frmWhatsappBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' frmWhatsappBot - WhatsApp Web Bot Demo
' czWebview placed at design-time (same pattern as SimpleTest)
'=========================================================================
Option Explicit

Private m_bLoggedIn     As Boolean
Private m_bBotInjected  As Boolean

'--- Safeguards ---
Private Const MAX_PER_SESSION   As Long = 10
Private Const MAX_PER_CONTACT   As Long = 2
Private Const BLACKLIST As String = "IM3|Indosat|WhatsApp|BONUS PULUHAN|Survey Kepuasan|Promo"

Private m_lTotalReplies As Long
Private m_sReplyLog     As String   '"contact=count;contact=count;..."

Private Sub Form_Load()
    czWebview1.UserDataFolder = App.Path & "\WAData"
    czWebview1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    m_lTotalReplies = 0
    m_sReplyLog = ""
    AddLog "Starting WhatsApp Web..."
    AddLog "Safeguards: max " & MAX_PER_SESSION & "/session, " & MAX_PER_CONTACT & "/contact"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lPanelW As Long
    lPanelW = fraBot.Width
    czWebview1.Move 60, 60, ScaleWidth - lPanelW - 180, ScaleHeight - 300
    fraBot.Move ScaleWidth - lPanelW - 60, 60, lPanelW, ScaleHeight - 300
    lblFooter.Move 0, ScaleHeight - 200, ScaleWidth
End Sub

'=========================================================================
' WebView2 Events (direct event subs, same as SimpleTest)
'=========================================================================

Private Sub czWebview1_InitComplete(ByVal Success As Boolean, ByVal ErrorCode As Long)
    If Success Then
        AddLog "WebView2 ready. Version: " & czWebview1.BrowserVersion
        lblStatus.Caption = "Status: Loading WhatsApp Web..."
        czWebview1.Navigate "https://web.whatsapp.com"
    Else
        lblStatus.Caption = "Status: FAILED (&H" & Hex$(ErrorCode) & ")"
        AddLog "Init failed: &H" & Hex$(ErrorCode)
    End If
End Sub

Private Sub czWebview1_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    If IsSuccess Then
        lblFooter.Caption = czWebview1.DocumentURL
        AddLog "Page loaded"
        '--- Check login with multiple selectors + retry every 2s
        czWebview1.ExecuteScriptAsync "(function chk(){" & _
            "var ok=document.querySelector('#pane-side')" & _
            "||document.querySelector('[data-testid=""chatlist-header""]')" & _
            "||document.querySelector('[data-testid=""chat-list""]')" & _
            "||document.querySelector('div[data-tab=""3""]');" & _
            "if(ok){chrome.webview.postMessage(JSON.stringify({type:'wa_status',loggedIn:true}));}" & _
            "else{setTimeout(chk,2000);}" & _
            "})()"
    End If
End Sub

Private Sub czWebview1_TitleChanged(ByVal Title As String)
    Me.Caption = Title & " - czWebview Bot"
End Sub

Private Sub czWebview1_WebMessageReceived(ByVal Message As String, ByVal IsJSON As Boolean)
    On Error Resume Next
    
    If InStr(1, Message, """type"":""wa_status""") > 0 Then
        If InStr(1, Message, """loggedIn"":true") > 0 Then
            m_bLoggedIn = True
            lblStatus.Caption = "Status: WhatsApp Connected!"
            lblStatus.ForeColor = &H8000&
            AddLog "WhatsApp Web logged in!"
        Else
            lblStatus.Caption = "Status: Scan QR Code to login..."
            lblStatus.ForeColor = &H404040
            AddLog "Waiting for QR scan..."
        End If
    ElseIf InStr(1, Message, """type"":""wa_chats""") > 0 Then
        ParseChats Message
    ElseIf InStr(1, Message, """type"":""wa_new_msg""") > 0 Then
        Dim sSender As String
        Dim sBody As String
        sSender = ExtractJsonValue(Message, "sender")
        sBody = ExtractJsonValue(Message, "body")
        AddLog "MSG [" & sSender & "]: " & sBody
        If chkAutoReply.Value = 1 And LenB(Trim$(txtAutoReply.Text)) <> 0 Then
            '--- Safeguard checks ---
            If IsBlacklisted(sSender) Then
                AddLog "SKIP [" & sSender & "]: blacklisted"
            ElseIf m_lTotalReplies >= MAX_PER_SESSION Then
                AddLog "SKIP: session limit (" & m_lTotalReplies & "/" & MAX_PER_SESSION & ")"
            ElseIf GetContactCount(sSender) >= MAX_PER_CONTACT Then
                AddLog "SKIP [" & sSender & "]: contact limit (" & MAX_PER_CONTACT & ")"
            Else
                SendAutoReply sSender, sBody
                m_lTotalReplies = m_lTotalReplies + 1
                IncrementContact sSender
                lblAutoReply.Caption = "Replies: " & m_lTotalReplies & "/" & MAX_PER_SESSION
            End If
        End If
    ElseIf InStr(1, Message, """type"":""wa_sent""") > 0 Then
        AddLog "Message sent!"
    ElseIf InStr(1, Message, """type"":""wa_error""") > 0 Then
        AddLog "Error: " & ExtractJsonValue(Message, "message")
    ElseIf InStr(1, Message, """type"":""wa_debug""") > 0 Then
        AddLog "DBG: " & ExtractJsonValue(Message, "message")
    End If
End Sub

'=========================================================================
' Bot Control Panel
'=========================================================================

Private Sub cmdInjectBot_Click()
    If Not m_bLoggedIn Then
        MsgBox "Please login to WhatsApp Web first!" & vbCrLf & "Scan the QR code.", vbExclamation
        Exit Sub
    End If
    czWebview1.ExecuteScriptAsync GetBotScript()
    m_bBotInjected = True
    cmdInjectBot.Caption = "Bot Injected!"
    cmdInjectBot.Enabled = False
    AddLog "Bot script injected. Listening..."
End Sub

Private Sub cmdGetChats_Click()
    If Not m_bLoggedIn Then
        MsgBox "Please login to WhatsApp Web first!", vbExclamation
        Exit Sub
    End If
    Dim sScript As String
    sScript = "(function(){" & _
              "var chats=document.querySelectorAll('[data-testid=""cell-frame-container""]');" & _
              "var names=[];" & _
              "chats.forEach(function(c,i){if(i<20){var t=c.querySelector('span[title]');if(t)names.push(t.getAttribute('title'));}});" & _
              "chrome.webview.postMessage(JSON.stringify({type:'wa_chats',names:names}));" & _
              "})()"
    czWebview1.ExecuteScriptAsync sScript
    AddLog "Fetching chats..."
End Sub

Private Sub cmdSend_Click()
    If Not m_bLoggedIn Then
        MsgBox "Please login to WhatsApp Web first!", vbExclamation
        Exit Sub
    End If
    If LenB(Trim$(txtContact.Text)) = 0 Then
        MsgBox "Enter a contact name!", vbExclamation: txtContact.SetFocus: Exit Sub
    End If
    If LenB(Trim$(txtMessage.Text)) = 0 Then
        MsgBox "Enter a message!", vbExclamation: txtMessage.SetFocus: Exit Sub
    End If
    SendWhatsAppMessage Trim$(txtContact.Text), Trim$(txtMessage.Text)
    AddLog "Sending to [" & txtContact.Text & "]..."
End Sub

Private Sub chkAutoReply_Click()
    tmrAutoReply.Enabled = (chkAutoReply.Value = 1)
    If chkAutoReply.Value = 1 Then
        AddLog "Auto-reply ENABLED"
        If Not m_bBotInjected And m_bLoggedIn Then cmdInjectBot_Click
    Else
        AddLog "Auto-reply DISABLED"
    End If
End Sub

Private Sub tmrAutoReply_Timer()
    If m_bLoggedIn And m_bBotInjected Then
        czWebview1.ExecuteScriptAsync "if(window.__czWABot)window.__czWABot.checkNew()"
    End If
End Sub

Private Sub cmdClearLog_Click()
    lstLog.Clear
End Sub

'=========================================================================
' WhatsApp Automation
'=========================================================================

Private Sub SendWhatsAppMessage(ByVal sContact As String, ByVal sMessage As String)
    Dim sScript As String
    sScript = "(function(){" & _
              "var sb=document.querySelector('[data-testid=""chat-list-search""]');" & _
              "if(!sb)sb=document.querySelector('[contenteditable=""true""][data-tab=""3""]');" & _
              "if(!sb){chrome.webview.postMessage(JSON.stringify({type:'wa_error',message:'Search box not found'}));return;}" & _
              "sb.focus();document.execCommand('selectAll');" & _
              "document.execCommand('insertText',false,'" & EscapeJS(sContact) & "');" & _
              "setTimeout(function(){" & _
              "  var c=document.querySelector('[data-testid=""cell-frame-container""] span[title=""" & EscapeJS(sContact) & """]');" & _
              "  if(!c){var cells=document.querySelectorAll('[data-testid=""cell-frame-container""] span[title]');for(var i=0;i<cells.length;i++){if(cells[i].title.indexOf('" & EscapeJS(sContact) & "')>=0){c=cells[i];break;}}}" & _
              "  if(c){c.click();" & _
              "    setTimeout(function(){" & _
              "      var mb=document.querySelector('[data-testid=""conversation-compose-box-input""]');" & _
              "      if(!mb)mb=document.querySelector('footer [contenteditable=""true""]');" & _
              "      if(mb){mb.focus();document.execCommand('insertText',false,'" & EscapeJS(sMessage) & "');" & _
              "        setTimeout(function(){var s=document.querySelector('[data-testid=""send""]');if(s){s.click();chrome.webview.postMessage(JSON.stringify({type:'wa_sent'}));}},500);}" & _
              "    },1500);" & _
              "  }else{chrome.webview.postMessage(JSON.stringify({type:'wa_error',message:'Contact not found'}));}" & _
              "},1500);" & _
              "})()"
    czWebview1.ExecuteScriptAsync sScript
End Sub

Private Sub SendAutoReply(ByVal sSender As String, ByVal sOriginalMsg As String)
    Dim sReply As String
    sReply = txtAutoReply.Text
    sReply = Replace(sReply, "{sender}", sSender)
    sReply = Replace(sReply, "{message}", sOriginalMsg)
    sReply = Replace(sReply, "{time}", Format$(Now, "hh:nn:ss"))
    AddLog "Auto-replying to [" & sSender & "]..."
    '--- Type directly in the current open chat then press Enter
    Dim sScript As String
    sScript = "(function(){"
    sScript = sScript & "var mb=document.querySelector('[data-testid=""conversation-compose-box-input""]');"
    sScript = sScript & "if(!mb)mb=document.querySelector('footer [contenteditable=""true""]');"
    sScript = sScript & "if(!mb)mb=document.querySelector('[contenteditable=""true""][data-tab=""10""]');"
    sScript = sScript & "if(!mb){chrome.webview.postMessage(JSON.stringify({type:'wa_error',message:'No compose box'}));return;}"
    sScript = sScript & "mb.focus();"
    sScript = sScript & "document.execCommand('insertText',false,'" & EscapeJS(sReply) & "');"
    sScript = sScript & "setTimeout(function(){"
    sScript = sScript & "var s=document.querySelector('[data-testid=""send""]');"
    sScript = sScript & "if(!s){var si=document.querySelector('span[data-icon=""send""]');if(si)s=si.closest('button')||si.parentElement;}"
    sScript = sScript & "if(!s)s=document.querySelector('button[aria-label=""Send""]');"
    sScript = sScript & "if(s){s.click();chrome.webview.postMessage(JSON.stringify({type:'wa_sent'}));}"
    sScript = sScript & "else{mb.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));"
    sScript = sScript & "chrome.webview.postMessage(JSON.stringify({type:'wa_sent'}));}"
    sScript = sScript & "},500);"
    sScript = sScript & "})()"
    czWebview1.ExecuteScriptAsync sScript
End Sub

Private Function GetBotScript() As String
    Dim s As String
    s = "(function(){"
    s = s & "if(window.__czWABot)return;"
    s = s & "var lastTxt='';var lastCnt=0;var cd=0;var waitTick=0;"
    s = s & "window.__czWABot={"
    s = s & "checkNew:function(){"
    '--- Cooldown after reply
    s = s & "if(cd>0){cd--;return;}"
    '--- Wait for chat to load after clicking sidebar
    s = s & "if(waitTick>0){waitTick--;return;}"
    '--- Step 1: Check current open chat for new messages
    s = s & "var panel=document.querySelector('#main');"
    s = s & "if(panel){"
    s = s & "var spans=panel.querySelectorAll('span.selectable-text');"
    s = s & "if(!spans.length)spans=panel.querySelectorAll('.copyable-text');"
    s = s & "if(!spans.length)spans=panel.querySelectorAll('span[dir]');"
    s = s & "var cnt=spans.length;"
    s = s & "if(cnt>lastCnt){"
    s = s & "lastCnt=cnt;var last=spans[cnt-1];"
    s = s & "var txt=last.innerText||'';"
    s = s & "if(txt&&txt!==lastTxt&&txt.indexOf('czWebview')<0){"
    s = s & "lastTxt=txt;"
    s = s & "var h=document.querySelector('header span[title]');"
    s = s & "var who=h?h.title:'?';"
    s = s & "cd=10;"
    s = s & "chrome.webview.postMessage(JSON.stringify({type:'wa_new_msg',sender:who,body:txt}));"
    s = s & "return;}}"
    s = s & "lastCnt=cnt;}"
    '--- Step 2: Scan sidebar for unread badges
    s = s & "var rows=document.querySelectorAll('[data-testid=""cell-frame-container""]');"
    s = s & "for(var i=0;i<rows.length;i++){"
    s = s & "var badge=rows[i].querySelector('[data-testid=""icon-unread-count""]');"
    s = s & "if(!badge)badge=rows[i].querySelector('span[aria-label*=""unread""]');"
    s = s & "if(!badge){var ss=rows[i].querySelectorAll('span');for(var j=0;j<ss.length;j++){var v=ss[j].innerText;"
    s = s & "if(v&&/^[0-9]+$/.test(v)&&parseInt(v)>0&&ss[j].offsetWidth<40){badge=ss[j];break;}}}"
    s = s & "if(badge){"
    s = s & "var title=rows[i].querySelector('span[title]');"
    s = s & "if(!title)continue;"
    s = s & "var name=title.title;"
    s = s & "if(/IM3|Indosat|WhatsApp|BONUS|Survey|Promo/i.test(name))continue;"
    s = s & "chrome.webview.postMessage(JSON.stringify({type:'wa_debug',message:'Open: '+name}));"
    s = s & "var clickEl=rows[i].querySelector('[data-testid=""cell-frame-title""]')||title;"
    s = s & "clickEl.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));"
    s = s & "clickEl.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));"
    s = s & "clickEl.dispatchEvent(new MouseEvent('click',{bubbles:true}));"
    s = s & "waitTick=2;lastCnt=0;lastTxt='';"
    s = s & "break;}}"
    s = s & "}};"
    '--- Initialize with current count
    s = s & "var p=document.querySelector('#main');"
    s = s & "if(p){lastCnt=p.querySelectorAll('span.selectable-text').length;}"
    s = s & "chrome.webview.postMessage(JSON.stringify({type:'wa_debug',message:'Bot OK'}));"
    s = s & "})()"
    GetBotScript = s
End Function

'=========================================================================
' Helper Functions
'=========================================================================

Private Sub AddLog(ByVal sMsg As String)
    lstLog.AddItem Format$(Now, "hh:nn:ss") & " " & sMsg
    If lstLog.ListCount > 0 Then lstLog.ListIndex = lstLog.ListCount - 1
End Sub

Private Function IsBlacklisted(ByVal sName As String) As Boolean
    Dim vList As Variant, i As Long
    vList = Split(BLACKLIST, "|")
    For i = 0 To UBound(vList)
        If InStr(1, sName, vList(i), vbTextCompare) > 0 Then
            IsBlacklisted = True: Exit Function
        End If
    Next
    IsBlacklisted = False
End Function

Private Function GetContactCount(ByVal sName As String) As Long
    Dim lPos As Long, lEq As Long, lSc As Long
    lPos = InStr(1, m_sReplyLog, sName & "=", vbTextCompare)
    If lPos = 0 Then GetContactCount = 0: Exit Function
    lEq = InStr(lPos, m_sReplyLog, "=")
    lSc = InStr(lEq, m_sReplyLog, ";")
    If lSc = 0 Then lSc = Len(m_sReplyLog) + 1
    GetContactCount = CLng(Mid$(m_sReplyLog, lEq + 1, lSc - lEq - 1))
End Function

Private Sub IncrementContact(ByVal sName As String)
    Dim lCur As Long
    lCur = GetContactCount(sName)
    If lCur = 0 Then
        m_sReplyLog = m_sReplyLog & sName & "=1;"
    Else
        Dim sOld As String
        sOld = sName & "=" & lCur
        m_sReplyLog = Replace(m_sReplyLog, sOld, sName & "=" & (lCur + 1))
    End If
End Sub

Private Sub ParseChats(ByVal sJson As String)
    Dim lStart As Long, lEnd As Long
    lStart = InStr(1, sJson, "[")
    lEnd = InStr(1, sJson, "]")
    If lStart > 0 And lEnd > lStart Then
        Dim sNames As String
        sNames = Mid$(sJson, lStart + 1, lEnd - lStart - 1)
        sNames = Replace(sNames, """", "")
        Dim vNames As Variant
        vNames = Split(sNames, ",")
        Dim i As Long
        AddLog "=== Active Chats ==="
        For i = 0 To UBound(vNames)
            If LenB(Trim$(vNames(i))) <> 0 Then AddLog "  " & (i + 1) & ". " & Trim$(vNames(i))
        Next
        AddLog "=== " & (UBound(vNames) + 1) & " chats ==="
    End If
End Sub

Private Function ExtractJsonValue(ByVal sJson As String, ByVal sKey As String) As String
    Dim lPos As Long, lStart As Long, lEnd As Long
    lPos = InStr(1, sJson, """" & sKey & """:""")
    If lPos > 0 Then
        lStart = lPos + Len(sKey) + 4
        lEnd = InStr(lStart, sJson, """")
        If lEnd > lStart Then ExtractJsonValue = Mid$(sJson, lStart, lEnd - lStart)
    End If
End Function

Private Function EscapeJS(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "\n")
    EscapeJS = s
End Function
