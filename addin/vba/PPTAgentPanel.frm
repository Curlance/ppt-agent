VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTAgentPanel
   Caption         =   "ppt-agent 监视台"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   StartUpPosition =   1  'CenterOwner
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblStatus
      AutoSize        =   -1  'True
      Caption         =   "正在连接守护进程…"
      Height          =   240
      Left            =   180
      Top             =   135
      Width           =   6240
   End
   Begin {8BD21D20-EC42-11CE-9E0D-00AA006002F3} lstFeed
      Height          =   4200
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   6240
   End
   Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkAuto
      Caption         =   "自动刷新"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   4740
      Width           =   1500
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnFast
      Caption         =   "极速"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   5100
      Width           =   1080
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnNormal
      Caption         =   "正常"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   5100
      Width           =   1080
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnSlow
      Caption         =   "慢速"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   5100
      Width           =   1080
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnRefresh
      Caption         =   "刷新"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   5100
      Width           =   1080
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnStop
      Caption         =   "急停"
      Height          =   375
      Left            =   4740
      TabIndex        =   6
      Top             =   5100
      Width           =   1500
   End
End
Attribute VB_Name = "PPTAgentPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' ppt-agent 监视台 · 窗体代码
'
' 控件 CLSID 不是凭记忆写的，是从注册表里指向 FM20.DLL 的条目枚举出来的：
'   Form          {C62A69F0-16DC-11CE-9E98-00AA00574A4F}
'   Label         {978C9E23-D4B0-11CE-BF2D-00AA003F40D0}
'   ListBox       {8BD21D20-EC42-11CE-9E0D-00AA006002F3}
'   CheckBox      {8BD21D40-EC42-11CE-9E0D-00AA006002F3}
'   CommandButton {D7053240-CE69-11CD-A777-00DD01143C57}
' 版本不匹配的 GUID 会让导入直接失败，所以这几个值不要手改。
'==============================================================================

Private Sub UserForm_Initialize()
    Me.Caption = "ppt-agent 监视台"
    chkAuto.Value = True
    RefreshFeed
    ScheduleNext
End Sub

Public Sub ScheduleNext()
    If chkAuto.Value = False Then Exit Sub
    PPTAgentScheduleTick
End Sub

' 读 /feed.txt，把头部压成一行摘要、把事件逐行塞进列表框
Public Sub RefreshFeed()
    Dim text As String
    If Not AgentGet("/feed.txt?limit=60", text) Then
        lblStatus.Caption = "× " & text
        Exit Sub
    End If

    Dim lines() As String, i As Long, line As String, inRows As Boolean
    lines = Split(text, vbLf)
    lstFeed.Clear
    inRows = False

    For i = LBound(lines) To UBound(lines)
        line = Trim$(Replace(lines(i), vbCr, ""))
        If line = "---" Then
            inRows = True
        ElseIf Len(line) > 0 Then
            If inRows Then
                lstFeed.AddItem line
            End If
        End If
    Next i

    lblStatus.Caption = AgentFormatHeader(text)
    If lstFeed.ListCount > 0 Then lstFeed.ListIndex = lstFeed.ListCount - 1
End Sub

Private Sub btnRefresh_Click()
    RefreshFeed
End Sub

Private Sub chkAuto_Click()
    If chkAuto.Value Then ScheduleNext
End Sub

'------------------------------------------------------------------------------
' 跟速与急停：只发 HTTP 指令，不自己动 PowerPoint
'------------------------------------------------------------------------------

Private Sub SetPace(ByVal millis As Long)
    Dim out As String
    If AgentPost("/pace", "{""pace_ms"":" & millis & "}", out) Then
        lblStatus.Caption = "跟速已设为 " & millis & " 毫秒"
    Else
        lblStatus.Caption = "× " & out
    End If
    RefreshFeed
End Sub

Private Sub btnFast_Click()
    SetPace 0
End Sub

Private Sub btnNormal_Click()
    SetPace 400
End Sub

Private Sub btnSlow_Click()
    SetPace 1500
End Sub

Private Sub btnStop_Click()
    If MsgBox("确定要停止 ppt-agent 守护进程吗？" & vbCrLf & _
              "PowerPoint 不会被关闭，已经打开的文件也不受影响。", _
              vbYesNo + vbQuestion, "ppt-agent") <> vbYes Then
        Exit Sub
    End If
    Dim out As String
    AgentPost "/shutdown", "", out
    lblStatus.Caption = "已请求停止守护进程（PowerPoint 未受影响）"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    chkAuto.Value = False
End Sub
