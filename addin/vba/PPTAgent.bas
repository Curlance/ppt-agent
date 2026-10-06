Attribute VB_Name = "PPTAgent"
Option Explicit

'==============================================================================
' ppt-agent · PowerPoint 端监视台（VBA 模块）
'
' 这个加载项只做两件事：
'   1) 读 pptd 守护进程的纯文本面板数据（GET /feed.txt）
'   2) 把用户的跟速 / 急停指令发回去（POST /pace、POST /shutdown）
'
' 它刻意**不做** COM 自动化 —— 那是 pptd 守护进程的职责。两个进程同时驱动同一个
' PowerPoint，会互相把对方的调用打成 RPC_E_CALL_REJECTED（"调用被拒绝"）。
' 这里只是一块显示屏 + 几个按钮。
'
' 为什么读纯文本而不是 JSON：VBA 里没有 JSON 解析器，手写一个又长又脆。
' /feed.txt 一行一条、用 " | " 分隔，Split 一句就够用。
'==============================================================================

Public Const AGENT_HOST As String = "127.0.0.1"
Public Const AGENT_PORT As Long = 8791

Private mToken As String

'------------------------------------------------------------------------------
' 基础
'------------------------------------------------------------------------------

Public Function AgentBaseUrl() As String
    AgentBaseUrl = "http://" & AGENT_HOST & ":" & AGENT_PORT
End Function

' 从 %LOCALAPPDATA%\ppt-agent\runtime.json 里取令牌。
' 读不到时面板会提示用 `pptctl url --mcp-token` 手动确认服务在跑。
Public Function AgentToken() As String
    If Len(mToken) > 0 Then
        AgentToken = mToken
        Exit Function
    End If

    Dim path As String
    path = Environ$("LOCALAPPDATA") & "\ppt-agent\runtime.json"

    On Error GoTo Done
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    mToken = JsonString(st.ReadText(-1), "token")
    st.Close
Done:
    AgentToken = mToken
End Function

' 够用的 JSON 取值：只认 "key" : "value" 这种字符串字段
Private Function JsonString(ByVal text As String, ByVal key As String) As String
    Dim needle As String, p As Long, q As Long
    needle = """" & key & """"
    p = InStr(1, text, needle, vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(needle), text, """")
    If p = 0 Then Exit Function
    q = InStr(p + 1, text, """")
    If q = 0 Then Exit Function
    JsonString = Mid$(text, p + 1, q - p - 1)
End Function

'------------------------------------------------------------------------------
' HTTP
'------------------------------------------------------------------------------

' 用 ServerXMLHTTP 而不是 XMLHTTP：只有它能设超时。
' 守护进程若正卡在 PowerPoint 的模态对话框上，XMLHTTP 会永久挂住整个 PowerPoint。
Public Function AgentGet(ByVal pathAndQuery As String, ByRef outText As String) As Boolean
    Dim http As Object
    On Error GoTo Failed
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 2000, 2000, 4000, 6000      ' 解析/连接/发送/接收（毫秒）
    http.Open "GET", AgentBaseUrl() & pathAndQuery, False
    If Len(AgentToken()) > 0 Then http.setRequestHeader "X-PPT-Token", AgentToken()
    http.send
    outText = http.responseText
    AgentGet = (http.Status = 200)
    Exit Function
Failed:
    outText = "连不上守护进程（" & Err.Description & "）。先运行：pptctl serve"
    AgentGet = False
End Function

Public Function AgentPost(ByVal pathAndQuery As String, ByVal body As String, ByRef outText As String) As Boolean
    Dim http As Object
    On Error GoTo Failed
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 2000, 2000, 4000, 6000
    http.Open "POST", AgentBaseUrl() & pathAndQuery, False
    http.setRequestHeader "Content-Type", "application/json"
    If Len(AgentToken()) > 0 Then http.setRequestHeader "X-PPT-Token", AgentToken()
    http.send body
    outText = http.responseText
    AgentPost = (http.Status = 200)
    Exit Function
Failed:
    outText = "发送失败（" & Err.Description & "）"
    AgentPost = False
End Function

'------------------------------------------------------------------------------
' 面板数据的格式化
'------------------------------------------------------------------------------

' 把 /feed.txt 的头部 key=value 拼成一行给用户看的中文摘要
Public Function AgentFormatHeader(ByVal text As String) As String
    Dim lines() As String, i As Long, line As String
    Dim alive As String, ver As String, deck As String, slide As String
    Dim pace As String, tools As String, visible As String

    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(Replace(lines(i), vbCr, ""))
        If line = "---" Then Exit For
        If InStr(line, "=") > 0 Then
            Dim k As String, v As String
            k = Left$(line, InStr(line, "=") - 1)
            v = Mid$(line, InStr(line, "=") + 1)
            Select Case k
                Case "alive": alive = v
                Case "version": ver = v
                Case "visible": visible = v
                Case "deck": deck = v
                Case "slide": slide = v
                Case "pace_ms": pace = v
                Case "tools": tools = v
            End Select
        End If
    Next i

    Dim parts As String
    If alive = "true" Then
        parts = "● PowerPoint " & ver
        If visible = "true" Then parts = parts & " · 窗口可见" Else parts = parts & " · 窗口不可见"
    Else
        parts = "× 守护进程未连接"
        AgentFormatHeader = parts
        Exit Function
    End If

    If Len(deck) > 0 Then parts = parts & " · " & deck
    If Len(slide) > 0 Then parts = parts & " · 第 " & slide & " 页"
    If Len(pace) > 0 Then parts = parts & " · 跟速 " & pace & "ms"
    If Len(tools) > 0 Then parts = parts & " · " & tools & " 个工具"

    AgentFormatHeader = parts
End Function

'------------------------------------------------------------------------------
' 自动刷新
'------------------------------------------------------------------------------

' PowerPoint 的 Application.OnTime 在某些版本/状态下不可用，
' 所以全程 On Error Resume Next —— 失败时用户还可以点「刷新」。
Public Sub PPTAgentScheduleTick()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 2), "PPTAgentTick"
End Sub

Public Sub PPTAgentTick()
    On Error Resume Next
    PPTAgentPanel.RefreshFeed
    PPTAgentPanel.ScheduleNext
End Sub

'------------------------------------------------------------------------------
' 入口：把宏挂到快速访问工具栏，点一下就开面板
'------------------------------------------------------------------------------

Public Sub PPTAgentShow()
    On Error GoTo Failed
    PPTAgentPanel.Show vbModeless
    Exit Sub
Failed:
    MsgBox "打不开面板：" & Err.Description & vbCrLf & vbCrLf & _
           "请确认 PPTAgentPanel 窗体已经导入本项目。", vbExclamation, "ppt-agent"
End Sub

' 打开浏览器里的完整观察台（面板只显示摘要，完整版在浏览器里）
Public Sub PPTAgentOpenObserver()
    Dim text As String
    If AgentGet("/health", text) Then
        Dim url As String
        url = AgentBaseUrl() & "/?token=" & AgentToken()
        On Error Resume Next
        ThisPresentation.FollowHyperlink url
    Else
        MsgBox text, vbExclamation, "ppt-agent"
    End If
End Sub
