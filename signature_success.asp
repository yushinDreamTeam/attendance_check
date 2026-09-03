<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Function StrValue(v)
    Dim value
    If IsObject(v) Then value = v.Value Else value = v
    If IsNull(value) Or IsEmpty(value) Then StrValue = "" Else StrValue = CStr(value)
End Function

Function H(v)
    H = Server.HTMLEncode(StrValue(v))
End Function

Function PositiveID(v)
    Dim s, re
    PositiveID = 0
    s = StrValue(v)
    Set re = New RegExp
    re.Pattern = "^[1-9][0-9]{0,9}$"
    If Not re.Test(s) Then Exit Function
    If CDbl(s) > 2147483647 Then Exit Function
    PositiveID = CLng(s)
End Function

Function ISODate(d)
    Dim value
    If IsObject(d) Then value = d.Value Else value = d
    ISODate = ""
    If IsNull(value) Then Exit Function
    If Not IsDate(value) Then Exit Function
    ISODate = Year(value) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
End Function

Function ParamID(v)
    ParamID = Array(3, 4, CLng(v))
End Function

Function Rows(c, sql, params)
    Dim cmd
    Set cmd = NewCommand(c, sql, params)
    Set Rows = cmd.Execute
End Function

Function SignatureParam(c, id)
    Dim r, t
    Set r = c.Execute("SELECT [사용자id] FROM [signature_data] WHERE 1=0")
    t = r.Fields(0).Type
    r.Close
    If t = 202 Or t = 200 Or t = 130 Or t = 129 Then
        SignatureParam = ParamText(CStr(id))
    ElseIf t = 2 Or t = 3 Then
        SignatureParam = ParamID(id)
    Else
        Err.Raise vbObjectError + 101, "DB", "signature_data의 사용자id 필드 형식을 확인해 주세요."
    End If
End Function

Function SafeSignaturePath(path, lessonID, teacherID)
    Dim expected
    expected = "signatures/sig_" & lessonID & "_" & teacherID & ".png"
    SafeSignaturePath = ""
    If LCase(Replace(StrValue(path), "\", "/")) = expected Then SafeSignaturePath = expected
End Function

Sub FailPage(status, message)
    Response.Status = status
    Response.Write "<meta charset='utf-8'><p>" & H(message) & "</p><a href='javascript:history.back()'>돌아가기</a>"
    Response.End
End Sub

Function ParamText(v)
    Dim s, kind
    s = StrValue(v)
    kind = 202
    If Len(s) > 255 Then kind = 203
    ParamText = Array(kind, Len(s) + 1, s)
End Function

Function NewCommand(c, sql, params)
    Dim cmd, p, i
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = c
    cmd.CommandType = 1
    cmd.CommandText = sql
    For i = 0 To UBound(params)
        p = params(i)
        cmd.Parameters.Append cmd.CreateParameter("p" & i, p(0), 1, p(1), p(2))
    Next
    Set NewCommand = cmd
End Function
%>

<%
Dim lessonID, teacherID, c, r, msg, title, lessonDateValue, path, fso
lessonID=PositiveID(Request.QueryString("id"))
teacherID=PositiveID(Session("signedTeacher"))
If lessonID=0 Or teacherID=0 Or PositiveID(Session("signedLesson"))<>lessonID Then FailPage "400 Bad Request", "제출 결과를 확인할 수 없습니다. 해당 연수의 QR에서 출석 화면을 열어 주세요."
On Error Resume Next
Set c=Server.CreateObject("ADODB.Connection"):c.Open "DSN=attendanceDB":msg=Err.Description
On Error GoTo 0
If msg<>"" Then FailPage "503 Service Unavailable", msg
Set r=Rows(c,"SELECT [파일경로] FROM [signature_data] WHERE [연수id]=? AND [사용자id]=?",Array(ParamID(lessonID),SignatureParam(c,teacherID)))
If r.EOF Then r.Close: c.Close: FailPage "404 Not Found","저장된 서명을 확인하지 못했습니다. 관리자에게 문의해 주세요."
path=SafeSignaturePath(r("파일경로"),lessonID,teacherID):r.Close
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If path="" Then c.Close: FailPage "404 Not Found","저장된 서명을 확인하지 못했습니다."
If Not fso.FileExists(Server.MapPath(path)) Then c.Close: FailPage "404 Not Found","서명 파일을 확인하지 못했습니다. 관리자에게 문의해 주세요."
Set r=Rows(c,"SELECT [연수제목],[연수일시] FROM [lesson_list] WHERE [ID]=?",Array(ParamID(lessonID)))
If r.EOF Then r.Close:c.Close:FailPage "404 Not Found","해당 연수가 없습니다."
title=StrValue(r("연수제목")):lessonDateValue=ISODate(r("연수일시")):r.Close:c.Close
%>
<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>서명 제출 완료</title><link rel="stylesheet" href="userpage.css"></head><body>
<main class="success-page"><section class="card success-card"><div class="success-mark" aria-hidden="true">✓</div><h1>서명 제출 완료</h1><p><strong><%=H(title)%></strong></p><p class="muted"><%=H(lessonDateValue)%></p><p>서명이 저장되었습니다.<br>이 창을 닫으셔도 됩니다.</p><a class="btn" href="userpage.asp?id=<%=lessonID%>">서명 다시 작성</a></section></main></body></html>
