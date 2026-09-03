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

Function BitIsSet(bits, id)
    BitIsSet = False
    If id < 1 Then Exit Function
    If Len(StrValue(bits)) < id Then Exit Function
    BitIsSet = (Mid(bits, id, 1) = "1")
End Function

Function ISODate(d)
    Dim value
    If IsObject(d) Then value = d.Value Else value = d
    ISODate = ""
    If IsNull(value) Then Exit Function
    If Not IsDate(value) Then Exit Function
    ISODate = Year(value) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
End Function

Function ParseDate(s)
    Dim re, p, d, n
    ParseDate = Empty
    Set re = New RegExp
    re.Pattern = "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
    If Not re.Test(s) Then Exit Function
    p = Split(s, "-")
    If CInt(p(0)) < 1900 Or CInt(p(0)) > 2100 Then Exit Function
    On Error Resume Next
    d = DateSerial(CInt(p(0)), CInt(p(1)), CInt(p(2)))
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n <> 0 Then Exit Function
    If ISODate(d) <> s Then Exit Function
    ParseDate = d
End Function

Function ParamText(v)
    Dim s, kind
    s = StrValue(v)
    kind = 202
    If Len(s) > 255 Then kind = 203
    ParamText = Array(kind, Len(s) + 1, s)
End Function

Function ParamID(v)
    ParamID = Array(3, 4, CLng(v))
End Function

Function ParamDate(v)
    ParamDate = Array(7, 8, v)
End Function

Function Rows(c, sql, params)
    Dim cmd
    Set cmd = NewCommand(c, sql, params)
    Set Rows = cmd.Execute
End Function

Function RunSQL(c, sql, params)
    Dim cmd, affected
    Set cmd = NewCommand(c, sql, params)
    cmd.Execute affected, , 128
    RunSQL = affected
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

Sub FailPage(status, message)
    Response.Status = status
    Response.Write "<meta charset='utf-8'><p>" & H(message) & "</p><a href='javascript:history.back()'>돌아가기</a>"
    Response.End
End Sub

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
Response.CodePage = 65001
Response.CharSet = "utf-8"

Dim conn, rs, sql
Dim lessonId, isEdit
Dim title, trainingDate, memo
Dim dateParts, dbDate
Dim selectedTeacherIds, i
Dim targetBits, oldTargetBits
Dim teacherRs, allTeacherRs, signatureTeacherRs
Dim teacherId, teacherName, teacherStatus
Dim bitValue, bitPosition, isChecked, showTeacher
Dim signatureExists, selected, msg, transactionOpen

lessonId = PositiveID(Request.QueryString("id"))
If StrValue(Request.QueryString("id")) <> "" And lessonId = 0 Then FailPage "400 Bad Request", "잘못된 연수 ID입니다."
isEdit = (lessonId > 0)
If Not isEdit Then lessonId = ""

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=attendanceDB"


' 저장 눌렀을 때
Sub SaveLesson()
    conn.BeginTrans
    transactionOpen = True
    oldTargetBits = ""
    If lessonId <> "" Then
        Set rs = Rows(conn, "SELECT [참여대상] FROM lesson_list WHERE ID=?", Array(ParamID(lessonId)))
        If rs.EOF Then Err.Raise vbObjectError+1, "연수", "해당 연수가 없습니다."
        oldTargetBits = StrValue(rs("참여대상"))
        rs.Close
    End If

    ' 체크한 사람 기준으로 참여대상 10101 이런거 만듦
    ' 기존 대상이던 휴직/퇴직도 수정화면에서 체크 풀면 빠집니다 네
    targetBits = ""
    Set allTeacherRs = conn.Execute("SELECT ID, 재직상태 FROM teachers ORDER BY ID")
    Do While Not allTeacherRs.EOF
        teacherId = CLng(allTeacherRs("ID"))
        If teacherId < 1 Or teacherId > 1000000 Then Err.Raise vbObjectError+2, "대상자", "교직원 ID를 확인해 주세요."
        teacherStatus = StrValue(allTeacherRs("재직상태"))
        If Len(targetBits) < teacherId-1 Then targetBits = targetBits & String(teacherId-1-Len(targetBits), "0")
        bitValue = "0"
        If selected.Exists(CStr(teacherId)) Then
            ' 휴직/퇴직인데 원래 대상이었던 사람만 수정 가능하게 거름
            If teacherStatus <> "재직" And Not BitIsSet(oldTargetBits, teacherId) Then Err.Raise vbObjectError+3, "대상자", "선택할 수 없는 대상자가 있습니다."
            bitValue = "1"
            selected.Remove CStr(teacherId)
        End If
        targetBits = targetBits & bitValue
        allTeacherRs.MoveNext
    Loop
    allTeacherRs.Close
    If selected.Count > 0 Then Err.Raise vbObjectError+4, "대상자", "존재하지 않는 대상자가 있습니다."

    ' 날짜 input 값 Access에 넣기 좋게 바꾸기
    dbDate = ParseDate(trainingDate)
    If lessonId = "" Then
        ' 새 연수 등록
        Call RunSQL(conn, "INSERT INTO lesson_list ([연수제목],[연수일시],[참여대상],[메모]) VALUES (?,?,?,?)", Array(ParamText(title),ParamDate(dbDate),ParamText(targetBits),ParamText(memo)))
        ' 방금 등록된 연수 ID 가져오기
        Set rs = conn.Execute("SELECT @@IDENTITY AS newID")
        lessonId = CLng(rs("newID"))
        rs.Close
    Else
        ' 기존 연수 수정
        Call RunSQL(conn, "UPDATE lesson_list SET [연수제목]=?,[연수일시]=?,[참여대상]=?,[메모]=? WHERE ID=?", Array(ParamText(title),ParamDate(dbDate),ParamText(targetBits),ParamText(memo),ParamID(lessonId)))
    End If

    ' 연수 대상이랑 signature_data 내용 맞춰두기
    Set signatureTeacherRs = conn.Execute("SELECT ID, 이름 FROM teachers ORDER BY ID")
    Do While Not signatureTeacherRs.EOF
        teacherId = CLng(signatureTeacherRs("ID"))
        If BitIsSet(targetBits, teacherId) Then
            Set rs = Rows(conn, "SELECT COUNT(*) AS cnt FROM signature_data WHERE [연수id]=? AND [사용자id]=?", Array(ParamID(lessonId),SignatureParam(conn,teacherId)))
            signatureExists = (CLng(rs("cnt")) > 0)
            rs.Close
            ' 새로 대상이 된 사람만 한 줄 추가, 기존 서명주소는 안건드림
            If Not signatureExists Then Call RunSQL(conn, "INSERT INTO signature_data ([연수id],[사용자id]) VALUES (?,?)", Array(ParamID(lessonId),SignatureParam(conn,teacherId)))
        Else
            ' 대상에서 빠지면 그 연수의 서명칸도 같이 뺌
            Call RunSQL(conn, "DELETE FROM signature_data WHERE [연수id]=? AND [사용자id]=?", Array(ParamID(lessonId),SignatureParam(conn,teacherId)))
        End If
        signatureTeacherRs.MoveNext
    Loop
    signatureTeacherRs.Close
    conn.CommitTrans
    transactionOpen = False
End Sub

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    lessonId = PositiveID(Request.Form("lessonId"))
    If StrValue(Request.Form("lessonId")) <> "" And lessonId = 0 Then FailPage "400 Bad Request", "잘못된 연수 ID입니다."
    If lessonId = 0 Then lessonId = ""
    title = Trim(StrValue(Request.Form("title")))
    trainingDate = StrValue(Request.Form("trainingDate"))
    memo = StrValue(Request.Form("memo"))
    If title = "" Or Len(title)>255 Or Len(memo)>8000 Or IsEmpty(ParseDate(trainingDate)) Then FailPage "400 Bad Request", "연수명·날짜·메모를 확인해 주세요."
    Set selected = Server.CreateObject("Scripting.Dictionary")
    For i = 1 To Request.Form("targetTeacher").Count
        teacherId = PositiveID(Request.Form("targetTeacher")(i))
        If teacherId = 0 Then FailPage "400 Bad Request", "대상자를 확인해 주세요."
        selected(CStr(teacherId)) = True
    Next
    Application.Lock
    transactionOpen = False
    On Error Resume Next
    SaveLesson
    msg = Err.Description
    If transactionOpen Then conn.RollbackTrans
    conn.Close
    Err.Clear
    On Error GoTo 0
    Application.UnLock
    If msg <> "" Then FailPage "400 Bad Request", "저장하지 못했습니다. " & msg
    Response.Redirect "administer.asp"
End If

' 수정으로 들어올때 기존 값 가져오는거
title = ""
trainingDate = ""
memo = ""
targetBits = ""

If isEdit Then

    sql = "SELECT * FROM lesson_list WHERE ID = " & lessonId
    Set rs = conn.Execute(sql)
    If rs.EOF Then FailPage "404 Not Found", "해당 연수가 없습니다."

    title = StrValue(rs("연수제목"))

    trainingDate = ISODate(rs("연수일시"))
    memo = StrValue(rs("메모"))
    targetBits = StrValue(rs("참여대상"))

    rs.Close
    Set rs = Nothing

End If


' 재직자는 전부, 수정이면 기존 대상이던 휴직/퇴직도 같이 불러옴
sql = "SELECT ID, 이름, 재직상태 FROM teachers ORDER BY 이름"
Set teacherRs = Server.CreateObject("ADODB.Recordset")
teacherRs.Open sql, conn

%>

<!DOCTYPE html>
<html lang="ko">

<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="style.css">

<% If isEdit Then %>
<title>연수 수정</title>
<% Else %>
<title>연수 등록</title>
<% End If %>

</head>

<body class="register-page">

<% If isEdit Then %>
<h1>연수 수정</h1>
<% Else %>
<h1>새 연수 등록</h1>
<% End If %>

<form method="post" action="lesson_register.asp">

<input type="hidden" name="lessonId" value="<%= H(lessonId) %>">

<p>
<label>연수명</label><br>
<input type="text" name="title" maxlength="255" value="<%= H(title) %>" required>
</p>

<p>
<label>날짜</label><br>
<input type="date" name="trainingDate" value="<%= H(trainingDate) %>" required>
</p>

<p>
<label>메모</label><br>
<textarea name="memo" maxlength="8000" rows="5" cols="50"><%= H(memo) %></textarea>
</p>

<details>

<summary>대상자 선택</summary>

<p>
<input type="checkbox" id="selectAll"<% If Not isEdit Then Response.Write " checked" %>>
<label for="selectAll">전체 선택</label>
</p>

<%
Do While Not teacherRs.EOF

    teacherId = CLng(teacherRs("ID"))
    teacherName = StrValue(teacherRs("이름"))
    teacherStatus = StrValue(teacherRs("재직상태"))

    showTeacher = False

    If teacherStatus = "재직" Then
        showTeacher = True

    ElseIf isEdit Then
        If Len(targetBits) >= teacherId Then
            If Mid(targetBits, teacherId, 1) = "1" Then
                showTeacher = True
            End If
        End If
    End If

    If showTeacher Then

        If isEdit Then
            isChecked = BitIsSet(targetBits, teacherId)
        Else
            isChecked = True
        End If
%>

<p>
<input
    type="checkbox"
    class="teacherCheck"
    name="targetTeacher"
    value="<%= teacherId %>"
    id="teacher_<%= teacherId %>"
    <% If isChecked Then Response.Write "checked" %>
>

<label for="teacher_<%= teacherId %>">
<%= H(teacherName) %><% If teacherStatus <> "재직" Then %> (<%= H(teacherStatus) %>)<% End If %>
</label>
</p>

<%
    End If

    teacherRs.MoveNext
Loop
%>

</details>

<p>
<% If isEdit Then %>
<button type="submit">수정 내용 저장</button>
<% Else %>
<button type="submit">연수 등록</button>
<% End If %>

<a href="administer.asp">취소</a>
</p>

</form>

<script>

const selectAll = document.getElementById("selectAll");
const teacherChecks = document.querySelectorAll(".teacherCheck");

selectAll.addEventListener("change", function() {

    teacherChecks.forEach(function(item) {
        item.checked = selectAll.checked;
    });

});

function updateSelectAll() {
    const checked = [...teacherChecks].filter(item => item.checked).length;
    selectAll.checked = teacherChecks.length > 0 && checked === teacherChecks.length;
    selectAll.indeterminate = checked > 0 && checked < teacherChecks.length;
}
teacherChecks.forEach(item => item.addEventListener("change", updateSelectAll));
updateSelectAll();
</script>

</body>
</html>

<%
teacherRs.Close
Set teacherRs = Nothing

conn.Close
Set conn = Nothing
%>
