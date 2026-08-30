<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

Dim conn, rs, sql
Dim lessonId, isEdit
Dim title, trainingDate, memo
Dim dateParts, dbDate
Dim selectedTeacherIds, i
Dim targetBits, oldTargetBits
Dim teacherRs, allTeacherRs
Dim teacherId, teacherName, teacherStatus
Dim bitValue, bitPosition, isChecked

lessonId = Request.QueryString("id")
isEdit = (lessonId <> "")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=attendanceDB"


' 저장 눌렀을 때
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    lessonId = Request.Form("lessonId")
    title = Request.Form("title")
    trainingDate = Request.Form("trainingDate")
    memo = Request.Form("memo")

    ' 날짜 input 값 Access에 넣기 좋게 바꾸기
    dateParts = Split(trainingDate, "-")
    dbDate = "#" & dateParts(1) & "/" & dateParts(2) & "/" & dateParts(0) & "#"

    selectedTeacherIds = ","

    For i = 1 To Request.Form("targetTeacher").Count
        selectedTeacherIds = selectedTeacherIds & Request.Form("targetTeacher")(i) & ","
    Next

    oldTargetBits = ""

    If lessonId <> "" Then
        Set rs = conn.Execute("SELECT [참여대상] FROM lesson_list WHERE ID = " & lessonId)

        If Not IsNull(rs("참여대상")) Then
            oldTargetBits = rs("참여대상")
        End If

        rs.Close
        Set rs = Nothing
    End If


    ' 체크한 사람 기준으로 참여대상 10101 이런거 만듦
    ' 수정이면 휴직/퇴직은 예전값 그대로 두고
    targetBits = ""

    Set allTeacherRs = conn.Execute("SELECT ID, 재직상태 FROM teachers ORDER BY ID")

    Do While Not allTeacherRs.EOF

        teacherId = CLng(allTeacherRs("ID"))
        teacherStatus = allTeacherRs("재직상태")

        Do While Len(targetBits) < teacherId - 1
            bitPosition = Len(targetBits) + 1

            If lessonId <> "" And Len(oldTargetBits) >= bitPosition Then
                targetBits = targetBits & Mid(oldTargetBits, bitPosition, 1)
            Else
                targetBits = targetBits & "0"
            End If
        Loop

        If teacherStatus = "재직" Then

            If InStr(selectedTeacherIds, "," & teacherId & ",") > 0 Then
                bitValue = "1"
            Else
                bitValue = "0"
            End If

        Else

            If lessonId <> "" And Len(oldTargetBits) >= teacherId Then
                bitValue = Mid(oldTargetBits, teacherId, 1)
            Else
                bitValue = "0"
            End If

        End If

        targetBits = targetBits & bitValue

        allTeacherRs.MoveNext
    Loop

    allTeacherRs.Close
    Set allTeacherRs = Nothing

    If lessonId <> "" And Len(oldTargetBits) > Len(targetBits) Then
        targetBits = targetBits & Mid(oldTargetBits, Len(targetBits) + 1)
    End If


    If lessonId = "" Then

        ' 새 연수 등록
        sql = "INSERT INTO lesson_list ([연수제목], [연수일시], [참여대상], [메모]) VALUES (" & _
              "'" & title & "', " & dbDate & ", '" & targetBits & "', '" & memo & "')"

        conn.Execute sql

    Else

        ' 기존 연수 수정
        sql = "UPDATE lesson_list SET " & _
              "[연수제목] = '" & title & "', " & _
              "[연수일시] = " & dbDate & ", " & _
              "[참여대상] = '" & targetBits & "', " & _
              "[메모] = '" & memo & "' " & _
              "WHERE [ID] = " & lessonId

        conn.Execute sql

    End If

    conn.Close
    Set conn = Nothing

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

    title = rs("연수제목")

    trainingDate = Year(rs("연수일시")) & "-" & _
                   Right("0" & Month(rs("연수일시")), 2) & "-" & _
                   Right("0" & Day(rs("연수일시")), 2)

    If Not IsNull(rs("메모")) Then
        memo = rs("메모")
    End If

    If Not IsNull(rs("참여대상")) Then
        targetBits = rs("참여대상")
    End If

    rs.Close
    Set rs = Nothing

End If


' 재직인 썜만
sql = "SELECT ID, 이름 FROM teachers WHERE 재직상태 = '재직' ORDER BY 이름"
Set teacherRs = Server.CreateObject("ADODB.Recordset")
teacherRs.Open sql, conn

%>

<!DOCTYPE html>
<html lang="ko">

<head>
<meta charset="UTF-8">

<% If isEdit Then %>
<title>연수 수정</title>
<% Else %>
<title>연수 등록</title>
<% End If %>

</head>

<body>

<% If isEdit Then %>
<h1>연수 수정</h1>
<% Else %>
<h1>새 연수 등록</h1>
<% End If %>

<form method="post" action="lesson_register.asp">

<input type="hidden" name="lessonId" value="<%= lessonId %>">

<p>
<label>연수명</label><br>
<input type="text" name="title" value="<%= title %>" required>
</p>

<p>
<label>날짜</label><br>
<input type="date" name="trainingDate" value="<%= trainingDate %>" required>
</p>

<p>
<label>메모</label><br>
<textarea name="memo" rows="5" cols="50"><%= memo %></textarea>
</p>

<details>

<summary>대상자 선택</summary>

<p>
<input type="checkbox" id="selectAll"<% If Not isEdit Then Response.Write " checked" %>>
<label for="selectAll">전체 선택</label>
</p>

<%
Do While Not teacherRs.EOF

    teacherId = teacherRs("ID")
    teacherName = teacherRs("이름")

    If isEdit Then
        isChecked = (Len(targetBits) >= teacherId And Mid(targetBits, teacherId, 1) = "1")
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
<%= teacherName %>
</label>
</p>

<%
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

</script>

</body>
</html>

<%
teacherRs.Close
Set teacherRs = Nothing

conn.Close
Set conn = Nothing
%>
