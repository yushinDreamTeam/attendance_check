<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

Dim conn, rs, sql
Dim lessonId, isEdit
Dim title, trainingDate, memo
Dim dateParts, dbDate
Dim selectedTeacherIds, i
Dim teacherIds, teacherNames

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

    ' 날짜 input 값 Access에 넣기 좋게 바꿈
    dateParts = Split(trainingDate, "-")
    dbDate = "#" & dateParts(1) & "/" & dateParts(2) & "/" & dateParts(0) & "#"

    ' 대상자 저장 방식 미정, 체크된 ID들만 받아두고
    selectedTeacherIds = ""

    For i = 1 To Request.Form("targetTeacher").Count
        If selectedTeacherIds <> "" Then
            selectedTeacherIds = selectedTeacherIds & ","
        End If

        selectedTeacherIds = selectedTeacherIds & Request.Form("targetTeacher")(i)
    Next

    If lessonId = "" Then

        ' 새 연수 등록
        sql = "INSERT INTO lesson_list ([연수제목], [연수일시]) VALUES (" & _
              "'" & title & "', " & dbDate & ")"

        conn.Execute sql

    Else

        ' 기존 연수 수정
        sql = "UPDATE lesson_list SET " & _
              "[연수제목] = '" & title & "', " & _
              "[연수일시] = " & dbDate & " " & _
              "WHERE [ID] = " & lessonId

        conn.Execute sql

    End If

    ' 메모 어디다 저장하는지 아직 몰라서 걍 받아만둠
    ' memo

    ' 대상자도 대강 똑가ㅇㅁ
    ' selectedTeacherIds

    conn.Close
    Set conn = Nothing

    Response.Redirect "administer.asp"
End If


' 수정으로 들어왔으면 기존 값 가져옴
title = ""
trainingDate = ""
memo = ""

If isEdit Then

    sql = "SELECT * FROM lesson_list WHERE ID = " & lessonId
    Set rs = conn.Execute(sql)

    title = rs("연수제목")

    trainingDate = Year(rs("연수일시")) & "-" & _
                   Right("0" & Month(rs("연수일시")), 2) & "-" & _
                   Right("0" & Day(rs("연수일시")), 2)

    ' 메모 컬럼 확인되면 여기서 가져오면 됨
    ' memo = rs("메모")

    rs.Close
    Set rs = Nothing

End If


' 선생님 DB 어떻게 대상자로 저장할지 아직 몰라서 일단 껍데기
teacherIds = Array(1001, 1002, 1003, 1004, 1005, 1006)
teacherNames = Array("강민식", "김민수", "박영희", "이성빈", "최진형", "홍길동")

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
<input type="checkbox" id="selectAll" checked>
<label for="selectAll">전체 선택</label>
</p>

<%
For i = 0 To UBound(teacherIds)
%>

<p>
<input
    type="checkbox"
    class="teacherCheck"
    name="targetTeacher"
    value="<%= teacherIds(i) %>"
    id="teacher_<%= teacherIds(i) %>"
    checked
>

<label for="teacher_<%= teacherIds(i) %>">
<%= teacherNames(i) %>
</label>
</p>

<%
Next
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
conn.Close
Set conn = Nothing
%>
