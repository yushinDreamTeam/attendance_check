<%@ Language="VBScript" CodePage="65001" %>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, editId
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    sql = "SELECT * FROM lesson_list ORDER BY id DESC"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>연수 관리</title>
</head>

<body>
    <h2>연수 관리</h2>

    <!--일단 지금 검색기능은 못할듯
    <input type="text" size="30">
    <button>검색</button>
    --->
    
    <a id = "teachersAdmin" href="teachers.asp">교직원 목록 관리</a>
    <div style="width:530px; height:450px; overflow-y:scroll;">
        <table>
            <thead>
                <tr>
                    <th>연수 목록</th>
                    <th>관리</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="2">등록된 연수가 없습니다.</td>
                    </tr>
                <%
                Else
                    Do While Not rs.EOF
                    editId = rs.Fields("ID").Value
                %>
                    <tr>
                        <td>
                            <b><%= rs("연수제목") %></b><br>
                            <%= rs("연수일시") %>
                        </td>
                        <td>
                            <button type="button" onclick="location.href='lesson_register.asp?id=<%= Server.URLEncode(editId) %>'">수정</button>
                        </td>
                    </tr>
                <%
                rs.MoveNext
                Loop
                End If
                %>
            </tbody>
        </table>
    </div>

    <br>

    <button style="font-size:25px;">+</button>

</body>
</html>

<%
    conn.Close
    set conn = nothing
    set rs = nothing
%>





#디자인

<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>연수 관리</title>

<style>
body {
    margin: 0;
    font-family: Arial;
    background-color: #f5faf7;
}

h2 {
    margin: 0;
    padding: 22px 30px;
    background-color: white;
    border-bottom: 1px solid #cccccc;
}

.main {
    width: 600px;
    margin: 30px auto;
}

.search {
    margin-bottom: 15px;
}

.search input {
    width: 470px;
    height: 35px;
    padding-left: 10px;
    border: 1px solid #bbbbbb;
}

.search button {
    width: 80px;
    height: 39px;
    background-color: #4c9b80;
    color: white;
    border: 0;
}

.list {
    width: 100%;
    height: 450px;
    overflow-y: scroll;
    background-color: white;
    border: 1px solid #bbbbbb;
}

.item {
    padding: 15px;
    height: 55px;
    border-bottom: 1px solid #dddddd;
}

.name {
    font-weight: bold;
    margin-bottom: 8px;
}

.date {
    font-size: 13px;
    color: #666666;
}

.buttons {
    float: right;
    margin-top: -37px;
}

.edit {
    padding: 6px 12px;
    background-color: white;
    border: 1px solid #4c9b80;
    color: #31745e;
}

.delete {
    padding: 6px 12px;
    background-color: white;
    border: 1px solid #cc7777;
    color: #a33f3f;
}

.add {
    float: right;
    margin-top: 15px;
    width: 50px;
    height: 50px;
    font-size: 28px;
    background-color: #4c9b80;
    color: white;
    border: 0;
    border-radius: 25px;
}
</style>

</head>

<body>

<h2>연수 관리</h2>

<div class="main">

<div class="search">
<input type="text" placeholder="연수명 검색">
<button>검색</button>
</div>

<div class="list">

<div class="item">
<div class="name">AI 활용 교직원 연수</div>
<div class="date">2026-08-24</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">개인정보 보호 교육</div>
<div class="date">2026-08-18</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">학교폭력 예방 교육</div>
<div class="date">2026-08-10</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">응급처치 및 안전교육</div>
<div class="date">2026-07-22</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">교직원 성희롱 예방교육</div>
<div class="date">2026-07-11</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">정보보안 연수</div>
<div class="date">2026-06-30</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

<div class="item">
<div class="name">아동학대 예방교육</div>
<div class="date">2026-06-20</div>

<div class="buttons">
<button class="edit">수정</button>
<button class="delete">삭제</button>
</div>
</div>

</div>

<button class="add">+</button>

</div>

</body>
</html>
