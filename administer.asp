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
    <link rel="stylesheet" href="admin.css">
</head>

<body>
    <h2>연수 관리</h2>

    <!--일단 지금 검색기능은 못할듯
    <input type="text" size="30">
    <button>검색</button>
    --->
    
    <a id = "teachersAdmin" href="teachers.asp">교직원 관리</a>
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
                            <button type="button" onclick="location.href='printing.asp?id=<%= Server.URLEncode(editId) %>'">출력</button>
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

    <button style="font-size:25px;" onclick="location.href='lesson_register.asp'">+</button>

</body>
</html>

<%
    conn.Close
    set conn = nothing
    set rs = nothing
%>

