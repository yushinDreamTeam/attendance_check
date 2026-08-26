<%@ Language=VBScript %>
<%
    Dim conn, sql, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendance"

    sql = "SELECT * FROM #여기에 DSN이름 ORDER BY id DESC"
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
    
    <div style="width:530px; height:450px; overflow-y:scroll;">
        <table>
            <thead>
                <tr>
                    <th>연수 제목</th>
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
                %>
                    <tr>
                        <td>
                            <b><%= rs("title") %></b><br>
                            <%= rs("date") %>
                        </td>
                        <td>
                            <button onclick="">수정</button>
                            <button onclick="">제거</button>
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
%>