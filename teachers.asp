<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, updateRs, teacherId, teacherName, status
    Dim action, deleteId, newName, newStatus
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    sql = "SELECT * FROM teachers ORDER BY 이름"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>교직원 관리</title>
</head>
<body>
    <h2>교직원 관리</h2>
    <a href="administer.asp">연수 목록으로</a>

    <p>교직원 추가시 ms access를 통해 새로운 행을 만들어 추가하면 됩니다(수정도 가능)</p>
    <p style="color: red;"><strong>*주의: 기존 교직원 정보는 절대로 수정하거나 삭제하지 마세요. 삭제 대신 "퇴직"으로 처리해주시면 됩니다</strong></p>

    <div style="width:530px; height:450px; overflow-y:scroll;">
        <table>
            <thead>
                <tr>
                    <th>교직원 목록</th>
                    <th>재직 상태</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="2">등록된 교직원이 없습니다.</td>
                    </tr>
                <%
                Else
                    Do While Not rs.EOF
                        teacherId = rs.Fields("ID").Value
                        teacherName = rs.Fields("이름").Value
                        status = rs.Fields("재직상태").Value
                %>
                    <tr>
                        <td>
                            <b><%= teacherName %></b><br>
                            재직상태: <%= status %>
                        </td>
                        <td>
                            <button type="button" onclick="location.href='teacher_edit.asp?id=<%= Server.URLEncode(teacherId) %>'">수정</button>
                        </td>
                    </tr>
                <%
                        rs.MoveNext
                    Loop
                End If
                %>
            </tbody>
        </table>

        <!-- 교직원 추가 버튼 -->
        <button type="button" onclick="location.href='teacher_add.asp'">교직원 추가</button>
    </div>
</body>

<%
conn.Close
Set rs = Nothing
Set conn = Nothing
%>