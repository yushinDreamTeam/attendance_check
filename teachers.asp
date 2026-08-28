<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, updateRs, teacherId, teacherName, employmentStatus
    Dim action, deleteId, newName, newStatus
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        action = Request.Form("action")
        deleteId = Request.Form("deleteId")

        If action = "add" Then
            newName = Trim(Request.Form("newName"))
            newStatus = Request.Form("newStatus")

            If newStatus <> "현역" And newStatus <> "휴가" And newStatus <> "은퇴" Then
                newStatus = "현역"
            End If

            If newName <> "" Then
                sql = "INSERT INTO teachers (이름, 재직상태) VALUES ('" & _
                      Replace(newName, "'", "''") & "', '" & _
                      Replace(newStatus, "'", "''") & "')"
                conn.Execute sql
            End If

            Response.Redirect "teachers.asp?added=1"
        ElseIf deleteId <> "" And IsNumeric(deleteId) Then
            conn.Execute "DELETE FROM teachers WHERE ID = " & CLng(deleteId)
            Response.Redirect "teachers.asp?deleted=1"
        End If

        Set updateRs = conn.Execute("SELECT ID FROM teachers")

        Do While Not updateRs.EOF
            teacherId = updateRs("ID")
            teacherName = Trim(Request.Form("teacherName_" & teacherId))
            employmentStatus = Request.Form("status_" & teacherId)

            If employmentStatus <> "현역" And employmentStatus <> "휴가" And employmentStatus <> "은퇴" Then
                employmentStatus = "현역"
            End If

            sql = "UPDATE teachers SET 이름 = '" & Replace(teacherName, "'", "''") & "', " & _
                  "재직상태 = '" & Replace(employmentStatus, "'", "''") & "' " & _
                  "WHERE ID = " & teacherId
            conn.Execute sql
            updateRs.MoveNext
        Loop

        updateRs.Close
        Set updateRs = Nothing
        Response.Redirect "teachers.asp?saved=1"
    End If

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
    <form method="post" action="teachers.asp">
        <input type="hidden" name="action" value="add">
        <input type="text" name="newName" placeholder="교직원 이름" required>
        <select name="newStatus">
            <option value="현역">현역</option>
            <option value="휴가">휴가</option>
            <option value="은퇴">은퇴</option>
        </select>
        <button type="submit">교직원 추가</button>
    </form>
    <form method="post" action="teachers.asp">
        <button type="submit" name="action" value="save" id="save">저장</button>
    <div style="width:530px; height:450px; overflow-y:scroll;">
        <table>
            <thead>
                <tr>
                    <th>교직원 이름</th>
                    <th>재직상태</th>
                    <th>관리</th>
                </tr>
            </thead>
            <tbody>
                <%
                IF rs.EOF Then
                            <input type="text" name="teacherName_<%= rs("ID") %>" value="<%= Server.HTMLEncode(rs("이름")) %>">
                    <tr>
                        <td colspan="2">등록된 교직원이 없습니다.</td>
                            <input type="radio" name="status_<%= rs("ID") %>" value="현역"<% If rs("재직상태") = "현역" Then Response.Write " checked" %>>
                            <label>현역</label>
                            <input type="radio" name="status_<%= rs("ID") %>" value="휴가"<% If rs("재직상태") = "휴가" Then Response.Write " checked" %>>
                            <label>휴가</label>
                            <input type="radio" name="status_<%= rs("ID") %>" value="은퇴"<% If rs("재직상태") = "은퇴" Then Response.Write " checked" %>>
                            <label>은퇴</label>
                        </td>
                        <td>
                            <button type="submit" name="deleteId" value="<%= rs("ID") %>" onclick="return confirm('이 교직원을 삭제하시겠습니까?');">삭제</button>
                    <tr>
                        <td>
                            <input type="text" value="<%= rs("이름") %>">
                        </td>
                        <td>
                            <form>
                                <input type="radio" name="status" value="재직" <%= IIf(rs("재직상태") = "재직", "checked", "") %>>
                                <label>재직</label>
                                <input type="radio" name="status" value="휴직" <%= IIf(rs("재직상태") = "휴직", "checked", "") %>>
                                <label>휴직</label>
                            </form>
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
    </form>
</body>

<%
conn.Close
Set rs = Nothing
Set conn = Nothing
%>