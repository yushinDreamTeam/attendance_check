<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, updateRs, teacherId, teacherName, status
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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>교직원 관리</title>
    <link rel="stylesheet" href="style.css">
</head>
<body class="teachers-page">
    <h2>교직원 관리</h2>
    <a href="administer.asp">연수 목록으로</a>

    <p>교직원 추가시 ms access를 통해 새로운 행을 만들어 추가하면 됩니다(수정도 가능)</p>
    <p style="color: red;"><strong>*주의: Access파일로 데이터 추가시 기존 교직원 정보는 절대로 수정하거나 삭제하지 마세요. 삭제 대신 "퇴직"으로 처리해주시면 됩니다</strong></p>

    <div style="width:530px; height:450px; overflow-y:scroll;">
        <form method="post" action="teachers.asp">
            <table>
                <thead>
                    <tr>
                        <th>이름</th>
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
                                <input type="hidden" name="teacherId_<%= teacherId %>" value="<%= Server.HTMLEncode(teacherId) %>">
                                <input type="text" name="teacherName_<%= teacherId %>" value="<%= Server.HTMLEncode(teacherName) %>">
                            </td>
                            <td>
                                <label>
                                    <input type="radio" name="status_<%= teacherId %>" value="재직"
                                        <% If status = "재직" Then Response.Write "checked" %>>
                                        재직
                                </label>
                                <label>
                                    <input type="radio" name="status_<%= teacherId %>" value="휴직"
                                        <% If status = "휴직" Then Response.Write "checked" %>>
                                        휴직
                                </label>
                                <label>
                                    <input type="radio" name="status_<%= teacherId %>" value="퇴직"
                                        <% If status = "퇴직" Then Response.Write "checked" %>>
                                        퇴직
                                </label>
                            </td>
                        </tr>
                    <%
                            rs.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>

            <button type="submit" name="save" value="1">저장</button>
        </form>

        <%
        If Request.Form("save") = "1" Then
            Dim key, itemId, newName, newStatus
            For Each key In Request.Form
                If Left(key, 12) = "teacherName_" Then
                    itemId = Mid(key, 13)
                    newName = Trim(Request.Form(key))
                    newStatus = Request.Form("status_" & itemId)

                    If newName <> "" Then
                        sql = "UPDATE teachers SET " & _
                        "이름 = '" & Replace(newName, "'", "''") & "', " & _
                        "재직상태 = '" & Replace(newStatus, "'", "''") & "' " & _
                        "WHERE ID = " & itemId
                        conn.Execute sql
                    End If

                End If
            Next

            Response.Redirect "teachers.asp"
        End If
        %>

        <!-- 교직원 추가 버튼
        <button type="button" onclick="location.href='teacher_add.asp'">교직원 추가</button> -->
    </div>
</body>

<%
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
