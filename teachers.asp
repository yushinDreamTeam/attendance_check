<%@ Language=VBScript CodePage="65001"%>
<%
Response.Buffer = True
Function StrValue(v)
    Dim value
    If IsObject(v) Then value = v.Value Else value = v
    If IsNull(value) Or IsEmpty(value) Then StrValue = "" Else StrValue = CStr(value)
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

Function RunSQL(c, sql, params)
    Dim cmd, affected
    Set cmd = NewCommand(c, sql, params)
    cmd.Execute affected, , 128
    RunSQL = affected
End Function

Sub FailPage(status, message)
    Response.Status = status
    Response.Write "<meta charset='utf-8'><p>" & H(message) & "</p><a href='javascript:history.back()'>돌아가기</a>"
    Response.End
End Sub

Function H(v)
    H = Server.HTMLEncode(StrValue(v))
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
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, updateRs, teacherId, teacherName, status
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    ' POST는 화면 출력 전에 처리
    Dim msg, transactionOpen, data, n, newName, newStatus
    Sub SaveTeachers()
        conn.BeginTrans
        transactionOpen = True
        Set updateRs = conn.Execute("SELECT ID FROM teachers ORDER BY ID")
        If Not updateRs.EOF Then
            data = updateRs.GetRows()
            updateRs.Close
            For n = 0 To UBound(data,2)
                teacherId = CLng(data(0,n))
                newName = Trim(StrValue(Request.Form("teacherName_" & teacherId)))
                newStatus = StrValue(Request.Form("status_" & teacherId))
                If newName = "" Or Len(newName)>255 Then Err.Raise vbObjectError+1,"교직원","이름을 확인해 주세요."
                If newStatus<>"재직" And newStatus<>"휴직" And newStatus<>"퇴직" Then Err.Raise vbObjectError+2,"교직원","재직 상태를 확인해 주세요."
                Call RunSQL(conn,"UPDATE teachers SET [이름]=?,[재직상태]=? WHERE ID=?",Array(ParamText(newName),ParamText(newStatus),ParamID(teacherId)))
            Next
        Else
            updateRs.Close
        End If
        conn.CommitTrans
        transactionOpen = False
    End Sub
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        Application.Lock
        On Error Resume Next
        SaveTeachers
        msg = Err.Description
        If transactionOpen Then conn.RollbackTrans
        conn.Close
        Err.Clear
        On Error GoTo 0
        Application.UnLock
        If msg<>"" Then FailPage "400 Bad Request","저장하지 못했습니다. " & msg
        Response.Redirect "teachers.asp"
    End If

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
                            teacherName = StrValue(rs.Fields("이름"))
                            status = StrValue(rs.Fields("재직상태"))
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



        <!-- 교직원 추가 버튼
        <button type="button" onclick="location.href='teacher_add.asp'">교직원 추가</button> -->
    </div>
</body>

<%
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
