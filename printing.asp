<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, lessonID
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    lessonId = Request.QueryString("id")
    If Not IsNumeric(lessonId) Then
        Response.Status = "400 Bad Request"
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    Dim tableName = "signature" & lessonId
    sql = "SELECT * FROM [" & tableName & "]"
    Set rs = conn.Execute(sql)
%>