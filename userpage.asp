<%@ Language="VBScript" CodePage="65001" %>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"
    Dim conn, sql, rs, lessonID
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    lessonID = Request.QueryString("id")
    If Not IsNumeric(lessonID) Then
        Response.Status = "400 Bad Request"
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    sql = "SELECT * FROM lesson_list WHERE ID = " & CLng(lessonID)
    Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>출석</title>

    <style>
        #signature-canvas {
            border: 2px solid black;
            background-color: white;
            touch-action: none;
        }
    </style>
</head>

<body>

    <h1>출석</h1>

    <form id="attendance-form">

        <label for="user-name">이름</label>

        <input
            type="text"
            id="user-name"
            required
        >

        <br><br>

        <label>서명</label>

        <br>

        <canvas
            id="signature-canvas"
            width="600"
            height="200">
        </canvas>

        <br>

        <button
            type="button"
            id="clear-signature-button">
            서명 지우기
        </button>

        <button type="submit">
            제출
        </button>

    </form>

    <script src="./userpage.js"></script>

</body>
</html>
