<%@ Language="VBScript" CodePage="65001" %>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>제출 완료</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background-color: #f5faf7;
        }

        .success-box {
            text-align: center;
            background: white;
            border: 1px solid #cfe3d9;
            border-radius: 12px;
            padding: 40px 60px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }

        h2 {
            margin-bottom: 10px;
            color: #2a5b4c;
        }

        p {
            margin: 0;
            color: #4d655f;
        }
    </style>
</head>
<body>
    <div class="success-box">
        <h2>제출 완료!</h2>
        <p>서명이 정상적으로 저장되었습니다.</p>
    </div>
</body>
</html>
