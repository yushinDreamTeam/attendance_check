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
    <title> 출석 </title>
</head>
<body>
    <main id="user-page">
        <!-- 연수 제목 표시 영역 -->
        <section id="training-info">
            <h1 id="training-title"><%= rs("연수제목") %></h1>
            <p id="training-date"><%= rs("연수일시") %></p>
        </section>

        <!-- 사용자 입력 -->
        <section id="user-input-section">
            <h2>참여자 이름 입력</h2>

            <!-- 이름 입력 폼 -->
            <form id="attendance-form">
                <div class="form-row">
                    <label for="user-name">이름</label>
                    <input 
                        type="text" 
                        id="user-name" 
                        name="userName" 
                        placeholder="이름을 입력하세요"
                        required
                    >
                </div>

                <!-- 디지털 서명 입력 -->
                <div class="form-row">
                    <label for="signature-canvas">디지털 서명</label>

                    <!-- 이쯤에 서명 -->

                    <!-- 서명 관련 버튼 -->
                    <div class="signature-buttons">
                        <button type="button" id="clear-signature-button">
                            서명 지우기
                        </button>
                    </div>
                </div>

                <!-- 제출 버튼 -->
                <div class="form-row">
                    <button type="submit" id="submit-button">
                        제출
                    </button>
                </div>
            </form>
        </section>

        <!-- 제출 결과 안내 -->
        <section id="submit-result" hidden>
            <h2>제출 결과</h2>
            <p id="submit-message">제출이 완료되었습니다.</p>
        </section>

    </main>
</body>
</html>
