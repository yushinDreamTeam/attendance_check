<%@ Language=VBScript %>

<%
'==================================================
' 연수 등록 페이지
' 기능:
' 1. 연수 등록 화면 출력
' 2. POST 요청 시 DB 저장
' 3. 대상자 출석부 생성 (예정)
'
' 실제 테이블명은 DB 담당자 확인 필요
'==================================================

Dim conn
Dim sql

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    Dim title
    Dim trainingDate
    Dim trainingTime
    Dim place
    Dim description
    Dim publicToken
    Dim expectedCount

    title = Request.Form("title")
    trainingDate = Request.Form("trainingDate")
    trainingTime = Request.Form("trainingTime")
    place = Request.Form("place")
    description = Request.Form("description")
    publicToken = Request.Form("publicToken")
    expectedCount = Request.Form("expectedCount")

    Set conn = Server.CreateObject("ADODB.Connection")

    ' 실제 환경에 맞게 수정 필요
    conn.Open "DSN=attendance"

    ' TODO:
    ' 실제 테이블명 확인 후 INSERT 작성
    sql = ""

    'conn.Execute sql

    ' TODO:
    ' 대상자 출석부 생성
    ' 1. 교직원 테이블 조회
    ' 2. 대상 조건 확인
    ' 3. Attendance 데이터 생성

    conn.Close
    Set conn = Nothing

    Response.Redirect("administer_page.asp")

End If
%>

<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>연수 등록 | 교직원 연수 출석</title>

<style>
body {
    font-family: Arial, sans-serif;
}

.page {
    max-width:900px;
    margin:auto;
}

.card {
    border:1px solid #ddd;
    padding:20px;
    border-radius:12px;
}

.field {
    margin-bottom:15px;
}

.control {
    width:100%;
    padding:10px;
}

button {
    padding:10px 15px;
}
</style>

</head>

<body>

<main class="page">

<h1>새 연수 등록</h1>

<form id="trainingCreateForm" class="card" method="post">

<h2>기본 정보</h2>

<div class="field">
<label>연수 제목</label>
<input class="control" name="title" required>
</div>

<div class="field">
<label>연수 날짜</label>
<input class="control" name="trainingDate" type="date" required>
</div>

<div class="field">
<label>시작 시간</label>
<input class="control" name="trainingTime" type="time">
</div>

<div class="field">
<label>장소</label>
<input class="control" name="place">
</div>

<div class="field">
<label>연수 내용</label>
<textarea class="control" name="description"></textarea>
</div>

<h2>참석 대상</h2>

<label>
<input type="checkbox" id="includeStaff" name="includeStaff" checked>
일반 교직원 전체
</label>

<br>

<label>
<input type="checkbox" id="includeLeadership" name="includeLeadership">
교장·교감 추가
</label>

<br>

<label>
<input type="checkbox" id="includeExternal" name="includeExternal">
외부 참석자 추가
</label>

<br><br>

<div class="field">
<label>외부 참석 인원</label>
<input class="control" id="externalCount" name="externalCount" type="number" value="0" min="0">
</div>

<input type="hidden" id="expectedCount" name="expectedCount" value="46">

<div>
예상 참석:
<strong id="totalCount">46</strong>명
</div>

<h2>모바일 서명 링크</h2>

<input type="hidden" id="publicToken" name="publicToken" value="tr-<%=Timer()%>">

<button type="button" id="makeLink">
URL·QR 만들기
</button>

<div id="qrArea" style="display:none;">
<p>참석 URL</p>
<input id="attendanceUrl" class="control" readonly>
<div id="qrCode"></div>
</div>

<br><br>

<button type="submit">
연수 등록
</button>

</form>

</main>

<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>

<script>

function updateCount(){

    let count = 0;

    if(document.getElementById("includeStaff").checked){
        count += 46;
    }

    if(document.getElementById("includeLeadership").checked){
        count += 2;
    }

    count += Number(document.getElementById("externalCount").value);

    document.getElementById("totalCount").innerHTML = count;
    document.getElementById("expectedCount").value = count;
}

document.querySelectorAll("input").forEach(function(item){
    item.addEventListener("change", updateCount);
});


document.getElementById("makeLink").addEventListener("click", function(){

    let token = document.getElementById("publicToken").value;

    let url = location.origin + "/training?token=" + token;

    document.getElementById("attendanceUrl").value = url;

    document.getElementById("qrArea").style.display = "block";

    document.getElementById("qrCode").innerHTML = "";

    new QRCode(
        document.getElementById("qrCode"),
        url
    );

});

</script>

</body>
</html>
