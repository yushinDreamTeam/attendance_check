<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

Dim trainingId, conn, rs, sql, trainingTitle
trainingId = Request.QueryString("id")

If Not IsNumeric(trainingId) Then
  Response.Status = "400 Bad Request"
  Response.Write "잘못된 연수 ID입니다."
  Response.End
End If

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=attendanceDB"
sql = "SELECT [연수제목] FROM [lesson_list] WHERE [ID] = " & CLng(trainingId)
Set rs = conn.Execute(sql)

If rs.EOF Then
  rs.Close
  conn.Close
  Set rs = Nothing
  Set conn = Nothing
  Response.Status = "404 Not Found"
  Response.Write "해당 연수를 찾을 수 없습니다."
  Response.End
End If

trainingTitle = rs("연수제목").Value
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>연수 상세 및 수정 | 교직원 연수 출석</title>
</head>
<body>
  <main class="page">
    <div class="page-heading">
      <div><h1>연수 상세</h1></div>
      <div class="detail-toolbar"><span class="status">진행 예정</span><a class="btn secondary" href="admin.asp">목록으로</a><button id="editButton" class="icon-btn" type="button" aria-label="연수 수정" aria-pressed="false" title="수정 잠금 해제">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 20h4l10.5-10.5a2.83 2.83 0 0 0-4-4L4 16v4Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="m13.5 6.5 4 4" stroke="currentColor" stroke-width="1.8"/></svg>
      </button></div>
    </div>

    <div class="detail-layout">
      <!-- 실제 DB 연결 시 trainingId로 조회한 값을 각 value와 data-view-for 영역에 출력합니다. -->
      <form id="trainingDetailForm" class="card detail-card" action="training-update.asp" method="post">
        <input name="trainingId" type="hidden" value="<%= Server.HTMLEncode(trainingId) %>">
        <input id="generalStaffCount" type="hidden" value="46">
        <input id="leadershipCount" type="hidden" value="2">
        <input id="expectedCount" name="expectedCount" type="hidden" value="50">
        <section class="section">
          <div class="section-title-row"><div><h2>기본 정보</h2></div></div>
          <div class="grid two">
            <div class="field full"><span class="field-label">연수 제목</span><div class="view-value" data-view-for="editTitle"><%= Server.HTMLEncode(trainingTitle) %></div><input class="control edit-control" id="editTitle" name="title" required value="<%= Server.HTMLEncode(trainingTitle) %>" hidden></div>
            <div class="field"><span class="field-label">연수 날짜</span><div class="view-value" data-view-for="editDate">2026-08-28</div><input class="control edit-control" id="editDate" name="trainingDate" type="date" value="2026-08-28" hidden></div>
            <div class="field"><span class="field-label">시작 시간</span><div class="view-value" data-view-for="editTime">15:30</div><input class="control edit-control" id="editTime" name="trainingTime" type="time" value="15:30" hidden></div>
            <div class="field full"><span class="field-label">장소</span><div class="view-value" data-view-for="editPlace">본관 2층 시청각실</div><input class="control edit-control" id="editPlace" name="place" value="본관 2층 시청각실" hidden></div>
            <div class="field full"><span class="field-label">연수 내용</span><div class="view-value" data-view-for="editDescription">아동학대 신고 의무자 교육 및 학교 현장 대응 절차 안내</div><textarea class="control edit-control" id="editDescription" name="description" hidden>아동학대 신고 의무자 교육 및 학교 현장 대응 절차 안내</textarea></div>
          </div>
        </section>

        <section class="section">
          <div class="section-title-row"><div><h2>참석 대상</h2><p class="section-help">예상 참석 <strong data-total-count>50</strong>명</p></div></div>
          <div class="view-only">
            <div class="notice"><strong>일반 교직원 46명, 교장·교감 2명</strong>과 외부 참석자 2명이 대상입니다.</div>
            <div class="summary-strip"><span class="summary-pill">교직원 <strong>48</strong>명</span><span class="summary-pill">외부인 <strong>2</strong>명</span><span class="summary-pill">대상 기준 <strong>2026-08-24 재직자</strong></span></div>
          </div>
          <div class="edit-only edit-grid" hidden>
            <div class="choice-grid" aria-label="참석 대상">
              <div class="choice"><input class="audience-check" id="includeStaff" name="includeStaff" type="checkbox" value="1" checked><label for="includeStaff"><strong>일반 교직원 전체</strong><span>교사와 일반 교직원</span></label></div>
              <div class="choice"><input class="audience-check" id="includeLeadership" name="includeLeadership" type="checkbox" value="1" checked><label for="includeLeadership"><strong>교장·교감 추가</strong><span>관리자 2명 포함</span></label></div>
              <div class="choice"><input class="audience-check" id="includeExternal" name="includeExternal" type="checkbox" value="1" checked><label for="includeExternal"><strong>외부인 추가</strong><span>인원과 메모 직접 입력</span></label></div>
            </div>
            <div id="externalFields" class="grid two"><div class="field"><label for="externalCount">외부 참석 인원</label><input class="control" id="externalCount" name="externalCount" type="number" min="0" value="2"></div><div class="field"><label for="externalNote">외부 참석자 메모</label><input class="control" id="externalNote" name="externalNote" value="교육지원청 장학사 2명"></div></div>
          </div>
        </section>

        <div class="actions edit-only" hidden><button id="cancelEdit" class="btn secondary" type="button">취소</button><button class="btn primary" type="submit">수정 내용 저장</button></div>
      </form>

      <aside class="card qr-card">
        <h2>모바일 서명 QR</h2>
        <p class="section-help">참석자가 휴대폰으로 스캔합니다.</p>
        <input id="baseUrl" type="hidden" value="http://192.168.0.10/training">
        <input id="publicToken" type="hidden" value="tr-20260824-a7f2">
        <div id="qrCode" class="qr-box" aria-label="모바일 서명용 QR 코드"></div>
        <div class="field" style="text-align:left"><label for="attendanceUrl">참석 URL</label><div class="url-box"><input class="control" id="attendanceUrl" readonly><button class="btn secondary" id="copyUrl" type="button">복사</button></div></div>
        <div class="qr-actions"><button class="btn ghost" id="refreshQr" type="button">QR 새로고침</button><button class="btn primary" id="downloadQr" type="button">PNG 저장</button></div>
      </aside>
    </div>
  </main>
  <div id="toast" class="toast" role="status"></div>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script src="assets/training-form.js"></script>
</body>
</html>
