<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>연수 등록 | 교직원 연수 출석</title>
</head>
<body>
  <main class="page">
    <div class="page-heading">
      <div><h1>새 연수 등록</h1></div>
      <a class="btn secondary" href="admin.asp">목록으로</a>
    </div>

    <!-- 실제 DB 연결 시 action을 training-save.asp로 두고, 시연용 JS submit 방지를 제거합니다. -->
    <form id="trainingCreateForm" class="card form-card" action="training-save.asp" method="post">
      <section class="section">
        <div class="section-title-row"><div><h2>기본 정보</h2></div></div>
        <div class="grid two">
          <div class="field full"><label for="title">연수 제목 <span class="required">*</span></label><input class="control" id="title" name="title" required maxlength="100" placeholder="예: 2026학년도 아동학대 예방교육"></div>
          <div class="field"><label for="trainingDate">연수 날짜 <span class="required">*</span></label><input class="control" id="trainingDate" name="trainingDate" type="date" required></div>
          <div class="field"><label for="trainingTime">시작 시간</label><input class="control" id="trainingTime" name="trainingTime" type="time"></div>
          <div class="field full"><label for="place">장소</label><input class="control" id="place" name="place" placeholder="예: 본관 2층 시청각실"></div>
          <div class="field full"><label for="description">연수 내용</label><textarea class="control" id="description" name="description" placeholder="연수 목적, 진행 내용, 안내사항 등을 입력하세요."></textarea></div>
        </div>
      </section>

      <section class="section">
        <div class="section-title-row"><div><h2>참석 대상</h2><p class="section-help">필요한 대상을 여러 개 선택할 수 있습니다.</p></div></div>
        <div class="choice-grid" aria-label="참석 대상">
          <div class="choice"><input class="audience-check" id="includeStaff" name="includeStaff" type="checkbox" value="1" checked><label for="includeStaff"><strong>일반 교직원 전체</strong><span>교사와 일반 교직원</span></label></div>
          <div class="choice"><input class="audience-check" id="includeLeadership" name="includeLeadership" type="checkbox" value="1"><label for="includeLeadership"><strong>교장·교감 추가</strong><span>관리자 2명 포함</span></label></div>
          <div class="choice"><input class="audience-check" id="includeExternal" name="includeExternal" type="checkbox" value="1"><label for="includeExternal"><strong>외부인 추가</strong><span>인원과 메모 직접 입력</span></label></div>
        </div>

        <!-- 나중에 DB가 연결되면 아래 인원수는 재직자 명단에서 가져오면 될 듯. -->
        <input id="generalStaffCount" type="hidden" value="46">
        <input id="leadershipCount" type="hidden" value="2">
        <input id="expectedCount" name="expectedCount" type="hidden" value="46">

        <div id="externalFields" class="grid two" style="margin-top:18px" hidden>
          <div class="field"><label for="externalCount">외부 참석 인원</label><input class="control" id="externalCount" name="externalCount" type="number" min="0" value="0"></div>
          <div class="field"><label for="externalNote">외부 참석자 메모</label><input class="control" id="externalNote" name="externalNote" placeholder="예: 교육지원청 장학사 2명"></div>
        </div>
        <div class="summary-strip" aria-live="polite"><span class="summary-pill">교직원 <strong data-staff-count>46</strong>명</span><span class="summary-pill">외부인 <strong data-external-count>0</strong>명</span><span class="summary-pill">예상 참석 <strong data-total-count>46</strong>명</span></div>
      </section>

      <section id="qrPreview" class="section">
        <div class="section-title-row"><div><h2>모바일 서명 링크</h2></div></div>
        <input id="baseUrl" type="hidden" value="http://192.168.0.10/training">
        <input id="publicToken" name="publicToken" type="hidden" value="tr-20260824-a7f2">
        <div class="link-maker">
          <div><strong>입력을 마쳤으면 URL과 QR을 만들어 주세요.</strong><p>만들어진 QR을 화면에 띄우면 참석자가 휴대폰으로 들어와 서명하는 방식입니다.</p></div>
          <button class="btn primary" id="makeLink" type="button">URL·QR 만들기</button>
        </div>
        <div id="linkResult" hidden>
          <div class="field"><label for="attendanceUrl">참석 URL</label><div class="url-box"><input class="control" id="attendanceUrl" readonly><button class="btn secondary" id="copyUrl" type="button">복사</button></div><span class="hint">서버 IP는 배포 환경의 고정 주소 또는 호스트명으로 설정하세요.</span></div>
          <div style="display:flex;justify-content:center"><div id="qrCode" class="qr-box" aria-label="모바일 서명용 QR 코드"></div></div>
        </div>
      </section>

      <div class="actions"><a class="btn secondary" href="admin.asp">취소</a><button class="btn primary" type="submit">연수 등록</button></div>
    </form>
  </main>
  <div id="toast" class="toast" role="status"></div>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script src="assets/training-form.js"></script>
</body>
</html>
