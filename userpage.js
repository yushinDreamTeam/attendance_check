const canvas = document.getElementById("signature-canvas");
const ctx = canvas.getContext("2d", {willReadFrequently:true});
const form = document.getElementById("attendance-form");
const submitButton = document.getElementById("submit-button");
const clearButton = document.getElementById("clear-signature-button");
const statusMessage = document.getElementById("submit-status");
let drawing = false;
let pointerId = null;
let saving = false;

// 선 설정
ctx.strokeStyle = "black";
ctx.lineWidth = 3;
ctx.lineCap = "round";
ctx.lineJoin = "round";

// 좌표 계산
function getPosition(event) {
    const rect = canvas.getBoundingClientRect();
    return {
        x: (event.clientX - rect.left) * canvas.width / rect.width,
        y: (event.clientY - rect.top) * canvas.height / rect.height
    };
}

// 서명 시작
canvas.addEventListener("pointerdown", function(event) {
    if(saving || drawing || !event.isPrimary || event.button!==0) return;
    event.preventDefault();
    drawing = true;
    pointerId = event.pointerId;
    canvas.setPointerCapture(event.pointerId);
    const position = getPosition(event);
    ctx.beginPath();
    ctx.moveTo(position.x, position.y);
});

// 서명 그리기
canvas.addEventListener("pointermove", function(event) {
    if (!drawing || event.pointerId!==pointerId || saving) return;
    const position = getPosition(event);
    ctx.lineTo(position.x, position.y);
    ctx.stroke();
});

// 서명 종료
function stopDrawing(event) {
    if(event.pointerId!==pointerId) return;
    drawing = false;
    pointerId = null;
    ctx.closePath();
    if(canvas.hasPointerCapture(event.pointerId)) canvas.releasePointerCapture(event.pointerId);
}
["pointerup","pointercancel","lostpointercapture"].forEach(type=>canvas.addEventListener(type,stopDrawing));

// 서명 지우기
clearButton.addEventListener("click", function() {
    if(saving) return;
    const captured=pointerId;
    drawing=false;pointerId=null;
    if(captured!==null && canvas.hasPointerCapture(captured)) canvas.releasePointerCapture(captured);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    statusMessage.textContent="";
});

function hasSignature() {
    const pixels=ctx.getImageData(0,0,canvas.width,canvas.height).data;
    let count=0;
    for(let i=0;i<pixels.length;i+=4){
        if(pixels[i+3]>32 && (pixels[i]<220 || pixels[i+1]<220 || pixels[i+2]<220) && ++count>=12) return true;
    }
    return false;
}

// 제출
form.addEventListener("submit", async function(event) {
    event.preventDefault();
    if(saving) return;
    const userName=document.getElementById("user-name").value.trim();
    if(!userName){statusMessage.textContent="이름을 입력해 주세요.";return;}
    if(!hasSignature()){statusMessage.textContent="서명을 입력해 주세요.";return;}
    saving=true;submitButton.disabled=true;clearButton.disabled=true;
    statusMessage.textContent="저장 중…";
    const controller=new AbortController();
    const timeout=setTimeout(()=>controller.abort(),45000);
    try {
        const body=new URLSearchParams({lessonId:form.dataset.lessonId,userName,signature:canvas.toDataURL("image/png")});
        if(new TextEncoder().encode(body.toString()).length>98000) throw new Error("서명이 너무 큽니다. 지운 뒤 다시 작성해 주세요.");
        const response=await fetch("save_signature.asp",{method:"POST",body,credentials:"same-origin",signal:controller.signal});
        let result;
        try { result=await response.json(); } catch { throw new Error("저장 응답을 확인하지 못했습니다. 다시 제출하거나 관리자에게 확인해 주세요."); }
        if(!response.ok || result.ok!==true) throw new Error(result.message || "저장하지 못했습니다.");
        window.location.assign("signature_success.asp?id="+encodeURIComponent(form.dataset.lessonId));
    } catch(error) {
        statusMessage.textContent=error.name==="AbortError"?"응답 시간이 초과되었습니다. 서명은 남아 있으니 저장 여부를 확인하거나 다시 제출해 주세요.":error.message;
        saving=false;submitButton.disabled=false;clearButton.disabled=false;
    } finally { clearTimeout(timeout); }
});
window.addEventListener("pageshow",()=>{saving=false;submitButton.disabled=false;clearButton.disabled=false;});
