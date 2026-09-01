const canvas = document.getElementById("signature-canvas");
const ctx = canvas.getContext("2d");

let drawing = false;
let hasSignature = false;


// 선 설정
ctx.strokeStyle = "black";
ctx.lineWidth = 3;
ctx.lineCap = "round";
ctx.lineJoin = "round";


// 좌표 계산
function getPosition(event) {

    const rect = canvas.getBoundingClientRect();

    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    return {
        x: (event.clientX - rect.left) * scaleX,
        y: (event.clientY - rect.top) * scaleY
    };
}


// 서명 시작
canvas.addEventListener("pointerdown", function(event) {

    drawing = true;
    hasSignature = true;

    canvas.setPointerCapture(event.pointerId);

    const position = getPosition(event);

    ctx.beginPath();

    ctx.moveTo(
        position.x,
        position.y
    );
});


// 서명 그리기
canvas.addEventListener("pointermove", function(event) {

    if (!drawing) {
        return;
    }

    const position = getPosition(event);

    ctx.lineTo(
        position.x,
        position.y
    );

    ctx.stroke();
});


// 서명 종료
canvas.addEventListener("pointerup", function(event) {

    drawing = false;

    ctx.closePath();

    canvas.releasePointerCapture(event.pointerId);
});


// 서명 지우기
document
    .getElementById("clear-signature-button")
    .addEventListener("click", function() {

        ctx.clearRect(
            0,
            0,
            canvas.width,
            canvas.height
        );

        hasSignature = false;
    });


// 제출
document
    .getElementById("attendance-form")
    .addEventListener("submit", function(event) {

        event.preventDefault();

        if (!hasSignature) {
            alert("서명을 입력해주세요.");
            return;
        }

        const userName = document.getElementById("user-name").value.trim();
        const lessonId = new URLSearchParams(window.location.search).get("id");

        if (!userName) {
            alert("이름을 입력해주세요.");
            return;
        }

        if (!lessonId) {
            alert("연수 ID가 없습니다.");
            return;
        }

        const signatureData = canvas.toDataURL("image/png");

        fetch("save_signature.asp", {
            method: "POST",
            headers: {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
            },
            body: new URLSearchParams({
                lessonId: lessonId,
                userName: userName,
                signatureData: signatureData
            })
        })
        .then(function(response) {
            return response.text();
        })
        .then(function(result) {
            if (result && result.indexOf("성공") >= 0) {
                window.location.href = "signature_success.asp";
                return;
            }

            if (result && result.indexOf("오류") >= 0) {
                alert(result);
                return;
            }

            window.location.href = "signature_success.asp";
        })
        .catch(function(error) {
            alert("서명 저장 중 오류가 발생했습니다.");
        });

    });
