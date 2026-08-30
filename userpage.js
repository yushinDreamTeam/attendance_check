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

        const userName =
            document.getElementById("user-name").value;

        canvas.toBlob(function(blob) {

            console.log("이름:", userName);
            console.log("서명:", blob);

            alert("제출되었습니다.");

        }, "image/png");

    });
