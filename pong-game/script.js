const canvas = document.getElementById("pong");
const ctx = canvas.getContext("2d");

const paddleWidth = 10, paddleHeight = 100, ballSize = 10;
let leftPaddleY = canvas.height / 2 - paddleHeight / 2;
let rightPaddleY = canvas.height / 2 - paddleHeight / 2;
let ballX = canvas.width / 2, ballY = canvas.height / 2;
let ballSpeedX = 4, ballSpeedY = 4;
let leftScore = 0, rightScore = 0;
const winningScore = 5;

function drawRect(x, y, w, h, color) {
  ctx.fillStyle = color;
  ctx.fillRect(x, y, w, h);
}

function drawCircle(x, y, r, color) {
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.arc(x, y, r, 0, Math.PI * 2);
  ctx.closePath();
  ctx.fill();
}

function drawText(text, x, y, color) {
  ctx.fillStyle = color;
  ctx.font = "40px Arial";
  ctx.fillText(text, x, y);
}

function resetBall() {
  ballX = canvas.width / 2;
  ballY = canvas.height / 2;
  ballSpeedX = -ballSpeedX;
  ballSpeedY = 4 * (Math.random() > 0.5 ? 1 : -1);
}

function draw() {
  drawRect(0, 0, canvas.width, canvas.height, "#111");
  drawText(leftScore, canvas.width / 4, 50, "#fff");
  drawText(rightScore, 3 * canvas.width / 4, 50, "#fff");

  drawRect(0, leftPaddleY, paddleWidth, paddleHeight, "#fff");
  drawRect(canvas.width - paddleWidth, rightPaddleY, paddleWidth, paddleHeight, "#fff");
  drawCircle(ballX, ballY, ballSize, "#fff");
}

function update() {
  ballX += ballSpeedX;
  ballY += ballSpeedY;

  // top/bottom bounce
  if (ballY <= 0 || ballY >= canvas.height) {
    ballSpeedY *= -1;
  }

  // paddle collision
  if (
    ballX <= paddleWidth &&
    ballY >= leftPaddleY &&
    ballY <= leftPaddleY + paddleHeight
  ) {
    ballSpeedX *= -1;
  }

  if (
    ballX >= canvas.width - paddleWidth &&
    ballY >= rightPaddleY &&
    ballY <= rightPaddleY + paddleHeight
  ) {
    ballSpeedX *= -1;
  }

  // scoring
  if (ballX < 0) {
    rightScore++;
    resetBall();
  } else if (ballX > canvas.width) {
    leftScore++;
    resetBall();
  }

  // win condition
  if (leftScore >= winningScore || rightScore >= winningScore) {
    alert(`${leftScore >= winningScore ? "Left" : "Right"} player wins!`);
    leftScore = 0;
    rightScore = 0;
    resetBall();
  }
}

function game() {
  update();
  draw();
}

setInterval(game, 1000 / 60);

document.addEventListener("keydown", (e) => {
  switch (e.key) {
    case "w":
      leftPaddleY -= 20;
      break;
    case "s":
      leftPaddleY += 20;
      break;
    case "ArrowUp":
      rightPaddleY -= 20;
      break;
    case "ArrowDown":
      rightPaddleY += 20;
      break;
  }

  // boundary check
  leftPaddleY = Math.max(Math.min(leftPaddleY, canvas.height - paddleHeight), 0);
  rightPaddleY = Math.max(Math.min(rightPaddleY, canvas.height - paddleHeight), 0);
});