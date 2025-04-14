const canvas = document.getElementById("pong");
const ctx = canvas.getContext("2d");
const bgMusic = document.getElementById("bg-music");

const paddleWidth = 20, paddleHeight = 100, ballSize = 20;
let leftPaddleY, rightPaddleY, ballX, ballY, ballSpeedX, ballSpeedY;
let leftScore = 0, rightScore = 0;
const winningScore = 5;
let gameInterval = null;
let gameMode = "pvp";
let highScore = localStorage.getItem("pongHighScore") || 0;
document.getElementById("high-score").innerText = highScore;

function resetPositions() {
  leftPaddleY = canvas.height / 2 - paddleHeight / 2;
  rightPaddleY = canvas.height / 2 - paddleHeight / 2;
  ballX = canvas.width / 2;
  ballY = canvas.height / 2;
  ballSpeedX = Math.random() > 0.5 ? 6 : -6;
  ballSpeedY = 4 * (Math.random() > 0.5 ? 1 : -1);
}

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

function drawText(text, x, y, size = 40, color = "#fff") {
  ctx.fillStyle = color;
  ctx.font = `${size}px Arial`;
  ctx.fillText(text, x, y);
}

function draw() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  drawText(leftScore, canvas.width / 4, 50);
  drawText(rightScore, 3 * canvas.width / 4, 50);
  drawRect(0, leftPaddleY, paddleWidth, paddleHeight, "#FFD700");
  drawRect(canvas.width - paddleWidth, rightPaddleY, paddleWidth, paddleHeight, "#FF69B4");
  drawCircle(ballX, ballY, ballSize / 2, "#00FFFF");
}

function update() {
  ballX += ballSpeedX;
  ballY += ballSpeedY;

  if (ballY <= 0 || ballY >= canvas.height) {
    ballSpeedY *= -1;
  }

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

  if (ballX < 0) {
    rightScore++;
    if (rightScore >= winningScore) endGame("RIGHT");
    resetPositions();
  }

  if (ballX > canvas.width) {
    leftScore++;
    if (leftScore >= winningScore) endGame("LEFT");
    resetPositions();
  }

  // AI control
  if (gameMode === "easy") {
    if (ballY > rightPaddleY + paddleHeight / 2) rightPaddleY += 3;
    else rightPaddleY -= 3;
  } else if (gameMode === "hard") {
    if (ballY > rightPaddleY + paddleHeight / 2) rightPaddleY += 6;
    else rightPaddleY -= 6;
  }
}

function gameLoop() {
  update();
  draw();
}

document.addEventListener("keydown", (e) => {
  switch (e.key) {
    case "w": leftPaddleY -= 20; break;
    case "s": leftPaddleY += 20; break;
    case "ArrowUp":
      if (gameMode === "pvp") rightPaddleY -= 20;
      break;
    case "ArrowDown":
      if (gameMode === "pvp") rightPaddleY += 20;
      break;
  }

  leftPaddleY = Math.max(0, Math.min(leftPaddleY, canvas.height - paddleHeight));
  rightPaddleY = Math.max(0, Math.min(rightPaddleY, canvas.height - paddleHeight));
});

function startGame(mode) {
  gameMode = mode;
  document.getElementById("main-menu").style.display = "none";
  bgMusic.play();
  leftScore = 0;
  rightScore = 0;
  resetPositions();
  gameInterval = setInterval(gameLoop, 1000 / 60);
}

function endGame(winner) {
  clearInterval(gameInterval);
  document.getElementById("winner-text").innerText = `${winner} PLAYER WINS!`;
  document.getElementById("game-over-screen").style.display = "flex";
  bgMusic.pause();
  bgMusic.currentTime = 0;
  const maxScore = Math.max(leftScore, rightScore);
  if (maxScore > highScore) {
    highScore = maxScore;
    localStorage.setItem("pongHighScore", highScore);
  }
}

function showMainMenu() {
  document.getElementById("game-over-screen").style.display = "none";
  document.getElementById("main-menu").style.display = "flex";
  document.getElementById("high-score").innerText = highScore;
}
