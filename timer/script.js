const dateElement = document.getElementById("date");
const timeElement = document.getElementById("time");

const dateFormatter = new Intl.DateTimeFormat("ko-KR", {
    year: "numeric",
    month: "long",
    day: "numeric",
    weekday: "long",
});

function pad(value) {
    return String(value).padStart(2, "0");
}

function updateClock() {
    const now = new Date();
    const hours = pad(now.getHours());
    const minutes = pad(now.getMinutes());
    const seconds = pad(now.getSeconds());

    dateElement.textContent = dateFormatter.format(now);
    timeElement.textContent = `${hours}:${minutes}:${seconds}`;
}

updateClock();
setInterval(updateClock, 1000);
