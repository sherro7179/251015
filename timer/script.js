const languageButtons = document.querySelectorAll(".language-switch__button");
const timezoneSelect = document.getElementById("timezone-select");
const dateElement = document.getElementById("date");
const hoursElement = document.getElementById("hours");
const minutesElement = document.getElementById("minutes");
const secondsElement = document.getElementById("seconds");
const sessionElement = document.getElementById("session");
const millisecondsElement = document.getElementById("milliseconds");
const cityCardsContainer = document.getElementById("city-cards");
const htmlElement = document.documentElement;
const languageSwitchGroup = document.querySelector(".language-switch");

const translations = {
    ko: {
        title: "세계 펄스 타이머",
        timezoneLabel: "타임존",
        milliseconds: "밀리초",
        citiesTitle: "주요 도시 반낮",
        day: "낮",
        night: "밤",
        languageGroup: "언어 선택",
    },
    en: {
        title: "World Pulse Timer",
        timezoneLabel: "Time Zone",
        milliseconds: "Milliseconds",
        citiesTitle: "Global Day & Night",
        day: "Day",
        night: "Night",
        languageGroup: "Language selection",
    },
};

const cities = [
    {
        id: "Asia/Seoul",
        name: { ko: "서울", en: "Seoul" },
        region: { ko: "대한민국", en: "South Korea" },
    },
    {
        id: "Asia/Tokyo",
        name: { ko: "도쿄", en: "Tokyo" },
        region: { ko: "일본", en: "Japan" },
    },
    {
        id: "Asia/Shanghai",
        name: { ko: "상하이", en: "Shanghai" },
        region: { ko: "중국", en: "China" },
    },
    {
        id: "Asia/Dubai",
        name: { ko: "두바이", en: "Dubai" },
        region: { ko: "아랍에미리트", en: "UAE" },
    },
    {
        id: "Europe/London",
        name: { ko: "런던", en: "London" },
        region: { ko: "영국", en: "UK" },
    },
    {
        id: "Europe/Paris",
        name: { ko: "파리", en: "Paris" },
        region: { ko: "프랑스", en: "France" },
    },
    {
        id: "America/New_York",
        name: { ko: "뉴욕", en: "New York" },
        region: { ko: "미국", en: "USA" },
    },
    {
        id: "America/Los_Angeles",
        name: { ko: "로스앤젤레스", en: "Los Angeles" },
        region: { ko: "미국", en: "USA" },
    },
    {
        id: "America/Sao_Paulo",
        name: { ko: "상파울루", en: "Sao Paulo" },
        region: { ko: "브라질", en: "Brazil" },
    },
    {
        id: "Australia/Sydney",
        name: { ko: "시드니", en: "Sydney" },
        region: { ko: "호주", en: "Australia" },
    },
];

const citySnapshots = [
    "Asia/Seoul",
    "Europe/London",
    "America/New_York",
    "America/Los_Angeles",
    "Asia/Dubai",
    "Australia/Sydney",
];

const formatterCache = new Map();

const impactTargets = {
    hour: hoursElement,
    minute: minutesElement,
    second: secondsElement,
};

const state = {
    locale: "ko",
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
    previous: {
        hour: null,
        minute: null,
        second: null,
    },
};

const cityElements = new Map();

function getFormatter({ locale, timeZone, options, cacheKey }) {
    const key = cacheKey ?? JSON.stringify({ locale, timeZone, options });
    if (!formatterCache.has(key)) {
        formatterCache.set(key, new Intl.DateTimeFormat(locale, { ...options, timeZone }));
    }
    return formatterCache.get(key);
}

function populateTimezoneOptions() {
    const optionSet = new Map();

    for (const city of cities) {
        optionSet.set(city.id, city);
    }

    if (!optionSet.has(state.timezone)) {
        optionSet.set(state.timezone, {
            id: state.timezone,
            name: { ko: state.timezone, en: state.timezone },
            region: { ko: "사용자 환경", en: "Current" },
        });
    }

    timezoneSelect.innerHTML = "";

    for (const [id, city] of optionSet) {
        const option = document.createElement("option");
        option.value = id;
        option.textContent = `${city.name[state.locale]} · ${city.region[state.locale]}`;
        if (id === state.timezone) {
            option.selected = true;
        }
        timezoneSelect.append(option);
    }
}

function buildCityCards() {
    cityCardsContainer.innerHTML = "";
    cityElements.clear();

    for (const zone of citySnapshots) {
        const city = cities.find((item) => item.id === zone);
        if (!city) {
            continue;
        }

        const card = document.createElement("article");
        card.className = "city-card";

        const name = document.createElement("h3");
        name.className = "city-card__name";

        const region = document.createElement("p");
        region.className = "city-card__region";

        const time = document.createElement("p");
        time.className = "city-card__time";

        const status = document.createElement("p");
        status.className = "city-card__status";

        card.append(name, region, time, status);
        cityCardsContainer.append(card);

        cityElements.set(zone, {
            card,
            name,
            region,
            time,
            status,
        });
    }
}

function updateImpact(segment) {
    const element = impactTargets[segment];
    if (!element) {
        return;
    }
    element.classList.remove("is-impact");
    // Trigger reflow to restart animation.
    void element.offsetWidth;
    element.classList.add("is-impact");
}

function formatTimeParts(date, locale, timeZone) {
    const formatter = getFormatter({
        locale,
        timeZone,
        options: {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
            hour12: true,
            fractionalSecondDigits: 3,
        },
    });

    const parts = formatter.formatToParts(date);
    const result = {};
    for (const part of parts) {
        result[part.type] = part.value;
    }
    return result;
}

function formatDate(date, locale, timeZone) {
    const formatter = getFormatter({
        locale,
        timeZone,
        options: { dateStyle: "full" },
        cacheKey: `date-${locale}-${timeZone}`,
    });
    return formatter.format(date);
}

function getHour24(date, timeZone) {
    const formatter = getFormatter({
        locale: "en-US",
        timeZone,
        options: { hour: "2-digit", hourCycle: "h23" },
        cacheKey: `hour24-${timeZone}`,
    });
    const parts = formatter.formatToParts(date);
    const hourPart = parts.find((part) => part.type === "hour");
    return hourPart ? Number(hourPart.value) : 0;
}

function updateLocaleTexts() {
    const map = translations[state.locale];
    document.querySelectorAll("[data-i18n]").forEach((element) => {
        const key = element.dataset.i18n;
        if (map[key]) {
            element.textContent = map[key];
        }
    });

    languageSwitchGroup.setAttribute("aria-label", map.languageGroup);
    htmlElement.lang = state.locale;
}

function updateTimezoneLabels() {
    for (const option of timezoneSelect.options) {
        const city = cities.find((item) => item.id === option.value);
        if (city) {
            option.textContent = `${city.name[state.locale]} · ${city.region[state.locale]}`;
        } else if (option.value === state.timezone) {
            option.textContent = `${state.timezone} · ${
                state.locale === "ko" ? "사용자 환경" : "Current"
            }`;
        }
    }
}

function updateCityCards(date) {
    for (const [zone, elements] of cityElements.entries()) {
        const city = cities.find((item) => item.id === zone);
        if (!city) {
            continue;
        }
        const localizedName = city.name[state.locale];
        const localizedRegion = city.region[state.locale];
        elements.name.textContent = localizedName;
        elements.region.textContent = localizedRegion;

        const parts = formatTimeParts(date, state.locale, zone);
        const hour24 = getHour24(date, zone);
        const session = hour24 >= 12 ? "PM" : "AM";
        const isDay = hour24 >= 6 && hour24 < 18;
        const statusLabel = isDay ? translations[state.locale].day : translations[state.locale].night;

        elements.time.textContent = `${parts.hour}:${parts.minute} ${session}`;
        elements.status.textContent = statusLabel;
        elements.status.classList.toggle("is-day", isDay);
        elements.status.classList.toggle("is-night", !isDay);
    }
}

function updateDayNightGradient(date) {
    const utcMinute = date.getUTCHours() * 60 + date.getUTCMinutes();
    const angle = ((utcMinute / (24 * 60)) * 360).toFixed(2);
    htmlElement.style.setProperty("--day-angle", `${angle}deg`);
}

function render(date) {
    const parts = formatTimeParts(date, state.locale, state.timezone);
    const hour24 = getHour24(date, state.timezone);

    if (state.previous.hour !== parts.hour) {
        updateImpact("hour");
        state.previous.hour = parts.hour;
    }
    if (state.previous.minute !== parts.minute) {
        updateImpact("minute");
        state.previous.minute = parts.minute;
    }
    if (state.previous.second !== parts.second) {
        updateImpact("second");
        state.previous.second = parts.second;
    }

    hoursElement.textContent = parts.hour;
    minutesElement.textContent = parts.minute;
    secondsElement.textContent = parts.second;
    sessionElement.textContent = hour24 >= 12 ? "PM" : "AM";
    millisecondsElement.textContent = parts.fractionalSecond ?? date.getMilliseconds().toString().padStart(3, "0");
    dateElement.textContent = formatDate(date, state.locale, state.timezone);

    updateCityCards(date);
    updateDayNightGradient(date);
}

function animationLoop() {
    const now = new Date();
    render(now);
    requestAnimationFrame(animationLoop);
}

function onLanguageChange(nextLocale) {
    if (state.locale === nextLocale) {
        return;
    }
    state.locale = nextLocale;
    syncLanguageButtons();
    updateLocaleTexts();
    populateTimezoneOptions();
    updateTimezoneLabels();
}

function onTimezoneChange(nextTimezone) {
    state.timezone = nextTimezone;
    populateTimezoneOptions();
    state.previous.hour = null;
    state.previous.minute = null;
    state.previous.second = null;
}

function bindEvents() {
    languageButtons.forEach((button) => {
        button.addEventListener("click", () => onLanguageChange(button.dataset.lang));
    });

    timezoneSelect.addEventListener("change", (event) => {
        onTimezoneChange(event.target.value);
    });

    document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") {
            state.previous.hour = null;
            state.previous.minute = null;
            state.previous.second = null;
        }
    });
}

function syncLanguageButtons() {
    languageButtons.forEach((button) => {
        const isActive = button.dataset.lang === state.locale;
        button.classList.toggle("is-active", isActive);
        button.setAttribute("aria-pressed", String(isActive));
    });
}

function init() {
    updateLocaleTexts();
    syncLanguageButtons();
    populateTimezoneOptions();
    buildCityCards();
    updateTimezoneLabels();
    bindEvents();
    requestAnimationFrame(animationLoop);
}

init();
