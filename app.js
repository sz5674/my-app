const VOCAB_STATE_KEY = "english-app-vocab-state-v2";
const READING_STATE_KEY = "english-app-reading-state-v2";
const CLOZE_STATE_KEY = "english-app-cloze-state-v1";
const PART6_STATE_KEY = "english-app-part6-state-v1";
const DAILY_GOAL_KEY = "english-app-daily-goal-v1";
const MOCK_STATE_KEY = "english-app-mock-state-v1";
const DAILY_RECORD_KEY = "english-app-daily-record-v1";
const DEFAULT_TIME_LIMIT = 20;
const DEFAULT_DAILY_GOAL = 30;

const CONFETTI_COLORS = [
  "#1d4ed8",
  "#2563eb",
  "#60a5fa",
  "#fbbf24",
  "#34d399",
  "#a78bfa",
  "#f472b6",
];

function ensureFxLayer() {
  let layer = document.getElementById("fx-layer");
  if (!layer) {
    layer = document.createElement("div");
    layer.id = "fx-layer";
    layer.className = "fx-layer";
    layer.setAttribute("aria-hidden", "true");
    document.body.appendChild(layer);
  }
  return layer;
}

function spawnConfettiAt(centerX, centerY, tier = "small") {
  const layer = ensureFxLayer();
  const counts = { small: 22, medium: 48, large: 80 };
  const count = counts[tier] || counts.small;
  const maxDurMs = tier === "large" ? 1350 : tier === "medium" ? 1150 : 950;
  const baseSpeed = tier === "large" ? 150 : tier === "medium" ? 120 : 95;
  for (let i = 0; i < count; i += 1) {
    const piece = document.createElement("span");
    piece.className = "confetti-piece";
    const angle = Math.random() * Math.PI * 2;
    const speed = baseSpeed + Math.random() * 100;
    const dx = Math.cos(angle) * speed;
    const dy = Math.sin(angle) * speed - 25 - Math.random() * 50;
    piece.style.left = `${centerX}px`;
    piece.style.top = `${centerY}px`;
    piece.style.setProperty("--dx", `${dx}px`);
    piece.style.setProperty("--dy", `${dy}px`);
    piece.style.setProperty("--rot", `${(Math.random() - 0.5) * 900}deg`);
    const durSec = 0.75 + Math.random() * 0.4;
    piece.style.setProperty("--burst-dur", `${durSec}s`);
    piece.style.background =
      CONFETTI_COLORS[Math.floor(Math.random() * CONFETTI_COLORS.length)];
    layer.appendChild(piece);
    window.setTimeout(() => piece.remove(), maxDurMs);
  }
}

function spawnConfettiFromElement(element, tier = "small") {
  if (!element || typeof element.getBoundingClientRect !== "function") {
    return;
  }
  const rect = element.getBoundingClientRect();
  spawnConfettiAt(
    rect.left + rect.width / 2,
    rect.top + rect.height / 2,
    tier
  );
}

function shuffleQuestion(question) {
  const optionsWithFlag = question.options.map((option, index) => ({
    option,
    isAnswer: index === question.answer,
  }));

  for (let i = optionsWithFlag.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [optionsWithFlag[i], optionsWithFlag[j]] = [
      optionsWithFlag[j],
      optionsWithFlag[i],
    ];
  }

  return {
    ...question,
    options: optionsWithFlag.map((item) => item.option),
    answer: optionsWithFlag.findIndex((item) => item.isAnswer),
  };
}

function pickRandomDistinctItems(list, count, excludedValue) {
  const filtered = list.filter((item) => item !== excludedValue);
  const picked = [];
  while (picked.length < count && filtered.length > 0) {
    const index = Math.floor(Math.random() * filtered.length);
    picked.push(filtered[index]);
    filtered.splice(index, 1);
  }
  return picked;
}

function sampleItems(list, count) {
  const copy = [...list];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy.slice(0, Math.min(count, copy.length));
}

function formatSeconds(totalSeconds) {
  const safe = Math.max(0, totalSeconds);
  const minutes = Math.floor(safe / 60);
  const seconds = String(safe % 60).padStart(2, "0");
  return `${minutes}:${seconds}`;
}

function formatPassageHtml(text) {
  if (!text) return "";
  return text.replace(/\n+/g, "<br>");
}

function createPhraseText(question) {
  if (question.phraseHint) {
    return question.phraseHint.replace(
      /^(頻出表現|読解キー句|読解根拠|本文表現|文中表現|文中熟語|文脈接続|定型表現|時制ポイント|Part 6ポイント)\s*:\s*/,
      ""
    );
  }
  return "特記事項なし";
}

function createExplanationText(question) {
  return question.explanation || "";
}

function normalizeQuestion(question, fallbackCategory) {
  return {
    ...question,
    category: question.category || fallbackCategory || "general",
  };
}

function getTodayKey() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function getEmptyDailyRecord() {
  return {
    vocab: { attempted: 0, correct: 0 },
    cloze: { attempted: 0, correct: 0 },
    part6: { attempted: 0, correct: 0 },
    reading: { attempted: 0, correct: 0 },
  };
}

function loadDailyRecord() {
  try {
    const raw = localStorage.getItem(DAILY_RECORD_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_error) {
    return {};
  }
}

function saveDailyRecord(recordMap) {
  localStorage.setItem(DAILY_RECORD_KEY, JSON.stringify(recordMap));
}

function updateDailyRecord(sectionName, isCorrect) {
  if (!["vocab", "cloze", "part6", "reading"].includes(sectionName)) return;
  const today = getTodayKey();
  const recordMap = loadDailyRecord();
  const todayRecord = recordMap[today] || getEmptyDailyRecord();
  if (!todayRecord[sectionName]) {
    todayRecord[sectionName] = { attempted: 0, correct: 0 };
  }
  todayRecord[sectionName].attempted += 1;
  if (isCorrect) {
    todayRecord[sectionName].correct += 1;
  }
  recordMap[today] = todayRecord;
  saveDailyRecord(recordMap);
}

function getAttemptedFromRecord(recordItem) {
  const item = recordItem || getEmptyDailyRecord();
  return (
    Number(item.vocab?.attempted || 0) +
    Number(item.cloze?.attempted || 0) +
    Number(item.part6?.attempted || 0) +
    Number(item.reading?.attempted || 0)
  );
}

function getCorrectFromRecord(recordItem) {
  const item = recordItem || getEmptyDailyRecord();
  return (
    Number(item.vocab?.correct || 0) +
    Number(item.cloze?.correct || 0) +
    Number(item.part6?.correct || 0) +
    Number(item.reading?.correct || 0)
  );
}

function computeCurrentStreak(sortedDates, recordMap) {
  let streak = 0;
  const cursor = new Date();
  cursor.setHours(0, 0, 0, 0);

  while (true) {
    const y = cursor.getFullYear();
    const m = String(cursor.getMonth() + 1).padStart(2, "0");
    const d = String(cursor.getDate()).padStart(2, "0");
    const key = `${y}-${m}-${d}`;
    if (!sortedDates.includes(key)) break;
    if (getAttemptedFromRecord(recordMap[key]) <= 0) break;
    streak += 1;
    cursor.setDate(cursor.getDate() - 1);
  }
  return streak;
}

function buildInfiniteQuestionSession(config) {
  const {
    containerId,
    stateKey,
    questionBank,
    titlePrefix,
    showPassage = false,
    randomizePrompt = true,
    enableUnitFilter = false,
    enableWeakPriority = false,
    enableTimedMode = false,
    enableReviewMode = false,
    fallbackCategory = "general",
    simplePromptMode = false,
    sectionName = "unknown",
  } = config;
  const showExplanationFeedback = sectionName !== "vocab";
  const container = document.getElementById(containerId);
  const normalizedQuestionBank = questionBank.map((question) =>
    normalizeQuestion(question, fallbackCategory)
  );
  let timerId = null;
  let session = loadSession();

  function getInitialSession() {
    return {
      attempted: 0,
      correct: 0,
      currentQuestion: null,
      isAnswered: false,
      weakStats: {},
      selectedUnit: "all",
      timedMode: false,
      timeLimit: DEFAULT_TIME_LIMIT,
      timeLeft: DEFAULT_TIME_LIMIT,
      reviewOnly: false,
      categoryStats: {},
      recentQuestionKeys: [],
      weakPriorityEnabled: enableWeakPriority,
      wrongOptionStats: {},
      dailyStats: {
        date: getTodayKey(),
        attempted: 0,
        correct: 0,
      },
    };
  }

  function loadSession() {
    try {
      const raw = localStorage.getItem(stateKey);
      if (!raw) return getInitialSession();
      const parsed = JSON.parse(raw);
      return {
        attempted: Number(parsed.attempted) || 0,
        correct: Number(parsed.correct) || 0,
        currentQuestion: parsed.currentQuestion || null,
        isAnswered: Boolean(parsed.isAnswered),
        weakStats: parsed.weakStats || {},
        selectedUnit: parsed.selectedUnit || "all",
        timedMode: Boolean(parsed.timedMode),
        timeLimit: Number(parsed.timeLimit) || DEFAULT_TIME_LIMIT,
        timeLeft:
          Number(parsed.timeLeft) ||
          Number(parsed.timeLimit) ||
          DEFAULT_TIME_LIMIT,
        reviewOnly: Boolean(parsed.reviewOnly),
        categoryStats: parsed.categoryStats || {},
        recentQuestionKeys: Array.isArray(parsed.recentQuestionKeys)
          ? parsed.recentQuestionKeys
          : [],
        weakPriorityEnabled:
          typeof parsed.weakPriorityEnabled === "boolean"
            ? parsed.weakPriorityEnabled
            : enableWeakPriority,
        wrongOptionStats: parsed.wrongOptionStats || {},
        dailyStats:
          parsed.dailyStats && parsed.dailyStats.date
            ? parsed.dailyStats
            : {
                date: getTodayKey(),
                attempted: 0,
                correct: 0,
              },
      };
    } catch (_error) {
      return getInitialSession();
    }
  }

  function saveSession() {
    const today = getTodayKey();
    if (!session.dailyStats || session.dailyStats.date !== today) {
      session.dailyStats = {
        date: today,
        attempted: 0,
        correct: 0,
      };
    }
    localStorage.setItem(stateKey, JSON.stringify(session));
  }

  function getQuestionPool() {
    let pool = normalizedQuestionBank;
    if (enableUnitFilter && session.selectedUnit !== "all") {
      const byUnit = pool.filter((question) => question.toeicUnit === session.selectedUnit);
      if (byUnit.length > 0) pool = byUnit;
    }
    if (enableReviewMode && session.reviewOnly) {
      const wrongOnly = pool.filter((question) => {
        const key = `${question.category}::${question.question}`;
        return Number(session.weakStats[key] || 0) > 0;
      });
      if (wrongOnly.length > 0) {
        return wrongOnly;
      }
    }

    return pool;
  }

  function pickQuestionByWeakness(pool) {
    if (!enableWeakPriority || !session.weakPriorityEnabled) {
      return pool[Math.floor(Math.random() * pool.length)];
    }

    const weightedPool = pool.flatMap((question) => {
      const key = `${question.category}::${question.question}`;
      const missCount = Number(session.weakStats[key] || 0);
      const weight = 1 + Math.min(Math.floor(missCount / 2), 2);
      return Array(weight).fill(question);
    });
    return weightedPool[Math.floor(Math.random() * weightedPool.length)];
  }

  function rememberQuestion(question) {
    const key = `${question.category}::${question.question}`;
    const next = [...session.recentQuestionKeys.filter((item) => item !== key), key];
    session.recentQuestionKeys = next.slice(-5);
  }

  function getNextQuestion() {
    const pool = getQuestionPool();
    const recentSet = new Set(session.recentQuestionKeys || []);
    let candidatePool = pool.filter((question) => {
      const key = `${question.category}::${question.question}`;
      return !recentSet.has(key);
    });
    if (candidatePool.length < Math.min(6, pool.length)) {
      candidatePool = pool;
    }
    const baseQuestion = pickQuestionByWeakness(candidatePool);
    const nextQuestion = randomizePrompt ? shuffleQuestion(baseQuestion) : baseQuestion;
    return {
      ...nextQuestion,
      category: baseQuestion.category,
    };
  }

  function ensureCurrentQuestion() {
    if (!session.currentQuestion) {
      session.currentQuestion = getNextQuestion();
      rememberQuestion(session.currentQuestion);
      session.isAnswered = false;
      session.timeLeft = session.timeLimit;
      saveSession();
      return;
    }

    if (session.isAnswered) {
      // If the last state was already answered, auto-prepare the next item.
      session.currentQuestion = getNextQuestion();
      rememberQuestion(session.currentQuestion);
      session.isAnswered = false;
      session.timeLeft = session.timeLimit;
      saveSession();
    }
  }

  function clearTimer() {
    if (timerId) {
      window.clearInterval(timerId);
      timerId = null;
    }
  }

  function moveToNextQuestion() {
    session.currentQuestion = getNextQuestion();
    rememberQuestion(session.currentQuestion);
    session.isAnswered = false;
    session.timeLeft = session.timeLimit;
    saveSession();
    render();
  }

  function markAttempt({ isCorrect, selectedOption, question }) {
    const today = getTodayKey();
    if (!session.dailyStats || session.dailyStats.date !== today) {
      session.dailyStats = {
        date: today,
        attempted: 0,
        correct: 0,
      };
    }

    session.attempted += 1;
    session.dailyStats.attempted += 1;
    updateDailyRecord(sectionName, isCorrect);
    const categoryKey = question.category || fallbackCategory;
    if (!session.categoryStats[categoryKey]) {
      session.categoryStats[categoryKey] = { attempted: 0, correct: 0 };
    }
    session.categoryStats[categoryKey].attempted += 1;

    if (isCorrect) {
      session.correct += 1;
      session.dailyStats.correct += 1;
      session.categoryStats[categoryKey].correct += 1;
      return;
    }

    if (selectedOption) {
      const wrongOptionKey = `${categoryKey}::${selectedOption}`;
      session.wrongOptionStats[wrongOptionKey] =
        Number(session.wrongOptionStats[wrongOptionKey] || 0) + 1;
    }
  }

  function startTimer(feedback) {
    clearTimer();
    if (
      !enableTimedMode ||
      !session.timedMode ||
      session.isAnswered
    ) {
      return;
    }
    timerId = window.setInterval(() => {
      session.timeLeft -= 1;
      if (session.timeLeft <= 0) {
        session.timeLeft = 0;
        session.isAnswered = true;
        const q = session.currentQuestion;
        markAttempt({
          isCorrect: false,
          selectedOption: null,
          question: q,
        });
        const key = `${q.category}::${q.question}`;
        session.weakStats[key] = Number(session.weakStats[key] || 0) + 1;
        const optionButtons = container.querySelectorAll(".option-btn");
        optionButtons.forEach((btn, idx) => {
          btn.disabled = true;
          if (idx === q.answer) {
            btn.classList.add("option-correct");
            const badge = document.createElement("span");
            badge.className = "option-result-badge option-result-correct";
            badge.textContent = "正解";
            btn.appendChild(badge);
          }
        });
        feedback.innerHTML =
          showExplanationFeedback && q.explanation
            ? `<strong>解説:</strong> ${createExplanationText(q)}`
            : "";
        saveSession();
        clearTimer();
        return;
      }
      saveSession();
      const timerLabel = document.getElementById(`${containerId}-timer`);
      if (timerLabel) {
        timerLabel.textContent = `残り ${session.timeLeft} 秒`;
      }
    }, 1000);
  }

  function render() {
    ensureCurrentQuestion();
    const q = session.currentQuestion;
    const toeicUnits = Array.from(
      new Set(
        normalizedQuestionBank
          .map((question) => question.toeicUnit)
          .filter((value) => Boolean(value))
      )
    );
    const shouldShowUnitFilter = enableUnitFilter && toeicUnits.length > 1;

    container.innerHTML = `
      ${
        shouldShowUnitFilter ||
        enableTimedMode ||
        enableWeakPriority ||
        enableReviewMode
          ? `<div class="control-row">
        ${
          shouldShowUnitFilter
            ? `<label class="control-item">TOEIC単元:
            <select id="${containerId}-toeic-unit">
              <option value="all" ${session.selectedUnit === "all" ? "selected" : ""}>すべて</option>
              ${toeicUnits
                .map(
                  (unit) =>
                    `<option value="${unit}" ${
                      session.selectedUnit === unit ? "selected" : ""
                    }>${unit}</option>`
                )
                .join("")}
            </select>
          </label>`
            : ""
        }
        ${
          enableWeakPriority
            ? `<label class="control-item">
            <input type="checkbox" id="${containerId}-weak-priority" ${
              session.weakPriorityEnabled ? "checked" : ""
            } />
            弱点優先出題
          </label>`
            : ""
        }
        ${
          enableReviewMode
            ? `<label class="control-item">
            <input type="checkbox" id="${containerId}-review-only" ${
              session.reviewOnly ? "checked" : ""
            } />
            復習モード（間違えた問題を優先）
          </label>`
            : ""
        }
        ${
          enableTimedMode
            ? `<label class="control-item">
            <input type="checkbox" id="${containerId}-timed-mode" ${
              session.timedMode ? "checked" : ""
            } />
            制限時間モード
          </label>
          <label class="control-item">秒数:
            <select id="${containerId}-time-limit" ${
              session.timedMode || session.attempted > 0 ? "" : ""
            }>
              ${[10, 15, 20, 30, 45]
                .map(
                  (sec) =>
                    `<option value="${sec}" ${
                      session.timeLimit === sec ? "selected" : ""
                    }>${sec}</option>`
                )
                .join("")}
            </select>
          </label>
          <span class="control-item" id="${containerId}-timer">${
              session.timedMode
                ? `残り ${session.timeLeft} 秒`
                : "制限時間モードOFF"
            }</span>`
            : ""
        }
      </div>`
          : ""
      }
      ${
        showPassage && q.passage
          ? `<p>${titlePrefix ? `【${titlePrefix}】<br>` : ""}${formatPassageHtml(
              q.passage
            )}</p>`
          : ""
      }
      ${
        simplePromptMode
          ? `<p class="simple-word"><strong>${q.question}</strong></p>`
          : `<p><strong>${q.question}</strong></p>`
      }
      <div class="options">
        ${q.options
          .map(
            (option, idx) =>
              `<button class="option-btn" data-idx="${idx}" ${
                session.isAnswered ? "disabled" : ""
              }>${option}</button>`
          )
          .join("")}
      </div>
      <div class="feedback" id="${containerId}-feedback"></div>
      <button class="next-btn" id="${containerId}-next" style="display:${
        session.isAnswered ? "inline-block" : "none"
      };">次の問題</button>
    `;

    const optionButtons = container.querySelectorAll(".option-btn");
    const feedback = document.getElementById(`${containerId}-feedback`);
    const nextBtn = document.getElementById(`${containerId}-next`);
    const unitSelect = document.getElementById(`${containerId}-toeic-unit`);
    const timedModeCheckbox = document.getElementById(`${containerId}-timed-mode`);
    const timeLimitSelect = document.getElementById(`${containerId}-time-limit`);
    const weakPriorityCheckbox = document.getElementById(
      `${containerId}-weak-priority`
    );
    const reviewOnlyCheckbox = document.getElementById(
      `${containerId}-review-only`
    );
    function markOptionResult(button, label, type) {
      const badge = document.createElement("span");
      badge.className = `option-result-badge option-result-${type}`;
      badge.textContent = label;
      button.appendChild(badge);
    }

    optionButtons.forEach((button) => {
      button.addEventListener("click", () => {
        if (session.isAnswered) return;
        const selected = Number(button.dataset.idx);
        const isCorrect = selected === q.answer;
        markAttempt({
          isCorrect,
          selectedOption: q.options[selected],
          question: q,
        });
        if (!isCorrect && enableWeakPriority) {
          const key = `${q.category}::${q.question}`;
          session.weakStats[key] = Number(session.weakStats[key] || 0) + 1;
        }
        session.isAnswered = true;
        optionButtons.forEach((btn, idx) => {
          btn.disabled = true;
          if (idx === q.answer) {
            btn.classList.add("option-correct");
            markOptionResult(btn, "正解", "correct");
          } else if (idx === selected) {
            btn.classList.add("option-wrong");
            markOptionResult(btn, "不正解", "wrong");
          }
        });
        if (isCorrect) {
          spawnConfettiFromElement(button, "small");
        }
        feedback.innerHTML =
          showExplanationFeedback && q.explanation
            ? `<strong>解説:</strong> ${createExplanationText(q)}`
            : "";
        nextBtn.style.display = "inline-block";
        clearTimer();
        saveSession();
      });
    });

    nextBtn.addEventListener("click", () => {
      moveToNextQuestion();
    });

    if (unitSelect) {
      unitSelect.addEventListener("change", (event) => {
        session.selectedUnit = event.target.value;
        session.currentQuestion = getNextQuestion();
        session.isAnswered = false;
        session.timeLeft = session.timeLimit;
        saveSession();
        render();
      });
    }

    if (timedModeCheckbox) {
      timedModeCheckbox.addEventListener("change", (event) => {
        session.timedMode = Boolean(event.target.checked);
        session.timeLeft = session.timeLimit;
        saveSession();
        render();
      });
    }

    if (weakPriorityCheckbox) {
      weakPriorityCheckbox.addEventListener("change", (event) => {
        session.weakPriorityEnabled = Boolean(event.target.checked);
        saveSession();
      });
    }

    if (reviewOnlyCheckbox) {
      reviewOnlyCheckbox.addEventListener("change", (event) => {
        session.reviewOnly = Boolean(event.target.checked);
        session.currentQuestion = getNextQuestion();
        session.isAnswered = false;
        session.timeLeft = session.timeLimit;
        saveSession();
        render();
      });
    }

    if (timeLimitSelect) {
      timeLimitSelect.addEventListener("change", (event) => {
        session.timeLimit = Number(event.target.value);
        session.timeLeft = session.timeLimit;
        saveSession();
        render();
      });
    }

    startTimer(feedback);
  }

  render();
}

function createVocabQuestionBank() {
  const vocabPhraseHints = {
    achieve: "頻出表現: achieve a goal / achieve strong results",
    require: "頻出表現: require approval / require additional documents",
    improve: "頻出表現: improve productivity / improve customer satisfaction",
    increase: "頻出表現: increase sales / increase market share",
    reduce: "頻出表現: reduce costs / reduce processing time",
    confirm: "頻出表現: confirm the schedule / confirm your attendance",
    schedule: "頻出表現: schedule a meeting / schedule a delivery",
    deliver: "頻出表現: deliver a presentation / deliver the package",
    announce: "頻出表現: announce a new policy / announce the results",
    approve: "頻出表現: approve the budget / approve the proposal",
    provide: "頻出表現: provide support / provide detailed information",
    purchase: "頻出表現: purchase equipment / purchase office supplies",
    delay: "頻出表現: delay the launch / delay payment",
    attend: "頻出表現: attend a workshop / attend the conference",
    assign: "頻出表現: assign a project / assign responsibilities",
    maintain: "頻出表現: maintain quality / maintain good relationships",
    replace: "頻出表現: replace outdated systems / replace damaged parts",
    respond: "頻出表現: respond quickly / respond to inquiries",
    inspect: "頻出表現: inspect the facility / inspect the equipment",
    prepare: "頻出表現: prepare a report / prepare for negotiations",
    submit: "頻出表現: submit an application / submit the final draft",
    negotiate: "頻出表現: negotiate a contract / negotiate better terms",
    estimate: "頻出表現: estimate total costs / estimate completion time",
    promote: "頻出表現: promote a product / promote from within",
    assist: "頻出表現: assist a client / assist with onboarding",
    expand: "頻出表現: expand operations / expand into new markets",
    operate: "頻出表現: operate machinery / operate efficiently",
    review: "頻出表現: review the document / review quarterly performance",
    train: "頻出表現: train new employees / train the sales team",
    arrange: "頻出表現: arrange transportation / arrange an interview",
    include: "頻出表現: include shipping costs / include key details",
    exclude: "頻出表現: exclude tax / exclude confidential data",
    reserve: "頻出表現: reserve a conference room / reserve seats",
    issue: "頻出表現: issue an invoice / issue a warning",
    contact: "頻出表現: contact technical support / contact the supplier",
    manage: "頻出表現: manage a project / manage cash flow",
    adopt: "頻出表現: adopt a strategy / adopt new technology",
    cancel: "頻出表現: cancel an order / cancel the subscription",
    extend: "頻出表現: extend the deadline / extend support hours",
    renew: "頻出表現: renew a contract / renew a license",
    allocate: "頻出表現: allocate resources / allocate a budget",
    authorize: "頻出表現: authorize payment / authorize access",
    coordinate: "頻出表現: coordinate with vendors / coordinate team schedules",
    clarify: "頻出表現: clarify expectations / clarify the requirements",
    comply: "頻出表現: comply with regulations / comply with company policy",
    conclude: "頻出表現: conclude negotiations / conclude the meeting",
    delegate: "頻出表現: delegate responsibilities / delegate routine tasks",
    distribute: "頻出表現: distribute materials / distribute workloads evenly",
    draft: "頻出表現: draft a proposal / draft a response letter",
    enforce: "頻出表現: enforce safety rules / enforce strict standards",
    evaluate: "頻出表現: evaluate performance / evaluate potential risks",
    facilitate: "頻出表現: facilitate communication / facilitate the transition",
    finalize: "頻出表現: finalize the schedule / finalize contract details",
    implement: "頻出表現: implement a new system / implement policy changes",
    inform: "頻出表現: inform stakeholders / inform the customer",
    launch: "頻出表現: launch a campaign / launch a new service",
    monitor: "頻出表現: monitor progress / monitor system performance",
    notify: "頻出表現: notify all participants / notify us in advance",
    obtain: "頻出表現: obtain approval / obtain accurate information",
    participate: "頻出表現: participate in training / participate actively",
    perform: "頻出表現: perform an analysis / perform routine checks",
    prioritize: "頻出表現: prioritize urgent requests / prioritize customer needs",
    process: "頻出表現: process applications / process refunds quickly",
    propose: "頻出表現: propose a solution / propose an alternative plan",
    reassign: "頻出表現: reassign tasks / reassign team members",
    recommend: "頻出表現: recommend improvements / recommend a candidate",
    reconcile: "頻出表現: reconcile accounts / reconcile monthly statements",
    register: "頻出表現: register for a seminar / register a complaint",
    relocate: "頻出表現: relocate the office / relocate to a new site",
    remind: "頻出表現: remind the team / remind customers about deadlines",
    resolve: "頻出表現: resolve an issue / resolve customer complaints",
    retain: "頻出表現: retain top talent / retain existing clients",
    secure: "頻出表現: secure funding / secure sensitive data",
    specify: "頻出表現: specify the requirements / specify delivery dates",
    streamline: "頻出表現: streamline operations / streamline internal workflows",
    supervise: "頻出表現: supervise trainees / supervise daily operations",
    suspend: "頻出表現: suspend service temporarily / suspend the account",
    terminate: "頻出表現: terminate a contract / terminate employment",
    update: "頻出表現: update the records / update the software",
    verify: "頻出表現: verify account details / verify the information",
  };

  const vocabItems = [
    ["achieve", "達成する"],
    ["require", "必要とする"],
    ["improve", "改善する"],
    ["increase", "増加する"],
    ["reduce", "減らす"],
    ["confirm", "確認する"],
    ["schedule", "予定を組む"],
    ["deliver", "配達する"],
    ["announce", "発表する"],
    ["approve", "承認する"],
    ["provide", "提供する"],
    ["purchase", "購入する"],
    ["delay", "遅らせる"],
    ["attend", "出席する"],
    ["assign", "割り当てる"],
    ["maintain", "維持する"],
    ["replace", "交換する"],
    ["respond", "返答する"],
    ["inspect", "点検する"],
    ["prepare", "準備する"],
    ["submit", "提出する"],
    ["negotiate", "交渉する"],
    ["estimate", "見積もる"],
    ["promote", "昇進させる"],
    ["assist", "支援する"],
    ["expand", "拡大する"],
    ["operate", "運営する"],
    ["review", "見直す"],
    ["train", "訓練する"],
    ["arrange", "手配する"],
    ["include", "含む"],
    ["exclude", "除外する"],
    ["reserve", "予約する"],
    ["issue", "発行する"],
    ["contact", "連絡する"],
    ["manage", "管理する"],
    ["adopt", "採用する"],
    ["cancel", "取り消す"],
    ["extend", "延長する"],
    ["renew", "更新する"],
    ["allocate", "割り当てる"],
    ["authorize", "許可する"],
    ["coordinate", "調整する"],
    ["clarify", "明確にする"],
    ["comply", "従う"],
    ["conclude", "結論づける"],
    ["delegate", "委任する"],
    ["distribute", "配布する"],
    ["draft", "下書きを作成する"],
    ["enforce", "実施する"],
    ["evaluate", "評価する"],
    ["facilitate", "促進する"],
    ["finalize", "確定する"],
    ["implement", "実行する"],
    ["inform", "通知する"],
    ["launch", "開始する"],
    ["monitor", "監視する"],
    ["notify", "知らせる"],
    ["obtain", "得る"],
    ["participate", "参加する"],
    ["perform", "実行する"],
    ["prioritize", "優先する"],
    ["process", "処理する"],
    ["propose", "提案する"],
    ["reassign", "再割り当てする"],
    ["recommend", "推薦する"],
    ["reconcile", "照合する"],
    ["register", "登録する"],
    ["relocate", "移転する"],
    ["remind", "念押しする"],
    ["resolve", "解決する"],
    ["retain", "保持する"],
    ["secure", "確保する"],
    ["specify", "明示する"],
    ["streamline", "効率化する"],
    ["supervise", "監督する"],
    ["suspend", "一時停止する"],
    ["terminate", "終了させる"],
    ["update", "更新する"],
    ["verify", "検証する"],
  ];

  const meanings = vocabItems.map((item) => item[1]);
  const bank = [];
  const variantsPerWord = 7;

  vocabItems.forEach(([word, meaning]) => {
    for (let i = 0; i < variantsPerWord; i += 1) {
      bank.push({
        question: word,
        options: [meaning, ...pickRandomDistinctItems(meanings, 3, meaning)],
        answer: 0,
        toeicUnit: "Part 5 語彙・コロケーション",
        phraseHint:
          vocabPhraseHints[word] || `頻出表現: ${word} effectively in context`,
      });
    }
  });

  return bank.slice(0, 500);
}

function createReadingQuestionBank() {
  const departments = [
    "Sales team",
    "Marketing team",
    "Customer support team",
    "Human resources team",
    "Accounting team",
    "IT team",
    "Training team",
    "Logistics team",
    "Product team",
    "Operations team",
    "Procurement team",
    "Legal team",
    "Planning team",
    "Quality assurance team",
    "Finance planning team",
  ];

  const actions = [
    "update the product catalog",
    "prepare a monthly report",
    "contact key clients",
    "organize internal training",
    "check inventory records",
    "review website content",
    "improve response templates",
    "confirm delivery schedules",
    "test the payment system",
    "schedule interview sessions",
    "revise the onboarding handbook",
    "prepare a client proposal",
    "analyze complaint data",
    "update supplier contracts",
    "check compliance documents",
  ];

  const reasons = [
    "several prices changed recently",
    "the director requested a quick status update",
    "many customers asked for clearer information",
    "the company wants faster onboarding",
    "a system update will go live soon",
    "the audit team needs accurate data",
    "recent feedback highlighted repeated mistakes",
    "the next campaign starts on Monday",
    "a new policy will begin next week",
    "management wants to reduce delays",
    "a major client requested changes",
    "new employees need clear guidelines",
    "the quarterly review is approaching",
    "a process audit is scheduled soon",
    "the support queue has grown quickly",
  ];

  const deadlines = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "this weekend",
  ];

  const meetingDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const bank = [];

  departments.forEach((department, deptIndex) => {
    actions.forEach((action, actionIndex) => {
      const reason = reasons[(deptIndex + actionIndex) % reasons.length];
      const deadline = deadlines[(deptIndex * 2 + actionIndex) % deadlines.length];
      const meetingDay =
        meetingDays[(deptIndex * 3 + actionIndex) % meetingDays.length];

      const passage =
        `The ${department} will ${action} this week. ` +
        `The manager asked everyone to finish by ${deadline} because ${reason}. ` +
        `A short check-in meeting is scheduled for ${meetingDay} morning.`;

      const actionDistractors = pickRandomDistinctItems(actions, 3, action);
      const reasonDistractors = pickRandomDistinctItems(reasons, 3, reason);

      bank.push({
        passage,
        question: "What is the team planning to do this week?",
        options: [action, ...actionDistractors],
        answer: 0,
        toeicUnit: "Part 7 シングルパッセージ",
        explanation: `本文1文目に「will ${action} this week」とあるため、行動内容は ${action} です。`,
        phraseHint: `読解根拠: 1文目の "will ${action} this week" が、今週の予定を直接示しています。`,
      });

      bank.push({
        passage,
        question: "Why did the manager set a deadline?",
        options: [reason, ...reasonDistractors],
        answer: 0,
        toeicUnit: "Part 7 シングルパッセージ",
        explanation: `本文2文目の「because ${reason}」が理由を直接示しています。`,
        phraseHint: `読解根拠: 2文目の "because ${reason}" が、締切設定の理由です。`,
      });
    });
  });

  const doublePassages = [
    {
      passage:
        "Email 1\nFrom: A. Sato\nTo: Operations Team\nSubject: Delivery Schedule Change\n\nDue to road repairs near the warehouse, tomorrow's morning deliveries will start one hour later than usual. Please update customer notifications before 6:00 p.m.\n\n---\n\nNotice\nTemporary Traffic Restriction\nThe city office announced lane closures around the warehouse district from 7:00 a.m. to 11:00 a.m. on Friday. Delivery trucks should use the east bypass route.",
      question:
        "What is the most likely reason for the delivery schedule change?",
      options: [
        "Road repairs near the warehouse",
        "A shortage of drivers",
        "A software update in the logistics system",
        "A delay in customs inspection",
      ],
      answer: 0,
      phraseHint:
        "読解根拠: lane closures と start one hour later が、変更理由を示しています。",
    },
    {
      passage:
        "Email 2\nFrom: HR Office\nTo: All Staff\nSubject: Benefits Seminar\n\nA benefits seminar will be held on May 12 at 3:00 p.m. in Conference Room C. Employees are encouraged to submit questions in advance.\n\n---\n\nMemo\nTo: Department Managers\n\nPlease share the seminar registration link with your teams by Monday noon. Attendance records will be used to plan follow-up sessions.",
      question:
        "What are department managers asked to do in the memo?",
      options: [
        "Share the registration link with their teams",
        "Move the seminar to a larger room",
        "Collect attendance fees",
        "Prepare presentation slides",
      ],
      answer: 0,
      phraseHint:
        "読解根拠: memo の share the registration link が、管理者への依頼内容です。",
    },
    {
      passage:
        "Email 3\nFrom: Procurement\nTo: Finance\nSubject: Printer Purchase Request\n\nOur current office printer has frequent paper jams. We would like approval to purchase a new model by next month.\n\n---\n\nReply\nFrom: Finance\nTo: Procurement\n\nPlease attach two vendor quotations and the expected maintenance cost. We can review your request at next Tuesday's budget meeting.",
      question:
        "What does Finance request before reviewing the purchase?",
      options: [
        "Two vendor quotations and expected maintenance cost",
        "A signed contract with a vendor",
        "An updated employee directory",
        "A training plan for new staff",
      ],
      answer: 0,
      phraseHint:
        "読解根拠: attach two vendor quotations が、審査前に必要な提出物を示しています。",
    },
  ];

  doublePassages.forEach((item) => {
    bank.push({
      passage: item.passage,
      question: item.question,
      options: item.options,
      answer: item.answer,
      toeicUnit: "Part 7 ダブルパッセージ",
      explanation: "2つの文書を照合して、依頼・理由・条件を特定する問題です。",
      phraseHint: item.phraseHint,
    });
  });

  return bank;
}

const rawReadingQuestionBank = createReadingQuestionBank();
const readingSingles = rawReadingQuestionBank.filter(
  (item) => item.toeicUnit === "Part 7 シングルパッセージ"
);
const readingDoubles = rawReadingQuestionBank.filter(
  (item) => item.toeicUnit === "Part 7 ダブルパッセージ"
);
const readingQuestionBank = [
  ...sampleItems(readingSingles, 117),
  ...readingDoubles,
];
const vocabQuestionBank = createVocabQuestionBank();

function createPart6QuestionBank() {
  const passages = [
    {
      passage:
        "Subject: Visitor Procedure Update\n\nStarting Monday, all visitors must [1] at the front desk before entering office areas. They should [2] a valid photo ID to receive a temporary badge. Security staff will check badges at each entrance. [3]\n\nThank you for your cooperation.",
      q1: { answer: "sign in", options: ["sign in", "signed in", "to sign", "signing"] },
      q2: { answer: "present", options: ["present", "presents", "presented", "presenting"] },
      q3: {
        answer: "This policy is designed to improve building safety.",
        options: [
          "This policy is designed to improve building safety.",
          "The marketing campaign starts next quarter.",
          "Please submit your travel receipts by Friday.",
          "Our cafeteria now serves seasonal desserts.",
        ],
      },
      q4: {
        answer: "verification",
        options: ["verification", "verify", "verified", "verifying"],
      },
    },
    {
      passage:
        "Dear Team,\n\nThe workshop location has changed to Room 402. Participants are asked to [1] at least ten minutes early. We will [2] with a short product overview and then move to hands-on practice. [3]\n\nBest regards,\nTraining Office",
      q1: { answer: "arrive", options: ["arrive", "arrived", "arriving", "arrival"] },
      q2: { answer: "begin", options: ["begin", "began", "begun", "beginning"] },
      q3: {
        answer: "If you cannot attend, notify the office by Thursday.",
        options: [
          "If you cannot attend, notify the office by Thursday.",
          "The printer in Meeting Room B is currently offline.",
          "Our annual sales report was published last month.",
          "A delivery truck is waiting near the loading dock.",
        ],
      },
      q4: {
        answer: "attendance",
        options: ["attendance", "attend", "attending", "attended"],
      },
    },
    {
      passage:
        "Notice: Portal Maintenance\n\nThe employee portal will be unavailable from 10:00 p.m. to 1:00 a.m. During this period, users cannot [1] expense reports online. Please [2] urgent requests before 6:00 p.m. [3]\n\nIT Support Team",
      q1: { answer: "submit", options: ["submit", "submits", "submitting", "submitted"] },
      q2: { answer: "complete", options: ["complete", "completes", "completed", "completing"] },
      q3: {
        answer:
          "After maintenance, clear your browser cache if pages do not load properly.",
        options: [
          "After maintenance, clear your browser cache if pages do not load properly.",
          "The conference room reservation was canceled yesterday.",
          "Please welcome our new intern in the finance department.",
          "The company picnic has been moved to next Sunday.",
        ],
      },
      q4: {
        answer: "unavailable",
        options: ["unavailable", "unavailability", "unavailably", "unavail"],
      },
    },
    {
      passage:
        "Memo: Visitor Parking\n\nBeginning next week, guests should [1] their vehicle details at reception before parking. Staff members will [2] temporary permits after confirming appointment information. [3]\n\nFacilities Management",
      q1: { answer: "register", options: ["register", "registers", "registered", "registering"] },
      q2: { answer: "issue", options: ["issue", "issues", "issued", "issuing"] },
      q3: {
        answer:
          "Vehicles without permits may be towed during business hours.",
        options: [
          "Vehicles without permits may be towed during business hours.",
          "The accounting team completed the annual audit.",
          "Please update your email signature this afternoon.",
          "Coffee coupons are available in the lounge.",
        ],
      },
      q4: {
        answer: "registration",
        options: ["registration", "register", "registered", "registering"],
      },
    },
    {
      passage:
        "Announcement: Training Portal\n\nThe new learning portal will [1] on June 1. Employees are encouraged to [2] the orientation video before using the system. [3]\n\nLearning and Development",
      q1: { answer: "launch", options: ["launch", "launches", "launched", "launching"] },
      q2: { answer: "watch", options: ["watch", "watches", "watched", "watching"] },
      q3: {
        answer:
          "A quick-start guide is attached for first-time users.",
        options: [
          "A quick-start guide is attached for first-time users.",
          "Our office plants are watered every Tuesday.",
          "The shipping invoice was sent to a different vendor.",
          "Most interviews are scheduled in the morning.",
        ],
      },
      q4: {
        answer: "orientation",
        options: ["orientation", "orient", "oriented", "orienting"],
      },
    },
    {
      passage:
        "Email: Client Meeting\n\nPlease [1] the revised presentation slides by noon so I can review them. We should [2] the key pricing updates during tomorrow's client meeting. [3]\n\nThanks,\nSales Director",
      q1: { answer: "send", options: ["send", "sends", "sent", "sending"] },
      q2: { answer: "highlight", options: ["highlight", "highlights", "highlighted", "highlighting"] },
      q3: {
        answer:
          "The client specifically requested a summary of cost-saving options.",
        options: [
          "The client specifically requested a summary of cost-saving options.",
          "The cafeteria closes at 6 p.m. on Fridays.",
          "Our website traffic dropped slightly last winter.",
          "Please return your access card after resigning.",
        ],
      },
      q4: {
        answer: "summary",
        options: ["summary", "summarize", "summarized", "summarizing"],
      },
    },
    {
      passage:
        "Notice: Security Check\n\nAll contractors must [1] their IDs at the main gate. Security staff will [2] the access list before allowing entry. [3]\n\nSecurity Office",
      q1: { answer: "present", options: ["present", "presents", "presented", "presenting"] },
      q2: { answer: "verify", options: ["verify", "verifies", "verified", "verifying"] },
      q3: {
        answer:
          "Please arrive 15 minutes early to avoid delays at the checkpoint.",
        options: [
          "Please arrive 15 minutes early to avoid delays at the checkpoint.",
          "Our legal team revised the policy document yesterday.",
          "The conference materials were printed in color.",
          "Employee surveys will be collected next quarter.",
        ],
      },
      q4: {
        answer: "verification",
        options: ["verification", "verify", "verified", "verifying"],
      },
    },
    {
      passage:
        "Internal Message: System Upgrade\n\nThe database upgrade will [1] after office hours tonight. Engineers will [2] performance logs throughout the process. [3]\n\nIT Operations",
      q1: { answer: "begin", options: ["begin", "begins", "began", "beginning"] },
      q2: { answer: "monitor", options: ["monitor", "monitors", "monitored", "monitoring"] },
      q3: {
        answer:
          "If any issue occurs, a follow-up notice will be sent immediately.",
        options: [
          "If any issue occurs, a follow-up notice will be sent immediately.",
          "The reception desk moved to the third floor last month.",
          "Most vendors prefer payment by bank transfer.",
          "The annual picnic is planned for September.",
        ],
      },
      q4: {
        answer: "monitoring",
        options: ["monitoring", "monitor", "monitored", "monitors"],
      },
    },
  ];

  const extraPart6Passages = [
    {
      passage:
        "Notice: Equipment Return\n\nAll temporary devices must be [1] to the IT desk by Friday. Team leads should [2] the return list before noon. [3]\n\nIT Administration",
      q1: { answer: "returned", options: ["returned", "return", "returning", "returns"] },
      q2: { answer: "submit", options: ["submit", "submits", "submitted", "submitting"] },
      q3: {
        answer: "Late returns may delay next week's equipment allocation.",
        options: [
          "Late returns may delay next week's equipment allocation.",
          "The marketing budget was approved last quarter.",
          "Please check the cafeteria menu for new items.",
          "Our website was redesigned by an outside agency.",
        ],
      },
      q4: {
        answer: "allocation",
        options: ["allocation", "allocate", "allocated", "allocating"],
      },
    },
    {
      passage:
        "Email: Meeting Materials\n\nPlease [1] the updated agenda before this afternoon's meeting. We need to [2] all handouts in advance. [3]\n\nProject Office",
      q1: { answer: "review", options: ["review", "reviews", "reviewed", "reviewing"] },
      q2: { answer: "print", options: ["print", "prints", "printed", "printing"] },
      q3: {
        answer: "Several attendees requested additional background data.",
        options: [
          "Several attendees requested additional background data.",
          "The parking area will be closed for maintenance.",
          "Our support team now works in two shifts.",
          "Please label all storage boxes clearly.",
        ],
      },
      q4: {
        answer: "background",
        options: ["background", "backgrounds", "backgrounded", "backgrounding"],
      },
    },
    {
      passage:
        "Memo: Policy Update\n\nThe revised travel policy will [1] effect on July 1. Employees should [2] the new reimbursement limits carefully. [3]\n\nFinance Department",
      q1: { answer: "take", options: ["take", "takes", "taken", "taking"] },
      q2: { answer: "review", options: ["review", "reviews", "reviewed", "reviewing"] },
      q3: {
        answer: "Claims that exceed the limits may require additional approval.",
        options: [
          "Claims that exceed the limits may require additional approval.",
          "The office printer was replaced last month.",
          "Staff members are invited to the sports event.",
          "A new coffee machine arrived on Monday.",
        ],
      },
      q4: {
        answer: "approval",
        options: ["approval", "approve", "approved", "approving"],
      },
    },
    {
      passage:
        "Announcement: Training Session\n\nNew hires are expected to [1] the onboarding seminar next Monday. Managers should [2] attendance by end of day. [3]\n\nHuman Resources",
      q1: { answer: "attend", options: ["attend", "attends", "attended", "attending"] },
      q2: { answer: "confirm", options: ["confirm", "confirms", "confirmed", "confirming"] },
      q3: {
        answer: "The session materials will be shared after registration closes.",
        options: [
          "The session materials will be shared after registration closes.",
          "The legal team updated the contract template.",
          "Please submit expense forms by Friday.",
          "A client meeting was postponed until next week.",
        ],
      },
      q4: {
        answer: "registration",
        options: ["registration", "register", "registered", "registering"],
      },
    },
    {
      passage:
        "Notice: Building Access\n\nVisitors must [1] at the main counter on arrival. Security will [2] each access pass before entry. [3]\n\nAdministration Office",
      q1: { answer: "sign in", options: ["sign in", "signed in", "signing in", "signs in"] },
      q2: { answer: "check", options: ["check", "checks", "checked", "checking"] },
      q3: {
        answer: "Please bring a valid photo ID to speed up the check-in process.",
        options: [
          "Please bring a valid photo ID to speed up the check-in process.",
          "The annual report will be distributed next month.",
          "Our inventory system is currently offline.",
          "A supplier has requested an earlier payment date.",
        ],
      },
      q4: {
        answer: "process",
        options: ["process", "processes", "processed", "processing"],
      },
    },
    {
      passage:
        "Internal Update: Website Maintenance\n\nThe customer portal will [1] temporarily unavailable tonight. Engineers will [2] system logs during maintenance. [3]\n\nWeb Operations",
      q1: { answer: "be", options: ["be", "is", "was", "being"] },
      q2: { answer: "monitor", options: ["monitor", "monitors", "monitored", "monitoring"] },
      q3: {
        answer: "Users may experience slower response times after the update.",
        options: [
          "Users may experience slower response times after the update.",
          "The finance team finalized next quarter's budget.",
          "Please send your profile photo to HR.",
          "A workshop on communication will be held tomorrow.",
        ],
      },
      q4: {
        answer: "response",
        options: ["response", "respond", "responded", "responding"],
      },
    },
    {
      passage:
        "Email: Vendor Contract\n\nPlease [1] the revised contract draft by Wednesday. We must [2] all clauses with the legal team first. [3]\n\nProcurement",
      q1: { answer: "submit", options: ["submit", "submits", "submitted", "submitting"] },
      q2: { answer: "review", options: ["review", "reviews", "reviewed", "reviewing"] },
      q3: {
        answer: "Any missing signatures will delay the final approval process.",
        options: [
          "Any missing signatures will delay the final approval process.",
          "The travel desk moved to the third floor.",
          "Please clean your workspace before leaving.",
          "Our support staff completed system training.",
        ],
      },
      q4: {
        answer: "approval",
        options: ["approval", "approve", "approved", "approving"],
      },
    },
    {
      passage:
        "Reminder: Invoice Processing\n\nAll invoices should be [1] in the shared folder by noon. Accountants will [2] each entry before payment is released. [3]\n\nAccounting Team",
      q1: { answer: "uploaded", options: ["uploaded", "upload", "uploading", "uploads"] },
      q2: { answer: "verify", options: ["verify", "verifies", "verified", "verifying"] },
      q3: {
        answer: "Incorrect account numbers can cause significant payment delays.",
        options: [
          "Incorrect account numbers can cause significant payment delays.",
          "The training room capacity is now 40 seats.",
          "A new vendor orientation starts next Monday.",
          "Please lock all cabinets after office hours.",
        ],
      },
      q4: {
        answer: "payment",
        options: ["payment", "pay", "paid", "paying"],
      },
    },
  ];
  passages.push(...extraPart6Passages);

  const bank = [];
  passages.forEach((item) => {
    bank.push({
      passage: item.passage.replace("[1]", "_____"),
      question: "Blank 1: Choose the best word or phrase.",
      options: item.q1.options,
      answer: item.q1.options.indexOf(item.q1.answer),
      category: "語法",
      toeicUnit: "Part 6 テキスト完成",
      explanation: `本文の文法・語法に合う語は「${item.q1.answer}」です。`,
      phraseHint: `本文表現: ${item.q1.answer}`,
    });
    bank.push({
      passage: item.passage.replace("[2]", "_____"),
      question: "Blank 2: Choose the best word or phrase.",
      options: item.q2.options,
      answer: item.q2.options.indexOf(item.q2.answer),
      category: "品詞/時制",
      toeicUnit: "Part 6 テキスト完成",
      explanation: `文脈と文型に最も自然に入る語は「${item.q2.answer}」です。`,
      phraseHint: `本文表現: ${item.q2.answer}`,
    });
    bank.push({
      passage: item.passage.replace("[3]", "_____"),
      question: "Blank 3: Which sentence best fits in the blank?",
      options: item.q3.options,
      answer: item.q3.options.indexOf(item.q3.answer),
      category: "文挿入",
      toeicUnit: "Part 6 テキスト完成",
      explanation:
        "前後の内容とつながる一文を選ぶ問題です。文脈に合う文を選ぶ必要があります。",
      phraseHint: "Part 6ポイント: 文脈に最も自然な文を選ぶ",
    });
    bank.push({
      passage:
        item.passage
          .replace(
            "\n\nThank you for your cooperation.",
            "\n\nID [4] must be completed before a badge is issued.\n\nThank you for your cooperation."
          )
          .replace(
            "\n\nBest regards,\nTraining Office",
            "\n\nPlease complete online [4] by 10:00 a.m.\n\nBest regards,\nTraining Office"
          )
          .replace(
            "\n\nIT Support Team",
            "\n\nThe service may remain [4] for a few minutes after restart.\n\nIT Support Team"
          )
          .replace("[4]", "_____"),
      question: "Blank 4: Choose the best word form for the context.",
      options: item.q4.options,
      answer: item.q4.options.indexOf(item.q4.answer),
      category: "語形",
      toeicUnit: "Part 6 テキスト完成",
      explanation: `空欄の文法位置に合う語形は「${item.q4.answer}」です。`,
      phraseHint: `本文表現: ${item.q4.answer}`,
    });
  });

  return bank;
}

function createClozeQuestionBank() {
  const bank = [];
  const verbItems = [
    {
      sentence: "Please _____ the report by 5 p.m. so the manager can review it.",
      answer: "submit",
      distractors: ["submitting", "submitted", "submission"],
      category: "品詞/語法",
    },
    {
      sentence: "The meeting was _____ because several team members were absent.",
      answer: "postponed",
      distractors: ["postpone", "postponing", "postponement"],
      category: "品詞/語法",
    },
    {
      sentence: "If you have any questions, you can _____ the support desk directly.",
      answer: "contact",
      distractors: ["contacted", "contacting", "contacts"],
      category: "品詞/語法",
    },
    {
      sentence: "The company plans to _____ a new branch in Osaka next year.",
      answer: "open",
      distractors: ["opened", "opening", "opens"],
      category: "品詞/語法",
    },
    {
      sentence: "All visitors must _____ at the front desk before entering.",
      answer: "sign in",
      distractors: ["signing in", "signed in", "signs in"],
      category: "品詞/語法",
    },
    {
      sentence: "Please _____ the updated invoice to the accounting team.",
      answer: "forward",
      distractors: ["forwarding", "forwarded", "forwards"],
      category: "品詞/語法",
    },
    {
      sentence: "The supervisor asked us to _____ the schedule immediately.",
      answer: "revise",
      distractors: ["revised", "revising", "revision"],
      category: "品詞/語法",
    },
    {
      sentence: "Employees should _____ their timesheets by Friday evening.",
      answer: "submit",
      distractors: ["submitted", "submitting", "submission"],
      category: "品詞/語法",
    },
  ];

  const adverbItems = [
    ["Our team needs to respond _____ to customer emails during business hours.", "quickly", ["quick", "quickness", "quicked"]],
    ["The instructions were written _____ so everyone could understand them.", "clearly", ["clear", "clearness", "cleared"]],
    ["She handled the complaint _____ and solved the issue right away.", "professionally", ["professional", "profession", "professioned"]],
    ["The engineer checked the system _____ before the launch.", "carefully", ["careful", "carefulness", "carefuled"]],
    ["Please fill out the form _____ to avoid mistakes.", "accurately", ["accurate", "accuracy", "accurated"]],
  ];

  const connectorItems = [
    ["_____ the budget was limited, the team completed the project on time.", "Although", ["Because", "Unless", "However"]],
    ["Sales increased; _____, the company hired two additional staff members.", "therefore", ["despite", "whereas", "meanwhile"]],
    ["_____ the proposal is approved, we will start development next week.", "If", ["Unless", "Despite", "However"]],
    ["He attended the seminar _____ he wanted to improve his presentation skills.", "because", ["although", "unless", "however"]],
    ["The package will be shipped today _____ payment is confirmed.", "once", ["while", "though", "despite"]],
    ["The office remained open, _____ many employees worked remotely.", "although", ["because", "unless", "therefore"]],
  ];

  const prepositionItems = [
    ["The conference will be held _____ April 15 at the city hall.", "on", ["at", "in", "for"]],
    ["Please send the document _____ email before noon.", "by", ["from", "with", "through"]],
    ["All applicants must submit the form _____ the deadline.", "before", ["between", "during", "until"]],
    ["The manager is responsible _____ training new staff members.", "for", ["to", "in", "at"]],
    ["We discussed the issue _____ detail during the meeting.", "in", ["on", "with", "to"]],
    ["Please arrive _____ time to set up the equipment.", "on", ["at", "for", "in"]],
    ["The discount is available _____ all members this month.", "to", ["for", "at", "with"]],
    ["They divided the tasks _____ three project teams.", "among", ["between", "for", "during"]],
  ];

  const tenseItems = [
    ["By the time he arrived, the meeting had already _____.", "started", ["start", "starting", "starts"]],
    ["The report _____ by Ms. Tanaka every Friday.", "is prepared", ["prepare", "prepared", "preparing"]],
    ["If sales continue to rise, the company _____ more staff next year.", "will hire", ["hires", "hired", "hiring"]],
    ["We _____ the final design when the client requested a revision.", "had completed", ["complete", "have completed", "completing"]],
    ["The CEO said that the new policy _____ next month.", "would begin", ["begins", "has begun", "beginning"]],
  ];

  for (let i = 0; i < 12; i += 1) {
    verbItems.forEach((template) => {
      bank.push({
        question: template.sentence,
        options: [template.answer, ...template.distractors],
        answer: 0,
        toeicUnit: "Part 5 文法・語法",
        category: template.category,
        explanation: `空欄には「${template.answer}」が入り、文法的に自然な形になります。`,
        phraseHint: `文中表現: ${template.sentence.replace("_____", template.answer)}`,
      });
    });
  }

  adverbItems.forEach(([question, answer, distractors]) => {
    bank.push({
      question,
      options: [answer, ...distractors],
      answer: 0,
      toeicUnit: "Part 5 文法・語法",
      category: "副詞",
      explanation: `動詞を修飾する副詞が必要なので「${answer}」が正解です。`,
      phraseHint: `文中熟語: ${question.replace("_____", answer)}`,
    });
  });

  connectorItems.forEach(([question, answer, distractors]) => {
    bank.push({
      question,
      options: [answer, ...distractors],
      answer: 0,
      toeicUnit: "Part 5 文法・語法",
      category: "接続詞",
      explanation: `前後の意味関係をつなぐ接続詞として「${answer}」が最適です。`,
      phraseHint: `文脈接続: ${question.replace("_____", answer)}`,
    });
  });

  prepositionItems.forEach(([question, answer, distractors]) => {
    bank.push({
      question,
      options: [answer, ...distractors],
      answer: 0,
      toeicUnit: "Part 5 文法・語法",
      category: "前置詞",
      explanation: `前置詞の定型用法により「${answer}」が適切です。`,
      phraseHint: `定型表現: ${question.replace("_____", answer)}`,
    });
  });

  tenseItems.forEach(([question, answer, distractors]) => {
    bank.push({
      question,
      options: [answer, ...distractors],
      answer: 0,
      toeicUnit: "Part 5 文法・語法",
      category: "時制/態",
      explanation: `文脈の時制・態に一致するのは「${answer}」です。`,
      phraseHint: `時制ポイント: ${question.replace("_____", answer)}`,
    });
  });

  return bank;
}

const clozeQuestionBank = createClozeQuestionBank();
const part6QuestionBank = createPart6QuestionBank();

function renderHomeQuestionCounts() {
  const box = document.getElementById("home-question-counts");
  if (!box) return;
  const total =
    vocabQuestionBank.length +
    clozeQuestionBank.length +
    part6QuestionBank.length +
    readingQuestionBank.length;
  box.innerHTML = `
    <p><strong>問題数（現在）</strong></p>
    <ul>
      <li>英単語: ${vocabQuestionBank.length}問</li>
      <li>短文穴埋め問題（Part 5）: ${clozeQuestionBank.length}問</li>
      <li>長文穴埋め問題（Part 6）: ${part6QuestionBank.length}問</li>
      <li>長文読解問題（Part 7）: ${readingQuestionBank.length}問</li>
      <li>合計: ${total}問</li>
    </ul>
  `;
}

function activateTab(tabName) {
  if (!tabName || !document.getElementById(tabName)) return;
  document.querySelectorAll(".tab-button").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-content").forEach((content) => {
    content.classList.toggle("active", content.id === tabName);
  });
  if (tabName === "records") {
    renderRecordsDashboard();
  }
}

function setupTabs() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      activateTab(button.dataset.tab);
    });
  });
  document.querySelectorAll("[data-open-tab]").forEach((el) => {
    el.addEventListener("click", () => {
      const tab = el.getAttribute("data-open-tab");
      if (tab) activateTab(tab);
    });
  });
}

function getStateSafely(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (_error) {
    return null;
  }
}

function formatRate(correct, attempted) {
  if (!attempted) return "0.0";
  return ((correct / attempted) * 100).toFixed(1);
}

function getDailyGoal() {
  const raw = localStorage.getItem(DAILY_GOAL_KEY);
  const value = Number(raw);
  if (!Number.isFinite(value) || value <= 0) {
    return DEFAULT_DAILY_GOAL;
  }
  return Math.floor(value);
}

function setDailyGoal(value) {
  localStorage.setItem(DAILY_GOAL_KEY, String(value));
}

function buildDailyGoalOptions(currentGoal) {
  const presets = [
    5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 120, 150, 200,
  ];
  const values = presets.includes(currentGoal)
    ? presets
    : [...presets, currentGoal].sort((a, b) => a - b);
  return values
    .map(
      (value) =>
        `<option value="${value}" ${
          value === currentGoal ? "selected" : ""
        }>${value} 問</option>`
    )
    .join("");
}

function topWeakCategories(state) {
  if (!state || !state.categoryStats) return [];
  const entries = Object.entries(state.categoryStats).map(([name, stats]) => {
    const attempted = Number(stats.attempted || 0);
    const correct = Number(stats.correct || 0);
    const wrong = Math.max(attempted - correct, 0);
    return { name, attempted, correct, wrong };
  });
  return entries
    .filter((item) => item.attempted > 0)
    .sort((a, b) => b.wrong - a.wrong)
    .slice(0, 3);
}

function renderRecordsDashboard() {
  const box = document.getElementById("records-dashboard");
  if (!box) return;

  const vocab = getStateSafely(VOCAB_STATE_KEY) || {};
  const cloze = getStateSafely(CLOZE_STATE_KEY) || {};
  const part6 = getStateSafely(PART6_STATE_KEY) || {};
  const reading = getStateSafely(READING_STATE_KEY) || {};
  const totalAttempted =
    Number(vocab.attempted || 0) +
    Number(cloze.attempted || 0) +
    Number(part6.attempted || 0) +
    Number(reading.attempted || 0);
  const totalCorrect =
    Number(vocab.correct || 0) +
    Number(cloze.correct || 0) +
    Number(part6.correct || 0) +
    Number(reading.correct || 0);
  const weak = topWeakCategories(cloze);
  const today = getTodayKey();
  const vocabToday =
    vocab.dailyStats && vocab.dailyStats.date === today
      ? Number(vocab.dailyStats.attempted || 0)
      : 0;
  const clozeToday =
    cloze.dailyStats && cloze.dailyStats.date === today
      ? Number(cloze.dailyStats.attempted || 0)
      : 0;
  const readingToday =
    reading.dailyStats && reading.dailyStats.date === today
      ? Number(reading.dailyStats.attempted || 0)
      : 0;
  const part6Today =
    part6.dailyStats && part6.dailyStats.date === today
      ? Number(part6.dailyStats.attempted || 0)
      : 0;
  const todayAttempted = vocabToday + clozeToday + part6Today + readingToday;
  const dailyGoal = getDailyGoal();
  const progressPercent = Math.min(
    100,
    Math.round((todayAttempted / dailyGoal) * 100)
  );
  const clozeWrongOptions = Object.entries(cloze.wrongOptionStats || {})
    .map(([key, count]) => ({ key, count: Number(count || 0) }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 3);
  const mock = getStateSafely(MOCK_STATE_KEY) || {};
  const recordMap = loadDailyRecord();
  const sortedDates = Object.keys(recordMap).sort();
  const recentDates = sortedDates.slice(-7).reverse();
  const totalStudyDays = sortedDates.filter(
    (key) => getAttemptedFromRecord(recordMap[key]) > 0
  ).length;
  const cumulativeAttempted = sortedDates.reduce(
    (sum, key) => sum + getAttemptedFromRecord(recordMap[key]),
    0
  );
  const cumulativeCorrect = sortedDates.reduce(
    (sum, key) => sum + getCorrectFromRecord(recordMap[key]),
    0
  );
  const currentStreak = computeCurrentStreak(sortedDates, recordMap);
  const firstStudyDate = sortedDates.length > 0 ? sortedDates[0] : null;
  const historyRows = recentDates
    .map((dateKey) => {
      const item = recordMap[dateKey] || getEmptyDailyRecord();
      const dayAttempted = getAttemptedFromRecord(item);
      const dayCorrect = getCorrectFromRecord(item);
      return `<li>${dateKey}: ${dayAttempted}問（正答率 ${formatRate(
        dayCorrect,
        dayAttempted
      )}%）</li>`;
    })
    .join("");

  box.innerHTML = `
    <h3>学習ダッシュボード</h3>
    <div class="control-row">
      <label class="control-item">今日の目標:
        <select id="daily-goal-input">
          ${buildDailyGoalOptions(dailyGoal)}
        </select>
      </label>
      <button id="save-daily-goal" class="next-btn">目標を保存</button>
    </div>
    <p><strong>今日の進捗:</strong> ${todayAttempted} / ${dailyGoal} 問 (${progressPercent}%)</p>
    <div class="progress-track">
      <div class="progress-fill" style="width:${progressPercent}%"></div>
    </div>
    <div class="dashboard-grid">
      <div class="dashboard-card">
        <p><strong>総回答数</strong></p>
        <p>${totalAttempted} 問</p>
      </div>
      <div class="dashboard-card">
        <p><strong>総正答率</strong></p>
        <p>${formatRate(totalCorrect, totalAttempted)}%</p>
      </div>
      <div class="dashboard-card">
        <p><strong>穴埋め正答率</strong></p>
        <p>${formatRate(Number(cloze.correct || 0), Number(cloze.attempted || 0))}%</p>
      </div>
      <div class="dashboard-card">
        <p><strong>長文正答率</strong></p>
        <p>${formatRate(Number(reading.correct || 0), Number(reading.attempted || 0))}%</p>
      </div>
      <div class="dashboard-card">
        <p><strong>Part 6 正答率</strong></p>
        <p>${formatRate(Number(part6.correct || 0), Number(part6.attempted || 0))}%</p>
      </div>
      <div class="dashboard-card">
        <p><strong>ミニ模試ベスト</strong></p>
        <p>${Number(mock.bestScore || 0)} / 10</p>
      </div>
    </div>
    <p><strong>穴埋めの弱点カテゴリ</strong></p>
    <p>${
      weak.length > 0
        ? weak
            .map(
              (item) =>
                `${item.name}（誤答 ${item.wrong} / ${item.attempted}）`
            )
            .join(" / ")
        : "まだデータがありません"
    }</p>
    <p><strong>穴埋めの誤答しやすい選択肢</strong></p>
    <p>${
      clozeWrongOptions.length > 0
        ? clozeWrongOptions
            .map((item) => `${item.key.split("::")[1]}（${item.count}回）`)
            .join(" / ")
        : "まだデータがありません"
    }</p>
    <p><strong>日別の学習記録（直近7日）</strong></p>
    <p><strong>学習継続:</strong> ${currentStreak}日連続 / <strong>累計学習日数:</strong> ${totalStudyDays}日</p>
    <p><strong>これまでの累計:</strong> ${cumulativeAttempted}問（正答率 ${formatRate(
      cumulativeCorrect,
      cumulativeAttempted
    )}%）</p>
    <p><strong>記録開始日:</strong> ${firstStudyDate || "まだ記録がありません"}</p>
    <ul class="daily-history">
      ${historyRows || "<li>まだ記録がありません</li>"}
    </ul>
  `;

  const saveGoalButton = document.getElementById("save-daily-goal");
  const dailyGoalInput = document.getElementById("daily-goal-input");
  if (saveGoalButton && dailyGoalInput) {
    saveGoalButton.addEventListener("click", () => {
      const nextGoal = Number(dailyGoalInput.value);
      if (Number.isFinite(nextGoal) && nextGoal > 0) {
        setDailyGoal(Math.floor(nextGoal));
        renderRecordsDashboard();
      }
    });
  }
}

function setupMiniMockTest(questionBanks) {
  const container = document.getElementById("mock-test");
  if (!container) return;

  let timerId = null;
  let session = {
    started: false,
    completed: false,
    isReviewMode: false,
    questions: [],
    index: 0,
    score: 0,
    timeLeft: 360,
    wrongQuestions: [],
    categoryStats: {},
    answerLog: [],
  };

  function readMockState() {
    const state = getStateSafely(MOCK_STATE_KEY) || {};
    return {
      bestScore: Number(state.bestScore || 0),
      attempts: Number(state.attempts || 0),
      recentWrongQuestions: Array.isArray(state.recentWrongQuestions)
        ? state.recentWrongQuestions
        : [],
      recentCategoryStats: state.recentCategoryStats || {},
    };
  }

  function saveMockState(score, wrongQuestions, categoryStats, countAsExam) {
    const prev = readMockState();
    const next = {
      bestScore: countAsExam ? Math.max(prev.bestScore, score) : prev.bestScore,
      attempts: countAsExam ? prev.attempts + 1 : prev.attempts,
      recentWrongQuestions: wrongQuestions,
      recentCategoryStats: categoryStats,
    };
    localStorage.setItem(MOCK_STATE_KEY, JSON.stringify(next));
  }

  function pickFromPool(pool, count) {
    const unique = [...pool];
    const picked = [];
    for (let i = 0; i < count && unique.length > 0; i += 1) {
      const idx = Math.floor(Math.random() * unique.length);
      picked.push(shuffleQuestion(unique[idx]));
      unique.splice(idx, 1);
    }
    return picked;
  }

  function pickMockQuestions() {
    const blueprint = [
      { key: "cloze", count: 2 },
      { key: "part6", count: 3 },
      { key: "readingSingle", count: 2 },
      { key: "readingDouble", count: 1 },
      { key: "vocab", count: 2 },
    ];
    const selected = [];

    blueprint.forEach(({ key, count }) => {
      const pool = Array.isArray(questionBanks[key]) ? questionBanks[key] : [];
      selected.push(...pickFromPool(pool, count));
    });

    while (selected.length < 10) {
      const fallbackPool = [
        ...(questionBanks.cloze || []),
        ...(questionBanks.part6 || []),
        ...(questionBanks.readingSingle || []),
        ...(questionBanks.readingDouble || []),
        ...(questionBanks.vocab || []),
      ];
      if (fallbackPool.length === 0) break;
      selected.push(...pickFromPool(fallbackPool, 1));
    }

    return selected.slice(0, 10);
  }

  function stopTimer() {
    if (timerId) {
      window.clearInterval(timerId);
      timerId = null;
    }
  }

  function finishExam() {
    stopTimer();
    const finalScore = session.score;
    const finalTotal = session.questions.length;
    session.completed = true;
    session.started = false;
    saveMockState(
      session.score,
      session.wrongQuestions,
      session.categoryStats,
      !session.isReviewMode
    );
    render();
    renderRecordsDashboard();
    if (finalTotal > 0) {
      const ratio = finalScore / finalTotal;
      window.setTimeout(() => {
        const cx = window.innerWidth / 2;
        const cy = window.innerHeight * 0.32;
        if (finalScore === finalTotal) {
          spawnConfettiAt(cx, cy, "large");
        } else if (ratio >= 0.8) {
          spawnConfettiAt(cx, cy, "medium");
        } else if (ratio >= 0.5) {
          spawnConfettiAt(cx, cy, "small");
        }
      }, 120);
    }
  }

  function getRecommendedCategory(categoryStats) {
    const entries = Object.entries(categoryStats || {}).map(([name, stats]) => {
      const attempted = Number(stats.attempted || 0);
      const correct = Number(stats.correct || 0);
      const rate = attempted > 0 ? correct / attempted : 1;
      return { name, attempted, rate };
    });
    const candidates = entries.filter((item) => item.attempted > 0);
    if (candidates.length === 0) return null;
    candidates.sort((a, b) => a.rate - b.rate);
    return candidates[0].name;
  }

  function startExam(seconds, reviewQuestions = []) {
    const isReviewMode = reviewQuestions.length > 0;
    const selectedQuestions = isReviewMode ? reviewQuestions : pickMockQuestions();
    const maxSeconds = isReviewMode ? Math.max(120, selectedQuestions.length * 25) : seconds;
    session = {
      started: true,
      completed: false,
      isReviewMode,
      questions: selectedQuestions,
      index: 0,
      score: 0,
      timeLeft: maxSeconds,
      wrongQuestions: [],
      categoryStats: {},
      answerLog: [],
    };
    stopTimer();
    timerId = window.setInterval(() => {
      session.timeLeft -= 1;
      if (session.timeLeft <= 0) {
        session.timeLeft = 0;
        finishExam();
        return;
      }
      const timer = document.getElementById("mock-timer");
      if (timer) timer.textContent = `残り時間: ${formatSeconds(session.timeLeft)}`;
    }, 1000);
    render();
  }

  function render() {
    const history = readMockState();
    if (!session.started && !session.completed) {
      container.innerHTML = `
        <p>TOEIC Reading縮小版の10問ミニ模試です（Part 5×2 / Part 6×3 / Part 7シングル×2 / Part 7ダブル×1 / 語彙×2）。全体タイマーで実戦感覚を鍛えます。</p>
        <p><strong>ベストスコア:</strong> ${history.bestScore} / 10（受験回数: ${history.attempts}）</p>
        <div class="control-row">
          <label class="control-item">制限時間:
            <select id="mock-duration">
              <option value="240">4分</option>
              <option value="360" selected>6分</option>
              <option value="480">8分</option>
            </select>
          </label>
          <button id="mock-start" class="next-btn">模試を開始</button>
          <button id="mock-review-start" class="next-btn" ${
            history.recentWrongQuestions.length === 0 ? "disabled" : ""
          }>解き直し（誤答のみ）</button>
        </div>
      `;
      const startBtn = document.getElementById("mock-start");
      const reviewBtn = document.getElementById("mock-review-start");
      const duration = document.getElementById("mock-duration");
      if (startBtn && duration) {
        startBtn.addEventListener("click", () => {
          startExam(Number(duration.value));
        });
      }
      if (reviewBtn) {
        reviewBtn.addEventListener("click", () => {
          const latest = readMockState();
          if (latest.recentWrongQuestions.length > 0) {
            startExam(240, latest.recentWrongQuestions.map((q) => shuffleQuestion(q)));
          }
        });
      }
      return;
    }

    if (session.completed) {
      const totalQuestions = session.questions.length;
      const accuracy = formatRate(session.score, totalQuestions);
      const recommendedCategory = getRecommendedCategory(session.categoryStats);
      const categoryRows = Object.entries(session.categoryStats)
        .map(([name, stats]) => {
          const attempted = Number(stats.attempted || 0);
          const correct = Number(stats.correct || 0);
          return `<li>${name}: ${correct}/${attempted}（${formatRate(
            correct,
            attempted
          )}%）</li>`;
        })
        .join("");
      const answerReviewRows = session.answerLog
        .map((item) => {
          const selectedText =
            typeof item.selectedIndex === "number"
              ? item.options[item.selectedIndex]
              : "未回答";
          const correctText = item.options[item.answerIndex];
          const resultLabel = item.isCorrect ? "正解" : "不正解";
          const optionRows = item.options
            .map((option, idx) => {
              const isSelected = idx === item.selectedIndex;
              const isCorrectOption = idx === item.answerIndex;
              const optionLabel = String.fromCharCode(65 + idx);
              const rowClass = isCorrectOption
                ? "review-option review-option-correct"
                : isSelected
                  ? "review-option review-option-selected"
                  : "review-option";
              return `<li class="${rowClass}">${optionLabel}. ${option}</li>`;
            })
            .join("");
          return `
            <div class="dashboard-card">
              <p><strong>Q${item.questionNo} (${item.toeicUnit || "TOEIC"})</strong></p>
              ${item.passage ? `<p>${formatPassageHtml(item.passage)}</p>` : ""}
              <p>${item.questionText}</p>
              <ul class="review-options">${optionRows}</ul>
              <p><strong>判定:</strong> ${resultLabel}</p>
              <p><strong>あなたの回答:</strong> ${selectedText}</p>
              <p><strong>正解:</strong> ${correctText}</p>
              ${
                item.explanation
                  ? `<p><strong>解説:</strong> ${item.explanation}</p>`
                  : ""
              }
              ${
                item.phraseHint
                  ? `<p><strong>ポイント:</strong> ${createPhraseText({
                      phraseHint: item.phraseHint,
                    })}</p>`
                  : ""
              }
            </div>
          `;
        })
        .join("");
      container.innerHTML = `
        <p class="result">模試終了: ${session.score} / ${totalQuestions}</p>
        <p><strong>正答率:</strong> ${accuracy}%</p>
        <p><strong>ベストスコア:</strong> ${readMockState().bestScore} / 10</p>
        <p><strong>カテゴリ別レビュー</strong></p>
        <ul>${categoryRows || "<li>データなし</li>"}</ul>
        <p><strong>次に解くべきカテゴリ:</strong> ${
          recommendedCategory || "データ不足"
        }</p>
        <p><strong>全問題の答え合わせ</strong></p>
        <div class="dashboard-grid">
          ${answerReviewRows || "<p>回答履歴がありません</p>"}
        </div>
        <button id="mock-retry" class="next-btn">もう一度挑戦</button>
        <button id="mock-review-retry" class="next-btn" ${
          session.wrongQuestions.length === 0 ? "disabled" : ""
        }>この回の誤答を解き直す</button>
      `;
      const retry = document.getElementById("mock-retry");
      const reviewRetry = document.getElementById("mock-review-retry");
      if (retry) {
        retry.addEventListener("click", () => {
          session.completed = false;
          render();
        });
      }
      if (reviewRetry) {
        reviewRetry.addEventListener("click", () => {
          if (session.wrongQuestions.length > 0) {
            startExam(240, session.wrongQuestions.map((q) => shuffleQuestion(q)));
          }
        });
      }
      return;
    }

    const q = session.questions[session.index];
    const totalQuestions = session.questions.length;
    container.innerHTML = `
      <div class="mock-header">
        <p><strong>${session.isReviewMode ? "解き直し" : "模試"} Q${session.index + 1} / ${totalQuestions}</strong></p>
        <p id="mock-timer"><strong>残り時間: ${formatSeconds(session.timeLeft)}</strong></p>
      </div>
      ${
        q.passage
          ? `<p><strong>本文（${q.toeicUnit || "TOEIC"}）</strong><br>${q.passage.replace(
              /\n/g,
              "<br>"
            )}</p>`
          : `<p><strong>出題単元:</strong> ${q.toeicUnit || "TOEIC Reading"}</p>`
      }
      <p><strong>${q.question}</strong></p>
      <div class="options">
        ${q.options
          .map(
            (option, idx) =>
              `<button class="option-btn" data-idx="${idx}">${option}</button>`
          )
          .join("")}
      </div>
      <button id="mock-quit" class="next-btn">終了</button>
    `;

    const buttons = container.querySelectorAll(".option-btn");
    const quit = document.getElementById("mock-quit");

    buttons.forEach((button) => {
      button.addEventListener("click", () => {
        const selected = Number(button.dataset.idx);
        const isCorrect = selected === q.answer;
        const category = q.category || "Part5";
        session.answerLog.push({
          questionNo: session.index + 1,
          toeicUnit: q.toeicUnit || "TOEIC Reading",
          passage: q.passage || "",
          questionText: q.question,
          options: [...q.options],
          selectedIndex: selected,
          answerIndex: q.answer,
          isCorrect,
          explanation: q.explanation || "",
          phraseHint: q.phraseHint || "",
        });
        if (!session.categoryStats[category]) {
          session.categoryStats[category] = { attempted: 0, correct: 0 };
        }
        session.categoryStats[category].attempted += 1;
        if (isCorrect) session.score += 1;
        if (isCorrect) {
          session.categoryStats[category].correct += 1;
        } else {
          session.wrongQuestions.push({
            question: q.question,
            options: [...q.options],
            answer: q.answer,
            category,
            explanation: q.explanation,
            phraseHint: q.phraseHint,
          });
        }
        buttons.forEach((btn) => (btn.disabled = true));
        window.setTimeout(() => {
          session.index += 1;
          if (session.index >= totalQuestions) {
            finishExam();
          } else {
            render();
          }
        }, 120);
      });
    });

    quit.addEventListener("click", () => {
      finishExam();
    });
  }

  render();
}

function setupReadingQuiz() {
  buildInfiniteQuestionSession({
    containerId: "reading-quiz",
    stateKey: READING_STATE_KEY,
    questionBank: readingQuestionBank,
    titlePrefix: "長文",
    showPassage: true,
    randomizePrompt: true,
    sectionName: "reading",
  });
}

setupTabs();
buildInfiniteQuestionSession({
  containerId: "vocab-quiz",
  stateKey: VOCAB_STATE_KEY,
  questionBank: vocabQuestionBank,
  titlePrefix: "",
  showPassage: false,
  randomizePrompt: true,
  fallbackCategory: "語彙",
  simplePromptMode: true,
  enableUnitFilter: true,
  sectionName: "vocab",
});
buildInfiniteQuestionSession({
  containerId: "cloze-quiz",
  stateKey: CLOZE_STATE_KEY,
  questionBank: clozeQuestionBank,
  titlePrefix: "",
  showPassage: false,
  randomizePrompt: true,
  enableUnitFilter: true,
  fallbackCategory: "Part5",
  sectionName: "cloze",
});
buildInfiniteQuestionSession({
  containerId: "part6-quiz",
  stateKey: PART6_STATE_KEY,
  questionBank: part6QuestionBank,
  titlePrefix: "Part 6",
  showPassage: true,
  randomizePrompt: true,
  enableUnitFilter: true,
  fallbackCategory: "Part6",
  sectionName: "part6",
});
setupReadingQuiz();
setupMiniMockTest({
  cloze: clozeQuestionBank,
  part6: part6QuestionBank,
  readingSingle: readingQuestionBank.filter(
    (item) => item.toeicUnit === "Part 7 シングルパッセージ"
  ),
  readingDouble: readingQuestionBank.filter(
    (item) => item.toeicUnit === "Part 7 ダブルパッセージ"
  ),
  vocab: vocabQuestionBank,
});
renderHomeQuestionCounts();
renderRecordsDashboard();

const refreshDashboardButton = document.getElementById(
  "refresh-records-dashboard"
);
if (refreshDashboardButton) {
  refreshDashboardButton.addEventListener("click", () => {
    renderRecordsDashboard();
  });
}
