const STORAGE_KEY = "rung-chuong-vang-builder";
const AUTO_ADVANCE_DELAY_MS = 2200;
const REVEAL_SOUND_DELAY_MS = 180;
const EXCEL_HEADER_ALIASES = {
  question: ["question", "cauhoi", "cau_hoi", "noi_dung", "noi_dung_cau_hoi"],
  type: ["type", "loai", "loai_cau_hoi"],
  media: ["media", "link_media", "url_media", "video_or_image", "video_hinh"],
  timer: ["timer", "time", "thoi_gian", "thoigian"],
  answerA: ["answer_a", "a", "dap_an_a", "dapan_a"],
  answerB: ["answer_b", "b", "dap_an_b", "dapan_b"],
  answerC: ["answer_c", "c", "dap_an_c", "dapan_c"],
  answerD: ["answer_d", "d", "dap_an_d", "dapan_d"],
  correctAnswer: ["correct_answer", "correct", "dap_an_dung", "dapan_dung"],
  awarded: ["awarded", "score_awarded", "cong_diem", "diem", "da_cong_diem"],
};

const LEGACY_DEFAULT_TEXT = {
  title: "Chuong Trinh Rung Chuong Vang",
  subtitle:
    "San choi giup hoc sinh tieu hoc cung co kien thuc qua cac cau hoi vui nhon, de nho.",
  rulesContent:
    "Moi cau hoi co 4 dap an A, B, C, D va thi sinh chon 1 dap an dung.\nSau khi doc hoac xem cau hoi, thi sinh tra loi trong thoi gian dong ho cat dem nguoc.\nKhi het gio, chuong trinh co the tu dong hien dap an dung de doi chieu.\nNguoi dan chuong trinh co the cong diem cho tung cau va tong ket o slide cuoi.",
  endingTitle: "Ket thuc chuong trinh",
  endingMessage:
    "Cam on cac em hoc sinh da tham gia. Day la bang tong ket diem va hanh trinh chinh phuc cac cau hoi.",
};

function makeId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }

  return `q-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

const sampleImage = `data:image/svg+xml;utf8,${encodeURIComponent(`
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 450">
    <defs>
      <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stop-color="#fff4bf"/>
        <stop offset="100%" stop-color="#d7edff"/>
      </linearGradient>
    </defs>
    <rect width="800" height="450" fill="url(#g)"/>
    <circle cx="120" cy="100" r="46" fill="#ffd051" opacity="0.8"/>
    <rect x="80" y="230" width="640" height="120" rx="28" fill="#ffffff" opacity="0.72"/>
    <path d="M366 120c0-38 31-69 69-69s69 31 69 69" fill="none" stroke="#d08d00" stroke-width="16" stroke-linecap="round"/>
    <path d="M360 128h152c0 72-152 72-152 0Z" fill="#ffcc4d" stroke="#b36d00" stroke-width="12" stroke-linejoin="round"/>
    <circle cx="436" cy="200" r="18" fill="#945300"/>
    <text x="400" y="292" text-anchor="middle" font-size="34" font-family="Arial, sans-serif" font-weight="700" fill="#234">Question illustration</text>
  </svg>
`)}`;

const defaultState = {
  settings: {
    title: "Golden Bell Challenge",
    subtitle:
      "A fun review game that helps primary students strengthen their knowledge through exciting quiz questions.",
    rulesContent:
      "Each question has 4 answer choices: A, B, C, and D, and players choose one correct answer.\nAfter reading or watching the question, players answer within the hourglass countdown.\nWhen time is up, the program can automatically reveal the correct answer.\nThe host can award points for each question and review the final results on the ending slide.",
    defaultTimer: 30,
    useQuestionBoard: true,
    randomQuestionSelection: true,
    autoAdvance: false,
    autoRevealOnTimeout: true,
    endingTitle: "End of Program",
    endingMessage:
      "Thank you to all students for joining. Here is the final score summary and the journey through all questions.",
    pointsPerQuestion: 10,
    backgroundVolume: 35,
    effectVolume: 85,
    sounds: {
      background: "",
      reveal: "",
      timeout: "",
    },
  },
  questions: [
    {
      id: makeId(),
      type: "text",
      prompt: "What is the capital city of Vietnam?",
      media: "",
      timer: 20,
      answers: ["Hanoi", "Hue", "Da Nang", "Can Tho"],
      correctIndex: 0,
      selectedIndex: null,
      asked: false,
      awarded: false,
    },
    {
      id: makeId(),
      type: "image",
      prompt: "Look at the picture and identify the musical instrument.",
      media: sampleImage,
      timer: 25,
      answers: ["Zither", "Bronze drum", "Trumpet", "Bamboo flute"],
      correctIndex: 1,
      selectedIndex: null,
      asked: false,
      awarded: false,
    },
    {
      id: makeId(),
      type: "text",
      prompt: "What is 5 x 4?",
      media: "",
      timer: 15,
      answers: ["18", "20", "22", "24"],
      correctIndex: 1,
      selectedIndex: null,
      asked: false,
      awarded: false,
    },
  ],
};

const state = loadState();

const elements = {
  heroTitle: document.getElementById("hero-title"),
  heroSubtitle: document.getElementById("hero-subtitle"),
  modeButtons: Array.from(document.querySelectorAll(".mode-button")),
  viewPanels: Array.from(document.querySelectorAll(".view-panel")),
  questionList: document.getElementById("question-list"),
  questionCountPill: document.getElementById("question-count-pill"),
  questionListItemTemplate: document.getElementById("question-list-item-template"),
  answerCardTemplate: document.getElementById("answer-card-template"),
  programTitleInput: document.getElementById("program-title-input"),
  programSubtitleInput: document.getElementById("program-subtitle-input"),
  rulesInput: document.getElementById("rules-input"),
  defaultTimerInput: document.getElementById("default-timer-input"),
  questionBoardInput: document.getElementById("question-board-input"),
  randomQuestionInput: document.getElementById("random-question-input"),
  autoAdvanceInput: document.getElementById("auto-advance-input"),
  endingTitleInput: document.getElementById("ending-title-input"),
  endingMessageInput: document.getElementById("ending-message-input"),
  autoRevealAnswerInput: document.getElementById("auto-reveal-answer-input"),
  pointsPerQuestionInput: document.getElementById("points-per-question-input"),
  backgroundVolumeInput: document.getElementById("background-volume-input"),
  effectVolumeInput: document.getElementById("effect-volume-input"),
  backgroundSoundInput: document.getElementById("background-sound-input"),
  revealSoundInput: document.getElementById("reveal-sound-input"),
  timeoutSoundInput: document.getElementById("timeout-sound-input"),
  backgroundSoundFileInput: document.getElementById("background-sound-file-input"),
  revealSoundFileInput: document.getElementById("reveal-sound-file-input"),
  timeoutSoundFileInput: document.getElementById("timeout-sound-file-input"),
  questionTextInput: document.getElementById("question-text-input"),
  questionTypeInput: document.getElementById("question-type-input"),
  questionTimerInput: document.getElementById("question-timer-input"),
  questionMediaInput: document.getElementById("question-media-input"),
  questionFileInput: document.getElementById("question-file-input"),
  correctAnswerInput: document.getElementById("correct-answer-input"),
  answerInputs: [
    document.getElementById("answer-0-input"),
    document.getElementById("answer-1-input"),
    document.getElementById("answer-2-input"),
    document.getElementById("answer-3-input"),
  ],
  questionOrderBadge: document.getElementById("question-order-badge"),
  questionTypeBadge: document.getElementById("question-type-badge"),
  scoreBadge: document.getElementById("score-badge"),
  questionTag: document.getElementById("question-tag"),
  timerLabel: document.getElementById("timer-label"),
  timerDisplay: document.getElementById("timer-display"),
  timerWidget: document.getElementById("timer-widget"),
  introScreen: document.getElementById("intro-screen"),
  introProgramTitle: document.getElementById("intro-program-title"),
  introProgramSubtitle: document.getElementById("intro-program-subtitle"),
  rulesList: document.getElementById("rules-list"),
  boardScreen: document.getElementById("board-screen"),
  boardNote: document.getElementById("board-note"),
  questionBoardGrid: document.getElementById("question-board-grid"),
  randomQuestionButton: document.getElementById("random-question-button"),
  questionContent: document.querySelector(".question-content"),
  summaryScreen: document.getElementById("summary-screen"),
  summaryTitle: document.getElementById("summary-title"),
  summaryNote: document.getElementById("summary-note"),
  summaryTotalScore: document.getElementById("summary-total-score"),
  summaryPointsNote: document.getElementById("summary-points-note"),
  summaryCorrectCount: document.getElementById("summary-correct-count"),
  summaryTotalCount: document.getElementById("summary-total-count"),
  summaryScorePerQuestion: document.getElementById("summary-score-per-question"),
  summaryList: document.getElementById("summary-list"),
  slideQuestion: document.getElementById("slide-question"),
  mediaFrame: document.getElementById("media-frame"),
  questionImage: document.getElementById("question-image"),
  questionVideo: document.getElementById("question-video"),
  questionYoutube: document.getElementById("question-youtube"),
  answersGrid: document.getElementById("answers-grid"),
  answerFeedback: document.getElementById("answer-feedback"),
  newQuestionButton: document.getElementById("new-question-button"),
  duplicateQuestionButton: document.getElementById("duplicate-question-button"),
  moveUpButton: document.getElementById("move-up-button"),
  moveDownButton: document.getElementById("move-down-button"),
  deleteQuestionButton: document.getElementById("delete-question-button"),
  previousQuestionButton: document.getElementById("previous-question-button"),
  nextQuestionButton: document.getElementById("next-question-button"),
  showSummaryButton: document.getElementById("show-summary-button"),
  toggleAnswerButton: document.getElementById("toggle-answer-button"),
  timerToggleButton: document.getElementById("timer-toggle-button"),
  resetTimerButton: document.getElementById("reset-timer-button"),
  fullscreenButton: document.getElementById("fullscreen-button"),
  musicToggleButton: document.getElementById("music-toggle-button"),
  closeSummaryButton: document.getElementById("close-summary-button"),
  restartQuizButton: document.getElementById("restart-quiz-button"),
  importExcelButton: document.getElementById("import-excel-button"),
  excelFileInput: document.getElementById("excel-file-input"),
  backgroundAudio: document.getElementById("background-audio"),
  revealAudio: document.getElementById("reveal-audio"),
  timeoutAudio: document.getElementById("timeout-audio"),
};

let selectedQuestionId = state.questions[0] ? state.questions[0].id : null;
let currentSlideIndex = 0;
let revealAnswer = false;
let introVisible = true;
let boardVisible = false;
let summaryVisible = false;
let stageBeforeSummary = "intro";
let timerIntervalId = null;
let pendingAdvanceId = null;
let pendingRevealSoundId = null;
let backgroundMusicRequested = false;
let remainingSeconds = getCurrentDuration();

bootstrap();

function bootstrap() {
  bindEvents();
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function bindEvents() {
  elements.modeButtons.forEach((button) => {
    button.addEventListener("click", () => setActiveView(button.dataset.view));
  });

  elements.programTitleInput.addEventListener("input", (event) => {
    state.settings.title = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.programSubtitleInput.addEventListener("input", (event) => {
    state.settings.subtitle = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.rulesInput.addEventListener("input", (event) => {
    state.settings.rulesContent = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.defaultTimerInput.addEventListener("input", (event) => {
    state.settings.defaultTimer = clampNumber(
      event.target.value,
      5,
      300,
      defaultState.settings.defaultTimer,
    );
    syncRemainingSeconds();
    persistAndRender({ keepReveal: true, keepSummary: true });
  });

  elements.questionBoardInput.addEventListener("change", (event) => {
    state.settings.useQuestionBoard = event.target.checked;
    if (!state.settings.useQuestionBoard) {
      state.settings.randomQuestionSelection = false;
      if (boardVisible) {
        boardVisible = false;
        introVisible = false;
      }
    }
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.randomQuestionInput.addEventListener("change", (event) => {
    state.settings.randomQuestionSelection = event.target.checked;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.autoAdvanceInput.addEventListener("change", (event) => {
    state.settings.autoAdvance = event.target.checked;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.endingTitleInput.addEventListener("input", (event) => {
    state.settings.endingTitle = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.endingMessageInput.addEventListener("input", (event) => {
    state.settings.endingMessage = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.autoRevealAnswerInput.addEventListener("change", (event) => {
    state.settings.autoRevealOnTimeout = event.target.checked;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.pointsPerQuestionInput.addEventListener("input", (event) => {
    state.settings.pointsPerQuestion = clampNumber(event.target.value, 1, 100, 10);
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.backgroundVolumeInput.addEventListener("input", (event) => {
    state.settings.backgroundVolume = clampNumber(event.target.value, 0, 100, 35);
    syncAudioElements();
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.effectVolumeInput.addEventListener("input", (event) => {
    state.settings.effectVolume = clampNumber(event.target.value, 0, 100, 85);
    syncAudioElements();
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.backgroundSoundInput.addEventListener("input", (event) => {
    updateSoundPath("background", event.target.value);
  });

  elements.revealSoundInput.addEventListener("input", (event) => {
    updateSoundPath("reveal", event.target.value);
  });

  elements.timeoutSoundInput.addEventListener("input", (event) => {
    updateSoundPath("timeout", event.target.value);
  });

  elements.backgroundSoundFileInput.addEventListener("change", (event) => {
    handleSoundUpload("background", event);
  });

  elements.revealSoundFileInput.addEventListener("change", (event) => {
    handleSoundUpload("reveal", event);
  });

  elements.timeoutSoundFileInput.addEventListener("change", (event) => {
    handleSoundUpload("timeout", event);
  });

  elements.questionTextInput.addEventListener("input", (event) => {
    const question = getSelectedQuestion();
    if (!question) {
      return;
    }

    question.prompt = event.target.value;
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.questionTypeInput.addEventListener("change", (event) => {
    const question = getSelectedQuestion();
    if (!question) {
      return;
    }

    question.type = event.target.value;
    if (question.type === "text") {
      question.media = "";
    }

    persistAndRender({ preserveTimer: true, keepReveal: true });
  });

  elements.questionTimerInput.addEventListener("input", (event) => {
    const question = getSelectedQuestion();
    if (!question) {
      return;
    }

    question.timer = clampNumber(
      event.target.value,
      5,
      300,
      state.settings.defaultTimer,
    );

    if (question.id === getCurrentQuestionId()) {
      syncRemainingSeconds();
    }

    persistAndRender({ keepReveal: true, keepSummary: true });
  });

  elements.questionMediaInput.addEventListener("input", (event) => {
    const question = getSelectedQuestion();
    if (!question) {
      return;
    }

    question.media = event.target.value.trim();
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.questionFileInput.addEventListener("change", handleQuestionMediaUpload);

  elements.correctAnswerInput.addEventListener("change", (event) => {
    const question = getSelectedQuestion();
    if (!question) {
      return;
    }

    question.correctIndex = clampNumber(event.target.value, 0, 3, 0);
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.answerInputs.forEach((input, index) => {
    input.addEventListener("input", (event) => {
      const question = getSelectedQuestion();
      if (!question) {
        return;
      }

      question.answers[index] = event.target.value;
      persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
    });
  });

  elements.newQuestionButton.addEventListener("click", createQuestion);
  elements.duplicateQuestionButton.addEventListener("click", duplicateQuestion);
  elements.moveUpButton.addEventListener("click", () => moveQuestion(-1));
  elements.moveDownButton.addEventListener("click", () => moveQuestion(1));
  elements.deleteQuestionButton.addEventListener("click", deleteQuestion);

  elements.previousQuestionButton.addEventListener("click", handlePreviousAction);
  elements.nextQuestionButton.addEventListener("click", handleNextAction);
  elements.showSummaryButton.addEventListener("click", () => {
    if (summaryVisible) {
      closeSummary();
      return;
    }

    openSummary();
  });

  elements.toggleAnswerButton.addEventListener("click", () => {
    if (!getCurrentQuestion() || getCurrentStage() !== "question") {
      return;
    }

    revealAnswer = !revealAnswer;
    if (revealAnswer) {
      playRevealSound();
    }

    renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.timerToggleButton.addEventListener("click", toggleTimer);
  elements.resetTimerButton.addEventListener("click", () => {
    clearPendingStageActions();
    stopTimer();
    syncRemainingSeconds();
    renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
  });

  elements.fullscreenButton.addEventListener("click", toggleFullscreenStage);
  elements.musicToggleButton.addEventListener("click", toggleBackgroundMusic);
  elements.closeSummaryButton.addEventListener("click", closeSummary);
  elements.restartQuizButton.addEventListener("click", restartQuiz);
  elements.randomQuestionButton.addEventListener("click", openRandomQuestion);
  elements.importExcelButton.addEventListener("click", () => elements.excelFileInput.click());
  elements.excelFileInput.addEventListener("change", importExcelDeck);

  document.addEventListener("fullscreenchange", () => {
    document.body.classList.toggle("fullscreen-stage", Boolean(document.fullscreenElement));
  });

  document.addEventListener("keydown", (event) => {
    if (event.target.matches("input, textarea, select")) {
      return;
    }

    if (event.key === "ArrowRight") {
      handleNextAction();
    }

    if (event.key === "ArrowLeft") {
      handlePreviousAction();
    }

    if (event.code === "Space") {
      event.preventDefault();
      toggleTimer();
    }

    if (event.key.toLowerCase() === "r" && getCurrentStage() === "question") {
      revealAnswer = !revealAnswer;
      if (revealAnswer) {
        playRevealSound();
      }
      renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
    }

    if (event.key.toLowerCase() === "m") {
      toggleBackgroundMusic();
    }
  });
}

function renderEverything(options = {}) {
  syncAudioElements();
  hydrateSettingsForm();
  renderQuestionList();
  hydrateQuestionForm();
  renderPresentation(options);
}

function hydrateSettingsForm() {
  elements.heroTitle.textContent = state.settings.title || defaultState.settings.title;
  elements.heroSubtitle.textContent =
    state.settings.subtitle || defaultState.settings.subtitle;
  elements.programTitleInput.value = state.settings.title;
  elements.programSubtitleInput.value = state.settings.subtitle;
  elements.rulesInput.value = state.settings.rulesContent;
  elements.defaultTimerInput.value = state.settings.defaultTimer;
  elements.questionBoardInput.checked = state.settings.useQuestionBoard;
  elements.randomQuestionInput.checked = state.settings.randomQuestionSelection;
  elements.randomQuestionInput.disabled = !state.settings.useQuestionBoard;
  elements.autoAdvanceInput.checked = state.settings.autoAdvance;
  elements.endingTitleInput.value = state.settings.endingTitle;
  elements.endingMessageInput.value = state.settings.endingMessage;
  elements.autoRevealAnswerInput.checked = state.settings.autoRevealOnTimeout;
  elements.pointsPerQuestionInput.value = state.settings.pointsPerQuestion;
  elements.backgroundVolumeInput.value = state.settings.backgroundVolume;
  elements.effectVolumeInput.value = state.settings.effectVolume;
  elements.backgroundSoundInput.value = state.settings.sounds.background;
  elements.revealSoundInput.value = state.settings.sounds.reveal;
  elements.timeoutSoundInput.value = state.settings.sounds.timeout;
}

function renderQuestionList() {
  elements.questionList.innerHTML = "";
  elements.questionCountPill.textContent = `${state.questions.length} questions`;

  if (!state.questions.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = "No questions yet. Click Add Question to get started.";
    elements.questionList.appendChild(emptyState);
    updateEditorButtons();
    return;
  }

  state.questions.forEach((question, index) => {
    const fragment = elements.questionListItemTemplate.content.cloneNode(true);
    const item = fragment.querySelector(".question-list-item");
    const order = fragment.querySelector(".question-index");
    const title = fragment.querySelector(".question-list-title");
    const meta = fragment.querySelector(".question-list-meta");
    const metaParts = [labelQuestionType(question.type), `${question.timer}s`];

    if (question.awarded) {
      metaParts.push(`+${state.settings.pointsPerQuestion} points`);
    }

    if (question.asked) {
      metaParts.push("Opened");
    }

    order.textContent = String(index + 1).padStart(2, "0");
    title.textContent = question.prompt || `Question ${index + 1}`;
    meta.textContent = metaParts.join(" - ");
    item.classList.toggle("active", question.id === selectedQuestionId);

    item.addEventListener("click", () => {
      clearPendingStageActions();
      stopTimer();
      selectedQuestionId = question.id;
      currentSlideIndex = index;
      revealAnswer = false;
      summaryVisible = false;
      syncRemainingSeconds();
      renderEverything();
      setActiveView("editor-view");
    });

    elements.questionList.appendChild(fragment);
  });

  updateEditorButtons();
}

function hydrateQuestionForm() {
  const question = getSelectedQuestion();

  elements.questionFileInput.value = "";

  if (!question) {
    elements.questionTextInput.value = "";
    elements.questionTypeInput.value = "text";
    elements.questionTimerInput.value = state.settings.defaultTimer;
    elements.questionMediaInput.value = "";
    elements.correctAnswerInput.value = "0";
    elements.answerInputs.forEach((input) => {
      input.value = "";
    });
    setMediaInputsEnabled(false);
    updateEditorButtons();
    return;
  }

  elements.questionTextInput.value = question.prompt;
  elements.questionTypeInput.value = question.type;
  elements.questionTimerInput.value = question.timer;
  elements.questionMediaInput.value = question.media;
  elements.correctAnswerInput.value = String(question.correctIndex);

  elements.answerInputs.forEach((input, index) => {
    input.value = question.answers[index] || "";
  });

  setMediaInputsEnabled(question.type !== "text");
  updateEditorButtons();
}

function renderPresentation(options = {}) {
  const question = getCurrentQuestion();
  const scoreStats = getScoreStats();
  const totalQuestions = state.questions.length;

  if (!options.preserveTimer) {
    syncRemainingSeconds();
  }

  if (!options.keepReveal) {
    revealAnswer = false;
  }

  if (!options.keepSummary) {
    summaryVisible = false;
  }

  const stage = getCurrentStage();

  elements.scoreBadge.textContent = `Diem ${scoreStats.totalScore}`;
  updateMusicButton();

  elements.introScreen.classList.add("hidden");
  elements.boardScreen.classList.add("hidden");
  elements.questionContent.classList.add("hidden");
  elements.summaryScreen.classList.add("hidden");

  if (stage === "summary") {
    cleanupVideoPlayers();
    elements.summaryScreen.classList.remove("hidden");
    elements.questionOrderBadge.textContent = "Ending Slide";
    elements.questionTypeBadge.textContent = "Summary";
    elements.timerLabel.textContent = "Completed";
    elements.timerDisplay.textContent = `${scoreStats.correctCount}/${totalQuestions}`;
    updateTimerArtwork(totalQuestions ? scoreStats.correctCount / totalQuestions : 1);
    renderSummary();
    updateStageButtons();
    return;
  }

  if (stage === "intro") {
    cleanupVideoPlayers();
    elements.introScreen.classList.remove("hidden");
    elements.questionOrderBadge.textContent = "Opening Slide";
    elements.questionTypeBadge.textContent = "Rules";
    elements.timerLabel.textContent = "Ready";
    elements.timerDisplay.textContent = `${totalQuestions} questions`;
    updateTimerArtwork(1);
    renderIntroScreen();
    updateStageButtons();
    return;
  }

  if (stage === "board") {
    cleanupVideoPlayers();
    const unaskedCount = getUnaskedQuestionIndices().length;
    elements.boardScreen.classList.remove("hidden");
    elements.questionOrderBadge.textContent = "Question Board";
    elements.questionTypeBadge.textContent = "Pick a Question";
    elements.timerLabel.textContent = "Remaining";
    elements.timerDisplay.textContent = `${unaskedCount}/${totalQuestions}`;
    updateTimerArtwork(totalQuestions ? unaskedCount / totalQuestions : 1);
    renderQuestionBoard();
    updateStageButtons();
    return;
  }

  elements.questionContent.classList.remove("hidden");
  elements.timerLabel.textContent = "Dong ho cat";

  if (!question) {
    cleanupVideoPlayers();
    elements.questionOrderBadge.textContent = "No Questions Yet";
    elements.questionTypeBadge.textContent = "Waiting for Data";
    elements.questionTag.textContent = "Create a Question";
    elements.slideQuestion.textContent = "Add a question in the builder panel to begin.";
    elements.toggleAnswerButton.textContent = "Show Answer";
    elements.answersGrid.innerHTML = "";
    elements.answerFeedback.className = "answer-feedback hidden";
    elements.answerFeedback.textContent = "";
    elements.mediaFrame.classList.add("hidden");
    elements.timerDisplay.textContent = `${state.settings.defaultTimer}s`;
    updateTimerArtwork(1);
    updateStageButtons();
    return;
  }

  elements.questionOrderBadge.textContent = `Question ${currentSlideIndex + 1} / ${totalQuestions}`;
  elements.questionTypeBadge.textContent = labelQuestionType(question.type);
  elements.questionTag.textContent = labelQuestionType(question.type);
  elements.slideQuestion.textContent =
    question.prompt || "Enter the question content in the builder panel.";
  elements.toggleAnswerButton.textContent = revealAnswer ? "Hide Answer" : "Show Answer";
  elements.timerToggleButton.textContent = timerIntervalId ? "Pause" : "Start Timer";

  renderMedia(question);
  renderAnswers(question);
  renderAnswerFeedback(question);
  updateTimerDisplay();
  updateStageButtons();
}

function renderIntroScreen() {
  elements.introProgramTitle.textContent =
    state.settings.title || defaultState.settings.title;
  elements.introProgramSubtitle.textContent =
    state.settings.subtitle || defaultState.settings.subtitle;
  renderRules();
}

function renderRules() {
  const lines = String(state.settings.rulesContent || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/^[-*•\s\d.()]+/, "").trim())
    .filter(Boolean);

  elements.rulesList.innerHTML = "";

  if (!lines.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = "Add game rules in the builder panel to display them on the opening slide.";
    elements.rulesList.appendChild(emptyState);
    return;
  }

  lines.forEach((line, index) => {
    const item = document.createElement("article");
    const order = document.createElement("span");
    const copy = document.createElement("p");

    item.className = "rule-item";
    order.className = "rule-item-index";
    order.textContent = String(index + 1).padStart(2, "0");
    copy.textContent = line;

    item.appendChild(order);
    item.appendChild(copy);
    elements.rulesList.appendChild(item);
  });
}

function renderQuestionBoard() {
  const totalQuestions = state.questions.length;
  const unaskedCount = getUnaskedQuestionIndices().length;

  elements.boardNote.textContent = state.settings.randomQuestionSelection
    ? "Click a number tile to open that exact question, or use the random button to choose one that has not been opened yet."
    : "Click a number tile to open that exact question. Each tile matches one question in the program.";

  elements.randomQuestionButton.classList.toggle(
    "hidden",
    !state.settings.useQuestionBoard || !state.settings.randomQuestionSelection,
  );
  elements.randomQuestionButton.disabled = !unaskedCount;

  elements.questionBoardGrid.innerHTML = "";

  if (!totalQuestions) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = "There are no questions on the board yet. Add questions in the builder panel first.";
    elements.questionBoardGrid.appendChild(emptyState);
    return;
  }

  state.questions.forEach((questionItem, index) => {
    const tile = document.createElement("button");
    const number = document.createElement("span");
    const status = document.createElement("span");

    tile.type = "button";
    tile.className = "question-board-tile";
    tile.style.setProperty("--tile-hue", String((index * 41 + 18) % 360));
    tile.classList.toggle("used", questionItem.asked);
    tile.classList.toggle("is-current", questionItem.id === selectedQuestionId);
    tile.addEventListener("click", () => {
      openQuestion(index, { markAsked: true });
    });

    number.className = "question-board-tile-number";
    number.textContent = String(index + 1);

    status.className = "question-board-tile-state";
    status.textContent = questionItem.asked ? "Opened" : "Ready";

    tile.appendChild(number);
    tile.appendChild(status);
    elements.questionBoardGrid.appendChild(tile);
  });
}

function renderMedia(question) {
  const showImage = question.type === "image" && Boolean(question.media);
  const youtubeEmbedUrl =
    question.type === "video" ? getYouTubeEmbedUrl(question.media) : "";
  const showYoutube = Boolean(youtubeEmbedUrl);
  const showVideo = question.type === "video" && Boolean(question.media) && !showYoutube;

  elements.mediaFrame.classList.toggle("hidden", !showImage && !showVideo && !showYoutube);
  elements.mediaFrame.classList.toggle("is-image", showImage);
  elements.mediaFrame.classList.toggle("is-video", showVideo || showYoutube);
  elements.questionImage.classList.toggle("hidden", !showImage);
  elements.questionVideo.classList.toggle("hidden", !showVideo);
  elements.questionYoutube.classList.toggle("hidden", !showYoutube);

  if (showImage) {
    elements.questionImage.src = question.media;
    cleanupVideoPlayers();
  }

  if (showVideo) {
    elements.questionYoutube.removeAttribute("src");
    elements.questionVideo.src = question.media;
    elements.questionImage.removeAttribute("src");
  }

  if (showYoutube) {
    elements.questionImage.removeAttribute("src");
    elements.questionVideo.pause();
    elements.questionVideo.removeAttribute("src");
    elements.questionVideo.load();
    elements.questionYoutube.src = youtubeEmbedUrl;
  }

  if (!showImage && !showVideo && !showYoutube) {
    elements.questionImage.removeAttribute("src");
    cleanupVideoPlayers();
  }
}

function cleanupVideoPlayers() {
  elements.questionVideo.pause();
  elements.questionVideo.removeAttribute("src");
  elements.questionVideo.load();
  elements.questionYoutube.removeAttribute("src");
}

function getYouTubeEmbedUrl(rawValue) {
  if (!rawValue) {
    return "";
  }

  const normalizedValue = /^(https?:)?\/\//i.test(rawValue)
    ? rawValue
    : `https://${rawValue}`;
  let parsedUrl;

  try {
    parsedUrl = new URL(normalizedValue);
  } catch (error) {
    return "";
  }

  const hostname = parsedUrl.hostname.replace(/^www\./i, "").toLowerCase();
  let videoId = "";

  if (hostname === "youtu.be") {
    videoId = parsedUrl.pathname.split("/").filter(Boolean)[0] || "";
  }

  if (hostname === "youtube.com" || hostname === "m.youtube.com") {
    if (parsedUrl.pathname === "/watch") {
      videoId = parsedUrl.searchParams.get("v") || "";
    } else if (parsedUrl.pathname.startsWith("/shorts/")) {
      videoId = parsedUrl.pathname.split("/").filter(Boolean)[1] || "";
    } else if (parsedUrl.pathname.startsWith("/embed/")) {
      videoId = parsedUrl.pathname.split("/").filter(Boolean)[1] || "";
    }
  }

  if (hostname === "youtube-nocookie.com") {
    if (parsedUrl.pathname.startsWith("/embed/")) {
      videoId = parsedUrl.pathname.split("/").filter(Boolean)[1] || "";
    }
  }

  if (!/^[a-zA-Z0-9_-]{6,}$/.test(videoId)) {
    return "";
  }

  return `https://www.youtube.com/embed/${videoId}?rel=0`;
}

function renderAnswers(question) {
  elements.answersGrid.innerHTML = "";
  const answered = Number.isInteger(question.selectedIndex);

  question.answers.forEach((answer, index) => {
    const fragment = elements.answerCardTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".answer-card");
    const letter = fragment.querySelector(".answer-letter");
    const text = fragment.querySelector(".answer-text");

    letter.textContent = answerLabel(index);
    text.textContent = answer || `Answer ${answerLabel(index)}`;
    card.disabled = answered;

    if (question.selectedIndex === index) {
      card.classList.add("selected");
      if (index !== question.correctIndex) {
        card.classList.add("incorrect");
      }
    }

    if (revealAnswer) {
      card.classList.add("revealed");
      if (index === question.correctIndex) {
        card.classList.add("correct");
      }
    }

    if (!answered && getCurrentStage() === "question") {
      card.addEventListener("click", () => {
        handleAnswerSelection(index);
      });
    }

    elements.answersGrid.appendChild(fragment);
  });
}

function renderAnswerFeedback(question) {
  if (!Number.isInteger(question.selectedIndex)) {
    elements.answerFeedback.className = "answer-feedback hidden";
    elements.answerFeedback.textContent = "";
    return;
  }

  const isCorrect = question.selectedIndex === question.correctIndex;
  elements.answerFeedback.className = `answer-feedback ${isCorrect ? "success" : "error"}`;
  elements.answerFeedback.textContent = isCorrect
    ? `Correct! ${state.settings.pointsPerQuestion} points were added automatically.`
    : "Incorrect answer. The correct option is now highlighted.";
}

function renderSummary() {
  const scoreStats = getScoreStats();
  const totalQuestions = state.questions.length;

  elements.summaryTitle.textContent =
    state.settings.endingTitle || defaultState.settings.endingTitle;
  elements.summaryTotalScore.textContent = String(scoreStats.totalScore);
  elements.summaryPointsNote.textContent =
    `${scoreStats.correctCount} scored questions`;
  elements.summaryCorrectCount.textContent = String(scoreStats.correctCount);
  elements.summaryTotalCount.textContent = `Out of ${totalQuestions} questions`;
  elements.summaryScorePerQuestion.textContent = String(state.settings.pointsPerQuestion);
  elements.summaryNote.textContent =
    state.settings.endingMessage || defaultState.settings.endingMessage;
  elements.closeSummaryButton.textContent =
    stageBeforeSummary === "board"
      ? "Back to Board"
      : stageBeforeSummary === "intro"
        ? "Back to Rules"
        : "Back to Slide";

  elements.summaryList.innerHTML = "";

  if (!totalQuestions) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = "No summary data is available yet.";
    elements.summaryList.appendChild(emptyState);
    return;
  }

  state.questions.forEach((question, index) => {
    const item = document.createElement("article");
    item.className = "summary-item";

    const copy = document.createElement("div");
    const title = document.createElement("p");
    const meta = document.createElement("small");
    const status = document.createElement("span");
    const correctAnswer =
      question.answers[question.correctIndex] || `Answer ${answerLabel(question.correctIndex)}`;

    title.textContent = `Question ${index + 1}: ${question.prompt || "Question content is empty"}`;
    meta.textContent =
      `${question.asked ? "Opened" : "Not opened"} - Correct answer: ${answerLabel(question.correctIndex)} - ${correctAnswer}`;
    status.className = `summary-status ${question.awarded ? "good" : "empty"}`;
    status.textContent = question.awarded
      ? `+${state.settings.pointsPerQuestion} points`
      : "0 points";

    copy.appendChild(title);
    copy.appendChild(meta);
    item.appendChild(copy);
    item.appendChild(status);
    elements.summaryList.appendChild(item);
  });
}

function updateTimerDisplay() {
  const duration = Math.max(1, getCurrentDuration());
  elements.timerDisplay.textContent = `${remainingSeconds}s`;
  updateTimerArtwork(remainingSeconds / duration);
}

function updateTimerArtwork(progress) {
  const safeProgress = Number.isFinite(progress)
    ? Math.max(0, Math.min(1, progress))
    : 1;
  elements.timerWidget.style.setProperty("--timer-progress", safeProgress.toFixed(3));
}

function updateEditorButtons() {
  const index = getSelectedQuestionIndex();
  const hasQuestion = index !== -1;

  elements.duplicateQuestionButton.disabled = !hasQuestion;
  elements.deleteQuestionButton.disabled = !hasQuestion;
  elements.moveUpButton.disabled = !hasQuestion || index === 0;
  elements.moveDownButton.disabled =
    !hasQuestion || index === state.questions.length - 1;
}

function updateStageButtons() {
  const question = getCurrentQuestion();
  const hasQuestion = Boolean(question);
  const stage = getCurrentStage();
  const isQuestionStage = stage === "question";
  const hasUnaskedQuestions = getUnaskedQuestionIndices().length > 0;

  elements.previousQuestionButton.disabled =
    stage === "intro" || (stage === "question" && !hasQuestion);
  elements.nextQuestionButton.disabled =
    stage === "summary" ||
    (stage === "question" && !hasQuestion) ||
    ((stage === "intro" || stage === "board") && !state.questions.length);
  elements.nextQuestionButton.textContent =
    stage === "intro"
      ? "Start"
      : stage === "board"
        ? hasUnaskedQuestions
          ? state.settings.randomQuestionSelection
            ? "Random Question"
            : "Open Next"
          : "Ending Slide"
        : state.settings.useQuestionBoard
          ? hasUnaskedQuestions
            ? "Back to Board"
            : "Ending Slide"
          : currentSlideIndex >= state.questions.length - 1
            ? "Ending Slide"
            : "Next";
  elements.showSummaryButton.disabled = false;
  elements.showSummaryButton.textContent = summaryVisible ? "Close Ending Slide" : "Ending Slide";
  elements.resetTimerButton.disabled = !isQuestionStage || !hasQuestion;
  elements.timerToggleButton.disabled = !isQuestionStage || !hasQuestion;
  elements.toggleAnswerButton.disabled = !isQuestionStage || !hasQuestion;
}

function updateMusicButton() {
  const hasMusic = Boolean(state.settings.sounds.background);
  const isPlaying = backgroundMusicRequested && !elements.backgroundAudio.paused;

  elements.musicToggleButton.disabled = !hasMusic;
  elements.musicToggleButton.textContent = !hasMusic
    ? "No Background Music"
    : isPlaying
      ? "Stop Music"
      : "Play Music";
}

function setMediaInputsEnabled(enabled) {
  elements.questionMediaInput.disabled = !enabled;
  elements.questionFileInput.disabled = !enabled;
}

function createQuestion() {
  clearPendingStageActions();
  stopTimer();
  introVisible = false;
  boardVisible = false;
  summaryVisible = false;

  const question = createBlankQuestion();
  state.questions.push(question);
  selectedQuestionId = question.id;
  currentSlideIndex = state.questions.length - 1;
  revealAnswer = false;
  syncRemainingSeconds();
  persistAndRender({ keepSummary: true });
  setActiveView("editor-view");
}

function duplicateQuestion() {
  const question = getSelectedQuestion();
  if (!question) {
    return;
  }

  clearPendingStageActions();
  stopTimer();
  introVisible = false;
  boardVisible = false;
  summaryVisible = false;

  const duplicate = deepClone(question);
  duplicate.id = makeId();
  duplicate.selectedIndex = null;
  duplicate.asked = false;
  duplicate.awarded = false;
  duplicate.prompt = question.prompt ? `${question.prompt} (copy)` : "Question copy";

  const index = getSelectedQuestionIndex();
  state.questions.splice(index + 1, 0, duplicate);
  selectedQuestionId = duplicate.id;
  currentSlideIndex = index + 1;
  revealAnswer = false;
  syncRemainingSeconds();
  persistAndRender({ keepSummary: true });
}

function moveQuestion(direction) {
  const index = getSelectedQuestionIndex();
  const targetIndex = index + direction;

  if (index === -1 || targetIndex < 0 || targetIndex >= state.questions.length) {
    return;
  }

  clearPendingStageActions();
  introVisible = false;
  boardVisible = false;
  summaryVisible = false;

  const [question] = state.questions.splice(index, 1);
  state.questions.splice(targetIndex, 0, question);
  selectedQuestionId = question.id;
  currentSlideIndex = targetIndex;
  persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function deleteQuestion() {
  const index = getSelectedQuestionIndex();
  if (index === -1) {
    return;
  }

  const question = state.questions[index];
  const shouldDelete = window.confirm(
    `Are you sure you want to delete this question?\n\n${question.prompt || "Untitled question"}`,
  );

  if (!shouldDelete) {
    return;
  }

  clearPendingStageActions();
  stopTimer();
  state.questions.splice(index, 1);

  if (state.questions.length) {
    const fallbackIndex = Math.max(0, index - 1);
    currentSlideIndex = Math.min(fallbackIndex, state.questions.length - 1);
    selectedQuestionId = state.questions[currentSlideIndex].id;
  } else {
    selectedQuestionId = null;
    currentSlideIndex = 0;
  }

  revealAnswer = false;
  introVisible = state.questions.length === 0;
  boardVisible = false;
  summaryVisible = false;
  syncRemainingSeconds();
  persistAndRender();
}

function handleNextAction() {
  const stage = getCurrentStage();

  if (stage === "summary") {
    return;
  }

  if (stage === "intro") {
    if (!state.questions.length) {
      return;
    }

    if (state.settings.useQuestionBoard) {
      openBoard();
      return;
    }

    openQuestion(0, { markAsked: true });
    return;
  }

  if (stage === "board") {
    if (!state.questions.length) {
      return;
    }

    if (state.settings.randomQuestionSelection) {
      openRandomQuestion();
      return;
    }

    const targetIndex = getDefaultBoardQuestionIndex();
    if (targetIndex === -1) {
      openSummary();
      return;
    }

    openQuestion(targetIndex, { markAsked: true });
    return;
  }

  if (!state.questions.length) {
    return;
  }

  if (state.settings.useQuestionBoard) {
    if (getUnaskedQuestionIndices().length) {
      openBoard();
      return;
    }

    openSummary();
    return;
  }

  if (currentSlideIndex >= state.questions.length - 1) {
    openSummary();
    return;
  }

  shiftSlide(1);
}

function shiftSlide(direction) {
  const nextIndex = currentSlideIndex + direction;

  if (summaryVisible || introVisible || boardVisible) {
    return;
  }

  if (nextIndex < 0 || nextIndex >= state.questions.length) {
    return;
  }

  openQuestion(nextIndex, { markAsked: direction > 0 });
}

function toggleTimer() {
  if (getCurrentStage() !== "question" || !getCurrentQuestion()) {
    return;
  }

  if (timerIntervalId) {
    stopTimer();
    renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
    return;
  }

  clearPendingStageActions();

  if (remainingSeconds <= 0) {
    syncRemainingSeconds();
  }

  timerIntervalId = window.setInterval(() => {
    remainingSeconds = Math.max(0, remainingSeconds - 1);
    updateTimerDisplay();

    if (remainingSeconds === 0) {
      stopTimer();
      handleTimeExpired();
    }
  }, 1000);

  renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function stopTimer() {
  if (timerIntervalId) {
    window.clearInterval(timerIntervalId);
    timerIntervalId = null;
  }
}

function handleTimeExpired() {
  const shouldReveal = state.settings.autoRevealOnTimeout && !revealAnswer;

  playEffect(elements.timeoutAudio);

  if (shouldReveal) {
    revealAnswer = true;
    playRevealSound(REVEAL_SOUND_DELAY_MS);
  }

  renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });

  if (!state.settings.autoAdvance) {
    return;
  }

  scheduleAutomaticStageAdvance();
}

function handleAnswerSelection(answerIndex) {
  const question = getCurrentQuestion();

  if (
    !question ||
    getCurrentStage() !== "question" ||
    Number.isInteger(question.selectedIndex)
  ) {
    return;
  }

  clearPendingStageActions();
  stopTimer();

  question.asked = true;
  question.selectedIndex = answerIndex;
  question.awarded = answerIndex === question.correctIndex;
  revealAnswer = true;

  playRevealSound(REVEAL_SOUND_DELAY_MS);
  persistStateQuietly();
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });

  if (state.settings.autoAdvance) {
    scheduleAutomaticStageAdvance();
  }
}

function scheduleAutomaticStageAdvance() {
  scheduleStageAdvance(() => {
    if (state.settings.useQuestionBoard) {
      if (getUnaskedQuestionIndices().length) {
        openBoard();
      } else {
        openSummary();
      }
      return;
    }

    if (currentSlideIndex >= state.questions.length - 1) {
      openSummary();
    } else {
      shiftSlide(1);
    }
  }, AUTO_ADVANCE_DELAY_MS);
}

function syncRemainingSeconds() {
  remainingSeconds = getCurrentDuration();
}

function openSummary() {
  const currentStage = getCurrentStage();
  clearPendingStageActions();
  stopTimer();
  if (currentStage !== "summary") {
    stageBeforeSummary = currentStage;
  }
  introVisible = false;
  boardVisible = false;
  summaryVisible = true;
  setActiveView("presentation-view");
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function closeSummary() {
  summaryVisible = false;

  if (stageBeforeSummary === "intro") {
    introVisible = true;
    boardVisible = false;
  } else if (stageBeforeSummary === "board" && state.settings.useQuestionBoard) {
    introVisible = false;
    boardVisible = true;
  } else {
    introVisible = false;
    boardVisible = false;
  }

  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: false });
}

function openIntro() {
  clearPendingStageActions();
  stopTimer();
  revealAnswer = false;
  introVisible = true;
  boardVisible = false;
  summaryVisible = false;
  setActiveView("presentation-view");
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function openBoard() {
  if (!state.settings.useQuestionBoard) {
    if (state.questions.length) {
      const fallbackIndex = getDefaultBoardQuestionIndex();
      if (fallbackIndex === -1) {
        openSummary();
      } else {
        openQuestion(fallbackIndex, { markAsked: false });
      }
    } else {
      openIntro();
    }
    return;
  }

  clearPendingStageActions();
  stopTimer();
  revealAnswer = false;
  introVisible = false;
  boardVisible = true;
  summaryVisible = false;
  setActiveView("presentation-view");
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function openQuestion(index, options = {}) {
  if (index < 0 || index >= state.questions.length) {
    return;
  }

  const { markAsked = false } = options;

  clearPendingStageActions();
  stopTimer();
  currentSlideIndex = index;
  selectedQuestionId = state.questions[index].id;
  if (markAsked) {
    state.questions[index].asked = true;
    persistStateQuietly();
  }
  revealAnswer = false;
  introVisible = false;
  boardVisible = false;
  summaryVisible = false;
  syncRemainingSeconds();
  setActiveView("presentation-view");
  renderEverything({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function openRandomQuestion() {
  const candidates = getUnaskedQuestionIndices();

  if (!candidates.length) {
    openSummary();
    return;
  }

  const randomIndex = candidates[Math.floor(Math.random() * candidates.length)];
  openQuestion(randomIndex, { markAsked: true });
}

function handlePreviousAction() {
  const stage = getCurrentStage();

  if (stage === "summary") {
    closeSummary();
    return;
  }

  if (stage === "intro") {
    return;
  }

  if (stage === "board") {
    openIntro();
    return;
  }

  if (state.settings.useQuestionBoard) {
    openBoard();
    return;
  }

  if (currentSlideIndex <= 0) {
    openIntro();
    return;
  }

  shiftSlide(-1);
}

function restartQuiz() {
  clearPendingStageActions();
  stopTimer();
  state.questions.forEach((question) => {
    question.selectedIndex = null;
    question.asked = false;
    question.awarded = false;
  });
  currentSlideIndex = 0;
  selectedQuestionId = state.questions[0] ? state.questions[0].id : null;
  revealAnswer = false;
  introVisible = true;
  boardVisible = false;
  summaryVisible = false;
  stageBeforeSummary = "intro";
  syncRemainingSeconds();
  persistAndRender();
  setActiveView("presentation-view");
}

function toggleFullscreenStage() {
  setActiveView("presentation-view");

  if (document.fullscreenElement) {
    document.exitFullscreen();
    return;
  }

  if (document.documentElement.requestFullscreen) {
    document.documentElement.requestFullscreen().catch(() => {
      document.body.classList.toggle("fullscreen-stage");
    });
    return;
  }

  document.body.classList.toggle("fullscreen-stage");
}

function toggleBackgroundMusic() {
  if (!state.settings.sounds.background) {
    window.alert("Please set background music in the Upgrade Panel first.");
    return;
  }

  if (backgroundMusicRequested && !elements.backgroundAudio.paused) {
    backgroundMusicRequested = false;
    elements.backgroundAudio.pause();
    renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
    return;
  }

  backgroundMusicRequested = true;
  const playPromise = elements.backgroundAudio.play();

  if (playPromise && typeof playPromise.catch === "function") {
    playPromise.catch(() => {
      backgroundMusicRequested = false;
      updateMusicButton();
      window.alert("Your browser requires a user interaction before audio can play.");
    });
  }

  renderPresentation({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function updateSoundPath(kind, value) {
  state.settings.sounds[kind] = value.trim();

  if (kind === "background" && !state.settings.sounds.background) {
    backgroundMusicRequested = false;
  }

  persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
}

function handleSoundUpload(kind, event) {
  const file = event.target.files ? event.target.files[0] : null;

  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    state.settings.sounds[kind] = String(reader.result || "");
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  };
  reader.onerror = () => {
    window.alert("Unable to read the audio file. Please try a different file.");
  };
  reader.readAsDataURL(file);
}

function handleQuestionMediaUpload(event) {
  const question = getSelectedQuestion();
  const file = event.target.files ? event.target.files[0] : null;

  if (!question || !file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    question.media = String(reader.result || "");
    if (file.type.startsWith("image/")) {
      question.type = "image";
    }
    if (file.type.startsWith("video/")) {
      question.type = "video";
    }
    persistAndRender({ preserveTimer: true, keepReveal: true, keepSummary: true });
  };
  reader.onerror = () => {
    window.alert("Unable to read the media file. Please try a different file.");
  };

  reader.readAsDataURL(file);
}

function scheduleStageAdvance(callback, delay) {
  clearPendingAdvance();
  pendingAdvanceId = window.setTimeout(() => {
    pendingAdvanceId = null;
    callback();
  }, delay);
}

function clearPendingAdvance() {
  if (pendingAdvanceId) {
    window.clearTimeout(pendingAdvanceId);
    pendingAdvanceId = null;
  }
}

function clearPendingRevealSound() {
  if (pendingRevealSoundId) {
    window.clearTimeout(pendingRevealSoundId);
    pendingRevealSoundId = null;
  }
}

function clearPendingStageActions() {
  clearPendingAdvance();
  clearPendingRevealSound();
}

function playRevealSound(delayMs = 0) {
  clearPendingRevealSound();

  if (!state.settings.sounds.reveal) {
    return;
  }

  if (delayMs > 0) {
    pendingRevealSoundId = window.setTimeout(() => {
      pendingRevealSoundId = null;
      playEffect(elements.revealAudio);
    }, delayMs);
    return;
  }

  playEffect(elements.revealAudio);
}

function playEffect(audioElement) {
  const source = audioElement.dataset.sourceValue || "";
  if (!source) {
    return;
  }

  try {
    audioElement.pause();
    audioElement.currentTime = 0;
    const playPromise = audioElement.play();
    if (playPromise && typeof playPromise.catch === "function") {
      playPromise.catch(() => {});
    }
  } catch (error) {
    // Ignore blocked audio playback.
  }
}

function syncAudioElements() {
  const backgroundChanged = setAudioSource(
    elements.backgroundAudio,
    state.settings.sounds.background,
  );
  setAudioSource(elements.revealAudio, state.settings.sounds.reveal);
  setAudioSource(elements.timeoutAudio, state.settings.sounds.timeout);

  elements.backgroundAudio.volume = clampNumber(
    state.settings.backgroundVolume,
    0,
    100,
    35,
  ) / 100;
  elements.revealAudio.volume = clampNumber(state.settings.effectVolume, 0, 100, 85) / 100;
  elements.timeoutAudio.volume = clampNumber(state.settings.effectVolume, 0, 100, 85) / 100;

  if (!state.settings.sounds.background) {
    backgroundMusicRequested = false;
    elements.backgroundAudio.pause();
  } else if (backgroundMusicRequested && (backgroundChanged || elements.backgroundAudio.paused)) {
    const playPromise = elements.backgroundAudio.play();
    if (playPromise && typeof playPromise.catch === "function") {
      playPromise.catch(() => {});
    }
  }
}

function setAudioSource(audioElement, source) {
  const safeSource = source || "";

  if (audioElement.dataset.sourceValue === safeSource) {
    return false;
  }

  audioElement.dataset.sourceValue = safeSource;
  audioElement.pause();
  audioElement.removeAttribute("src");

  if (safeSource) {
    audioElement.src = safeSource;
  }

  audioElement.load();
  return true;
}

async function importExcelDeck(event) {
  const file = event.target.files ? event.target.files[0] : null;
  if (!file) {
    return;
  }

  try {
    const rows = await readExcelRows(file);
    const importedQuestions = mapExcelRowsToQuestions(rows);

    state.questions = importedQuestions;
    clearPendingStageActions();
    stopTimer();
    selectedQuestionId = state.questions[0] ? state.questions[0].id : null;
    currentSlideIndex = 0;
    revealAnswer = false;
    introVisible = true;
    boardVisible = false;
    summaryVisible = false;
    stageBeforeSummary = "intro";
    syncRemainingSeconds();
    persistAndRender();
  } catch (error) {
    window.alert(
      error instanceof Error && error.message
        ? error.message
        : "Unable to read the Excel file. Please use the .xlsx template to import questions.",
    );
  } finally {
    elements.excelFileInput.value = "";
  }
}

async function readExcelRows(file) {
  if (!window.DecompressionStream) {
    throw new Error("This browser does not support reading .xlsx Excel files. Please use a newer version of Edge or Chrome.");
  }

  const arrayBuffer = await file.arrayBuffer();
  const zipEntries = await unzipXlsxEntries(arrayBuffer);
  const worksheetXml = resolveWorksheetXml(zipEntries);

  if (!worksheetXml) {
    throw new Error("The question sheet could not be found in the Excel file.");
  }

  const sharedStrings = parseSharedStringsXml(zipEntries.get("xl/sharedStrings.xml") || "");
  return parseWorksheetRows(worksheetXml, sharedStrings);
}

async function unzipXlsxEntries(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const eocdOffset = findEndOfCentralDirectory(view);

  if (eocdOffset === -1) {
    throw new Error("The selected file is not a valid Excel .xlsx file.");
  }

  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const decoder = new TextDecoder();
  const entries = new Map();
  let offset = centralDirectoryOffset;

  for (let index = 0; index < totalEntries; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      break;
    }

    const compressionMethod = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const fileName = decoder.decode(
      bytes.slice(offset + 46, offset + 46 + fileNameLength),
    );

    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    const compressedBytes = bytes.slice(dataStart, dataEnd);

    let outputBytes = compressedBytes;
    if (compressionMethod === 8) {
      outputBytes = await inflateRawBytes(compressedBytes);
    } else if (compressionMethod !== 0) {
      throw new Error("The Excel file is using an unsupported compression method.");
    }

    entries.set(fileName, decoder.decode(outputBytes));
    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(view) {
  const minimumOffset = Math.max(0, view.byteLength - 65557);

  for (let offset = view.byteLength - 22; offset >= minimumOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      return offset;
    }
  }

  return -1;
}

async function inflateRawBytes(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function resolveWorksheetXml(entries) {
  if (entries.has("xl/worksheets/sheet1.xml")) {
    return entries.get("xl/worksheets/sheet1.xml");
  }

  const workbookXml = entries.get("xl/workbook.xml");
  const workbookRelsXml = entries.get("xl/_rels/workbook.xml.rels");

  if (workbookXml && workbookRelsXml) {
    const workbookDocument = parseSpreadsheetXml(workbookXml);
    const relsDocument = parseSpreadsheetXml(workbookRelsXml);
    const firstSheet = workbookDocument.getElementsByTagName("sheet")[0];

    if (firstSheet) {
      const relationshipId =
        firstSheet.getAttribute("r:id") ||
        firstSheet.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id",
        ) ||
        "";

      if (relationshipId) {
        const relationships = Array.from(relsDocument.getElementsByTagName("Relationship"));
        const matchedRelationship = relationships.find(
          (relationship) => relationship.getAttribute("Id") === relationshipId,
        );

        if (matchedRelationship) {
          const target = matchedRelationship.getAttribute("Target") || "";
          const normalizedPath = target.startsWith("/")
            ? target.replace(/^\/+/, "")
            : `xl/${target.replace(/^\.?\//, "")}`;

          if (entries.has(normalizedPath)) {
            return entries.get(normalizedPath);
          }
        }
      }
    }
  }

  const firstSheetKey = Array.from(entries.keys()).find(
    (key) => key.startsWith("xl/worksheets/") && key.endsWith(".xml"),
  );

  return firstSheetKey ? entries.get(firstSheetKey) : "";
}

function parseSharedStringsXml(xmlText) {
  if (!xmlText) {
    return [];
  }

  const documentNode = parseSpreadsheetXml(xmlText);
  return Array.from(documentNode.getElementsByTagName("si")).map((stringItem) =>
    Array.from(stringItem.getElementsByTagName("t"))
      .map((textNode) => textNode.textContent || "")
      .join(""),
  );
}

function parseWorksheetRows(xmlText, sharedStrings) {
  const documentNode = parseSpreadsheetXml(xmlText);
  const rows = Array.from(documentNode.getElementsByTagName("row")).map((rowNode) => {
    const rowValues = [];

    Array.from(rowNode.getElementsByTagName("c")).forEach((cellNode) => {
      const cellReference = cellNode.getAttribute("r") || "A1";
      const columnLabel = cellReference.replace(/[0-9]/g, "") || "A";
      const columnIndex = columnLabelToIndex(columnLabel);
      rowValues[columnIndex] = parseWorksheetCell(cellNode, sharedStrings);
    });

    return rowValues;
  });

  return rows.filter((row) =>
    row.some((value) => String(value === undefined ? "" : value).trim() !== ""),
  );
}

function parseWorksheetCell(cellNode, sharedStrings) {
  const type = cellNode.getAttribute("t") || "";

  if (type === "inlineStr") {
    return Array.from(cellNode.getElementsByTagName("t"))
      .map((textNode) => textNode.textContent || "")
      .join("");
  }

  const rawValueNode = cellNode.getElementsByTagName("v")[0];
  const rawValue = rawValueNode ? rawValueNode.textContent || "" : "";

  if (type === "s") {
    return sharedStrings[Number(rawValue)] || "";
  }

  if (type === "b") {
    return rawValue === "1" ? "TRUE" : "FALSE";
  }

  return rawValue;
}

function parseSpreadsheetXml(xmlText) {
  const documentNode = new DOMParser().parseFromString(xmlText, "application/xml");
  const parseError = documentNode.getElementsByTagName("parsererror")[0];

  if (parseError) {
    throw new Error("The Excel file content could not be parsed.");
  }

  return documentNode;
}

function columnLabelToIndex(label) {
  let result = 0;

  label
    .toUpperCase()
    .split("")
    .forEach((character) => {
      result = result * 26 + (character.charCodeAt(0) - 64);
    });

  return Math.max(0, result - 1);
}

function mapExcelRowsToQuestions(rows) {
  if (rows.length < 2) {
    throw new Error("The Excel file must include a header row and at least one question row.");
  }

  const headers = rows[0].map(normalizeExcelHeader);
  const columns = {
    question: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.question),
    type: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.type),
    media: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.media),
    timer: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.timer),
    answerA: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.answerA),
    answerB: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.answerB),
    answerC: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.answerC),
    answerD: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.answerD),
    correctAnswer: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.correctAnswer),
    awarded: findExcelColumnIndex(headers, EXCEL_HEADER_ALIASES.awarded),
  };

  const importedQuestions = rows
    .slice(1)
    .map((row) => {
      const prompt = getExcelCellValue(row, columns.question);
      const media = getExcelCellValue(row, columns.media);
      const answers = [
        getExcelCellValue(row, columns.answerA),
        getExcelCellValue(row, columns.answerB),
        getExcelCellValue(row, columns.answerC),
        getExcelCellValue(row, columns.answerD),
      ].map((answer) => String(answer || "").trim());
      const hasContent = [prompt, media, ...answers].some(
        (value) => String(value || "").trim() !== "",
      );

      if (!hasContent) {
        return null;
      }

      return {
        id: makeId(),
        type: normalizeImportedQuestionType(getExcelCellValue(row, columns.type), media),
        prompt: String(prompt || "").trim(),
        media: String(media || "").trim(),
        timer: clampNumber(
          getExcelCellValue(row, columns.timer),
          5,
          300,
          state.settings.defaultTimer,
        ),
        answers,
        correctIndex: parseCorrectAnswerValue(
          getExcelCellValue(row, columns.correctAnswer),
        ),
        selectedIndex: null,
        asked: false,
        awarded: parseAwardedValue(getExcelCellValue(row, columns.awarded)),
      };
    })
    .filter(Boolean);

  if (!importedQuestions.length) {
    throw new Error("The Excel file does not contain any valid question rows. Please use the provided template.");
  }

  return importedQuestions;
}

function findExcelColumnIndex(headers, aliases) {
  return headers.findIndex((header) => aliases.includes(header));
}

function getExcelCellValue(row, columnIndex) {
  if (columnIndex === -1) {
    return "";
  }

  return row[columnIndex] === undefined ? "" : row[columnIndex];
}

function normalizeExcelHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeImportedQuestionType(rawType, media) {
  const normalizedType = normalizeExcelHeader(rawType);
  const mediaText = String(media || "").trim();

  if (["image", "hinh_anh", "anh"].includes(normalizedType)) {
    return "image";
  }

  if (["video", "youtube"].includes(normalizedType)) {
    return "video";
  }

  if (["text", "chu", "cau_hoi_chu"].includes(normalizedType)) {
    return "text";
  }

  if (getYouTubeEmbedUrl(mediaText)) {
    return "video";
  }

  if (/\.(mp4|webm|ogg|mov|m4v)(\?.*)?$/i.test(mediaText) || mediaText.startsWith("data:video/")) {
    return "video";
  }

  if (
    /\.(jpg|jpeg|png|gif|webp|bmp|svg)(\?.*)?$/i.test(mediaText) ||
    mediaText.startsWith("data:image/")
  ) {
    return "image";
  }

  return "text";
}

function parseCorrectAnswerValue(value) {
  const normalizedValue = String(value || "").trim().toUpperCase();

  if (["A", "B", "C", "D"].includes(normalizedValue)) {
    return normalizedValue.charCodeAt(0) - 65;
  }

  const numericValue = Number(normalizedValue);
  if (Number.isFinite(numericValue)) {
    if (numericValue >= 1 && numericValue <= 4) {
      return numericValue - 1;
    }

    if (numericValue >= 0 && numericValue <= 3) {
      return numericValue;
    }
  }

  return 0;
}

function parseAwardedValue(value) {
  const normalizedValue = normalizeExcelHeader(value);
  return [
    "1",
    "true",
    "yes",
    "y",
    "co",
    "dung",
    "dat",
    "da_cong",
    "da_cong_diem",
  ].includes(normalizedValue);
}

function persistAndRender(options = {}) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (error) {
    window.alert(
      "Unable to save data in the browser. Media or audio files may be too large, so file paths are recommended for large assets.",
    );
  }

  renderEverything(options);
}

function persistStateQuietly() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (error) {
    // Ignore storage quota issues during live presentation flow.
  }
}

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      return deepClone(defaultState);
    }

    return sanitizeState(JSON.parse(saved));
  } catch (error) {
    return deepClone(defaultState);
  }
}

function sanitizeState(value) {
  const settings = value && value.settings ? value.settings : {};
  const sounds = settings && settings.sounds ? settings.sounds : {};
  const questions = Array.isArray(value && value.questions) ? value.questions : [];
  const defaultTimer = clampNumber(
    settings.defaultTimer,
    5,
    300,
    defaultState.settings.defaultTimer,
  );

  return {
    settings: {
      title: migrateLegacyDefaultText(
        settings.title,
        LEGACY_DEFAULT_TEXT.title,
        defaultState.settings.title,
      ),
      subtitle: migrateLegacyDefaultText(
        settings.subtitle,
        LEGACY_DEFAULT_TEXT.subtitle,
        defaultState.settings.subtitle,
      ),
      rulesContent: migrateLegacyDefaultText(
        settings.rulesContent,
        LEGACY_DEFAULT_TEXT.rulesContent,
        defaultState.settings.rulesContent,
      ),
      defaultTimer,
      useQuestionBoard:
        typeof settings.useQuestionBoard === "boolean"
          ? settings.useQuestionBoard
          : defaultState.settings.useQuestionBoard,
      randomQuestionSelection:
        typeof settings.randomQuestionSelection === "boolean"
          ? settings.randomQuestionSelection
          : defaultState.settings.randomQuestionSelection,
      autoAdvance: Boolean(settings.autoAdvance),
      autoRevealOnTimeout:
        typeof settings.autoRevealOnTimeout === "boolean"
          ? settings.autoRevealOnTimeout
          : defaultState.settings.autoRevealOnTimeout,
      endingTitle: migrateLegacyDefaultText(
        settings.endingTitle,
        LEGACY_DEFAULT_TEXT.endingTitle,
        defaultState.settings.endingTitle,
      ),
      endingMessage: migrateLegacyDefaultText(
        settings.endingMessage,
        LEGACY_DEFAULT_TEXT.endingMessage,
        defaultState.settings.endingMessage,
      ),
      pointsPerQuestion: clampNumber(
        settings.pointsPerQuestion,
        1,
        100,
        defaultState.settings.pointsPerQuestion,
      ),
      backgroundVolume: clampNumber(
        settings.backgroundVolume,
        0,
        100,
        defaultState.settings.backgroundVolume,
      ),
      effectVolume: clampNumber(
        settings.effectVolume,
        0,
        100,
        defaultState.settings.effectVolume,
      ),
      sounds: {
        background: String(sounds.background || ""),
        reveal: String(sounds.reveal || ""),
        timeout: String(sounds.timeout || ""),
      },
    },
    questions: questions.map((question) => sanitizeQuestion(question, defaultTimer)),
  };
}

function sanitizeQuestion(question, defaultTimer) {
  const answers = Array.isArray(question && question.answers)
    ? question.answers.slice(0, 4)
    : [];

  while (answers.length < 4) {
    answers.push("");
  }

  return {
    id: String((question && question.id) || makeId()),
    type: ["text", "image", "video"].includes(question && question.type)
      ? question.type
      : "text",
    prompt: String((question && question.prompt) || ""),
    media: String((question && question.media) || ""),
    timer: clampNumber(question && question.timer, 5, 300, defaultTimer),
    answers: answers.map((answer) => String(answer || "")),
    correctIndex: clampNumber(question && question.correctIndex, 0, 3, 0),
    selectedIndex: Number.isInteger(question && question.selectedIndex)
      ? clampNumber(question.selectedIndex, 0, 3, 0)
      : null,
    asked: Boolean(question && question.asked),
    awarded: Boolean(question && question.awarded),
  };
}

function migrateLegacyDefaultText(value, legacyText, fallbackText) {
  const normalizedValue = String(value || "");

  if (!normalizedValue) {
    return fallbackText;
  }

  return normalizedValue === legacyText ? fallbackText : normalizedValue;
}

function createBlankQuestion() {
  return {
    id: makeId(),
    type: "text",
    prompt: "",
    media: "",
    timer: state.settings.defaultTimer,
    answers: ["", "", "", ""],
    correctIndex: 0,
    selectedIndex: null,
    asked: false,
    awarded: false,
  };
}

function getSelectedQuestion() {
  return state.questions.find((question) => question.id === selectedQuestionId) || null;
}

function getSelectedQuestionIndex() {
  return state.questions.findIndex((question) => question.id === selectedQuestionId);
}

function getCurrentQuestion() {
  return state.questions[currentSlideIndex] || null;
}

function getCurrentStage() {
  if (summaryVisible) {
    return "summary";
  }

  if (introVisible) {
    return "intro";
  }

  if (boardVisible) {
    return "board";
  }

  return "question";
}

function getCurrentQuestionId() {
  const question = getCurrentQuestion();
  return question ? question.id : null;
}

function getCurrentDuration() {
  const question = getCurrentQuestion();
  return question ? question.timer : state.settings.defaultTimer;
}

function getScoreStats() {
  const correctCount = state.questions.filter((question) => question.awarded).length;

  return {
    correctCount,
    totalScore: correctCount * state.settings.pointsPerQuestion,
  };
}

function getUnaskedQuestionIndices() {
  return state.questions.reduce((indexes, question, index) => {
    if (!question.asked) {
      indexes.push(index);
    }

    return indexes;
  }, []);
}

function getDefaultBoardQuestionIndex() {
  if (!state.questions.length) {
    return -1;
  }

  const unaskedIndexes = getUnaskedQuestionIndices();
  if (unaskedIndexes.length) {
    return unaskedIndexes[0];
  }

  return -1;
}

function setActiveView(viewId) {
  elements.modeButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.view === viewId);
  });

  elements.viewPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === viewId);
  });
}

function labelQuestionType(type) {
  const labels = {
    text: "Text Question",
    image: "Image Question",
    video: "Video Question",
  };

  return labels[type] || "Question";
}

function answerLabel(index) {
  return String.fromCharCode(65 + index);
}

function clampNumber(value, min, max, fallback) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }

  return Math.min(max, Math.max(min, Math.round(parsed)));
}
