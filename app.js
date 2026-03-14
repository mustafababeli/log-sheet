const HOURS = [
  "8:00",
  "9:00",
  "10:00",
  "11:00",
  "12:00",
  "13:00",
  "14:00",
  "15:00",
  "16:00",
  "17:00",
  "18:00",
  "19:00",
  "20:00",
  "21:00",
  "22:00",
  "23:00",
  "24:00",
  "1:00",
  "2:00",
  "3:00",
  "4:00",
  "5:00",
  "6:00",
  "7:00",
];

const HOUR_OPTIONS = [
  "08:00",
  "09:00",
  "10:00",
  "11:00",
  "12:00",
  "13:00",
  "14:00",
  "15:00",
  "16:00",
  "17:00",
  "18:00",
  "19:00",
  "20:00",
  "21:00",
  "22:00",
  "23:00",
  "24:00",
  "01:00",
  "02:00",
  "03:00",
  "04:00",
  "05:00",
  "06:00",
  "07:00",
];

const L_GROUPS = Array.from({ length: 8 }, (_, index) => ({
  group: `L${index + 1}`,
  fields: ["KV", "A", "MVR"],
}));

const MT_GROUPS = ["M. TR1", "M. TR2", "M. TR3"].map((group) => ({
  group,
  fields: [
    { label: "S.TAP", unit: "" },
    { label: "KV", unit: "KV" },
    { label: "A", unit: "A" },
    { label: "ACTV", unit: "MW" },
    { label: "REACTOR", unit: "MVA" },
  ],
}));

const FEEDERS_33 = Array.from({ length: 15 }, (_, index) => String(index + 1));
const FEEDERS_11 = Array.from({ length: 18 }, (_, index) => String(index + 1));

const FIELD_GROUPS = [
  ...L_GROUPS,
  ...MT_GROUPS,
  {
    group: "FEEDERS 33 KV",
    fields: FEEDERS_33,
  },
  {
    group: "FEEDERS 11 KV",
    fields: FEEDERS_11,
  },
  {
    group: "TEMPRETURE",
    fields: ["TR.1 W", "TR.1 O", "TR.2 W", "TR.2 O", "TR.3 W", "TR.3 O"],
  },
];

const FLAT_FIELDS = [
  ...L_GROUPS.flatMap((group) =>
    group.fields.map((field, fieldIndex) => ({
      key: `${group.group}__${field}`,
      group: group.group,
      label: field,
      fullLabel: `${group.group} - ${field}`,
      groupHeaderId: `group-${group.group}`,
      labelHeaderId: `l-${group.group}-${fieldIndex}`,
    })),
  ),
  ...MT_GROUPS.flatMap((group, groupIndex) =>
    group.fields.map((field, fieldIndex) => ({
      key: `${group.group}__${field.label}`,
      group: group.group,
      label: field.label,
      fullLabel: `${group.group} - ${field.label}`,
      groupHeaderId: `group-${group.group}`,
      labelHeaderId: `mt-label-${groupIndex}-${fieldIndex}`,
    })),
  ),
  ...FEEDERS_33.map((field, fieldIndex) => ({
    key: `FEEDERS 33 KV__${field}`,
    group: "FEEDERS 33 KV",
    label: field,
    fullLabel: `FEEDERS 33 KV - ${field}`,
    groupHeaderId: "feeders-33",
    labelHeaderId: `feed33-${fieldIndex}`,
  })),
  ...FEEDERS_11.map((field, fieldIndex) => ({
    key: `FEEDERS 11 KV__${field}`,
    group: "FEEDERS 11 KV",
    label: field,
    fullLabel: `FEEDERS 11 KV - ${field}`,
    groupHeaderId: "feeders-11",
    labelHeaderId: `feed11-${fieldIndex}`,
  })),
  ...[
    {
      key: "TEMPRETURE__TR.1 W",
      label: "TR.1 W",
      subgroupHeaderId: "temperature-tr1",
      labelHeaderId: "temperature-unit-0",
    },
    {
      key: "TEMPRETURE__TR.1 O",
      label: "TR.1 O",
      subgroupHeaderId: "temperature-tr1",
      labelHeaderId: "temperature-unit-1",
    },
    {
      key: "TEMPRETURE__TR.2 W",
      label: "TR.2 W",
      subgroupHeaderId: "temperature-tr2",
      labelHeaderId: "temperature-unit-2",
    },
    {
      key: "TEMPRETURE__TR.2 O",
      label: "TR.2 O",
      subgroupHeaderId: "temperature-tr2",
      labelHeaderId: "temperature-unit-3",
    },
    {
      key: "TEMPRETURE__TR.3 W",
      label: "TR.3 W",
      subgroupHeaderId: "temperature-tr3",
      labelHeaderId: "temperature-unit-4",
    },
    {
      key: "TEMPRETURE__TR.3 O",
      label: "TR.3 O",
      subgroupHeaderId: "temperature-tr3",
      labelHeaderId: "temperature-unit-5",
    },
  ].map((field) => ({
    ...field,
    group: "TEMPRETURE",
    fullLabel: `TEMPRETURE - ${field.label}`,
    groupHeaderId: "temperature-group",
  })),
];

const SF6_COLUMNS = [
  "L1",
  "L2",
  "L3",
  "L4",
  "L5",
  "L6",
  "L7",
  "L8",
  "Bas.c",
  "M.Tr1",
  "M.Tr2",
  "M.Tr3",
];
const STORAGE_PREFIX = "electricity-log-sheet";
const HEADER_STORAGE_KEY = `${STORAGE_PREFIX}-header-overrides`;
const SETTINGS_STORAGE_KEY = `${STORAGE_PREFIX}-settings`;
const UI_TEXT = {
  en: {
    sheetDate: "Sheet Date",
    activeHour: "Active Hour",
    currentField: "Current Field",
    value: "Value",
    changeDate: "Change Date",
    resetSheet: "Reset Sheet",
    settings: "Settings",
    startTalking: "Start Talking",
    stopTalking: "Stop Talking",
    openCamera: "Open Camera",
    back: "Back",
    saveAndNext: "Save And Next",
    skipSpots: "Skip Spots",
    exportExcel: "Export Excel",
    shareSheet: "Share Sheet",
    settingsTitle: "Settings",
    language: "Language",
    close: "Close",
    saveSettings: "Save Settings",
    valuePlaceholder: "Type or speak the reading",
    noDateSelected: "No date selected",
    ready: "Ready",
  },
  ar: {
    sheetDate: "تاريخ الورقة",
    activeHour: "الساعة الحالية",
    currentField: "الحقل الحالي",
    value: "القيمة",
    changeDate: "تغيير التاريخ",
    resetSheet: "إعادة ضبط الورقة",
    settings: "الإعدادات",
    startTalking: "ابدأ التحدث",
    stopTalking: "إيقاف التحدث",
    openCamera: "فتح الكاميرا",
    back: "رجوع",
    saveAndNext: "حفظ والانتقال",
    skipSpots: "تخطي خانات",
    exportExcel: "تصدير إكسل",
    shareSheet: "مشاركة الورقة",
    settingsTitle: "الإعدادات",
    language: "اللغة",
    close: "إغلاق",
    saveSettings: "حفظ الإعدادات",
    valuePlaceholder: "اكتب أو انطق القراءة",
    noDateSelected: "لم يتم اختيار تاريخ",
    ready: "جاهز",
  },
};

const state = {
  date: "",
  activeHourIndex: 0,
  activeFieldIndex: 0,
  entries: {},
  skippedFields: [],
  headerOverrides: {},
  editingHeaderId: null,
  language: "en",
};

const setupDialog = document.getElementById("setupDialog");
const skipDialog = document.getElementById("skipDialog");
const headerDialog = document.getElementById("headerDialog");
const setupForm = document.getElementById("setupForm");
const skipForm = document.getElementById("skipForm");
const headerForm = document.getElementById("headerForm");
const sheetDateInput = document.getElementById("sheetDate");
const startHourSelect = document.getElementById("startHour");
const cancelSetupButton = document.getElementById("cancelSetupButton");
const cancelSkipButton = document.getElementById("cancelSkipButton");
const cancelHeaderButton = document.getElementById("cancelHeaderButton");
const openSetupButton = document.getElementById("openSetupButton");
const manageSkipsButton = document.getElementById("manageSkipsButton");
const resetSheetButton = document.getElementById("resetSheetButton");
const settingsButton = document.getElementById("settingsButton");
const resetHeaderButton = document.getElementById("resetHeaderButton");
const cameraButton = document.getElementById("cameraButton");
const closeCameraButton = document.getElementById("closeCameraButton");
const captureCameraButton = document.getElementById("captureCameraButton");
const retakeCameraButton = document.getElementById("retakeCameraButton");
const saveFieldButton = document.getElementById("saveFieldButton");
const exportExcelButton = document.getElementById("exportExcelButton");
const shareSheetButton = document.getElementById("shareSheetButton");
const fieldNameInput = document.getElementById("fieldName");
const fieldValueInput = document.getElementById("fieldValue");
const voiceButton = document.getElementById("voiceButton");
const voiceStatus = document.getElementById("voiceStatus");
const backFieldButton = document.getElementById("backFieldButton");
const sheetDateText = document.getElementById("sheetDateText");
const activeHourText = document.getElementById("activeHourText");
const currentFieldText = document.getElementById("currentFieldText");
const valueText = document.getElementById("valueText");
const sheetDateLabel = document.getElementById("sheetDateLabel");
const activeHourLabel = document.getElementById("activeHourLabel");
const progressLabel = document.getElementById("progressLabel");
const completionLabel = document.getElementById("completionLabel");
const saveStatus = document.getElementById("saveStatus");
const logTable = document.getElementById("logTable");
const sf6Table = document.getElementById("sf6Table");
const dateDay = document.getElementById("dateDay");
const dateMonth = document.getElementById("dateMonth");
const dateYear = document.getElementById("dateYear");
const skipFieldList = document.getElementById("skipFieldList");
const headerTextInput = document.getElementById("headerTextInput");
const headerSubTextInput = document.getElementById("headerSubTextInput");
const headerHorizontalButton = document.getElementById(
  "headerHorizontalButton",
);
const headerVerticalButton = document.getElementById("headerVerticalButton");
const cameraDialog = document.getElementById("cameraDialog");
const cameraVideo = document.getElementById("cameraVideo");
const cameraCanvas = document.getElementById("cameraCanvas");
const cameraSnapshot = document.getElementById("cameraSnapshot");
const cameraStatus = document.getElementById("cameraStatus");
const settingsDialog = document.getElementById("settingsDialog");
const settingsForm = document.getElementById("settingsForm");
const settingsTitle = document.getElementById("settingsTitle");
const languageLabel = document.getElementById("languageLabel");
const languageSelect = document.getElementById("languageSelect");
const closeSettingsButton = document.getElementById("closeSettingsButton");
const saveSettingsButton = document.getElementById("saveSettingsButton");
const SpeechRecognitionApi =
  window.SpeechRecognition || window.webkitSpeechRecognition;
const currentFieldGroup = currentFieldText.closest(".control-group");
const infoRow = currentFieldGroup.closest(".info-row");
const actionRow = document.querySelector(".action-row");
const valueGroup = document.querySelector(".value-group");

let recognition = null;
let isListening = false;
let cameraStream = null;
let selectedHeaderOrientation = "horizontal";
let currentFieldInActionRow = false;

function t(key) {
  return UI_TEXT[state.language]?.[key] || UI_TEXT.en[key] || key;
}

function loadSettings() {
  const saved = localStorage.getItem(SETTINGS_STORAGE_KEY);
  if (!saved) {
    return { language: "en" };
  }

  try {
    const parsed = JSON.parse(saved);
    return { language: parsed.language === "ar" ? "ar" : "en" };
  } catch {
    return { language: "en" };
  }
}

function saveSettings() {
  localStorage.setItem(
    SETTINGS_STORAGE_KEY,
    JSON.stringify({
      language: state.language,
    }),
  );
}

function applyLanguage() {
  document.documentElement.lang = state.language === "ar" ? "ar" : "en";
  document.documentElement.dir = state.language === "ar" ? "rtl" : "ltr";
  sheetDateText.textContent = t("sheetDate");
  activeHourText.textContent = t("activeHour");
  currentFieldText.textContent = t("currentField");
  valueText.textContent = t("value");
  openSetupButton.textContent = t("changeDate");
  resetSheetButton.textContent = t("resetSheet");
  voiceButton.textContent = isListening ? t("stopTalking") : t("startTalking");
  cameraButton.textContent = t("openCamera");
  backFieldButton.textContent = t("back");
  saveFieldButton.textContent = t("saveAndNext");
  manageSkipsButton.textContent = t("skipSpots");
  exportExcelButton.textContent = t("exportExcel");
  shareSheetButton.textContent = t("shareSheet");
  settingsTitle.textContent = t("settingsTitle");
  languageLabel.textContent = t("language");
  closeSettingsButton.textContent = t("close");
  saveSettingsButton.textContent = t("saveSettings");
  fieldValueInput.placeholder = t("valuePlaceholder");

  if (!state.date) {
    sheetDateLabel.textContent = t("noDateSelected");
  }
}

function syncMobileFieldLayout() {
  const shouldMoveCurrentField = window.innerWidth <= 640;

  if (shouldMoveCurrentField && !currentFieldInActionRow) {
    actionRow.insertBefore(currentFieldGroup, valueGroup);
    currentFieldInActionRow = true;
    return;
  }

  if (!shouldMoveCurrentField && currentFieldInActionRow) {
    infoRow.insertBefore(currentFieldGroup, openSetupButton);
    currentFieldInActionRow = false;
  }
}

function setVoiceStatus(message, status = "idle") {
  voiceStatus.textContent = message;
  voiceStatus.classList.remove("listening", "error");

  if (status === "listening") {
    voiceStatus.classList.add("listening");
  }

  if (status === "error") {
    voiceStatus.classList.add("error");
  }
}

function setListeningState(listening) {
  isListening = listening;
  voiceButton.textContent = listening ? t("stopTalking") : t("startTalking");
}

function normalizeTranscript(transcript) {
  return transcript
    .replace(/\bzero\b/gi, "0")
    .replace(/\bone\b/gi, "1")
    .replace(/\btwo\b/gi, "2")
    .replace(/\bthree\b/gi, "3")
    .replace(/\bfour\b/gi, "4")
    .replace(/\bfive\b/gi, "5")
    .replace(/\bsix\b/gi, "6")
    .replace(/\bseven\b/gi, "7")
    .replace(/\beight\b/gi, "8")
    .replace(/\bnine\b/gi, "9")
    .replace(/\bpoint\b/gi, ".")
    .replace(/\bdot\b/gi, ".")
    .replace(/\bminus\b/gi, "-")
    .replace(/(\d)\s+(?=\d)/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeValueInput(value) {
  return value
    .replace(/(\d)\s+(?=\d)/g, "$1")
    .replace(/[^0-9.\-]/g, "")
    .replace(/(?!^)-/g, "")
    .replace(/(\..*)\./g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function setupVoiceRecognition() {
  if (!SpeechRecognitionApi) {
    voiceButton.disabled = true;
    setVoiceStatus(
      "Microphone input is not supported in this browser. Typing still works.",
      "error",
    );
    return;
  }

  recognition = new SpeechRecognitionApi();
  recognition.lang = "en-US";
  recognition.interimResults = true;
  recognition.continuous = false;
  recognition.maxAlternatives = 1;

  recognition.addEventListener("start", () => {
    setListeningState(true);
    setVoiceStatus("Listening... speak the value now.", "listening");
  });

  recognition.addEventListener("result", (event) => {
    const transcript = Array.from(event.results)
      .map((result) => result[0].transcript)
      .join(" ");

    fieldValueInput.value = normalizeTranscript(transcript);

    if (event.results[event.results.length - 1].isFinal) {
      setVoiceStatus("Voice captured. Review it, then press Save And Next.");
    }
  });

  recognition.addEventListener("end", () => {
    setListeningState(false);
    if (voiceStatus.classList.contains("listening")) {
      setVoiceStatus("Voice stopped. Review the text, then save.");
    }
  });

  recognition.addEventListener("error", (event) => {
    setListeningState(false);

    if (event.error === "not-allowed") {
      setVoiceStatus(
        "Microphone permission was blocked. Allow microphone access and try again.",
        "error",
      );
      return;
    }

    if (event.error === "no-speech") {
      setVoiceStatus(
        "No speech was heard. Try again and speak a little closer to the microphone.",
        "error",
      );
      return;
    }

    setVoiceStatus(
      `Voice input error: ${event.error}. You can still type manually.`,
      "error",
    );
  });
}

function startVoiceCapture() {
  if (!recognition || isListening) {
    return;
  }

  try {
    recognition.start();
  } catch {
    window.setTimeout(() => {
      if (!isListening) {
        try {
          recognition.start();
        } catch {
          setVoiceStatus(
            "Microphone is busy. Wait a moment and try again.",
            "error",
          );
        }
      }
    }, 150);
  }
}

function getToday() {
  const now = new Date();
  const year = now.getFullYear();
  const month = `${now.getMonth() + 1}`.padStart(2, "0");
  const day = `${now.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function createEmptyEntries() {
  return Object.fromEntries(
    HOURS.map((hour) => [
      hour,
      Object.fromEntries(FLAT_FIELDS.map((field) => [field.key, ""])),
    ]),
  );
}

function populateHourOptions() {
  startHourSelect.innerHTML = "";
  HOUR_OPTIONS.forEach((hour, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = hour;
    startHourSelect.append(option);
  });
}

function getStorageKey(date) {
  return `${STORAGE_PREFIX}-${date}`;
}

function getSkipStorageKey(date) {
  return `${STORAGE_PREFIX}-skips-${date}`;
}

function saveEntries() {
  if (!state.date) {
    return;
  }

  localStorage.setItem(
    getStorageKey(state.date),
    JSON.stringify(state.entries),
  );
  saveStatus.textContent = "Saved";
}

function loadEntries(date) {
  const saved = localStorage.getItem(getStorageKey(date));
  const emptyEntries = createEmptyEntries();

  if (!saved) {
    return emptyEntries;
  }

  try {
    const parsed = JSON.parse(saved);
    for (const hour of HOURS) {
      emptyEntries[hour] = {
        ...emptyEntries[hour],
        ...(parsed[hour] || {}),
      };
    }
    return emptyEntries;
  } catch {
    return emptyEntries;
  }
}

function loadSkippedFields(date) {
  const saved = localStorage.getItem(getSkipStorageKey(date));
  if (!saved) {
    return [];
  }

  try {
    const parsed = JSON.parse(saved);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveSkippedFields() {
  if (!state.date) {
    return;
  }

  localStorage.setItem(
    getSkipStorageKey(state.date),
    JSON.stringify(state.skippedFields),
  );
}

function loadHeaderOverrides() {
  const saved = localStorage.getItem(HEADER_STORAGE_KEY);
  if (!saved) {
    return {};
  }

  try {
    const parsed = JSON.parse(saved);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveHeaderOverrides() {
  localStorage.setItem(
    HEADER_STORAGE_KEY,
    JSON.stringify(state.headerOverrides),
  );
}

function buildExportFileName() {
  return `daily-operation-log-${state.date || "sheet"}.xls`;
}

function buildExcelHtml() {
  const headerHtml = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
      <div style="font-weight:700;font-size:18px;">Date: ${state.date || ""}</div>
      <div style="font-weight:700;font-size:22px;">DAILY OPERATION LOG SHEET</div>
    </div>
  `;

  return `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, Helvetica, sans-serif; }
          table { border-collapse: collapse; }
          th, td { border: 1px solid #444; text-align: center; }
        </style>
      </head>
      <body>
        ${headerHtml}
        ${logTable.outerHTML}
        <br>
        ${sf6Table.outerHTML}
      </body>
    </html>
  `;
}

function createExcelFile() {
  const html = buildExcelHtml();
  const blob = new Blob([html], {
    type: "application/vnd.ms-excel",
  });

  return new File([blob], buildExportFileName(), {
    type: "application/vnd.ms-excel",
  });
}

function downloadExcelFile() {
  const file = createExcelFile();
  const url = URL.createObjectURL(file);
  const link = document.createElement("a");
  link.href = url;
  link.download = file.name;
  link.click();
  URL.revokeObjectURL(url);
}

async function shareSheetFile() {
  if (!state.date) {
    saveStatus.textContent = "Choose a date first";
    return;
  }

  const file = createExcelFile();

  if (!navigator.share) {
    downloadExcelFile();
    saveStatus.textContent = "Share not supported, downloaded instead";
    return;
  }

  try {
    const shareData = {
      title: "Daily Operation Log Sheet",
      text: `Daily sheet for ${state.date}`,
      files: [file],
    };

    if (navigator.canShare && !navigator.canShare(shareData)) {
      downloadExcelFile();
      saveStatus.textContent = "File share not supported, downloaded instead";
      return;
    }

    await navigator.share(shareData);
    saveStatus.textContent = "Sheet shared";
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }

    downloadExcelFile();
    saveStatus.textContent = "Share failed, downloaded instead";
  }
}

function setHeaderOrientationSelection(orientation) {
  selectedHeaderOrientation = orientation;
  headerHorizontalButton.classList.toggle(
    "active",
    orientation === "horizontal",
  );
  headerVerticalButton.classList.toggle("active", orientation === "vertical");
}

async function openCamera() {
  if (!navigator.mediaDevices?.getUserMedia) {
    cameraStatus.textContent = "Camera is not supported in this browser.";
    cameraStatus.classList.add("error");
    cameraDialog.showModal();
    return;
  }

  try {
    cameraStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: "environment",
      },
      audio: false,
    });

    cameraVideo.srcObject = cameraStream;
    cameraVideo.hidden = false;
    cameraSnapshot.hidden = true;
    cameraStatus.textContent =
      "Camera opened. Capture a photo of the meter number.";
    cameraStatus.classList.remove("error");
    cameraDialog.showModal();
  } catch {
    cameraStatus.textContent = "Camera access was blocked or is unavailable.";
    cameraStatus.classList.add("error");
    cameraDialog.showModal();
  }
}

function stopCamera() {
  if (!cameraStream) {
    return;
  }

  cameraStream.getTracks().forEach((track) => track.stop());
  cameraStream = null;
  cameraVideo.srcObject = null;
}

function captureCameraFrame() {
  if (!cameraStream) {
    cameraStatus.textContent = "Open the camera first.";
    cameraStatus.classList.add("error");
    return;
  }

  const width = cameraVideo.videoWidth;
  const height = cameraVideo.videoHeight;
  if (!width || !height) {
    cameraStatus.textContent =
      "Camera is still loading. Try again in a moment.";
    cameraStatus.classList.add("error");
    return;
  }

  cameraCanvas.width = width;
  cameraCanvas.height = height;
  const context = cameraCanvas.getContext("2d");
  context.drawImage(cameraVideo, 0, 0, width, height);

  cameraSnapshot.src = cameraCanvas.toDataURL("image/png");
  cameraSnapshot.hidden = false;
  cameraVideo.hidden = true;
  cameraStatus.textContent =
    "Photo captured. OCR is not installed yet, so this is camera-only for now.";
  cameraStatus.classList.remove("error");
}

function getCurrentHour() {
  return HOURS[state.activeHourIndex];
}

function getCurrentField() {
  return FLAT_FIELDS[state.activeFieldIndex];
}

function isSkippedField(fieldKey) {
  return state.skippedFields.includes(fieldKey);
}

function isFeederField(groupName) {
  return groupName === "FEEDERS 33 KV" || groupName === "FEEDERS 11 KV";
}

function getHourCompletion(hour) {
  const values = Object.values(state.entries[hour] || {});
  const filled = values.filter((value) => String(value).trim() !== "").length;
  return { filled, total: values.length };
}

function createCell(tag, className, text = "", rowSpan = 1, colSpan = 1) {
  const cell = document.createElement(tag);
  if (className) {
    cell.className = className;
  }
  if (text) {
    cell.textContent = text;
  }
  if (rowSpan > 1) {
    cell.rowSpan = rowSpan;
  }
  if (colSpan > 1) {
    cell.colSpan = colSpan;
  }
  return cell;
}

function getHeaderOverride(
  headerId,
  fallbackText,
  fallbackOrientation = "horizontal",
) {
  const override = state.headerOverrides[headerId];
  if (!override) {
    return {
      text: fallbackText,
      subText: "",
      orientation: fallbackOrientation,
    };
  }

  return {
    text: override.text || fallbackText,
    subText: override.subText || "",
    orientation: override.orientation || fallbackOrientation,
  };
}

function getHeaderDisplayText(
  headerId,
  fallbackText,
  fallbackOrientation = "horizontal",
) {
  const config = getHeaderOverride(headerId, fallbackText, fallbackOrientation);
  return config.subText
    ? `${config.text} ${config.subText}`.trim()
    : config.text;
}

function getFieldDisplayLabel(field) {
  const parts = [];

  if (field.groupHeaderId) {
    parts.push(getHeaderDisplayText(field.groupHeaderId, field.group));
  }

  if (field.subgroupHeaderId) {
    const subgroupFallback = field.label.split(" ")[0];
    parts.push(getHeaderDisplayText(field.subgroupHeaderId, subgroupFallback));
  }

  if (field.labelHeaderId) {
    parts.push(getHeaderDisplayText(field.labelHeaderId, field.label));
  } else if (field.label) {
    parts.push(field.label);
  }

  return parts.join(" - ");
}

function appendHeaderContent(cell, text, subText, orientation, large = false) {
  if (orientation === "vertical") {
    if (subText) {
      const stack = document.createElement("div");
      stack.className = "vertical-stack";

      const mainSpan = document.createElement("div");
      mainSpan.className = large ? "vertical-word large" : "vertical-word";
      mainSpan.textContent = text;

      const subSpan = document.createElement("div");
      subSpan.className = "vertical-word";
      subSpan.textContent = subText;

      stack.append(mainSpan, subSpan);
      cell.append(stack);
    } else {
      const span = document.createElement("div");
      span.className = large ? "vertical-word large" : "vertical-word";
      span.textContent = text;
      cell.append(span);
    }
    return;
  }

  if (subText) {
    const mainSpan = document.createElement("div");
    mainSpan.className = "header-text horizontal";
    if (large) {
      mainSpan.classList.add("large");
    }
    mainSpan.textContent = text;

    const subSpan = document.createElement("div");
    subSpan.className = "header-subtext";
    subSpan.textContent = subText;

    mainSpan.textContent = "";

    const mainText = document.createElement("div");
    mainText.textContent = text;

    mainSpan.append(mainText, subSpan);
    cell.append(mainSpan);
  } else {
    const span = document.createElement("div");
    span.className = "header-text horizontal";
    if (large) {
      span.classList.add("large");
    }
    span.textContent = text;
    cell.append(span);
  }
}

function createEditableHeaderCell(
  tag,
  className,
  headerId,
  text,
  rowSpan = 1,
  colSpan = 1,
  defaultOrientation = "horizontal",
  large = false,
) {
  const cell = createCell(tag, className, "", rowSpan, colSpan);
  const config = getHeaderOverride(headerId, text, defaultOrientation);
  cell.classList.add("editable-header");
  cell.dataset.headerId = headerId;
  cell.dataset.defaultText = text;
  cell.dataset.defaultOrientation = defaultOrientation;
  appendHeaderContent(
    cell,
    config.text,
    config.subText,
    config.orientation,
    large,
  );
  return cell;
}

function renderMainTable() {
  const thead = document.createElement("thead");
  const row1 = document.createElement("tr");
  const row2 = document.createElement("tr");
  const row3 = document.createElement("tr");

  row1.append(
    createEditableHeaderCell(
      "th",
      "time-head group",
      "time",
      "Time",
      3,
      1,
      "vertical",
      true,
    ),
  );

  L_GROUPS.forEach((group) => {
    row1.append(
      createEditableHeaderCell(
        "th",
        "group",
        `group-${group.group}`,
        group.group,
        2,
        group.fields.length,
      ),
    );
    group.fields.forEach((field, fieldIndex) => {
      row3.append(
        createEditableHeaderCell(
          "th",
          "unit-row l-cell",
          `l-${group.group}-${fieldIndex}`,
          field,
          1,
          1,
          "vertical",
        ),
      );
    });
  });

  MT_GROUPS.forEach((group, groupIndex) => {
    row1.append(
      createEditableHeaderCell(
        "th",
        "group mt-group",
        `group-${group.group}`,
        group.group,
        1,
        group.fields.length,
      ),
    );
    group.fields.forEach((field, fieldIndex) => {
      row2.append(
        createEditableHeaderCell(
          "th",
          "mt-cell",
          `mt-label-${groupIndex}-${fieldIndex}`,
          field.label,
          1,
          1,
          "vertical",
        ),
      );
      row3.append(
        createEditableHeaderCell(
          "th",
          "unit-row mt-cell",
          `mt-unit-${groupIndex}-${fieldIndex}`,
          field.unit || "",
          1,
          1,
          "vertical",
        ),
      );
    });
  });

  row1.append(
    createEditableHeaderCell(
      "th",
      "group",
      "feeders-33",
      "FEEDERS 33 KV",
      1,
      FEEDERS_33.length,
    ),
  );
  FEEDERS_33.forEach((field, fieldIndex) => {
    row2.append(
      createEditableHeaderCell(
        "th",
        "feeder-cell",
        `feed33-${fieldIndex}`,
        field,
        2,
      ),
    );
  });

  row1.append(
    createEditableHeaderCell(
      "th",
      "group",
      "feeders-11",
      "FEEDERS 11 KV",
      1,
      FEEDERS_11.length,
    ),
  );
  FEEDERS_11.forEach((field, fieldIndex) => {
    row2.append(
      createEditableHeaderCell(
        "th",
        "feeder-cell",
        `feed11-${fieldIndex}`,
        field,
        2,
      ),
    );
  });

  row1.append(
    createEditableHeaderCell(
      "th",
      "group",
      "temperature-group",
      "TEMPRETURE",
      1,
      6,
    ),
  );
  row2.append(
    createEditableHeaderCell("th", "subgroup", "temperature-tr1", "TR.1", 1, 2),
  );
  row2.append(
    createEditableHeaderCell("th", "subgroup", "temperature-tr2", "TR.2", 1, 2),
  );
  row2.append(
    createEditableHeaderCell("th", "subgroup", "temperature-tr3", "TR.3", 1, 2),
  );
  ["W", "O", "W", "O", "W", "O"].forEach((field, fieldIndex) => {
    row3.append(
      createEditableHeaderCell(
        "th",
        "unit-row temp-cell",
        `temperature-unit-${fieldIndex}`,
        field,
      ),
    );
  });

  thead.append(row1, row2, row3);

  const tbody = document.createElement("tbody");

  HOURS.forEach((hour, hourIndex) => {
    const row = document.createElement("tr");
    row.dataset.hourIndex = String(hourIndex);

    if (hourIndex === state.activeHourIndex) {
      row.classList.add("active-row");
    }

    const completion = getHourCompletion(hour);
    if (completion.filled === completion.total) {
      row.classList.add("completed-row");
    }

    row.append(createCell("td", "hour-cell", hour));

    FLAT_FIELDS.forEach((field, fieldIndex) => {
      const value = state.entries[hour]?.[field.key] || "";
      const cell = createCell("td", "value-cell", value);
      cell.dataset.hourIndex = String(hourIndex);
      cell.dataset.fieldIndex = String(fieldIndex);

      if (isFeederField(field.group)) {
        cell.classList.add("feeder-value-cell");
      }

      if (String(value).trim() === "0") {
        cell.classList.add("zero-value");
      }

      if (isSkippedField(field.key)) {
        cell.classList.add("skipped-cell");
      }

      if (
        hourIndex === state.activeHourIndex &&
        fieldIndex === state.activeFieldIndex
      ) {
        cell.classList.add("active-cell");
      }

      row.append(cell);
    });

    tbody.append(row);
  });

  logTable.innerHTML = "";
  logTable.append(thead, tbody);
}

function renderSf6Table() {
  const row1 = document.createElement("tr");
  row1.append(createEditableHeaderCell("th", "sf6-title", "sf6-title", "Sf6"));
  SF6_COLUMNS.forEach((column, index) => {
    const className = index >= 8 ? "sf6-small" : "sf6-head";
    row1.append(
      createEditableHeaderCell("th", className, `sf6-col-${index}`, column),
    );
  });
  row1.append(createCell("td", "sf6-blank", "", 4, 30));

  const row2 = document.createElement("tr");
  row2.append(createEditableHeaderCell("th", "sf6-time", "sf6-time", "Time"));
  SF6_COLUMNS.forEach(() => {
    row2.append(createCell("td", "sf6-value", ""));
  });

  const row3 = document.createElement("tr");
  row3.append(createCell("td", "sf6-time", ""));
  SF6_COLUMNS.forEach(() => {
    row3.append(createCell("td", "sf6-value", ""));
  });

  const row4 = document.createElement("tr");
  row4.append(createCell("td", "sf6-time", ""));
  SF6_COLUMNS.forEach(() => {
    row4.append(createCell("td", "sf6-value", ""));
  });

  sf6Table.innerHTML = "";
  sf6Table.append(row1, row2, row3, row4);
}

function renderTable() {
  renderMainTable();
  renderSf6Table();
}

function updateDateReplica() {
  if (!state.date) {
    dateDay.textContent = "";
    dateMonth.textContent = "";
    dateYear.textContent = "";
    sheetDateLabel.textContent = t("noDateSelected");
    return;
  }

  const [year, month, day] = state.date.split("-");
  dateDay.textContent = day;
  dateMonth.textContent = month;
  dateYear.textContent = year;
}

function updateEntryPanel() {
  if (!state.date) {
    return;
  }

  const currentHour = getCurrentHour();
  const currentField = getCurrentField();
  const currentValue = state.entries[currentHour][currentField.key] || "";
  const completion = getHourCompletion(currentHour);

  sheetDateLabel.textContent = new Date(
    `${state.date}T00:00:00`,
  ).toLocaleDateString(undefined, {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  activeHourLabel.textContent = currentHour;
  fieldNameInput.value = getFieldDisplayLabel(currentField);
  fieldValueInput.value = currentValue;
  fieldValueInput.disabled = isSkippedField(currentField.key);
  saveFieldButton.disabled = isSkippedField(currentField.key);
  progressLabel.textContent = `Field ${state.activeFieldIndex + 1} of ${FLAT_FIELDS.length}`;
  completionLabel.textContent = `${Math.round((completion.filled / completion.total) * 100)}% complete`;
  fieldValueInput.focus();
  fieldValueInput.select();
  updateDateReplica();

  if (!isListening) {
    setVoiceStatus("You can type or use the microphone.");
  }

  if (isSkippedField(currentField.key)) {
    setVoiceStatus("This spot is skipped for the whole day.");
  }
}

function setActiveField(fieldIndex) {
  const boundedIndex = Math.max(
    0,
    Math.min(fieldIndex, FLAT_FIELDS.length - 1),
  );
  state.activeFieldIndex = boundedIndex;
  refreshUi();
}

function setActiveHour(index) {
  const boundedIndex = Math.max(0, Math.min(index, HOURS.length - 1));
  state.activeHourIndex = boundedIndex;
  const firstEmptyField = FLAT_FIELDS.findIndex((field) => {
    if (isSkippedField(field.key)) {
      return false;
    }
    const value = state.entries[getCurrentHour()][field.key];
    return String(value).trim() === "";
  });
  state.activeFieldIndex = firstEmptyField === -1 ? 0 : firstEmptyField;
  refreshUi();
}

function refreshUi() {
  renderTable();
  updateEntryPanel();
}

function openSheet(date, hourIndex) {
  state.date = date;
  state.entries = loadEntries(date);
  state.skippedFields = loadSkippedFields(date);
  saveStatus.textContent = t("ready");
  setActiveHour(hourIndex);
}

function moveField(step) {
  let nextIndex = state.activeFieldIndex + step;
  if (nextIndex < 0) {
    return;
  }

  while (
    nextIndex >= 0 &&
    nextIndex < FLAT_FIELDS.length &&
    isSkippedField(FLAT_FIELDS[nextIndex].key)
  ) {
    nextIndex += step;
  }

  if (nextIndex >= FLAT_FIELDS.length) {
    if (state.activeHourIndex < HOURS.length - 1) {
      setActiveHour(state.activeHourIndex + 1);
    }
    return;
  }

  state.activeFieldIndex = nextIndex;
  refreshUi();
}

function saveCurrentField() {
  const hour = getCurrentHour();
  const field = getCurrentField();
  const cleanedValue = normalizeValueInput(fieldValueInput.value);
  fieldValueInput.value = cleanedValue;
  state.entries[hour][field.key] = cleanedValue;
  saveEntries();
}

function renderSkipFieldList() {
  skipFieldList.innerHTML = "";

  FLAT_FIELDS.forEach((field) => {
    const label = document.createElement("label");
    label.className = "skip-option";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "skippedFields";
    checkbox.value = field.key;
    checkbox.checked = isSkippedField(field.key);

    const text = document.createElement("span");
    text.textContent = getFieldDisplayLabel(field);

    label.append(checkbox, text);
    skipFieldList.append(label);
  });
}

function openHeaderEditor(headerId, defaultText, defaultOrientation) {
  const config = getHeaderOverride(headerId, defaultText, defaultOrientation);
  state.editingHeaderId = headerId;
  headerTextInput.value = config.text;
  headerSubTextInput.value = config.subText;
  setHeaderOrientationSelection(config.orientation);
  headerDialog.showModal();
  window.setTimeout(() => {
    headerSubTextInput.focus();
    headerSubTextInput.select();
  }, 0);
}

function resetSheet() {
  if (!state.date) {
    return;
  }

  const confirmed = window.confirm(`Reset the whole sheet for ${state.date}?`);
  if (!confirmed) {
    return;
  }

  state.entries = createEmptyEntries();
  state.activeHourIndex = 0;
  state.activeFieldIndex = 0;
  saveEntries();
  refreshUi();
}

setupForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const date = sheetDateInput.value;
  const hourIndex = Number(startHourSelect.value);
  openSheet(date, hourIndex);
  setupDialog.close();
});

cancelSetupButton.addEventListener("click", () => {
  setupDialog.close();
});

openSetupButton.addEventListener("click", () => {
  sheetDateInput.value = state.date || getToday();
  startHourSelect.value = String(state.activeHourIndex || 0);
  setupDialog.showModal();
});

manageSkipsButton.addEventListener("click", () => {
  if (!state.date) {
    return;
  }

  renderSkipFieldList();
  skipDialog.showModal();
});

exportExcelButton.addEventListener("click", () => {
  if (!state.date) {
    saveStatus.textContent = "Choose a date first";
    return;
  }

  downloadExcelFile();
  saveStatus.textContent = "Excel exported";
});

shareSheetButton.addEventListener("click", () => {
  shareSheetFile();
});

cameraButton.addEventListener("click", () => {
  openCamera();
});

voiceButton.addEventListener("click", () => {
  if (!recognition) {
    return;
  }

  if (isListening) {
    recognition.stop();
    return;
  }

  startVoiceCapture();
});

saveFieldButton.addEventListener("click", () => {
  saveCurrentField();
  moveField(1);
  startVoiceCapture();
});

backFieldButton.addEventListener("click", () => {
  saveCurrentField();
  moveField(-1);
});

resetSheetButton.addEventListener("click", resetSheet);

cancelSkipButton.addEventListener("click", () => {
  skipDialog.close();
});

skipForm.addEventListener("submit", (event) => {
  event.preventDefault();
  state.skippedFields = Array.from(
    skipFieldList.querySelectorAll('input[name="skippedFields"]:checked'),
  ).map((input) => input.value);
  saveSkippedFields();
  setActiveHour(state.activeHourIndex);
  skipDialog.close();
});

fieldValueInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    saveFieldButton.click();
    return;
  }

  const allowedKeys = [
    "Backspace",
    "Delete",
    "Tab",
    "ArrowLeft",
    "ArrowRight",
    "ArrowUp",
    "ArrowDown",
    "Home",
    "End",
    "Enter",
  ];

  if (allowedKeys.includes(event.key) || event.ctrlKey || event.metaKey) {
    return;
  }

  if (!/[0-9.-]/.test(event.key)) {
    event.preventDefault();
  }
});

fieldValueInput.addEventListener("input", () => {
  const cleanedValue = normalizeValueInput(fieldValueInput.value);
  if (fieldValueInput.value !== cleanedValue) {
    fieldValueInput.value = cleanedValue;
  }
});

logTable.addEventListener("click", (event) => {
  const headerCell = event.target.closest("th[data-header-id]");
  if (headerCell) {
    openHeaderEditor(
      headerCell.dataset.headerId,
      headerCell.dataset.defaultText || "",
      headerCell.dataset.defaultOrientation || "horizontal",
    );
    return;
  }

  const cell = event.target.closest("td[data-hour-index][data-field-index]");
  if (cell) {
    state.activeHourIndex = Number(cell.dataset.hourIndex);
    setActiveField(Number(cell.dataset.fieldIndex));
    return;
  }

  const row = event.target.closest("tr[data-hour-index]");
  if (row) {
    setActiveHour(Number(row.dataset.hourIndex));
  }
});

sf6Table.addEventListener("click", (event) => {
  const headerCell = event.target.closest("th[data-header-id]");
  if (!headerCell) {
    return;
  }

  openHeaderEditor(
    headerCell.dataset.headerId,
    headerCell.dataset.defaultText || "",
    headerCell.dataset.defaultOrientation || "horizontal",
  );
});

cancelHeaderButton.addEventListener("click", () => {
  headerDialog.close();
});

closeCameraButton.addEventListener("click", () => {
  stopCamera();
  cameraDialog.close();
});

captureCameraButton.addEventListener("click", () => {
  captureCameraFrame();
});

retakeCameraButton.addEventListener("click", () => {
  if (!cameraStream) {
    openCamera();
    return;
  }

  cameraSnapshot.hidden = true;
  cameraVideo.hidden = false;
  cameraStatus.textContent = "Camera reopened. Capture again when ready.";
  cameraStatus.classList.remove("error");
});

resetHeaderButton.addEventListener("click", () => {
  if (!state.editingHeaderId) {
    return;
  }

  delete state.headerOverrides[state.editingHeaderId];
  saveHeaderOverrides();
  renderTable();
  headerDialog.close();
});

headerForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!state.editingHeaderId) {
    return;
  }

  state.headerOverrides[state.editingHeaderId] = {
    text: headerTextInput.value.trim(),
    subText: headerSubTextInput.value.trim(),
    orientation: selectedHeaderOrientation,
  };
  saveHeaderOverrides();
  renderTable();
  headerDialog.close();
});

headerHorizontalButton.addEventListener("click", () => {
  setHeaderOrientationSelection("horizontal");
});

headerVerticalButton.addEventListener("click", () => {
  setHeaderOrientationSelection("vertical");
});

cameraDialog.addEventListener("close", () => {
  stopCamera();
});

settingsButton.addEventListener("click", () => {
  languageSelect.value = state.language;
  settingsDialog.showModal();
});

closeSettingsButton.addEventListener("click", () => {
  settingsDialog.close();
});

settingsForm.addEventListener("submit", (event) => {
  event.preventDefault();
  state.language = languageSelect.value === "ar" ? "ar" : "en";
  saveSettings();
  applyLanguage();
  refreshUi();
  settingsDialog.close();
});

window.addEventListener("resize", syncMobileFieldLayout);

populateHourOptions();
state.language = loadSettings().language;
state.headerOverrides = loadHeaderOverrides();
setupVoiceRecognition();
applyLanguage();
syncMobileFieldLayout();
renderTable();
const initialDate = getToday();
sheetDateInput.value = initialDate;
startHourSelect.value = "0";
openSheet(initialDate, 0);
