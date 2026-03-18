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

const FIELD_INDEX_BY_KEY = new Map(
  FLAT_FIELDS.map((field, index) => [field.key, index]),
);
const FAST_INPUT_ALLOWED_GROUPS = new Set([
  ...L_GROUPS.map((group) => group.group),
  "FEEDERS 33 KV",
  "FEEDERS 11 KV",
]);
const FAST_INPUT_FIELDS = FLAT_FIELDS.filter((field) =>
  FAST_INPUT_ALLOWED_GROUPS.has(field.group),
);
const FAST_INPUT_FIELD_KEYS = new Set(
  FAST_INPUT_FIELDS.map((field) => field.key),
);

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
const FACILITIES_STORAGE_KEY = `${STORAGE_PREFIX}-facilities`;
const MONTHLY_STORAGE_KEY = `${STORAGE_PREFIX}-monthly-readings`;
const MONTHLY_PHOTO_STORAGE_KEY = `${STORAGE_PREFIX}-monthly-photos`;
const MONTHLY_FEEDER_GROUPS = [
  { key: "feeders-a", title: "مغذيات جهة 33", start: 1, count: 24 },
  { key: "feeders-b", title: "مغذيات جهة 11", start: 1, count: 24 },
];
const MONTHLY_RESIDENTIAL_GROUPS = [
  { key: "residential-a", start: 1, count: 5 },
  { key: "residential-b", start: 6, count: 5 },
];
const MONTHLY_LINE_GROUPS = [
  { key: "lines-a", lineNumbers: [1, 2, 3, 4] },
  { key: "lines-b", lineNumbers: [5, 6, 7, 8] },
];
const MONTHLY_TRANSFORMER_GROUPS = [
  {
    key: "transformers-33",
    title: "قراءات المحولات الرئيسية جهة 33 kV",
    labels: [
      "المحولة الرئيسية رقم / 1",
      "المحولة الرئيسية رقم / 2",
      "المحولة الرئيسية رقم / 3",
      "المحولة الرئيسية رقم / 4",
    ],
  },
  {
    key: "transformers-132",
    title: "قراءات المحولات الرئيسية جهة 132 kV",
    labels: [
      "المحولة الرئيسية رقم / 1",
      "المحولة الرئيسية رقم / 2",
      "المحولة الرئيسية رقم / 3",
      "المحولة الرئيسية رقم / 4",
    ],
  },
  {
    key: "service-transformers",
    title: "قراءات محولات الخدمة",
    labels: [
      "محولة الخدمة رقم / 1",
      "محولة الخدمة رقم / 2",
      "محولة الخدمة رقم / 3",
      "محولة الخدمة رقم / 4",
    ],
  },
  {
    key: "transformers-11",
    title: "قراءات المحولات الرئيسية جهة 11 kV",
    labels: [
      "المحولة الرئيسية رقم / 1",
      "المحولة الرئيسية رقم / 2",
      "المحولة الرئيسية رقم / 3",
      "المحولة الرئيسية رقم / 4",
    ],
  },
];
const UI_TEXT = {
  en: {
    homeTitle: "Logbook Hub",
    facilities: "Facilities",
    newFacility: "New Facility",
    next: "Next",
    workAreas: "Work Areas",
    gateway: "Daily Log Sheet",
    attendance: "Attendance",
    monthlyReadings: "Monthly Readings",
    addFacility: "Add Facility",
    facilityName: "Facility Name",
    saveFacility: "Save Facility",
    addOperator: "Add Operator",
    operatorName: "Operator Name",
    saveOperator: "Save Operator",
    operators: "Operators",
    noFacilities: "No facilities yet. Add your first facility.",
    noOperators: "No operators yet.",
    edit: "Edit",
    editFacility: "Edit Facility",
    saveChanges: "Save Changes",
    delete: "Delete",
    deleteFacility: "Delete Facility",
    backToFacilities: "Back To Facilities",
    loginTitle: "Log In",
    username: "Username",
    password: "Password",
    enter: "Enter",
    sheetDate: "Sheet Date",
    activeHour: "Active Hour",
    currentField: "Current Field",
    value: "Value",
    changeDate: "Change Date",
    resetSheet: "Reset Sheet",
    settings: "Settings",
    startTalking: "Start Talking",
    stopTalking: "Stop Talking",
    fastInput: "Fast Input",
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
    homeTitle: "مركز السجل",
    facilities: "المنشآت",
    newFacility: "منشأة جديدة",
    addFacility: "إضافة منشأة",
    facilityName: "اسم المنشأة",
    saveFacility: "حفظ المنشأة",
    addOperator: "إضافة مشغل",
    operatorName: "اسم المشغل",
    saveOperator: "حفظ المشغل",
    operators: "المشغلون",
    noFacilities: "لا توجد منشآت بعد. أضف أول منشأة.",
    noOperators: "لا يوجد مشغلون بعد.",
    edit: "تعديل",
    editFacility: "تعديل المنشأة",
    saveChanges: "حفظ التغييرات",
    delete: "حذف",
    deleteFacility: "حذف المنشأة",
    backToFacilities: "العودة إلى المنشآت",
    sheetDate: "تاريخ الورقة",
    activeHour: "الساعة الحالية",
    currentField: "الحقل الحالي",
    value: "القيمة",
    changeDate: "تغيير التاريخ",
    resetSheet: "إعادة ضبط الورقة",
    settings: "الإعدادات",
    startTalking: "ابدأ التحدث",
    stopTalking: "إيقاف التحدث",
    fastInput: "ادخال سريع",
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
  facilities: [],
  selectedFacilityId: "",
  selectedOperatorId: "",
  editingFacilityId: "",
  editingFacilityDialogId: "",
  monthlyPeriod: "",
  monthlyEntries: {},
  monthlyPhotos: {},
  monthlyCameraMode: false,
  pendingMonthlyPhotoField: "",
};

const homeScreen = document.getElementById("homeScreen");
const loginScreen = document.getElementById("loginScreen");
const moduleScreen = document.getElementById("moduleScreen");
const monthlyScreen = document.getElementById("monthlyScreen");
const appShell = document.getElementById("appShell");
const homeTitle = document.getElementById("homeTitle");
const facilitiesHeading = document.getElementById("facilitiesHeading");
const homeLanguageButton = document.getElementById("homeLanguageButton");
const facilityList = document.getElementById("facilityList");
const homeNextButton = document.getElementById("homeNextButton");
const moduleTitle = document.getElementById("moduleTitle");
const gatewayButton = document.getElementById("gatewayButton");
const attendanceButton = document.getElementById("attendanceButton");
const monthlyReadingsButton = document.getElementById("monthlyReadingsButton");
const monthlyBackButton = document.getElementById("monthlyBackButton");
const monthlyCameraButton = document.getElementById("monthlyCameraButton");
const monthlyExportImageButton = document.getElementById(
  "monthlyExportImageButton",
);
const monthlyCameraInput = document.getElementById("monthlyCameraInput");
const monthlyMonthInput = document.getElementById("monthlyMonthInput");
const monthlyMonthText = document.getElementById("monthlyMonthText");
const monthlyFacilityText = document.getElementById("monthlyFacilityText");
const monthlyOperatorText = document.getElementById("monthlyOperatorText");
const monthlyFacilityLabel = document.getElementById("monthlyFacilityLabel");
const monthlyOperatorLabel = document.getElementById("monthlyOperatorLabel");
const monthlySaveStatus = document.getElementById("monthlySaveStatus");
const monthlyPages = document.getElementById("monthlyPages");
const gatewayTitle = document.getElementById("gatewayTitle");
const attendanceTitle = document.getElementById("attendanceTitle");
const monthlyReadingsTitle = document.getElementById("monthlyReadingsTitle");
const loginTitle = document.getElementById("loginTitle");
const loginForm = document.getElementById("loginForm");
const loginUsernameText = document.getElementById("loginUsernameText");
const loginPasswordText = document.getElementById("loginPasswordText");
const loginUsernameInput = document.getElementById("loginUsernameInput");
const loginPasswordInput = document.getElementById("loginPasswordInput");
const enterFacilitiesButton = document.getElementById("enterFacilitiesButton");
const openFacilityDialogButton = document.getElementById(
  "openFacilityDialogButton",
);
const facilityDialog = document.getElementById("facilityDialog");
const facilityForm = document.getElementById("facilityForm");
const facilityNameInput = document.getElementById("facilityNameInput");
const cancelFacilityButton = document.getElementById("cancelFacilityButton");
const operatorDialog = document.getElementById("operatorDialog");
const operatorForm = document.getElementById("operatorForm");
const operatorNameInput = document.getElementById("operatorNameInput");
const cancelOperatorButton = document.getElementById("cancelOperatorButton");
const editFacilityDialog = document.getElementById("editFacilityDialog");
const editFacilityForm = document.getElementById("editFacilityForm");
const editFacilityNameInput = document.getElementById("editFacilityNameInput");
const editOperatorList = document.getElementById("editOperatorList");
const cancelEditFacilityButton = document.getElementById(
  "cancelEditFacilityButton",
);
const deleteFacilityButton = document.getElementById("deleteFacilityButton");
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
const backToHomeButton = document.getElementById("backToHomeButton");
const resetHeaderButton = document.getElementById("resetHeaderButton");
const fastInputButton = document.getElementById("fastInputButton");
const saveFieldButton = document.getElementById("saveFieldButton");
const exportExcelButton = document.getElementById("exportExcelButton");
const exportImageButton = document.getElementById("exportImageButton");
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
const sheetStage = document.querySelector(".sheet-stage");
const sheetTopScroller = document.getElementById("sheetTopScroller");
const sheetTopScrollerInner = document.getElementById("sheetTopScrollerInner");
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
const fastInputDialog = document.getElementById("fastInputDialog");
const fastInputForm = document.getElementById("fastInputForm");
const fastInputFieldName = document.getElementById("fastInputFieldName");
const fastInputValue = document.getElementById("fastInputValue");
const closeFastInputButton = document.getElementById("closeFastInputButton");
const fastInputBackButton = document.getElementById("fastInputBackButton");
const fastInputSaveButton = document.getElementById("fastInputSaveButton");
const settingsDialog = document.getElementById("settingsDialog");
const settingsForm = document.getElementById("settingsForm");
const settingsTitle = document.getElementById("settingsTitle");
const languageLabel = document.getElementById("languageLabel");
const languageSelect = document.getElementById("languageSelect");
const closeSettingsButton = document.getElementById("closeSettingsButton");
const saveSettingsButton = document.getElementById("saveSettingsButton");
const settingsChangeDateButton = document.getElementById(
  "settingsChangeDateButton",
);
const settingsResetSheetButton = document.getElementById(
  "settingsResetSheetButton",
);
const settingsSkipSpotsButton = document.getElementById(
  "settingsSkipSpotsButton",
);
const SpeechRecognitionApi =
  window.SpeechRecognition || window.webkitSpeechRecognition;
const currentFieldGroup = currentFieldText.closest(".control-group");
const infoRow = currentFieldGroup.closest(".info-row");
const actionRow = document.querySelector(".action-row");
const valueGroup = document.querySelector(".value-group");

let recognition = null;
let isListening = false;
let selectedHeaderOrientation = "horizontal";
let currentFieldInActionRow = false;
let syncingSheetScroll = false;

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

function loadFacilities() {
  const saved = localStorage.getItem(FACILITIES_STORAGE_KEY);
  if (!saved) {
    return [];
  }

  try {
    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .filter((facility) => facility && typeof facility.name === "string")
      .map((facility) => ({
        id:
          facility.id ||
          `facility-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        name: facility.name.trim(),
        operators: Array.isArray(facility.operators)
          ? facility.operators
              .filter(
                (operator) => operator && typeof operator.name === "string",
              )
              .map((operator) => ({
                id:
                  operator.id ||
                  `operator-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
                name: operator.name.trim(),
              }))
          : [],
      }));
  } catch {
    return [];
  }
}

function saveFacilities() {
  localStorage.setItem(
    FACILITIES_STORAGE_KEY,
    JSON.stringify(state.facilities),
  );
}

function createId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function renderFacilities() {
  facilityList.innerHTML = "";
  homeNextButton.disabled =
    !state.selectedFacilityId || !state.selectedOperatorId;

  if (state.facilities.length === 0) {
    const empty = document.createElement("p");
    empty.className = "facility-empty";
    empty.textContent = t("noFacilities");
    facilityList.append(empty);
    return;
  }

  state.facilities.forEach((facility) => {
    const card = document.createElement("article");
    card.className = "facility-card";

    const header = document.createElement("div");
    header.className = "facility-card-header";

    const title = document.createElement("h3");
    title.textContent = facility.name;

    const addOperatorButton = document.createElement("button");
    addOperatorButton.type = "button";
    addOperatorButton.className = "ghost-button";
    addOperatorButton.textContent = t("addOperator");
    addOperatorButton.dataset.facilityId = facility.id;

    const editFacilityButton = document.createElement("button");
    editFacilityButton.type = "button";
    editFacilityButton.className = "ghost-button";
    editFacilityButton.textContent = t("edit");
    editFacilityButton.dataset.editFacilityId = facility.id;

    const actionGroup = document.createElement("div");
    actionGroup.className = "facility-card-actions";
    actionGroup.append(addOperatorButton, editFacilityButton);

    header.append(title, actionGroup);

    const operatorsTitle = document.createElement("p");
    operatorsTitle.className = "facility-subtitle";
    operatorsTitle.textContent = t("operators");

    const operatorsList = document.createElement("div");
    operatorsList.className = "operator-list";

    if (facility.operators.length === 0) {
      const emptyOperators = document.createElement("p");
      emptyOperators.className = "operator-empty";
      emptyOperators.textContent = t("noOperators");
      operatorsList.append(emptyOperators);
    } else {
      facility.operators.forEach((operator) => {
        const operatorButton = document.createElement("button");
        operatorButton.type = "button";
        operatorButton.className = "operator-chip";
        operatorButton.textContent = operator.name;
        operatorButton.dataset.facilityId = facility.id;
        operatorButton.dataset.operatorId = operator.id;
        if (
          state.selectedFacilityId === facility.id &&
          state.selectedOperatorId === operator.id
        ) {
          operatorButton.classList.add("selected");
        }
        operatorsList.append(operatorButton);
      });
    }

    card.append(header, operatorsTitle, operatorsList);
    facilityList.append(card);
  });
}

function openFacilityEditor(facilityId) {
  const facility = state.facilities.find((item) => item.id === facilityId);
  if (!facility) {
    return;
  }

  state.editingFacilityDialogId = facilityId;
  editFacilityNameInput.value = facility.name;
  editOperatorList.innerHTML = "";

  if (facility.operators.length === 0) {
    const empty = document.createElement("p");
    empty.className = "operator-empty";
    empty.textContent = t("noOperators");
    editOperatorList.append(empty);
  } else {
    facility.operators.forEach((operator) => {
      const row = document.createElement("div");
      row.className = "edit-operator-row";
      row.dataset.operatorId = operator.id;

      const input = document.createElement("input");
      input.type = "text";
      input.value = operator.name;
      input.required = true;
      input.dataset.operatorNameInput = "true";

      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "ghost-button danger-button";
      deleteButton.textContent = t("delete");
      deleteButton.dataset.deleteOperatorId = operator.id;

      row.append(input, deleteButton);
      editOperatorList.append(row);
    });
  }

  editFacilityDialog.showModal();
  window.setTimeout(() => {
    editFacilityNameInput.focus();
    editFacilityNameInput.select();
  }, 0);
}

function showHomeScreen() {
  homeScreen.classList.remove("hidden");
  loginScreen.classList.add("hidden");
  moduleScreen.classList.add("hidden");
  monthlyScreen.classList.add("hidden");
  appShell.classList.add("hidden");
}

function showLoginScreen() {
  homeScreen.classList.add("hidden");
  loginScreen.classList.remove("hidden");
  moduleScreen.classList.add("hidden");
  monthlyScreen.classList.add("hidden");
  appShell.classList.add("hidden");
}

function showModuleScreen() {
  homeScreen.classList.add("hidden");
  loginScreen.classList.add("hidden");
  moduleScreen.classList.remove("hidden");
  monthlyScreen.classList.add("hidden");
  appShell.classList.add("hidden");
}

function showAppShell() {
  homeScreen.classList.add("hidden");
  loginScreen.classList.add("hidden");
  moduleScreen.classList.add("hidden");
  monthlyScreen.classList.add("hidden");
  appShell.classList.remove("hidden");
}

function showMonthlyScreen() {
  homeScreen.classList.add("hidden");
  loginScreen.classList.add("hidden");
  moduleScreen.classList.add("hidden");
  monthlyScreen.classList.remove("hidden");
  appShell.classList.add("hidden");
}

function getSelectedFacility() {
  return (
    state.facilities.find((facility) => facility.id === state.selectedFacilityId) ||
    null
  );
}

function getSelectedOperator() {
  const facility = getSelectedFacility();
  return (
    facility?.operators?.find((operator) => operator.id === state.selectedOperatorId) ||
    null
  );
}

function selectOperator(facilityId, operatorId) {
  state.selectedFacilityId = facilityId;
  state.selectedOperatorId = operatorId;
  renderFacilities();
}

function openSelectedOperatorSheet() {
  if (!state.selectedFacilityId || !state.selectedOperatorId) {
    return;
  }

  showAppShell();
  const initialDate = getToday();
  sheetDateInput.value = initialDate;
  startHourSelect.value = "0";
  openSheet(initialDate, 0);
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
  homeTitle.textContent = t("homeTitle");
  facilitiesHeading.textContent = t("facilities");
  openFacilityDialogButton.textContent = t("newFacility");
  homeNextButton.textContent = state.language === "ar" ? "التالي" : t("next");
  homeLanguageButton.textContent = state.language === "ar" ? "EN" : "ع";
  homeLanguageButton.setAttribute("aria-label", t("language"));
  homeLanguageButton.setAttribute("title", t("language"));
  moduleTitle.textContent =
    state.language === "ar" ? "مساحات العمل" : t("workAreas");
  gatewayTitle.textContent =
    state.language === "ar" ? "سجل اليوم" : t("gateway");
  attendanceTitle.textContent =
    state.language === "ar" ? "الحضور" : t("attendance");
  monthlyReadingsTitle.textContent =
    state.language === "ar" ? "القراءات الشهرية" : t("monthlyReadings");
  monthlyBackButton.textContent = t("back");
  monthlyMonthText.textContent = state.language === "ar" ? "الشهر" : "Month";
  monthlyFacilityText.textContent =
    state.language === "ar" ? "المنشأة" : "Facility";
  monthlyOperatorText.textContent =
    state.language === "ar" ? "المشغل" : "Operator";
  monthlyExportImageButton.textContent =
    state.language === "ar" ? "تصدير صورة" : "Export Photo";
  loginTitle.textContent =
    state.language === "ar" ? "تسجيل الدخول" : t("loginTitle");
  loginUsernameText.textContent =
    state.language === "ar" ? "اسم المستخدم" : t("username");
  loginPasswordText.textContent =
    state.language === "ar" ? "كلمة المرور" : t("password");
  enterFacilitiesButton.textContent =
    state.language === "ar" ? "دخول" : t("enter");
  sheetDateText.textContent = t("sheetDate");
  activeHourText.textContent = t("activeHour");
  currentFieldText.textContent = t("currentField");
  valueText.textContent = t("value");
  openSetupButton.textContent = t("changeDate");
  resetSheetButton.textContent = t("resetSheet");
  voiceButton.textContent = isListening ? t("stopTalking") : t("startTalking");
  fastInputButton.textContent = t("fastInput");
  backFieldButton.textContent = t("back");
  saveFieldButton.textContent = t("saveAndNext");
  manageSkipsButton.textContent = t("skipSpots");
  exportExcelButton.textContent = t("exportExcel");
  exportImageButton.textContent =
    state.language === "ar" ? "تصدير صورة" : "Export Photo";
  shareSheetButton.textContent = t("shareSheet");
  backToHomeButton.textContent = t("backToFacilities");
  settingsTitle.textContent = t("settingsTitle");
  languageLabel.textContent = t("language");
  closeSettingsButton.textContent = t("close");
  saveSettingsButton.textContent = t("saveSettings");
  settingsChangeDateButton.textContent = t("changeDate");
  settingsResetSheetButton.textContent = t("resetSheet");
  settingsSkipSpotsButton.textContent = t("skipSpots");
  fieldValueInput.placeholder = t("valuePlaceholder");
  fastInputValue.placeholder = t("valuePlaceholder");
  facilityDialog.querySelector("h2").textContent = t("addFacility");
  facilityDialog.querySelector('label[for="facilityNameInput"]').textContent =
    t("facilityName");
  facilityDialog.querySelector('.primary-button[type="submit"]').textContent =
    t("saveFacility");
  operatorDialog.querySelector("h2").textContent = t("addOperator");
  operatorDialog.querySelector('label[for="operatorNameInput"]').textContent =
    t("operatorName");
  operatorDialog.querySelector('.primary-button[type="submit"]').textContent =
    t("saveOperator");
  editFacilityDialog.querySelector("h2").textContent = t("editFacility");
  editFacilityDialog.querySelector(
    'label[for="editFacilityNameInput"]',
  ).textContent = t("facilityName");
  editFacilityDialog.querySelector(
    ".edit-operators-block .strip-label",
  ).textContent = t("operators");
  deleteFacilityButton.textContent = t("deleteFacility");
  editFacilityDialog.querySelector(
    '.primary-button[type="submit"]',
  ).textContent = t("saveChanges");
  renderFacilities();
  updateMonthlyHeader();

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

function getCurrentMonth() {
  return getToday().slice(0, 7);
}

function loadAllMonthlyEntries() {
  const saved = localStorage.getItem(MONTHLY_STORAGE_KEY);
  if (!saved) {
    return {};
  }

  try {
    return JSON.parse(saved);
  } catch {
    return {};
  }
}

function loadAllMonthlyPhotos() {
  const saved = localStorage.getItem(MONTHLY_PHOTO_STORAGE_KEY);
  if (!saved) {
    return {};
  }

  try {
    return JSON.parse(saved);
  } catch {
    return {};
  }
}

function getMonthlyStorageBucketKey() {
  return [
    state.selectedFacilityId || "facility",
    state.selectedOperatorId || "operator",
    state.monthlyPeriod || getCurrentMonth(),
  ].join("__");
}

function loadMonthlyEntriesForCurrentSelection() {
  const allEntries = loadAllMonthlyEntries();
  return allEntries[getMonthlyStorageBucketKey()] || {};
}

function saveMonthlyEntries() {
  const allEntries = loadAllMonthlyEntries();
  allEntries[getMonthlyStorageBucketKey()] = state.monthlyEntries;
  localStorage.setItem(MONTHLY_STORAGE_KEY, JSON.stringify(allEntries));
  monthlySaveStatus.textContent =
    state.language === "ar" ? "تم حفظ القراءات الشهرية" : "Monthly readings saved";
}

function loadMonthlyPhotosForCurrentSelection() {
  const allPhotos = loadAllMonthlyPhotos();
  return allPhotos[getMonthlyStorageBucketKey()] || {};
}

function saveMonthlyPhotos() {
  const allPhotos = loadAllMonthlyPhotos();
  allPhotos[getMonthlyStorageBucketKey()] = state.monthlyPhotos;
  localStorage.setItem(MONTHLY_PHOTO_STORAGE_KEY, JSON.stringify(allPhotos));
  monthlySaveStatus.textContent =
    state.language === "ar"
      ? "تم حفظ صورة الحقل"
      : "Field photo saved";
}

function formatMonthlyPeriodLabel() {
  const period = state.monthlyPeriod || getCurrentMonth();
  const [year, month] = period.split("-");
  return `${month} / ${year}`;
}

function createMonthlyFieldKey(...segments) {
  return segments.join("__");
}

function isMonthlyNextReadingField(fieldKey) {
  return fieldKey.endsWith("__next");
}

function createMonthlyInput(fieldKey, options = {}) {
  const input = document.createElement("input");
  input.type = "text";
  input.className = "monthly-cell-input";
  input.dataset.monthlyField = fieldKey;
  if (isMonthlyNextReadingField(fieldKey)) {
    input.dataset.photoField = fieldKey;
  }

  if (options.numeric) {
    input.inputMode = "decimal";
    input.dataset.monthlyType = "number";
  }

  return input;
}

function createMonthlyInputCell(fieldKey, options = {}) {
  const cell = document.createElement("td");
  cell.append(createMonthlyInput(fieldKey, options));
  return cell;
}

function createMonthlyNumberedTable(config) {
  const card = document.createElement("section");
  card.className = "monthly-card";

  const title = document.createElement("div");
  title.className = "monthly-card-title";
  title.textContent = config.title;
  card.append(title);

  const table = document.createElement("table");
  table.className = "monthly-table";

  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  [
    config.numberLabel,
    config.nameLabel,
    "القراءة السابقة",
    "القراءة اللاحقة",
  ].forEach((label) => {
    const cell = document.createElement("th");
    cell.textContent = label;
    headRow.append(cell);
  });
  head.append(headRow);

  const body = document.createElement("tbody");
  for (let index = 0; index < config.count; index += 1) {
    const rowNumber = config.start + index;
    const row = document.createElement("tr");

    const numberCell = document.createElement("td");
    numberCell.className = "monthly-seq-cell";
    numberCell.textContent = config.blankNumbers ? "" : String(rowNumber);
    row.append(numberCell);
    row.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(config.key, rowNumber, "name"),
      ),
    );
    row.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(config.key, rowNumber, "previous"),
        { numeric: true },
      ),
    );
    row.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(config.key, rowNumber, "next"),
        { numeric: true },
      ),
    );
    body.append(row);
  }

  table.append(head, body);
  card.append(table);
  return card;
}

function createMonthlySectionTitle(text) {
  const title = document.createElement("div");
  title.className = "monthly-section-title";
  title.textContent = text;
  return title;
}

function createMonthlyLineReadingsTable(group) {
  const table = document.createElement("table");
  table.className = "monthly-table monthly-line-table";

  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["اسم الخط", "", "القراءة السابقة", "القراءة اللاحقة"].forEach((label) => {
    const cell = document.createElement("th");
    cell.textContent = label;
    headRow.append(cell);
  });
  head.append(headRow);

  const body = document.createElement("tbody");
  group.lineNumbers.forEach((lineNumber) => {
    const impRow = document.createElement("tr");
    const expRow = document.createElement("tr");

    const lineNameCell = document.createElement("td");
    lineNameCell.className = "monthly-line-name-cell";
    lineNameCell.rowSpan = 2;
    lineNameCell.append(
      createMonthlyInput(createMonthlyFieldKey(group.key, lineNumber, "name")),
    );
    impRow.append(lineNameCell);

    const impTypeCell = document.createElement("td");
    impTypeCell.className = "monthly-line-type";
    impTypeCell.textContent = "IMP";
    impRow.append(impTypeCell);
    impRow.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, lineNumber, "imp", "previous"),
        { numeric: true },
      ),
    );
    impRow.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, lineNumber, "imp", "next"),
        { numeric: true },
      ),
    );

    const expTypeCell = document.createElement("td");
    expTypeCell.className = "monthly-line-type";
    expTypeCell.textContent = "EXP";
    expRow.append(expTypeCell);
    expRow.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, lineNumber, "exp", "previous"),
        { numeric: true },
      ),
    );
    expRow.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, lineNumber, "exp", "next"),
        { numeric: true },
      ),
    );

    body.append(impRow, expRow);
  });

  table.append(head, body);
  return table;
}

function createMonthlyTransformerTable(group) {
  const card = document.createElement("section");
  card.className = "monthly-card";

  const title = document.createElement("div");
  title.className = "monthly-card-title";
  title.textContent = group.title;
  card.append(title);

  const table = document.createElement("table");
  table.className = "monthly-table";

  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["رقم المحولة", "القراءة السابقة", "القراءة اللاحقة"].forEach((label) => {
    const cell = document.createElement("th");
    cell.textContent = label;
    headRow.append(cell);
  });
  head.append(headRow);

  const body = document.createElement("tbody");
  group.labels.forEach((label, index) => {
    const row = document.createElement("tr");

    const nameCell = document.createElement("td");
    nameCell.className = "monthly-label-cell";
    nameCell.textContent = label;
    row.append(nameCell);
    row.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, index + 1, "previous"),
        { numeric: true },
      ),
    );
    row.append(
      createMonthlyInputCell(
        createMonthlyFieldKey(group.key, index + 1, "next"),
        { numeric: true },
      ),
    );

    body.append(row);
  });

  table.append(head, body);
  card.append(table);
  return card;
}

function createMonthlyMetaSection() {
  const section = document.createElement("section");
  section.className = "monthly-meta";

  const notesRow = document.createElement("div");
  notesRow.className = "monthly-notes-row";

  const notesLabel = document.createElement("span");
  notesLabel.className = "monthly-notes-label";
  notesLabel.textContent = "الملاحظات:";

  const notesInput = document.createElement("textarea");
  notesInput.className = "monthly-notes-input";
  notesInput.rows = 3;
  notesInput.dataset.monthlyField = createMonthlyFieldKey("meta", "notes");

  notesRow.append(notesLabel, notesInput);

  const footerGrid = document.createElement("div");
  footerGrid.className = "monthly-footer-grid";

  ["left", "right"].forEach((side) => {
    const block = document.createElement("div");
    block.className = "monthly-footer-block";

    [
      ["اسم المراقب:", "observer"],
      ["التاريخ:", "date"],
      ["الوقت:", "time"],
    ].forEach(([label, key]) => {
      const row = document.createElement("label");
      row.className = "monthly-footer-row";

      const text = document.createElement("span");
      text.textContent = label;

      const input = createMonthlyInput(createMonthlyFieldKey("meta", side, key));
      row.append(text, input);
      block.append(row);
    });

    footerGrid.append(block);
  });

  section.append(notesRow, footerGrid);
  return section;
}

function renderMonthlyReadingsSheet() {
  monthlyPages.innerHTML = "";

  const pageOne = document.createElement("section");
  pageOne.className = "monthly-page";
  pageOne.dir = "rtl";

  const pageOneHeader = document.createElement("div");
  pageOneHeader.className = "monthly-page-header";
  pageOneHeader.innerHTML =
    '<span class="monthly-page-code">202</span><span class="monthly-page-date">/ /</span><span class="monthly-page-month">الشهر <strong data-monthly-period-label></strong></span>';
  pageOne.append(pageOneHeader);

  const feederGrid = document.createElement("div");
  feederGrid.className = "monthly-dual-grid";
  MONTHLY_FEEDER_GROUPS.forEach((group) => {
    feederGrid.append(
      createMonthlyNumberedTable({
        ...group,
        numberLabel: "رقم المغذي",
        nameLabel: "اسم المغذي",
      }),
    );
  });
  pageOne.append(feederGrid);

  pageOne.append(createMonthlySectionTitle("قراءات الدور السكنية"));

  const residentialGrid = document.createElement("div");
  residentialGrid.className = "monthly-dual-grid monthly-dual-grid--compact";
  MONTHLY_RESIDENTIAL_GROUPS.forEach((group) => {
    residentialGrid.append(
      createMonthlyNumberedTable({
        ...group,
        title: "قراءات الدور السكنية",
        numberLabel: "رقم الدار",
        nameLabel: "اسم شاغل الدار",
        blankNumbers: true,
      }),
    );
  });
  pageOne.append(residentialGrid);

  const pageTwo = document.createElement("section");
  pageTwo.className = "monthly-page";
  pageTwo.dir = "rtl";

  const pageTwoHeader = document.createElement("div");
  pageTwoHeader.className = "monthly-sheet-heading";
  pageTwoHeader.textContent = "قراءات مقاييس الطاقة لمحطة";
  pageTwo.append(pageTwoHeader);

  const lineSection = document.createElement("section");
  lineSection.className = "monthly-lines-section";
  lineSection.append(createMonthlySectionTitle("قراءات الخطوط"));

  const lineGrid = document.createElement("div");
  lineGrid.className = "monthly-dual-grid monthly-dual-grid--lines";
  MONTHLY_LINE_GROUPS.forEach((group) => {
    lineGrid.append(createMonthlyLineReadingsTable(group));
  });
  lineSection.append(lineGrid);
  pageTwo.append(lineSection);

  const transformerGrid = document.createElement("div");
  transformerGrid.className = "monthly-transformer-grid";
  MONTHLY_TRANSFORMER_GROUPS.forEach((group) => {
    transformerGrid.append(createMonthlyTransformerTable(group));
  });
  pageTwo.append(transformerGrid);
  pageTwo.append(createMonthlyMetaSection());

  monthlyPages.append(pageTwo, pageOne);
  populateMonthlyInputs();
  updateMonthlyHeader();
}

function populateMonthlyInputs() {
  monthlyPages.querySelectorAll("[data-monthly-field]").forEach((input) => {
    const key = input.dataset.monthlyField;
    input.value = state.monthlyEntries[key] || "";
  });
  applyMonthlyPhotoPreviews();
}

function applyMonthlyPhotoPreviews() {
  monthlyPages.querySelectorAll("[data-photo-field]").forEach((field) => {
    const photo = state.monthlyPhotos[field.dataset.photoField];
    if (photo) {
      field.classList.add("photo-attached");
      field.style.setProperty(
        "--monthly-photo",
        `linear-gradient(rgba(17, 17, 17, 0.28), rgba(17, 17, 17, 0.28)), url("${photo}")`,
      );
      return;
    }

    field.classList.remove("photo-attached");
    field.style.removeProperty("--monthly-photo");
  });
}

function updateMonthlyHeader() {
  if (!state.monthlyPeriod) {
    state.monthlyPeriod = getCurrentMonth();
  }

  const facilityName = getSelectedFacility()?.name || "-";

  monthlyMonthInput.value = state.monthlyPeriod;
  monthlyFacilityLabel.textContent = facilityName;
  monthlyOperatorLabel.textContent = getSelectedOperator()?.name || "-";
  monthlyCameraButton.classList.toggle("active", state.monthlyCameraMode);
  monthlyCameraButton.textContent =
    state.language === "ar"
      ? state.monthlyCameraMode
        ? "الكاميرا مفعلة"
        : "الكاميرا"
      : state.monthlyCameraMode
        ? "Camera On"
        : "Camera";
  document
    .querySelectorAll("[data-monthly-period-label]")
    .forEach((node) => (node.textContent = formatMonthlyPeriodLabel()));
  const monthlyHeading = monthlyPages.querySelector(".monthly-sheet-heading");
  if (monthlyHeading) {
    monthlyHeading.textContent = `قراءات مقاييس الطاقة لمحطة ${facilityName}`;
  }
}

function openMonthlyReadingsScreen() {
  if (!state.selectedFacilityId || !state.selectedOperatorId) {
    return;
  }

  state.monthlyPeriod =
    monthlyMonthInput.value || state.monthlyPeriod || getCurrentMonth();
  state.monthlyEntries = loadMonthlyEntriesForCurrentSelection();
  state.monthlyPhotos = loadMonthlyPhotosForCurrentSelection();
  renderMonthlyReadingsSheet();
  monthlySaveStatus.textContent = t("ready");
  showMonthlyScreen();
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

function drawCenteredText(context, text, rect, font, color) {
  if (!text) {
    return;
  }

  context.save();
  context.font = font;
  context.fillStyle = color;
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText(text, rect.x + rect.width / 2, rect.y + rect.height / 2);
  context.restore();
}

function drawVerticalText(context, text, rect, font, color) {
  if (!text) {
    return;
  }

  context.save();
  context.font = font;
  context.fillStyle = color;
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.translate(rect.x + rect.width / 2, rect.y + rect.height / 2);
  context.rotate(-Math.PI / 2);
  context.fillText(text, 0, 0);
  context.restore();
}

function drawCellContent(context, cell, paperRect, scale) {
  const rect = cell.getBoundingClientRect();
  const x = (rect.left - paperRect.left) * scale;
  const y = (rect.top - paperRect.top) * scale;
  const width = rect.width * scale;
  const height = rect.height * scale;
  const computed = window.getComputedStyle(cell);

  context.fillStyle = computed.backgroundColor || "#ffffff";
  context.fillRect(x, y, width, height);
  context.strokeStyle = computed.borderTopColor || "#444444";
  context.lineWidth = 1;
  context.strokeRect(x, y, width, height);

  const mixedHeader = cell.querySelector(".mixed-header-stack");
  if (mixedHeader) {
    const mainLabel = mixedHeader.querySelector(".main-vertical-label");
    const verticalSubLabel = mixedHeader.querySelector(".sub-label-vertical");
    const horizontalSubLabel = mixedHeader.querySelector(
      ".sub-label-horizontal",
    );

    if (horizontalSubLabel) {
      drawVerticalText(
        context,
        mainLabel?.textContent?.trim() || "",
        {
          x,
          y,
          width,
          height: height * 0.62,
        },
        `${Math.max(10, Math.round(11 * scale))}px Arial`,
        computed.color || "#111111",
      );

      drawCenteredText(
        context,
        horizontalSubLabel.textContent.trim(),
        {
          x,
          y: y + height * 0.62,
          width,
          height: height * 0.32,
        },
        `700 ${Math.max(9, Math.round(9 * scale))}px Arial`,
        computed.color || "#111111",
      );
      return;
    }

    const labels = [mainLabel, verticalSubLabel].filter(Boolean);
    const sectionHeight = height / Math.max(labels.length, 1);
    labels.forEach((node, index) => {
      drawVerticalText(
        context,
        node.textContent.trim(),
        {
          x,
          y: y + index * sectionHeight,
          width,
          height: sectionHeight,
        },
        `${Math.max(10, Math.round(11 * scale))}px Arial`,
        computed.color || "#111111",
      );
    });
    return;
  }

  const verticalWords = Array.from(cell.querySelectorAll(".vertical-word"));
  if (verticalWords.length > 0) {
    const sectionHeight = height / verticalWords.length;
    verticalWords.forEach((node, index) => {
      drawVerticalText(
        context,
        node.textContent.trim(),
        {
          x,
          y: y + index * sectionHeight,
          width,
          height: sectionHeight,
        },
        `${Math.max(10, Math.round(11 * scale))}px Arial`,
        computed.color || "#111111",
      );
    });
    return;
  }

  const headerText = cell.querySelector(".header-text");
  if (headerText) {
    const mainText =
      headerText.firstElementChild?.textContent?.trim() ||
      headerText.textContent.trim();
    const subText =
      headerText.querySelector(".header-subtext")?.textContent?.trim() || "";

    if (subText) {
      drawCenteredText(
        context,
        mainText,
        { x, y, width, height: height * 0.52 },
        `700 ${Math.max(10, Math.round(11 * scale))}px Arial`,
        computed.color || "#111111",
      );

      drawCenteredText(
        context,
        subText,
        { x, y: y + height * 0.48, width, height: height * 0.42 },
        `700 ${Math.max(9, Math.round(9 * scale))}px Arial`,
        computed.color || "#111111",
      );
      return;
    }
  }

  const namesBox = cell.querySelector(".names-sheet-box");
  if (namesBox) {
    const nameItems = Array.from(namesBox.querySelectorAll(".sheet-name-item"));
    const availableHeight = Math.max(height - 8 * scale, 0);
    const rowHeight =
      nameItems.length > 0
        ? availableHeight / nameItems.length
        : availableHeight;

    nameItems.forEach((item, index) => {
      const itemFontWeight = item.classList.contains("selected")
        ? "700"
        : "400";
      drawCenteredText(
        context,
        item.textContent.trim(),
        {
          x,
          y: y + index * rowHeight,
          width,
          height: rowHeight,
        },
        `${itemFontWeight} ${Math.max(10, Math.round(11 * scale))}px Arial`,
        computed.color || "#111111",
      );
    });
    return;
  }

  drawCenteredText(
    context,
    cell.textContent.trim(),
    { x, y, width, height },
    `${computed.fontWeight || "400"} ${Math.max(10, Math.round(parseFloat(computed.fontSize) * scale))}px Arial`,
    computed.color || "#111111",
  );
}

function drawHeaderText(context, element, paperRect, scale) {
  const text = element.textContent.trim();
  if (!text) {
    return;
  }

  const rect = element.getBoundingClientRect();
  const computed = window.getComputedStyle(element);

  drawCenteredText(
    context,
    text,
    {
      x: (rect.left - paperRect.left) * scale,
      y: (rect.top - paperRect.top) * scale,
      width: rect.width * scale,
      height: rect.height * scale,
    },
    `${computed.fontWeight || "700"} ${Math.max(10, Math.round(parseFloat(computed.fontSize) * scale))}px Arial`,
    computed.color || "#111111",
  );
}

async function downloadSheetAsImage() {
  const target = document.querySelector(".sheet-paper");
  if (!target) {
    return;
  }

  const paperRect = target.getBoundingClientRect();
  const scale = window.devicePixelRatio > 1 ? 2 : 1;
  const width = Math.ceil(paperRect.width * scale);
  const height = Math.ceil(paperRect.height * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);

  target
    .querySelectorAll(".paper-header .date-line span, .paper-header h1")
    .forEach((element) => {
      drawHeaderText(context, element, paperRect, scale);
    });

  target.querySelectorAll("th, td").forEach((cell) => {
    drawCellContent(context, cell, paperRect, scale);
  });

  const pngUrl = canvas.toDataURL("image/png");
  const link = document.createElement("a");
  link.href = pngUrl;
  link.download = `daily-operation-log-${state.date || "sheet"}.png`;
  link.click();
}

function inlineElementStyles(source, target) {
  const computed = window.getComputedStyle(source);
  const styleText = Array.from(computed).map(
    (property) => `${property}: ${computed.getPropertyValue(property)};`,
  );
  target.setAttribute("style", styleText.join(" "));

  const sourceChildren = Array.from(source.children);
  const targetChildren = Array.from(target.children);
  sourceChildren.forEach((child, index) => {
    if (targetChildren[index]) {
      inlineElementStyles(child, targetChildren[index]);
    }
  });
}

function buildMonthlyExportClone(target) {
  const clone = target.cloneNode(true);
  inlineElementStyles(target, clone);

  const originalFields = target.querySelectorAll("input, textarea");
  const clonedFields = clone.querySelectorAll("input, textarea");

  originalFields.forEach((field, index) => {
    const cloneField = clonedFields[index];
    if (!cloneField) {
      return;
    }

    const replacement = document.createElement("div");
    replacement.textContent = field.value || "";
    replacement.className = cloneField.className;
    replacement.setAttribute("style", cloneField.getAttribute("style") || "");

    if (field.classList.contains("photo-attached")) {
      replacement.classList.add("photo-attached");
    }

    replacement.style.display = "flex";
    replacement.style.alignItems = "center";
    replacement.style.justifyContent = "center";
    replacement.style.whiteSpace = "pre-wrap";
    replacement.style.minHeight = cloneField.style.minHeight || "100%";

    cloneField.replaceWith(replacement);
  });

  return clone;
}

async function downloadMonthlyPagesAsImage() {
  const pages = Array.from(monthlyPages.querySelectorAll(".monthly-page"));
  if (pages.length === 0) {
    return;
  }

  if (typeof window.html2canvas !== "function") {
    throw new Error("html2canvas is not available");
  }

  const scale = window.devicePixelRatio > 1 ? 2 : 1;
  for (const [index, page] of pages.entries()) {
    const canvas = await window.html2canvas(page, {
      backgroundColor: "#ffffff",
      scale,
      useCORS: true,
      logging: false,
    });
    const link = document.createElement("a");
    link.href = canvas.toDataURL("image/png");
    link.download = `monthly-readings-${state.monthlyPeriod || "sheet"}-page-${index + 1}.png`;
    link.click();

    await new Promise((resolve) => {
      window.setTimeout(resolve, 150);
    });
  }
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

function getFastInputFieldIndex(fieldKey) {
  return FAST_INPUT_FIELDS.findIndex((field) => field.key === fieldKey);
}

function alignActiveFieldToFastInput() {
  const currentField = getCurrentField();
  if (currentField && FAST_INPUT_FIELD_KEYS.has(currentField.key)) {
    return currentField;
  }

  const nextFastField =
    FAST_INPUT_FIELDS.find(
      (field) => FIELD_INDEX_BY_KEY.get(field.key) >= state.activeFieldIndex,
    ) || FAST_INPUT_FIELDS[FAST_INPUT_FIELDS.length - 1];

  if (!nextFastField) {
    return null;
  }

  state.activeFieldIndex = FIELD_INDEX_BY_KEY.get(nextFastField.key);
  return nextFastField;
}

function updateFastInputPanel() {
  if (!state.date) {
    fastInputFieldName.value = "";
    fastInputValue.value = "";
    fastInputValue.disabled = true;
    fastInputSaveButton.disabled = true;
    return;
  }

  const currentHour = getCurrentHour();
  const currentField = alignActiveFieldToFastInput();
  if (!currentField) {
    fastInputFieldName.value = "";
    fastInputValue.value = "";
    fastInputValue.disabled = true;
    fastInputSaveButton.disabled = true;
    return;
  }

  fastInputFieldName.value = getFieldDisplayLabel(currentField);
  fastInputValue.value = state.entries[currentHour][currentField.key] || "";
  fastInputValue.disabled = isSkippedField(currentField.key);
  fastInputSaveButton.disabled = isSkippedField(currentField.key);
}

function saveFastInputField() {
  fieldValueInput.value = fastInputValue.value;
  saveCurrentField();
  updateFastInputPanel();
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
  const relevantFields = FLAT_FIELDS.filter(
    (field) => !isSkippedField(field.key),
  );
  const filled = relevantFields.filter((field) => {
    const value = state.entries[hour]?.[field.key] || "";
    return String(value).trim() !== "";
  }).length;

  return {
    filled,
    total: relevantFields.length,
  };
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

function getCurrentFieldHeaderText(
  headerId,
  fallbackText,
  fallbackOrientation = "horizontal",
) {
  const config = getHeaderOverride(headerId, fallbackText, fallbackOrientation);
  return config.subText ? `${config.text} (${config.subText})` : config.text;
}

function getFieldDisplayLabel(field) {
  const parts = [];

  if (field.groupHeaderId) {
    parts.push(getCurrentFieldHeaderText(field.groupHeaderId, field.group));
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

function appendHeaderContent(
  cell,
  text,
  subText,
  orientation,
  large = false,
  labelRole = "default",
) {
  const shouldStayHorizontal = labelRole === "main";
  const shouldStayVertical = labelRole === "sub";

  if (shouldStayVertical) {
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
      return;
    }

    const span = document.createElement("div");
    span.className = large ? "vertical-word large" : "vertical-word";
    span.textContent = text;
    cell.append(span);
    return;
  }

  if (shouldStayHorizontal || subText) {
    const stack = document.createElement("div");
    stack.className = "header-text horizontal";

    if (!subText) {
      stack.textContent = text;
      cell.append(stack);
      return;
    }

    const mainSpan = document.createElement("div");
    mainSpan.className = "main-horizontal-label";
    mainSpan.textContent = text;

    const subSpan = document.createElement("div");
    subSpan.className =
      orientation === "vertical"
        ? "vertical-word sub-label-vertical"
        : "header-subtext sub-label-horizontal";
    subSpan.textContent = subText;

    stack.append(mainSpan, subSpan);
    cell.append(stack);
    return;
  }

  const span = document.createElement("div");
  span.className = large ? "vertical-word large" : "vertical-word";
  span.textContent = text;
  cell.append(span);
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
  labelRole = "default",
  editable = true,
) {
  const cell = createCell(tag, className, "", rowSpan, colSpan);
  const config = getHeaderOverride(headerId, text, defaultOrientation);
  if (editable) {
    cell.classList.add("editable-header");
    cell.dataset.headerId = headerId;
    cell.dataset.defaultText = text;
    cell.dataset.defaultOrientation = defaultOrientation;
  }
  appendHeaderContent(
    cell,
    config.text,
    config.subText,
    config.orientation,
    large,
    labelRole,
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
      "default",
      false,
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
        "horizontal",
        false,
        "main",
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
          false,
          "sub",
          false,
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
        "horizontal",
        false,
        "main",
        false,
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
          false,
          "sub",
          false,
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
          false,
          "sub",
          false,
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
      "horizontal",
      false,
      "main",
      false,
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
        1,
        "vertical",
        false,
        "main",
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
      "horizontal",
      false,
      "main",
      false,
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
        1,
        "vertical",
        false,
        "main",
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
      "horizontal",
      false,
      "main",
      false,
    ),
  );
  row2.append(
    createEditableHeaderCell(
      "th",
      "subgroup",
      "temperature-tr1",
      "TR.1",
      1,
      2,
      "horizontal",
      false,
      "main",
      false,
    ),
  );
  row2.append(
    createEditableHeaderCell(
      "th",
      "subgroup",
      "temperature-tr2",
      "TR.2",
      1,
      2,
      "horizontal",
      false,
      "main",
      false,
    ),
  );
  row2.append(
    createEditableHeaderCell(
      "th",
      "subgroup",
      "temperature-tr3",
      "TR.3",
      1,
      2,
      "horizontal",
      false,
      "main",
      false,
    ),
  );
  ["W", "O", "W", "O", "W", "O"].forEach((field, fieldIndex) => {
    row3.append(
      createEditableHeaderCell(
        "th",
        "unit-row temp-cell",
        `temperature-unit-${fieldIndex}`,
        field,
        1,
        1,
        "horizontal",
        false,
        "main",
        false,
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
  row1.append(
    createEditableHeaderCell(
      "th",
      "sf6-title",
      "sf6-title",
      "Sf6",
      1,
      1,
      "horizontal",
      false,
      "main",
      false,
    ),
  );
  SF6_COLUMNS.forEach((column, index) => {
    const className = index >= 8 ? "sf6-small" : "sf6-head";
    row1.append(
      createEditableHeaderCell(
        "th",
        className,
        `sf6-col-${index}`,
        column,
        1,
        1,
        "horizontal",
        false,
        "main",
        false,
      ),
    );
  });
  row1.append(buildNamesSheetCell());

  const row2 = document.createElement("tr");
  row2.append(
    createEditableHeaderCell(
      "th",
      "sf6-time",
      "sf6-time",
      "Time",
      1,
      1,
      "horizontal",
      false,
      "main",
      false,
    ),
  );
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
  syncSheetScrollbars();
}

function syncSheetScrollbars() {
  const sheetPaper = document.querySelector(".sheet-paper");
  if (!sheetPaper) {
    return;
  }

  sheetTopScrollerInner.style.width = `${sheetPaper.scrollWidth}px`;
}

function buildNamesSheetCell() {
  const cell = createCell("td", "names-sheet-cell", "", 4, 30);
  const box = document.createElement("div");
  box.className = "names-sheet-box";

  const facility = state.facilities.find(
    (item) => item.id === state.selectedFacilityId,
  );
  const names = facility
    ? facility.operators
        .map((operator) => ({
          name: operator.name,
          selected: operator.id === state.selectedOperatorId,
        }))
        .filter((operator) => operator.name)
    : [];

  if (names.length === 0) {
    const empty = document.createElement("div");
    empty.className = "sheet-name-item";
    empty.textContent = "";
    box.append(empty);
    cell.append(box);
    return cell;
  }

  names.forEach((operator) => {
    const item = document.createElement("div");
    item.className = "sheet-name-item";
    if (operator.selected) {
      item.classList.add("selected");
    }
    item.textContent = operator.name;
    box.append(item);
  });

  cell.append(box);
  return cell;
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
  const completionPercent =
    completion.total === 0
      ? 100
      : Math.round((completion.filled / completion.total) * 100);
  completionLabel.textContent = `${completionPercent}% complete`;
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
  updateFastInputPanel();
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

function moveFastInputField(step) {
  const currentField = alignActiveFieldToFastInput();
  if (!currentField) {
    return;
  }

  let nextIndex = getFastInputFieldIndex(currentField.key) + step;
  if (nextIndex < 0) {
    return;
  }

  while (
    nextIndex >= 0 &&
    nextIndex < FAST_INPUT_FIELDS.length &&
    isSkippedField(FAST_INPUT_FIELDS[nextIndex].key)
  ) {
    nextIndex += step;
  }

  if (nextIndex >= FAST_INPUT_FIELDS.length) {
    if (state.activeHourIndex < HOURS.length - 1) {
      setActiveHour(state.activeHourIndex + 1);
      alignActiveFieldToFastInput();
      refreshUi();
    }
    return;
  }

  state.activeFieldIndex = FIELD_INDEX_BY_KEY.get(FAST_INPUT_FIELDS[nextIndex].key);
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

openFacilityDialogButton.addEventListener("click", () => {
  facilityNameInput.value = "";
  facilityDialog.showModal();
  window.setTimeout(() => {
    facilityNameInput.focus();
  }, 0);
});

homeLanguageButton.addEventListener("click", () => {
  state.language = state.language === "ar" ? "en" : "ar";
  saveSettings();
  applyLanguage();
  refreshUi();
});

homeNextButton.addEventListener("click", () => {
  if (!state.selectedFacilityId || !state.selectedOperatorId) {
    return;
  }

  showModuleScreen();
});

gatewayButton.addEventListener("click", () => {
  openSelectedOperatorSheet();
});

attendanceButton.addEventListener("click", () => {
  saveStatus.textContent =
    state.language === "ar"
      ? "قسم الحضور قادم لاحقاً"
      : "Attendance is coming soon";
});

monthlyReadingsButton.addEventListener("click", () => {
  openMonthlyReadingsScreen();
});

monthlyBackButton.addEventListener("click", () => {
  showModuleScreen();
});

monthlyCameraButton.addEventListener("click", () => {
  state.monthlyCameraMode = !state.monthlyCameraMode;
  state.pendingMonthlyPhotoField = "";
  updateMonthlyHeader();
  monthlySaveStatus.textContent =
    state.language === "ar"
      ? state.monthlyCameraMode
        ? "اضغط على أي حقل شهري لفتح الكاميرا"
        : "تم إيقاف وضع الكاميرا"
      : state.monthlyCameraMode
        ? "Tap a monthly field to open the camera"
        : "Camera mode turned off";
});

monthlyMonthInput.addEventListener("change", () => {
  state.monthlyPeriod = monthlyMonthInput.value || getCurrentMonth();
  state.monthlyEntries = loadMonthlyEntriesForCurrentSelection();
  state.monthlyPhotos = loadMonthlyPhotosForCurrentSelection();
  populateMonthlyInputs();
  updateMonthlyHeader();
  monthlySaveStatus.textContent = t("ready");
});

monthlyScreen.addEventListener("click", (event) => {
  if (!state.monthlyCameraMode) {
    return;
  }

  const target = event.target.closest(".monthly-cell-input");
  if (!target || !target.dataset.photoField) {
    return;
  }

  event.preventDefault();
  state.pendingMonthlyPhotoField = target.dataset.photoField;
  monthlyCameraInput.value = "";
  monthlyCameraInput.click();
});

monthlyScreen.addEventListener("input", (event) => {
  const target = event.target.closest("[data-monthly-field]");
  if (!target) {
    return;
  }

  let nextValue = target.value;
  if (target.dataset.monthlyType === "number") {
    nextValue = normalizeValueInput(target.value);
    if (target.value !== nextValue) {
      target.value = nextValue;
    }
  }

  state.monthlyEntries[target.dataset.monthlyField] = nextValue;
  saveMonthlyEntries();
});

monthlyCameraInput.addEventListener("change", () => {
  const file = monthlyCameraInput.files?.[0];
  const fieldKey = state.pendingMonthlyPhotoField;
  state.pendingMonthlyPhotoField = "";

  if (!file || !fieldKey) {
    return;
  }

  const reader = new FileReader();
  reader.addEventListener("load", () => {
    const photoData = typeof reader.result === "string" ? reader.result : "";
    if (!photoData) {
      return;
    }

    state.monthlyPhotos[fieldKey] = photoData;
    saveMonthlyPhotos();
    applyMonthlyPhotoPreviews();
  });
  reader.readAsDataURL(file);
});

monthlyExportImageButton.addEventListener("click", async () => {
  try {
    await downloadMonthlyPagesAsImage();
    monthlySaveStatus.textContent =
      state.language === "ar" ? "تم تصدير الصورة" : "Photo exported";
  } catch {
    monthlySaveStatus.textContent =
      state.language === "ar" ? "فشل تصدير الصورة" : "Photo export failed";
  }
});

cancelFacilityButton.addEventListener("click", () => {
  facilityDialog.close();
});

facilityForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const name = facilityNameInput.value.trim();
  if (!name) {
    return;
  }

  state.facilities.push({
    id: createId("facility"),
    name,
    operators: [],
  });
  saveFacilities();
  renderFacilities();
  facilityDialog.close();
});

cancelOperatorButton.addEventListener("click", () => {
  operatorDialog.close();
});

cancelEditFacilityButton.addEventListener("click", () => {
  editFacilityDialog.close();
});

deleteFacilityButton.addEventListener("click", () => {
  const facility = state.facilities.find(
    (item) => item.id === state.editingFacilityDialogId,
  );
  if (!facility) {
    return;
  }

  const confirmed = window.confirm(
    `Delete facility "${facility.name}" and all its operators?`,
  );
  if (!confirmed) {
    return;
  }

  state.facilities = state.facilities.filter((item) => item.id !== facility.id);

  if (state.selectedFacilityId === facility.id) {
    state.selectedFacilityId = "";
    state.selectedOperatorId = "";
  }

  saveFacilities();
  renderFacilities();
  editFacilityDialog.close();
});

operatorForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const name = operatorNameInput.value.trim();
  if (!name || !state.editingFacilityId) {
    return;
  }

  const facility = state.facilities.find(
    (item) => item.id === state.editingFacilityId,
  );
  if (!facility) {
    return;
  }

  facility.operators.push({
    id: createId("operator"),
    name,
  });

  operatorNameInput.value = "";
  state.editingFacilityId = "";
  operatorDialog.close();
  saveFacilities();
  renderFacilities();
  renderSheetNames();
});

facilityList.addEventListener("click", (event) => {
  const editFacilityButton = event.target.closest(
    "button[data-edit-facility-id]",
  );
  if (editFacilityButton) {
    openFacilityEditor(editFacilityButton.dataset.editFacilityId || "");
    return;
  }

  const addOperatorButton = event.target.closest(
    "button[data-facility-id]:not([data-operator-id])",
  );
  if (addOperatorButton) {
    state.editingFacilityId = addOperatorButton.dataset.facilityId || "";
    operatorNameInput.value = "";
    operatorDialog.showModal();
    window.setTimeout(() => {
      operatorNameInput.focus();
    }, 0);
    return;
  }

  const operatorButton = event.target.closest(
    "button[data-facility-id][data-operator-id]",
  );
  if (operatorButton) {
    selectOperator(
      operatorButton.dataset.facilityId || "",
      operatorButton.dataset.operatorId || "",
    );
  }
});

editOperatorList.addEventListener("click", (event) => {
  const deleteButton = event.target.closest("button[data-delete-operator-id]");
  if (!deleteButton) {
    return;
  }

  const row = deleteButton.closest(".edit-operator-row");
  if (row) {
    row.remove();
  }

  if (!editOperatorList.querySelector(".edit-operator-row")) {
    const empty = document.createElement("p");
    empty.className = "operator-empty";
    empty.textContent = "No operators yet.";
    editOperatorList.innerHTML = "";
    editOperatorList.append(empty);
  }
});

editFacilityForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const facility = state.facilities.find(
    (item) => item.id === state.editingFacilityDialogId,
  );
  if (!facility) {
    return;
  }

  const nextName = editFacilityNameInput.value.trim();
  if (!nextName) {
    return;
  }

  const operatorRows = Array.from(
    editOperatorList.querySelectorAll(".edit-operator-row"),
  );
  facility.name = nextName;
  facility.operators = operatorRows
    .map((row) => {
      const input = row.querySelector('input[data-operator-name-input="true"]');
      const name = input ? input.value.trim() : "";
      if (!name) {
        return null;
      }

      return {
        id: row.dataset.operatorId || createId("operator"),
        name,
      };
    })
    .filter(Boolean);

  saveFacilities();
  renderFacilities();
  editFacilityDialog.close();
});

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

exportImageButton.addEventListener("click", async () => {
  if (!state.date) {
    saveStatus.textContent = "Choose a date first";
    return;
  }

  try {
    await downloadSheetAsImage();
    saveStatus.textContent =
      state.language === "ar" ? "تم تصدير الصورة" : "Photo exported";
  } catch {
    saveStatus.textContent =
      state.language === "ar" ? "فشل تصدير الصورة" : "Photo export failed";
  }
});

shareSheetButton.addEventListener("click", () => {
  shareSheetFile();
});

fastInputButton.addEventListener("click", () => {
  alignActiveFieldToFastInput();
  updateFastInputPanel();
  fastInputDialog.showModal();
  window.setTimeout(() => {
    fastInputValue.focus();
    fastInputValue.select();
  }, 0);
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
});

backFieldButton.addEventListener("click", () => {
  saveCurrentField();
  moveField(-1);
});

resetSheetButton.addEventListener("click", resetSheet);
backToHomeButton.addEventListener("click", () => {
  showModuleScreen();
});

loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  showHomeScreen();
});

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

fastInputValue.addEventListener("input", () => {
  const cleanedValue = normalizeValueInput(fastInputValue.value);
  if (fastInputValue.value !== cleanedValue) {
    fastInputValue.value = cleanedValue;
  }
});

fastInputValue.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    fastInputSaveButton.click();
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

closeFastInputButton.addEventListener("click", () => {
  fastInputDialog.close();
});

fastInputBackButton.addEventListener("click", () => {
  fieldValueInput.value = fastInputValue.value;
  saveCurrentField();
  moveFastInputField(-1);
  updateFastInputPanel();
  window.setTimeout(() => {
    fastInputValue.focus();
    fastInputValue.select();
  }, 0);
});

fastInputForm.addEventListener("submit", (event) => {
  event.preventDefault();
  saveFastInputField();
  moveFastInputField(1);
  updateFastInputPanel();
  window.setTimeout(() => {
    fastInputValue.focus();
    fastInputValue.select();
  }, 0);
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

settingsButton.addEventListener("click", () => {
  languageSelect.value = state.language;
  settingsDialog.showModal();
});

closeSettingsButton.addEventListener("click", () => {
  settingsDialog.close();
});

settingsChangeDateButton.addEventListener("click", () => {
  settingsDialog.close();
  openSetupButton.click();
});

settingsResetSheetButton.addEventListener("click", () => {
  settingsDialog.close();
  resetSheet();
});

settingsSkipSpotsButton.addEventListener("click", () => {
  settingsDialog.close();
  manageSkipsButton.click();
});

settingsForm.addEventListener("submit", (event) => {
  event.preventDefault();
  state.language = languageSelect.value === "ar" ? "ar" : "en";
  saveSettings();
  applyLanguage();
  refreshUi();
  settingsDialog.close();
});

sheetTopScroller.addEventListener("scroll", () => {
  if (syncingSheetScroll) {
    return;
  }

  syncingSheetScroll = true;
  sheetStage.scrollLeft = sheetTopScroller.scrollLeft;
  syncingSheetScroll = false;
});

sheetStage.addEventListener("scroll", () => {
  if (syncingSheetScroll) {
    return;
  }

  syncingSheetScroll = true;
  sheetTopScroller.scrollLeft = sheetStage.scrollLeft;
  syncingSheetScroll = false;
});

window.addEventListener("resize", syncMobileFieldLayout);
window.addEventListener("resize", syncSheetScrollbars);

populateHourOptions();
state.language = loadSettings().language;
state.headerOverrides = loadHeaderOverrides();
state.facilities = loadFacilities();
state.monthlyPeriod = getCurrentMonth();
monthlyMonthInput.value = state.monthlyPeriod;
setupVoiceRecognition();
renderMonthlyReadingsSheet();
applyLanguage();
syncMobileFieldLayout();
renderTable();
syncSheetScrollbars();
renderFacilities();
showLoginScreen();
