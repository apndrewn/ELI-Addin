// --- COLGATE EI ASSISTANT ---

// !!! IMPORTANT: PASTE YOUR API KEY BELOW !!!
// ✅ SECURE BACKEND CONNECTION
// Points to your Vercel deployment
const BACKEND_URL = "https://eli-backend.vercel.app/api/generate";

// --- SESSION MANAGEMENT (Per Instance) ---
// This ID stays the same only while the taskpane remains open.
// Closing and re-opening the pane creates a NEW Session ID.
// --- SESSION MANAGEMENT (Standard Per-Instance) ---
var CURRENT_SESSION_ID = sessionStorage.getItem("eli_session_id");
if (!CURRENT_SESSION_ID) {
  CURRENT_SESSION_ID = "sess_" + Date.now();
  sessionStorage.setItem("eli_session_id", CURRENT_SESSION_ID);
}

//console.log("📅 Session ID (Daily):", CURRENT_SESSION_ID);
// We keep this variable defined as a placeholder so we don't have to delete it 
// from every function call below. It is no longer used for authentication.
const API_KEY = "SECURE_PROXY_MODE";

// 🧠 ELI KNOWLEDGE BASE (EDIT THIS TO TEACH THE BOT)

const CLAUSE_DB_NAME = "ELI_ClauseDB";
const CLAUSE_STORE = "clauses";

// Initialize the Clause Database
function initClauseDB() {
  return new Promise((resolve, reject) => {
    var request = indexedDB.open(CLAUSE_DB_NAME, 1);

    request.onupgradeneeded = function (event) {
      var db = event.target.result;
      if (!db.objectStoreNames.contains(CLAUSE_STORE)) {
        db.createObjectStore(CLAUSE_STORE, { keyPath: "id" });
      }
    };

    request.onsuccess = function (event) {
      resolve(event.target.result);
    };

    request.onerror = function (event) {
      console.error("Clause DB Error:", event);
      reject("DB Error");
    };
  });
}

// ASYNC: Save entire library
function saveLibraryToDB(libraryArray) {
  return initClauseDB().then(db => {
    return new Promise((resolve, reject) => {
      var tx = db.transaction([CLAUSE_STORE], "readwrite");
      var store = tx.objectStore(CLAUSE_STORE);

      // 1. Clear old data (simplest way to sync full array)
      store.clear();

      // 2. Add all items
      libraryArray.forEach(item => store.put(item));

      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  });
}

// ASYNC: Load entire library
function loadLibraryFromDB() {
  return initClauseDB().then(db => {
    return new Promise((resolve, reject) => {
      var tx = db.transaction([CLAUSE_STORE], "readonly");
      var store = tx.objectStore(CLAUSE_STORE);
      var request = store.getAll();

      request.onsuccess = () => {
        resolve(request.result || []);
      };
      request.onerror = () => reject(request.error);
    });
  });
}
const FEATURE_DB_NAME = "ELI_FeatureDB";
const STORE_TEMPLATES = "templates";
const STORE_PLAYBOOKS = "playbooks";
const STORE_REQUESTERS = "requesters"; // <--- NEW CONSTANT

function initFeatureDB() {
  return new Promise((resolve, reject) => {
    // INCREASED VERSION TO 2 to trigger upgrade
    var req = indexedDB.open(FEATURE_DB_NAME, 2);

    req.onupgradeneeded = function (e) {
      var db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_TEMPLATES)) db.createObjectStore(STORE_TEMPLATES, { keyPath: "id" });
      if (!db.objectStoreNames.contains(STORE_PLAYBOOKS)) db.createObjectStore(STORE_PLAYBOOKS, { keyPath: "id" });

      // --- NEW STORE ---
      if (!db.objectStoreNames.contains(STORE_REQUESTERS)) db.createObjectStore(STORE_REQUESTERS, { keyPath: "id" });
    };

    req.onsuccess = function (e) { resolve(e.target.result); };
    req.onerror = function (e) { reject("Feature DB Error"); };
  });
}

function loadFeatureData(storeName) {
  return initFeatureDB().then(db => {
    return new Promise((resolve) => {
      var tx = db.transaction([storeName], "readonly");
      var req = tx.objectStore(storeName).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => resolve([]);
    });
  });
}

function saveFeatureData(storeName, data) {
  return initFeatureDB().then(db => {
    return new Promise((resolve) => {
      var tx = db.transaction([storeName], "readwrite");
      var store = tx.objectStore(storeName);
      store.clear(); // Overwrite logic
      data.forEach(item => {
        if (!item.id) item.id = "id_" + Date.now() + "_" + Math.random();
        store.put(item);
      });
      tx.oncomplete = () => resolve();
    });
  });
}
function migrateRequestersToDB() {
  var oldData = localStorage.getItem("colgate_saved_requesters");
  if (oldData) {
    try {
      var parsed = JSON.parse(oldData);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Save to DB
        saveFeatureData(STORE_REQUESTERS, parsed).then(() => {
          console.log("✅ Requesters migrated to DB.");
          localStorage.removeItem("colgate_saved_requesters"); // Cleanup
        });
      }
    } catch (e) { console.error("Migration Error", e); }
  }
}

const ELI_KNOWLEDGE_BASE = {
  OVERVIEW:
    "ELI (Enhanced Legal Intelligence) is an AI-powered Word Add-in designed to help Colgate legal teams review, draft, and negotiate contracts faster and safer.",

  ERROR_CHECK: {
    "What it does":
      "Scans for typos, grammar errors, and defined term inconsistencies.",
    "Advanced Features":
      "Can check for broken cross-references (e.g., 'Error! Source not found') and mismatched Footer References (CP Ref #).",
    Settings:
      "Supports US vs UK dialect enforcement and 'Strict Mode' for passive voice detection.",
  },

  LEGAL_ASSIST: {
    "What it does":
      "Identifies legal risks in the document (Indemnity, Liability, Termination) and rates them High/Medium/Low.",
    Training:
      "You can click 'Ignore' on specific risks to train ELI to stop flagging them for your session.",
  },

  PLAYBOOK: {
    "What it does":
      "Compares the document against Colgate's playbook and provides fallbacks. (NDA, Services, JDA).",
    "Approval Workflow":
      "If a mandatory clause is missing or changed, ELI flags it. You can request approval via email directly from the tool. Once approved, the clause is unlocked.",
    Fallbacks:
      "Offers pre-approved 'Fallback' clauses if the counterparty rejects the standard terms.",
  },

  SUMMARY: {
    "What it does":
      "Extracts key deal points (Parties, Term, Liability Cap, Payment Terms) into a clean grid.",
    "Audience Modes":
      "You can tailor the summary for 'Legal Counsel' (detailed), 'Business Stakeholders' (simple), or 'Finance' (money-focused).",
    Export:
      "Can export the summary to CSV or draft a cover email automatically.",
  },

  ASK_ELI_CHAT: {
    "What it does": "A conversational assistant for the document.",
    Tools: [
      "Rewrite: Select text and ask ELI to rephrase it.",
      "Defend: Generates negotiation arguments to justify your position.",
      "Translate: Translates selection while keeping formatting.",
      "Draft: Creates new clauses from scratch.",
    ],
    "Scope Toggle":
      "Switch between 'Strict' (Document facts only) and 'Broad' (General legal knowledge) modes.",
  },

  COMPARE: {
    "What it does":
      "Smart comparison between the open document and another file (DOCX/PDF/TXT).",
    Modes:
      "Can compare 'Old vs New' (Redlines) or 'Template vs Draft' (Spotting deviations from standard forms).",
  },

  AUTO_FILL: {
    "What it does":
      "Finds placeholder brackets like [Client Name] and fills them instantly using a data file or manual form.",
    "Smart Paste":
      "You can paste messy text (like an email signature) and ELI will auto-extract the data to fill the form.",
  },

  EMAIL_TOOL: {
    "What it does":
      "Scans the document to identify the Third Party, Agreement Type, and Contact Info, then drafts a ready-to-send Outlook/Gmail email.",
    Templates:
      "Includes templates for sending to the Counterparty (External) or the Internal Requestor.",
  },

  SETTINGS: {
    "Laptop Mode": "Pins the toolbar and shrinks icons for small screens.",
    "Audit Trail":
      "Logs every AI change (Draft, Fix, Rewrite) at the bottom of the document for compliance.",
    Language:
      "Switches the entire UI into Spanish, French, German, Chinese, etc.",
  },
};

// --- DOCUMENT CLEANER CONFIGURATION ---
const CLEANER_RULES = {
  "Expenses": {
    title: "💰 Expenses Clause",
    options: [
      { label: "Colgate Standard Clause ($500)", text: "Colgate will reimubrse Company as per Colgate's travel and expense policy after receipt of invoice." },
      { label: "No Reimbursement", text: "Supplier shall be responsible for all expenses incurred in the performance of this Agreement.", desc: "Supplier pays all their own expenses." }
    ]
  },
  // Example of scalability:
  "Governing Law": {
    title: "⚖️ Governing Law",
    options: [
      { label: "New York", text: "This Agreement shall be governed by the laws of the State of New York.", desc: "Standard US preference." },
      { label: "Delaware", text: "This Agreement shall be governed by the laws of the State of Delaware.", desc: "Corporate standard." }
    ]
  }
};
// --- CONSTANTS: DEFAULT WORKFLOWS (Single Source of Truth) ---
const DEFAULT_WORKFLOW_MAPS = {
  default: ["error", "legal", "summary", "chat", "playbook", "compare", "template", "email"],
  draft: ["template", "error", "chat", "email"],
  review: ["playbook", "compare", "legal", "chat", "email", "error"]
};
// 🏪 ELI APP STORE CATALOG (CORE TOOLS + NEW APPS)
// 🏪 ELI APP STORE CATALOG (CATEGORIZED)
const ELI_APP_STORE_CATALOG = [
  // --- NEW EXTENSIONS ---
  {
    id: "app_docu_scrub",
    key: "app1",
    label: "Docu-Scrub",
    icon: "🧼",
    desc: "Format Cleaner: Whitens fonts, flosses spacing, and removes metadata plaque.",
    category: "Featured Extensions",
    action: function () { showDocuScrubUI(); }
  },
  {
    id: "app_term_matrix",
    key: "app2",
    label: "Defined Term Matrix (BETA)",
    icon: "🔠",
    desc: "Definition Scanner: Creates a live index of defined terms & flags undefined 'ghost' terms.",
    category: "Featured Extensions",
    action: function () { runTermMatrix(); } // <--- CHANGED
  },
  {
    id: "app_timeline",
    key: "app3",
    label: "Key Date Timeline",
    icon: "🗓️",
    desc: "Visualizes critical dates (Effective, Expiry, Renewal) in a chronological timeline.",
    category: "Featured Extensions",
    action: function () { runTimelineAnalysis(); }
  },
  {
    id: "app_clause_lib",
    key: "app_clause_lib", // Unique ID
    label: "Clause Library",
    icon: "📚",
    desc: "AI-powered repository. Auto-extract, categorize, and insert preferred clauses.",
    category: "Featured Extensions",
    action: function () {
      // This calls the module we defined above
      showClauseLibraryInterface();
    }
  },

  // --- CORE TOOLS (Restorable) ---
  {
    id: "btnError", key: "error", label: "Error Check", icon: "📝",
    desc: "Scans for typos, grammar mistakes, and broken refs.",
    category: "Core System Tools", // <--- NEW
    action: null
  },
  {
    id: "btnAssist", key: "legal", label: "Legal Issues", icon: "⚖️",
    desc: "AI analysis of high-risk clauses.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnSummary", key: "summary", label: "Summary", icon: "📋",
    desc: "Extracts key deal points into a grid.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnQA", key: "chat", label: "Ask ELI", icon: "💬",
    desc: "Chat assistant. Rewrite, Defend, Translate.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnPlaybook", key: "playbook", label: "Playbook", icon: "📖",
    desc: "Compares document against Colgate standards.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnCompare", key: "compare", label: "Compare", icon: "↔️",
    desc: "Redline analysis against other files.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnTemplate", key: "template", label: "Build Draft", icon: "⚡",
    desc: "Builds a draft from templates.",
    category: "Core System Tools",
    action: null
  },
  {
    id: "btnEmailTool", key: "email", label: "Email Tool", icon: "✉️",
    desc: "Drafts Outlook-ready cover emails.",
    category: "Core System Tools",
    action: null
  }
];
// Paste your key above
// ==========================================
// START: GLOBAL THEME LOGIC
// ==========================================
var BTN_THEMES = ["Default", "Midnight", "Flat", "Pastel", "Platinum", "Royal", "Glass", "Emerald", "Outline", "Executive"];
var BTN_THEME_CLASSES = ["", "theme-midnight", "theme-flat", "theme-pastel", "theme-platinum", "theme-vantage", "theme-glass", "theme-material", "theme-outline", "theme-chrome"];

window.toggleButtonTheme = function () {
  if (!userSettings.buttonThemeIndex) userSettings.buttonThemeIndex = 0;
  userSettings.buttonThemeIndex = (userSettings.buttonThemeIndex + 1) % BTN_THEMES.length;

  applyButtonTheme();
  saveCurrentSettings();
};

// ==========================================
// START: HEADER THEME LOGIC
// ==========================================
var HEADER_THEMES = ["Default", "Classic White", "Robot AI", "Executive", "Midnight", "Oceanic", "Emerald", "Royal", "Copper", "Minimalist"];
var HEADER_THEME_CLASSES = ["", "theme-classic-white", "theme-robot", "theme-executive", "theme-midnight", "theme-oceanic", "theme-emerald", "theme-royal", "theme-copper", "theme-minimalist"];

window.toggleHeaderTheme = function () {
  if (!userSettings.headerThemeIndex) userSettings.headerThemeIndex = 0;
  userSettings.headerThemeIndex = (userSettings.headerThemeIndex + 1) % HEADER_THEMES.length;

  applyHeaderTheme();
  saveCurrentSettings();
};

window.applyHeaderTheme = function () {
  var idx = userSettings.headerThemeIndex || 0;
  var name = HEADER_THEMES[idx];
  var cls = HEADER_THEME_CLASSES[idx];

  // 1. Update Label
  var lbl = document.getElementById("lblHeaderTheme");
  if (lbl) lbl.innerText = name;

  // 2. Apply Class to Header
  var header = document.querySelector(".header");
  if (header) {
    header.classList.remove(
      "theme-classic-white", "theme-robot", "theme-executive", "theme-midnight",
      "theme-oceanic", "theme-emerald", "theme-royal", "theme-copper", "theme-minimalist"
    );
    if (cls) header.classList.add(cls);
  }

  // 3. Robot Icon Logic
  var logoDefault = document.getElementById("logoDefault");
  var logoRobot = document.getElementById("logoRobot");
  var longTitle = document.querySelector(".long-title");

  if (logoDefault && logoRobot) {
    if (name === "Robot AI") {
      // Show Robot, Hide Default
      logoDefault.classList.add("hidden");
      logoRobot.classList.remove("hidden");

      // Hide Long Title text
      if (longTitle) longTitle.style.display = "none";
    } else {
      // Show Default, Hide Robot
      logoDefault.classList.remove("hidden");
      logoRobot.classList.add("hidden");

      if (longTitle) longTitle.style.display = ""; // clears inline style
    }
  }
};

window.resetViewSettings = function () {
  // 1. Reset Settings
  userSettings.pinToolbar = true;
  userSettings.manualLaptopMode = false;
  userSettings.buttonThemeIndex = 0;
  userSettings.headerThemeIndex = 0;

  // 2. Update UI Checkboxes
  var chkPin = document.getElementById("chkPinToolbar");
  if (chkPin) chkPin.checked = true;

  var chkLap = document.getElementById("chkLaptopMode");
  if (chkLap) chkLap.checked = false;
  // Note: We aren't forcing Laptop Mode OFF logic here, just resetting the manual override flag.
  // The user can re-toggle if they want.

  // 3. Apply Themes
  applyButtonTheme();
  applyHeaderTheme();

  // 4. Save
  saveCurrentSettings();

  if (typeof showToast === "function") showToast("View settings reset.");
};

window.applyButtonTheme = function () {
  var idx = userSettings.buttonThemeIndex || 0;
  var name = BTN_THEMES[idx];
  var cls = BTN_THEME_CLASSES[idx];

  console.log("🎨 Applying Theme:", name);

  // 1. Update Settings Menu Label
  var lbl = document.getElementById("lblBtnTheme");
  if (lbl) lbl.innerText = name;

  // 2. Apply to All Toolbar Buttons
  var buttons = document.querySelectorAll(".btn-tile");
  buttons.forEach(function (btn) {
    // Remove ALL possible theme classes first
    btn.classList.remove(
      "theme-midnight", "theme-flat", "theme-pastel", "theme-neon", "theme-platinum",
      "theme-vantage", "theme-glass", "theme-material", "theme-outline", "theme-chrome"
    );

    // Add the active class (if not default)
    if (cls) btn.classList.add(cls);
  });
};
// ==========================================
// END: GLOBAL THEME LOGIC
// ==========================================

// - REPLACEMENT
// --- DEBUG: CONFIRM SCRIPT EXECUTION ---
var eliClauseLibrary = [];
// GLOBAL SETTINGS OBJECT
var userSettings = {
  workflowMaps: null,  // Stores your custom playlists (Draft, Review, Default)
  toolOrder: null,     // Stores the currently active button order
  // Error Check Settings
  strictMode: false,
  currentWorkflow: "default",
  installedStoreApps: [],
  disableTrackChanges: false,
  matchThreshold: 0.50, // <--- Add this
  extractionTerms: [
    "Indemnity", "Liability", "Termination", "Confidentiality",
    "Governing Law", "Warranty", "Force Majeure", "IP Rights"
  ],  // <--- Add this
  hideMatches: false,
  dialect: "US",
  checkStyle: "Standard",
  autoSnapshotOnLoad: false,
  sensitivity: "Balanced",
  // Data Integration
  defaultDataSource: "local",
  dataSourceId: "Onit",
  userGroup: "External Innovation", // <--- NEW: Group Setting
  hasOnboarded: false, // Default to false for new users
  pinToolbar: true, // <--- ADD THIS LINE
  manualLaptopMode: false, // Track if user manually enabled Laptop Mode
  checkBrokenRefs: false,
  checkFooterRef: true,
  highlightFill: true, // Default to true as per user request
  enableAuditTrail: false, // Default Offe
  userEmail: "", // Setup persistence
  fileNamePatterns: null, // File Naming persistence
  requesters: [], // Imported Requesters List

  // --- NEW: Playbook Settings ---
  playbookSource: "hardcoded", // 'hardcoded', 'local', 'online'
  playbookUrl: "",
  approvalEmail: "legal.approvals@colgate.com", // Fallback Default
  enableJustificationComments: true, // <--- NEW SETTING
  buttonThemeIndex: 0,
  headerThemeIndex: 0,
  emailTemplates: {
    myName: "",
    // Template 1: Sending to Third Party
    thirdParty: {
      subject: "[Entity] and [TP Entity] -- [Type] ([CP Ref]) ([Date])",
      body: "Hello [TP Contact],\n\nAt the request of [Requester], attached please find a [Type] between [Entity] and [TP Entity]. Please review and let us know if it is acceptable.\n\nThanks,\n[My Name]",
    },
    // Template 2: Sending to Requester
    requester: {
      subject: "FYI: [Entity] and [TP Entity] -- [Type] ([CP Ref])",
      body: "Hello [Requester],\n\nPlease find attached a draft [Type] with [TP Entity]. Can you please review and let me know if you have any questions or concerns?\n\nThanks,\n[My Name]",
    },
  },
  visibleButtons: {
    error: true,
    legal: true,
    summary: true,
    playbook: true,
    chat: true,
    template: true,
    compare: true,
  },

  // Legal Check Settings
  riskSensitivity: "All",
  showRiskScore: false,
  model: "gemini-2.5-flash-preview-09-2025",
  customPrompts: {
    error: "",
    legal: "",
  },
};


// --- NEW DATA STORES ---
var approvalQueue = [];
var loadedPlaybook = [];

// Load queue from memory

// HELPER: Centralized Save (Safe)
// ==========================================
// START: saveCurrentSettings
// ==========================================
// ==========================================
// START: saveCurrentSettings (SILENT OPTION)
// ==========================================
function saveCurrentSettings(silent) {
  try {
    localStorage.setItem("colgate_user_settings", JSON.stringify(userSettings));
    console.log("💾 Settings Saved:", userSettings);
    // Only show toast if 'silent' is false or undefined
    if (!silent && typeof showToast === 'function') showToast("✅ Settings Saved");
  } catch (e) {
    console.error("❌ Save Failed (LocalStorage blocked?):", e);
  }
}
// ==========================================
// START: loadSavedSettings
// ==========================================
function loadSavedSettings() {
  console.log("📂 Attempting to load settings...");
  try {
    var saved = localStorage.getItem("colgate_user_settings");
    if (saved) {
      var parsed = JSON.parse(saved);
      // Merge saved settings into the defaults safely
      for (var key in parsed) {
        // Only load keys that exist in our default object (Security)
        if (userSettings.hasOwnProperty(key)) {
          userSettings[key] = parsed[key];
        }
      }
      console.log("✅ Settings Loaded Successfully:", userSettings);
    } else {
      console.log("No saved settings found.Using defaults.");
    }
  } catch (e) {
    console.warn("⚠️ Error loading settings (using defaults):", e);
  }
}
// ==========================================
// END: loadSavedSettings
// ==========================================

var currentController = null;
var loadingInterval = null;

// --- ADD THESE TWO LINES HERE: ---
var userFeedback = [];
var legalBlacklist = [];
var chatHistory = [];
var AGREEMENT_DATA = null; // Init Global Store
// GLOBAL SETTINGS
// --- COLGATE PLAYBOOK STANDARDS ---
// --- COLGATE PLAYBOOK STANDARDS (UPDATED) ---
// --- COLGATE PLAYBOOK STANDARDS (MERGED & UPDATED) ---
// [Search for "const STANDARD_PLAYBOOK = [" around line 291 and add this BEFORE it]
const dictionary = {
  en: {
    eaderSettings: "SETTINGS",
    lblPreferences: "PREFERENCES",
    lblLaptop: "Laptop Mode",
    lblLanguage: "LANGUAGE",
    btnError: "Error Check",
    btnLegal: "Legal Issues",
    btnSummary: "Summary",
    btnPlaybook: "Playbook",
    btnCompare: "Compare",
    btnAutoFill: "Build Draft",
    btnEmail: "Email", // Ensure this is here
    lblChat: "Ask ELI", // <--- ADD THIS LINE
    msgSelectText: "Click a button above to scan entire document.",
    msgScanLimit: "If text is selected, only that text will be scanned.",
    lblCustomize: "Customize Toolbar",
    lblAudit: "<strong>Audit Trail</strong> (Log Changes)",
    msgWarnTitle: "Partial Selection Active",
    msgWarnDesc:
      "You are analyzing <strong>only</strong> the highlighted text.<br>(Click anywhere else to reset to Full Document)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Ask a question or click a tool below to start.",
  },
  es: {
    headerSettings: "CONFIGURACIÓN",
    lblPreferences: "PREFERENCIAS",
    lblLaptop: "Modo Laptop",
    lblLanguage: "IDIOMA",
    btnError: "Revisar Error",
    btnLegal: "Asuntos Legales",
    btnSummary: "Resumen",
    btnPlaybook: "Estrategia",
    btnCompare: "Comparar",
    btnAutoFill: "Auto-Relleno",
    btnEmail: "Correo",
    lblChat: "Preguntar a ELI",
    msgSelectText: "Seleccione texto en su documento o haga clic en un botón.",
    msgScanLimit: "Solo se escaneará el texto seleccionado.",
    lblCustomize: "Personalizar Barra",
    lblAudit: "<strong>Historial</strong> (Registrar Cambios)",
    msgWarnTitle: "Selección Parcial Activa",
    msgWarnDesc:
      "Está analizando <strong>solo</strong> el texto resaltado.<br>(Haga clic en otro lugar para restablecer)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Haga una pregunta o elija una herramienta para comenzar.",
  },
  fr: {
    headerSettings: "PARAMÈTRES",
    lblPreferences: "PRÉFÉRENCES",
    lblLaptop: "Mode Portable",
    lblLanguage: "LANGUE",
    btnError: "Erreurs",
    btnLegal: "Juridique",
    btnSummary: "Résumé",
    btnPlaybook: "Playbook",
    btnCompare: "Comparer",
    btnAutoFill: "Remplissage",
    btnEmail: "Email",
    lblChat: "Demander à ELI",
    msgSelectText: "Sélectionnez du texte ou cliquez sur un bouton.",
    msgScanLimit: "Seul le texte sélectionné sera analysé.",
    lblCustomize: "Personnaliser la barre",
    lblAudit: "<strong>Piste d'audit</strong> (Journal)",
    msgWarnTitle: "Sélection Partielle Active",
    msgWarnDesc:
      "Vous analysez <strong>uniquement</strong> le texte surligné.<br>(Cliquez ailleurs pour tout réinitialiser)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Posez une question ou cliquez sur un outil pour commencer.",
  },
  de: {
    headerSettings: "EINSTELLUNGEN",
    lblPreferences: "PRÄFERENZEN",
    lblLaptop: "Laptop-Modus",
    lblLanguage: "SPRACHE",
    btnError: "Fehlerprüfung",
    btnLegal: "Rechtliches",
    btnSummary: "Zusammenfassung",
    btnPlaybook: "Handbuch",
    btnCompare: "Vergleichen",
    btnAutoFill: "Auto-Ausfüllen",
    btnEmail: "E-Mail",
    lblChat: "Frag ELI",
    msgSelectText: "Markieren Sie Text oder klicken Sie auf eine Schaltfläche.",
    msgScanLimit: "Nur markierter Text wird gescannt.",
    lblCustomize: "Symbolleiste anpassen",
    lblAudit: "<strong>Audit-Protokoll</strong> (Log)",
    msgWarnTitle: "Teilauswahl Aktiv",
    msgWarnDesc:
      "Sie analysieren <strong>nur</strong> den markierten Text.<br>(Klicken Sie woanders, um zurückzusetzen)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Stellen Sie eine Frage oder wählen Sie ein Tool aus.",
  },
  it: {
    headerSettings: "IMPOSTAZIONI",
    lblPreferences: "PREFERENZE",
    lblLaptop: "Modalità Laptop",
    lblLanguage: "LINGUA",
    btnError: "Controllo Errori",
    btnLegal: "Legale",
    btnSummary: "Riepilogo",
    btnPlaybook: "Playbook",
    btnCompare: "Confronta",
    btnAutoFill: "Compilazione",
    btnEmail: "Email",
    lblChat: "Chiedi a ELI",
    msgSelectText: "Seleziona il testo o clicca su un pulsante.",
    msgScanLimit: "Verrà scansionato solo il testo selezionato.",
    lblCustomize: "🎨 Personalizza barra",
    lblAudit: "<strong>Audit Trail</strong> (Registro)",
    msgWarnTitle: "Selezione Parziale Attiva",
    msgWarnDesc:
      "Stai analizzando <strong>solo</strong> il testo evidenziato.<br>(Clicca altrove per ripristinare)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Fai una domanda o clicca su uno strumento per iniziare.",
  },
  pt: {
    headerSettings: "CONFIGURAÇÕES",
    lblPreferences: "PREFERÊNCIAS",
    lblLaptop: "Modo Laptop",
    lblLanguage: "IDIOMA",
    btnError: "Verificar Erros",
    btnLegal: "Jurídico",
    btnSummary: "Resumo",
    btnPlaybook: "Manual",
    btnCompare: "Comparar",
    btnAutoFill: "Preenchimento",
    btnEmail: "Email",
    lblChat: "Perguntar ao ELI",
    msgSelectText: "Selecione o texto ou clique num botão.",
    msgScanLimit: "Apenas o texto selecionado será analisado.",
    lblCustomize: "Personalizar Barra",
    lblAudit: "<strong>Trilha de Auditoria</strong>",
    msgWarnTitle: "Seleção Parcial Ativa",
    msgWarnDesc:
      "Você está analisando <strong>apenas</strong> o texto destacado.<br>(Clique em outro lugar para redefinir)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Faça uma pergunta ou clique em uma ferramenta para começar.",
  },
  zh: {
    headerSettings: "设置",
    lblPreferences: "首选项",
    lblLaptop: "笔记本模式",
    lblLanguage: "语言",
    btnError: "错误检查",
    btnLegal: "法律风险",
    btnSummary: "摘要",
    btnPlaybook: "攻略手册",
    btnCompare: "比较",
    btnAutoFill: "自动填充",
    btnEmail: "邮件",
    lblChat: "询问 ELI",
    msgSelectText: "请在文档中选择文本或点击上方按钮。",
    msgScanLimit: "仅扫描选定的文本。",
    lblCustomize: "自定义工具栏",
    lblAudit: "<strong>审计跟踪</strong> (日志)",
    msgWarnTitle: "部分选择处于活动状态",
    msgWarnDesc:
      "您正在分析<strong>仅</strong>突出显示的文本。<br>（单击其他任何位置以重置）",
    msgChatWelcome: "<strong>👋 ELI</strong>: 请提问或点击下方工具开始。",
  },
  ja: {
    headerSettings: "設定",
    lblPreferences: "環境設定",
    lblLaptop: "ラップトップモード",
    lblLanguage: "言語",
    btnError: "エラーチェック",
    btnLegal: "法的リスク",
    btnSummary: "要約",
    btnPlaybook: "プレイブック",
    btnCompare: "比較",
    btnAutoFill: "自動入力",
    btnEmail: "メール",
    lblChat: "ELIに聞く",
    msgSelectText: "テキストを選択するか、ボタンをクリックしてください。",
    msgScanLimit: "選択されたテキストのみスキャンされます。",
    lblCustomize: "ツールバーのカスタマイズ",
    lblAudit: "<strong>監査証跡</strong> (ログ)",
    msgWarnTitle: "部分選択アクティブ",
    msgWarnDesc:
      "ハイライトされたテキスト<strong>のみ</strong>を分析しています。<br>（リセットするには他の場所をクリック）",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: 質問するか、下のツールをクリックして開始します。",
  },
  ru: {
    headerSettings: "НАСТРОЙКИ",
    lblPreferences: "ПРЕДПОЧТЕНИЯ",
    lblLaptop: "Режим ноутбука",
    lblLanguage: "ЯЗЫК",
    btnError: "Ошибки",
    btnLegal: "Юридический",
    btnSummary: "Резюме",
    btnPlaybook: "Сценарий",
    btnCompare: "Сравнить",
    btnAutoFill: "Автозаполнение",
    btnEmail: "Email",
    lblChat: "Спросить ELI",
    msgSelectText: "Выберите текст или нажмите кнопку.",
    msgScanLimit: "Будет сканироваться только выбранный текст.",
    lblCustomize: "🎨 Настроить панель",
    lblAudit: "<strong>Аудит</strong> (Журнал)",
    msgWarnTitle: "Активен частичный выбор",
    msgWarnDesc:
      "Вы анализируете <strong>только</strong> выделенный текст.<br>(Нажмите в любом месте для сброса)",
    msgChatWelcome:
      "<strong>👋 ELI</strong>: Задайте вопрос или выберите инструмент, чтобы начать.",
  },
}; // 2. Initialize Language Logic
document.addEventListener("DOMContentLoaded", () => {
  const langSelector = document.getElementById("selLanguage");

  // Load saved language or default to English
  const savedLang = localStorage.getItem("eli_language") || "en";
  langSelector.value = savedLang;
  setLanguage(savedLang);

  // Apply Settings & Theme
  loadSavedSettings();
  renderToolbar(); // Re-render to ensure buttons exist before theming
  applyButtonTheme();
  applyHeaderTheme();

  // Listen for changes
  langSelector.addEventListener("change", (e) => {
    setLanguage(e.target.value);
  });

  // --- NEW: User Group Logic ---
  const groupSelector = document.getElementById("selUserGroup");
  if (groupSelector) {
    // Load saved group (default to 'External Innovation' if missing)
    groupSelector.value = userSettings.userGroup || "External Innovation";

    // Save on change
    groupSelector.addEventListener("change", (e) => {
      userSettings.userGroup = e.target.value;
      saveCurrentSettings();
    });
  }

  // --- RESTORED SETTINGS BUTTON LOGIC ---
  var btnSettings = document.getElementById("btnHeaderSettings");
  var settingsMenu = document.getElementById("settingsMenu");
  var btnCloseSettings = document.getElementById("btnCloseSettingsMain");

  if (btnSettings && settingsMenu) {
    btnSettings.onclick = function (e) {
      e.stopPropagation(); // Prevent bubbling
      if (settingsMenu.classList.contains("hidden")) {
        settingsMenu.classList.remove("hidden");
        settingsMenu.style.display = "block";
      } else {
        settingsMenu.classList.add("hidden");
        settingsMenu.style.display = "none";
      }
    };
  }

  if (btnCloseSettings && settingsMenu) {
    btnCloseSettings.onclick = function () {
      settingsMenu.classList.add("hidden");
      settingsMenu.style.display = "none";
    };
  }
}); // End DOMContentLoaded (Original)

// --- ROBUST INITIALIZATION ---
// Ensure Settings Logic attaches even if DOMContentLoaded fired early
(function () {
  function initSettingsListeners() {
    var btnSettings = document.getElementById("btnHeaderSettings");
    var settingsMenu = document.getElementById("settingsMenu");
    var btnCloseSettings = document.getElementById("btnCloseSettingsMain");

    if (btnSettings && settingsMenu) {
      btnSettings.onclick = function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (settingsMenu.classList.contains("hidden")) {
          settingsMenu.classList.remove("hidden");
          settingsMenu.style.display = "block";
        } else {
          settingsMenu.classList.add("hidden");
          settingsMenu.style.display = "none";
        }
      };
    }
    if (btnCloseSettings && settingsMenu) {
      btnCloseSettings.onclick = function (e) {
        e.preventDefault();
        settingsMenu.classList.add("hidden");
        settingsMenu.style.display = "none";
      };
    }

    // --- NEW: Import Requesters Logic ---
    var btnImport = document.getElementById("btnImportRequesters");
    var fileInput = document.getElementById("fileImportRequesters");

    if (btnImport && fileInput) {
      btnImport.onclick = function (e) {
        e.preventDefault(); // Stop any button default
        fileInput.click();
      };
      fileInput.onchange = function (e) {
        if (!e.target.files || e.target.files.length === 0) return;
        var file = e.target.files[0];
        var reader = new FileReader();

        reader.onload = function (event) {
          try {
            var content = event.target.result;
            var json = JSON.parse(content);

            if (Array.isArray(json)) {
              // SCHEMA SAFETY CHECK
              var validItems = json.filter(item => item && typeof item === 'object' && item.email && item.name);
              var invalidCount = json.length - validItems.length;

              // Load existing from DB first
              loadFeatureData(STORE_REQUESTERS).then(function (existing) {
                var combined = (existing || []).concat(validItems);

                // Save combined to DB
                saveFeatureData(STORE_REQUESTERS, combined).then(() => {
                  if (validItems.length > 0) showToast("✅ " + validItems.length + " Requesters Imported");
                  if (invalidCount > 0) showToast("⚠️ Skipped " + invalidCount + " invalid items");
                });
              });

            } else {
              showToast("⚠️ Invalid Data: Expected Array");
            }
          } catch (err) {
            console.error("JSON Import Error:", err);
            showToast("⚠️ Bad JSON File. Please check format.");
          }
          fileInput.value = ""; // Reset
        };
        reader.readAsText(file);
      };
    }
    var btnImportTmpl = document.getElementById("btnImportTemplates");
    var fileInputTmpl = document.getElementById("fileImportTemplates");

    if (btnImportTmpl && fileInputTmpl) {
      btnImportTmpl.onclick = function (e) {
        e.preventDefault();
        fileInputTmpl.click();
      };

      fileInputTmpl.onchange = function (e) {
        if (!e.target.files || e.target.files.length === 0) return;
        var file = e.target.files[0];
        var reader = new FileReader();

        reader.onload = function (event) {
          try {
            var content = event.target.result;
            var json = JSON.parse(content);

            if (Array.isArray(json)) {

              loadFeatureData(STORE_TEMPLATES).then(function (existing) {
                var combined = (existing || []).concat(json);
                saveFeatureData(STORE_TEMPLATES, combined).then(function () {
                  showToast("✅ Templates Imported (" + json.length + ")");
                });
              });



            } else {
              // Warning, not crash
              showToast("⚠️ Invalid JSON: Expected an Array.");
            }
          } catch (err) {
            console.error("Template Import Error:", err);
            // Non-intrusive warning
            showToast("⚠️ Bad JSON File. Please check format.");
          }
          fileInputTmpl.value = ""; // Reset
        };
        reader.readAsText(file);
      };
    }
  }

  if (document.readyState === "complete" || document.readyState === "interactive") {
    initSettingsListeners();
  } else {
    document.addEventListener("DOMContentLoaded", initSettingsListeners);
  }
})();

// ==========================================
// START: setLanguage
// ==========================================
// ==========================================
// START: setLanguage (SAFE VERSION)
// ==========================================
function setLanguage(lang) {
  // 1. Safety Checks (Prevents Crash)
  if (typeof dictionary === "undefined") {
    console.warn("⚠️ Dictionary object is missing.");
    return;
  }

  // 2. Fallback to English if lang is invalid or missing from dictionary
  if (!lang || !dictionary[lang]) {
    // console.log("Language '" + lang + "' not found. Defaulting to English.");
    lang = "en";
  }

  // 3. Save Preference
  localStorage.setItem("eli_language", lang);

  // 4. Update UI
  document.querySelectorAll("[data-key]").forEach(function (el) {
    var key = el.getAttribute("data-key");
    if (dictionary[lang] && dictionary[lang][key]) {
      el.innerHTML = dictionary[lang][key];
    }
  });
}
// ==========================================
// END: setLanguage
// ==========================================
// ==========================================
// END: setLanguage
// ==========================================

// --- ENTERPRISE GUARDRAILS (KILL SWITCH) ---
const STANDARD_PLAYBOOK = [
  {
    id: "confidentiality_def",
    name: "Confidentiality Definition",
    types: ["NDA", "Services", "JDA", "Other"],
    standard:
      "Confidential Information means (a) any information or data, whether technical or nontechnical, whether written (including graphic or electronic) and marked confidential or oral, disclosed by such Party or on such Party’s behalf to the Receiving Party, either directly or indirectly, on or after the Effective Date, and (b) the existence and terms of this Agreement.",
    fallbacks: [
      {
        label: "Marking Required (Strict)",
        text: "Confidential Information shall mean information that is disclosed by such Party and that is identified in writing as Confidential Information hereunder no later than the date of its disclosure hereunder.",
        desc: "Limits scope to marked items only.",
      },
      {
        label: "Reasonable Person (Broad)",
        text: "Confidential Information includes information that is identified in writing as Confidential or that a reasonable person would understand to be confidential given the nature of the information and the circumstances of its disclosure.",
        desc: "Protects inadvertent disclosure.",
      },
    ],
  },
  {
    // <--- THIS WAS MISSING IN YOUR SNIPPET
    id: "gov_law",
    name: "Governing Law",
    types: ["All", "NDA", "Services", "Clinical", "JDA", "Other"],
    standard:
      "This Agreement shall be governed by, and construed and enforced in accordance with, the laws of State of New York, without giving effect to any choice of law or conflicts of law rules or provisions.",
    fallbacks: [
      {
        label: "England & Wales",
        text: "<u>Governing Law.</u> This Agreement shall be governed by, and construed and enforced in accordance with, the laws of England & Wales, without giving effect to any choice of law or conflicts of law rules or provisions.",
        desc: "International / Neutral preference.",
      },
      {
        label: "New Jersey",
        text: "This Agreement shall be governed by, and construed and enforced in accordance with, the laws of State of New Jersey, without giving effect to any choice of law or conflicts of law rules or provisions.",
        desc: "Colgate HQ Alternative.",
      },
      {
        label: "Silent",
        text: "Intentionally Omitted.",
        desc: "Remove clause to avoid jurisdiction dispute.",
      },
    ],
  },
  {
    id: "residual_info",
    name: "Residual Information",
    types: ["NDA", "Services"],
    approvalRequired: true,
    // The standard you WANT (if it were there)
    standard:
      "Notwithstanding the confidentiality obligations contained in this Section, the Receiving Party may, during and after the term hereof, use in its business any Residual Information. 'Residual Information' means information of a general nature, such as general knowledge, professional skills, know-how, work experience or techniques, that would be retained in the unaided memory of an ordinary person skilled in the art, not intent on appropriating the proprietary information of the Disclosing Party, as a result of such person’s access to, use, review, evaluation, or testing of the Confidential Information of the Disclosing Party. Memory shall be considered unaided if the person has not intentionally memorized the information contained within the Confidential Information for the purpose of retaining and subsequently using or disclosing same.   Nothing in this Section, however, shall be deemed to grant to the Receiving Party a license under the Disclosing Party’s patents or other intellectual property rights.",
    fallbacks: [
      {
        label: "Silent / Deleted",
        text: "", // Empty string represents deletion
        desc: "If no risk or very low risk (one-off project; we don't have any expertise in this field), consider removing the clause.",
      },
    ],
  },
  {
    id: "assignment",
    name: "Assignment",
    types: ["All", "NDA", "Services", "JDA"],
    standard:
      "Neither this Agreement nor any of the Parties’ rights or obligations hereunder may be assigned, transferred or delegated (collectively “to assign” or “assignment”) without the prior written consent of the other Party; provided, however, that Colgate may assign this Agreement or any of its rights or obligations hereunder to any of its Affiliates without obtaining the Company’s prior written consent. This Agreement shall be binding on, and inure to the benefit of, each Party’s successors and permitted assigns.",
    fallbacks: [
      {
        label: "Mutual Affiliate Assignment",
        text: "Neither this Agreement nor any of the Parties’ rights or obligations hereunder may be assigned, transferred or delegated (collectively “to assign” or “assignment”) without the prior written consent of the other Party; provided, however, that either Party may assign this Agreement or any of its rights or obligations hereunder to any of its Affiliates without obtaining the other Party’s prior written consent. This Agreement shall be binding on, and inure to the benefit of, each Party’s successors and permitted assigns.",
        desc: "Allows both parties to assign to Affiliates without consent.",
      },
    ],
  },
  {
    id: "trade_secrets",
    name: "Trade Secrets (No Disclosure)",
    types: ["NDA", "Services", "JDA", "All"],
    standard:
      "No Disclosing Party shall disclose any “trade secrets” to the Receiving Party pursuant to this Agreement. For purposes of this Agreement “trade secrets” shall have the meaning ascribed to them in the Defend Trade Secrets Act of 2016.",
    fallbacks: [
      {
        label: "Accept Disclosure (VP Sign-off)",
        text: "",
        desc: "VP Sign-off is needed to accept trade secrets (Clause Deleted).",
      },
      {
        label: "Restore Prohibition",
        text: "<u>Trade Secrets.</u> No Disclosing Party shall disclose any “trade secrets” to the Receiving Party pursuant to this Agreement. For purposes of this Agreement “trade secrets” shall have the meaning ascribed to them in the Defend Trade Secrets Act of 2016.",
        desc: "Return the clause (Restrict trade secrets).",
      },
    ],
  },
  {
    id: "term_length",
    name: "Term",
    types: ["NDA"],
    approvalRequired: true,
    standard:
      "The term of this Agreement shall commence on the Effective Date and continue in effect for one (1) year.",
    fallbacks: [
      {
        label: "2 Year Term",
        text: "The term of this Agreement shall commence on the Effective Date and continue in effect for two (2) years.",
        desc: "Matches standard template default.",
      },
      {
        label: "3 Year Term",
        text: "The term of this Agreement shall commence on the Effective Date and continue in effect for three (3) years.",
        desc: "Longer duration.",
      },
    ],
  },
  {
    id: "survival",
    name: "Survival",
    types: ["NDA"],
    standard:
      "The obligations of confidentiality and non-use set forth herein shall apply during the term of this Agreement and survive for a period of three (3) years after the expiration or termination of this Agreement.",
    fallbacks: [
      {
        label: "5 Year Survival",
        text: "The obligations of confidentiality... shall survive for a period of five (5) years after the expiration or termination.",
        desc: "Extended protection.",
      },
      {
        label: "Perpetual (Trade Secrets)",
        text: "The obligations of confidentiality shall survive for three (3) years, provided that obligations with respect to Trade Secrets shall survive perpetually.",
        desc: "Stronger IP protection.",
      },
    ],
  },
  {
    id: "liability_cap",
    name: "Limitation of Liability",
    types: ["NDA", "Services", "JDA"],
    standard:
      "IN NO EVENT SHALL EITHER PARTY BE LIABLE TO THE OTHER PARTY... FOR INCIDENTAL, CONSEQUENTIAL, INDIRECT, SPECIAL, PUNITIVE OR EXEMPLARY DAMAGES OR LIABILITIES OF ANY KIND OR FOR LOSS OF REVENUE OR PROFITS OR LOSS OF BUSINESS.",
    fallbacks: [
      {
        label: "Monetary Cap",
        text: "Liability shall be limited to an AGGREGATE AMOUNT IN EXCESS OF $[AMOUNT].",
        desc: "Sets a specific dollar limit on risk.",
      },
    ],
  },
  {
    id: "payment",
    name: "Payment Terms",
    types: ["Services", "Clinical", "JDA", "Other"],
    standard:
      "Colgate shall pay undisputed invoices within ninety (90) days of receipt.",
    fallbacks: [
      {
        label: "Net 60 Days",
        text: "Colgate shall pay undisputed invoices within sixty (60) days of receipt.",
        desc: "Requires VP Approval.",
      },
      {
        label: "Net 45 (2% Disc)",
        text: "Colgate shall pay undisputed invoices within forty-five (45) days of receipt, provided a 2% discount is applied.",
        desc: "Critical Vendors Only.",
      },
    ],
  },
  {
    id: "ai_services",
    name: "AI Policy (Services/Non-Agency)",
    types: ["Services", "JDA", "Other"],
    standard: `(a) "AI Technologies" means deep learning, machine learning, and other artificial intelligence technologies. Prior to making use of any AI Technologies in providing the Deliverables or performing the Services under this Agreement (except for incidental use of AI Technologies that is not material to the Services), Supplier shall disclose in writing to Company a description of the AI Technologies to be used and the purpose for using them. (b) Supplier shall not use Company Confidential Information to train, test or verify ("Train") any AI models, unless (i) approved by Company in writing or (ii) specifically developed for Company’s exclusive use. All Company inputs and outputs specifically for Company constitute Company’s Confidential Information. (c) Supplier warrants that (i) Supplier AI Technologies comply with all applicable laws and industry guidelines; (ii) use of AI Technologies to perform Services shall comply with applicable laws; and (iii) Supplier shall not commingle Company’s data with third-party data to facilitate collusive behavior or "hub and spoke" activities.`,
  },
  {
    id: "ai_supply",
    name: "AI Policy (Supply/Trans/Equip)",
    types: ["Services", "Other"],
    standard: `(a) "AI Technologies" means deep learning, machine learning, and other artificial intelligence technologies. Upon Company’s request, Supplier shall disclose in writing to Company a description of the AI Technologies used in performing its obligations under this Agreement and the purpose for using the AI Technologies (except for incidental use). (b) Supplier shall not use Company Confidential Information to train, test or verify ("Train") any AI models, unless (i) approved by Company in writing or (ii) specifically developed for Company’s exclusive use. (c) Supplier warrants that (i) Supplier AI Technologies comply with all applicable laws; (ii) Supplier’s use of AI Technologies to provide goods/services shall comply with applicable laws; and (iii) Supplier shall not commingle Company’s data with third-party data to facilitate collusive behavior.`,
  },
  {
    id: "clinical_std",
    name: "Clinical Standards",
    types: ["Clinical"],
    standard:
      "Supplier shall comply with Good Clinical Practice (GCP) and ICH guidelines in the performance of the Study.",
    fallbacks: [
      {
        label: "Local Law Only",
        text: "Supplier shall comply with all applicable local laws and regulations regarding clinical trials.",
        desc: "Less specific standard.",
      },
    ],
  },
]; // Store custom clauses added during the session here
var sessionClauses = [];
// --- TEMPLATE LIBRARY (EDITABLE DIRECTORY) ---

// NEW: Default System Personas (This is what loads if you haven't edited it)
const DEFAULT_PERSONAS = {
  error:
    "You are a professional editor for Colgate. Scan the text for grammatical errors, spelling mistakes, and awkward phrasing.",
  legal:
    "You are a Senior Legal Assistant for Colgate. Review this agreement and identify specific legal risks.",
};
// --- CONSTANTS: LOADING MESSAGES (Dental Remix) ---
const LOADING_MSGS = {
  default: [
    "Analyzing document...",
    "🦷 Scrubbing away errors...",
    "Processing...",
    "✨ Polishing the details..."
  ],
  error: [
    "📄 Scanning text...",
    "🪥 Brushing up grammar...",
    "🔍 Identifying typos...",
    "🦷 Flossing out mistakes...",
    "🧠 Checking grammar rules...",
    "✨ Whitening the syntax...",
    "🧐 Looking for awkward phrasing...",
    "🧴 Squeezing out fresh edits...",
    "✍️ Formulating suggestions...",
    "✅ Minty fresh finish..."
  ],
  legal: [
    "⚖️ Reading agreement...",
    "🛡️ Fighting legal plaque...",
    "🔎 Checking definitions...",
    "🦷 Drilling into details...",
    "🧐 Analyzing liability clauses...",
    "✨ Polishing the fine print...",
    "📝 Drafting risk report...",
    "🦷 Rinse and repeat..."
  ],
  summary: [
    "📖 Reading full text...",
    "✨ Scrubbing for key points...",
    "🗝️ Extracting key clauses...",
    "🦷 Polishing the summary...",
    "🧠 Analyzing terms...",
    "🪥 Brushing up the abstract...",
    "📝 Summarizing points...",
    "✨ Minty fresh results..."
  ],
  template: [
    "📄 Reading file...",
    "🧴 Squeezing in data...",
    "🧠 Mapping fields...",
    "🦷 Filling cavities in the doc...",
    "✍️ Filling placeholders...",
    "✨ Polishing final draft..."
  ],
  redline: [
    "📝 Tracking changes...",
    "🦷 Flossing between the edits...",
    "🔍 Identifying deletions...",
    "✨ Whitening the redlines...",
    "🧠 Analyzing insertions...",
    "🛡️ Enamel protection for your doc...",
    "📊 Calculating risk score...",
    "✅ Smiles all around..."
  ]
};
// --- CONSTANTS: CENTRALIZED PROMPTS ---

// ==========================================
// START: renderToolbar (CRASH PROOF VERSION)
// ==========================================
function renderToolbar() {
  var container = document.getElementById("mainToolbar");
  if (!container) return;
  if (typeof TOOL_DEFINITIONS === 'undefined' || !Array.isArray(TOOL_DEFINITIONS)) {
    setTimeout(renderToolbar, 100);
    return;
  }
  container.innerHTML = "";

  var currentLang = localStorage.getItem('eli_language') || 'en';

  // --- SAFETY CHECK: Initialize Defaults if Missing ---
  if (!userSettings.visibleButtons) {
    userSettings.visibleButtons = {
      error: true, legal: true, summary: true, playbook: true,
      chat: true, template: true, compare: true, email: true
    };
  }

  if (!userSettings.buttonPages) {
    userSettings.buttonPages = {
      error: 1, legal: 1, summary: 1, chat: 1,
      email: 2, template: 2, playbook: 2, compare: 2
    };
  }
  // ----------------------------------------------------

  // 1. Determine Correct Order
  var sortOrder = userSettings.toolOrder && userSettings.toolOrder.length > 0
    ? userSettings.toolOrder
    : ["error", "legal", "summary", "chat", "playbook", "compare", "template", "email"];
  // --- FIX: Force visibility/paging for active tools ---
  sortOrder.forEach(function (key, index) {
    if (!userSettings.visibleButtons) userSettings.visibleButtons = {};
    if (!userSettings.buttonPages) userSettings.buttonPages = {};

    userSettings.visibleButtons[key] = true;
    userSettings.buttonPages[key] = (index < 4) ? 1 : 2;
  });
  // ----------------------------------------------------
  var pageButtons = [];

  sortOrder.forEach(function (key) {
    var def = TOOL_DEFINITIONS.find(function (t) { return t.key === key; });
    if (!def) return;

    // Safe Access
    var isVisible = userSettings.visibleButtons[key];
    var assignedPage = userSettings.buttonPages[key] || 1;

    if (isVisible && assignedPage == toolbarPage) {
      pageButtons.push(def);
    }
  });

  // LEFT ARROW
  var leftArrow = document.createElement("div");
  leftArrow.className = "nav-arrow" + (toolbarPage === 1 ? " hidden-arrow" : "");
  leftArrow.innerHTML = "◀";
  leftArrow.onclick = function () { toolbarPage = 1; renderToolbar(); };
  container.appendChild(leftArrow);

  // RENDER BUTTONS
  for (var i = 0; i < 4; i++) {
    if (pageButtons[i]) {
      (function (t) {
        var btn = document.createElement("button");
        btn.id = t.id;
        btn.className = "btn-tile";

        var displayText = t.label;
        if (t.i18n && dictionary && dictionary[currentLang] && dictionary[currentLang][t.i18n]) {
          displayText = dictionary[currentLang][t.i18n];
        }

        var gearHtml = t.configAction ? '<span class="btn-config-icon" title="Settings">⚙️</span>' : "";
        var infoHtml = '<span class="btn-info-icon" title="Ask AI">?</span>';

        btn.innerHTML = infoHtml + gearHtml +
          '<div class="tile-icon">' + t.icon + '</div>' +
          '<div class="tile-label" data-key="' + (t.i18n || '') + '">' + displayText + '</div>';

        // Main Button Click
        btn.onclick = function (e) {
          if (e.target.classList.contains('btn-config-icon') || e.target.classList.contains('btn-info-icon')) return;
          t.action();
        };

        // Settings Gear
        if (t.configAction) {
          var gear = btn.querySelector('.btn-config-icon');
          if (gear) {
            gear.onclick = function (e) {
              e.preventDefault();
              e.stopPropagation();
              t.configAction();
            };
          }
        }

        // Info Icon
        var infoBtn = btn.querySelector('.btn-info-icon');
        if (infoBtn) {
          infoBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            if (typeof createSmartTooltip === "function") {
              createSmartTooltip(e, t);
            } else {
              showToast(t.desc || t.label);
            }
          };
        }

        container.appendChild(btn);
      })(pageButtons[i]);
    } else {
      var empty = document.createElement("div");
      container.appendChild(empty);
    }
  }

  // RIGHT ARROW
  var rightArrow = document.createElement("div");
  rightArrow.className = "nav-arrow" + (toolbarPage === 2 ? " hidden-arrow" : "");
  rightArrow.innerHTML = "▶";
  rightArrow.onclick = function () { toolbarPage = 2; renderToolbar(); };
  container.appendChild(rightArrow);

  if (typeof applyButtonTheme === "function") applyButtonTheme();
}
// START: init
// ==========================================
function init() {
  // --- SAFE LAYOUT FIX ---
  document.documentElement.style.overflow = "hidden";  // Lock the outer HTML frame
  document.body.style.overflow = "hidden";             // Lock the Body frame

  var out = document.getElementById("output");
  if (out) {
    out.style.height = "100vh";       // Force content to fill the screen
    out.style.overflowY = "auto";     // Move the scrollbar to here
  }
  // -----------------------

  console.log("🚀 Init() triggered");
  window.AGREEMENT_DATA = { status: "loading" };
  console.log("🚀 Init() triggered");
  window.AGREEMENT_DATA = { status: "loading" };
  // 1. Load Data
  loadSavedSettings();
  applyButtonTheme();
  applyHeaderTheme();
  loadClauseLibrary();
  migrateRequestersToDB();
  // Initialize DB and try to restore session
  PROMPTS.initDocDB().then(() => {
    window.AGREEMENT_DATA = { status: "loading" };
    getDocumentGuid().then((docId) => {
      // Set global context for other functions to use
      window.CURRENT_DOC_ID = docId;

      console.log("📂 Current Document ID: " + docId);
      // Check if we have a persisted footer file
      // Check if we have a persisted footer file
      PROMPTS.loadDocFromDB(docId).then((record) => {
        if (record && record.data) {
          console.log("Restoring Document Session: " + record.name);

          // --- FIX: RESTORE CONTENT SO SCANNER CAN SEE IT ---
          if (record.name.toLowerCase().endsWith(".docx")) {
            // WORD DOC: Extract Text immediately using Mammoth
            if (typeof mammoth !== "undefined") {
              mammoth.convertToHtml({ arrayBuffer: record.data })
                .then(function (result) {
                  // 1. Get the text
                  var cleanText = result.value.replace(/<[^>]*>/g, ' ');

                  // 2. SAVE IT TO GLOBAL MEMORY
                  window.AGREEMENT_DATA = { name: record.name, text: cleanText };

                  // 3. Update Footer UI
                  var btn = document.querySelector("#footerBtnViewFile span");
                  if (btn) btn.innerText = "⚡";
                })
                .catch(function (e) { console.error("Mammoth Restore Error", e); });
            }
          }
          else if (record.name.toLowerCase().endsWith(".pdf")) {
            // PDF: Load Binary
            var blob = new Blob([record.data]);
            var reader = new FileReader();
            reader.onloadend = function () {
              var b64 = reader.result.split(',')[1];
              // SAVE TO GLOBAL MEMORY
              window.AGREEMENT_DATA = {
                name: record.name,
                isBase64: true,
                mime: "application/pdf",
                data: b64
              };
              var btn = document.querySelector("#footerBtnViewFile span");
              if (btn) btn.innerText = "⚡";
            };
            reader.readAsDataURL(blob);
          }
          else {
            // Text File
            var dec = new TextDecoder("utf-8");
            window.AGREEMENT_DATA = { name: record.name, text: dec.decode(record.data) };
          }
        } else {
          // No file found? Clear the loading state so scanner doesn't wait forever
          window.AGREEMENT_DATA = null;
        }
      });
    });
  });

  // Apply Sticky Laptop Mode (Restore State)
  if (userSettings.manualLaptopMode) {
    document.body.classList.add("compact-mode");
    console.log("💻 Restoring Manual Laptop Mode");
  }

  loadApprovalQueue(); // Ensure approvals are loaded
  refreshPlaybookData(); // Ensure playbook is loaded

  // 2. Setup Defaults if missing
  if (!userSettings.visibleButtons) {
    userSettings.visibleButtons = {
      error: true, legal: true, summary: true, playbook: true,
      chat: true, template: true, compare: true, email: true
    };
  }

  if (!userSettings.buttonPages) {
    userSettings.buttonPages = {
      error: 1, legal: 1, summary: 1, chat: 1,
      email: 2, template: 2, playbook: 2, compare: 2
    };
  }
  if (typeof userSettings.visibleButtons.email === "undefined")
    userSettings.visibleButtons.email = true;

  // 3. Render UI Components
  // 3. Render UI Components
  renderMainFooter();

  // --- SURGICAL FIX: Force 'All' Toolbar on every reload ---
  setTimeout(function () {
    toolbarPage = 1; // Reset to Page 1
    setWorkflowMode('default', true); // Force "All" Toolbar
  }, 100);
  setTimeout(function () {
    var footerLoader = document.getElementById("footerLoaderLocal");
    // Only highlight if it's visible (i.e. not connected to Onit)
    if (footerLoader && footerLoader.style.display !== 'none') {
      // Apply Glow
      footerLoader.style.transition = "all 0.5s ease";
      footerLoader.style.boxShadow = "0 0 10px 2px #29b6f6"; // Blue Glow
      footerLoader.style.backgroundColor = "#e3f2fd";       // Light Blue BG
      footerLoader.style.borderRadius = "4px";
      footerLoader.style.padding = "0 6px"; // Add padding for box look

      // Fade out after 4 seconds
      setTimeout(function () {
        footerLoader.style.boxShadow = "none";
        footerLoader.style.backgroundColor = "transparent";
        footerLoader.style.padding = "0";
      }, 4000);
    }
  }, 800); // Wait for footer to fully render
  // --- 4. TOOLBAR EXPANDER (Restored) ---
  // This bar appears when the toolbar collapses so you can get it back
  if (!document.getElementById("toolbarExpander")) {
    var expander = document.createElement("div");
    expander.id = "toolbarExpander";
    expander.innerHTML = `<span style="font-size: 10px; margin-right: 6px;">▼</span><span>SHOW TOOLBAR</span>`;

    // Logic: Click to show toolbar
    expander.onclick = function () {
      var tb = document.getElementById("mainToolbar");
      if (tb) tb.classList.remove("collapsed");
      this.classList.remove("visible");
      this.style.display = ""; // Clear inline styles
    };

    // Insert into Top Zone
    var topZone = document.getElementById("topZone");
    if (topZone) topZone.appendChild(expander);
    else document.body.appendChild(expander);
  }

  // --- 5. SETTINGS MENU TOGGLES ---
  var btnSettings = document.getElementById("btnHeaderSettings");
  var settingsMenu = document.getElementById("settingsMenu");

  if (btnSettings && settingsMenu) {
    btnSettings.title = "Open Settings"; // Override header title
    btnSettings.onclick = function (e) {
      e.stopPropagation();
      settingsMenu.classList.toggle("hidden");
    };

    // Close Button (X) inside the new menu
    var btnCloseX = document.getElementById("btnCloseSettingsMain");
    if (btnCloseX)
      btnCloseX.onclick = function () {
        settingsMenu.classList.add("hidden");
      };

    // Click outside to close
    document.addEventListener("click", function (e) {
      if (!settingsMenu.contains(e.target) && !btnSettings.contains(e.target)) {
        settingsMenu.classList.add("hidden");
      }
    });
    settingsMenu.onclick = function (e) {
      e.stopPropagation();
    };
  }

  // --- 6. BIND SETTINGS CONTROLS ---

  // A. Customize Toolbar
  var btnCust = document.getElementById("btnCustomizeToolbar");
  if (btnCust)
    btnCust.onclick = function () {
      settingsMenu.classList.add("hidden");
      showInterfaceConfig();
    };
  var chkSnap = document.getElementById("chkAutoSnap");



  // Bind Logic
  if (chkSnap) {
    chkSnap.checked = userSettings.autoSnapshotOnLoad;
    chkSnap.onchange = function () {
      userSettings.autoSnapshotOnLoad = this.checked;
      saveCurrentSettings();
      showToast(this.checked ? "✅ Auto-Snapshot Enabled" : "🚫 Auto-Snapshot Disabled");
    };
  }

  // B. Email Templates
  var btnEmail = document.getElementById("btnEmailSettings");
  if (btnEmail)
    btnEmail.onclick = function () {
      settingsMenu.classList.add("hidden");
      showConfigModal("email");
    };

  // B2. File Name Settings
  var btnFileName = document.getElementById("btnFileNameSettings");
  if (btnFileName)
    btnFileName.onclick = function () {
      settingsMenu.classList.add("hidden");
      showConfigModal("filename");
    };

  // C. Reset App
  var btnReset = document.getElementById("btnGlobalReset");
  if (btnReset)
    btnReset.onclick = function () {
      settingsMenu.classList.add("hidden");
      showResetConfirmationModal();
    };

  // D. Audit Trail
  var chkAudit = document.getElementById("chkAuditTrail");
  if (chkAudit) {
    chkAudit.checked = userSettings.enableAuditTrail;
    chkAudit.onchange = function () {
      userSettings.enableAuditTrail = this.checked;
      saveCurrentSettings();
      showToast(this.checked ? "📝 Audit Recording..." : "🛑 Audit Stopped");
    };
  }

  // D2. Data Integration
  var selData = document.getElementById("selDefaultDataSource");
  if (selData) {
    selData.value = userSettings.defaultDataSource || "local";
    selData.onchange = function () {
      userSettings.defaultDataSource = this.value;
      saveCurrentSettings();
      // Force Footer Refresh
      var footer = document.getElementById("main-footer");
      if (footer) footer.remove();
      renderMainFooter();
    };
  }

  var txtDataId = document.getElementById("txtDataSourceId");
  if (txtDataId) {
    txtDataId.value = userSettings.dataSourceId || "Onit";
    txtDataId.onchange = function () {
      userSettings.dataSourceId = this.value;
      saveCurrentSettings();
      // Force Footer Refresh
      var footer = document.getElementById("main-footer");
      if (footer) footer.remove();
      renderMainFooter();
    };
  }

  // E. Laptop Mode (With Forced Pin Logic)
  var chkLaptop = document.getElementById("chkLaptopMode");
  if (chkLaptop) {
    chkLaptop.checked = document.body.classList.contains("compact-mode"); // Sync state
    chkLaptop.onchange = function () {
      userSettings.manualLaptopMode = this.checked; // Track Manual Choice
      saveCurrentSettings();

      var pinChk = document.getElementById("chkPinToolbar");

      if (this.checked) {
        // ENABLE LAPTOP MODE
        document.body.classList.add("compact-mode");
        showToast("💻 Laptop Mode Enabled");

        // FORCE UNPIN (Auto-Hide)
        if (pinChk) {
          if (pinChk.checked) {
            pinChk.checked = true;
            if (typeof pinChk.onchange === "function") pinChk.onchange();
          }
          // Optional: Lock it or leave user choice?
          // User said "it should unpin". Doesn't say "lock".
          // Previous code locked it. Unlocking might be friendlier.
          // But consistency suggests locking if "Laptop Mode" implies "Unpinned".
          // I'll leave it unlocked so they CAN pin if they really want, but default is Unpinned.
          pinChk.disabled = false;
          pinChk.parentElement.style.opacity = "1";
          pinChk.parentElement.title = "";
        }
      } else {
        // DISABLE LAPTOP MODE
        document.body.classList.remove("compact-mode");
        showToast("🖥️ Standard Mode Restored");

        // Restore Normal Pin State (Default to Pinned?)
        if (pinChk) {
          // Maybe don't force change? Or restore user pref?
          // If they leave Laptop Mode, they might want Pin back.
          // I'll leave it as is or default Pin?
          // Previous code released lock.
          pinChk.disabled = false;
          pinChk.parentElement.style.opacity = "1";
          pinChk.parentElement.title = "";
        }
      }
    };

  }

  // F. Pin Toolbar Logic
  var chkPin = document.getElementById("chkPinToolbar");
  if (chkPin) {
    chkPin.checked = userSettings.pinToolbar;
    chkPin.onchange = function () {
      userSettings.pinToolbar = this.checked;
      saveCurrentSettings();

      var tb = document.getElementById("mainToolbar");
      var exp = document.getElementById("toolbarExpander");

      if (this.checked) {
        // PINNED: Show toolbar, Hide expander
        if (tb) tb.classList.remove("collapsed");
        if (exp) {
          exp.classList.remove("visible");
          exp.style.display = "";
        }
      }
    };
  }
  // --- NEW: INJECT "AUTO-SNAPSHOT" INTO MAIN MENU ---
  // G. Language Selector
  var langSel = document.getElementById("selLanguage");
  if (langSel) {
    langSel.value = localStorage.getItem("eli_language") || "en";
    langSel.onchange = function (e) {
      setLanguage(e.target.value);
    };
  }

  // H. Model Selector (Settings Menu)
  var modelSelect = document.getElementById("selAiModel");
  if (modelSelect) {
    if (userSettings.model) modelSelect.value = userSettings.model;
    modelSelect.onchange = function (e) {
      userSettings.model = e.target.value;
      saveCurrentSettings();
      showToast("🤖 Model Switched: " + e.target.options[e.target.selectedIndex].text);
    };
  }

  // 7. Header & Help Interactions
  var header = document.querySelector(".header");
  var helpIconBox = document.querySelector(".brand-icon-box");

  // A. Bind the Main Header (Home Button)
  if (header) {
    header.style.cursor = "pointer";
    header.title = "Return to Home Screen";
    header.onclick = function (e) {
      // Only go home if they didn't click the icon (Safety check)
      if (e.target.closest(".brand-icon-box")) return;

      header.style.opacity = "0.8";
      setTimeout(function () {
        header.style.opacity = "1";
      }, 150);
      showHome();
    };
  }
  // B. Bind the Icon (Force it to top)
  if (helpIconBox) {
    // 1. Force CSS to ensure it catches clicks
    helpIconBox.style.cursor = "pointer";
    helpIconBox.style.position = "relative";
    helpIconBox.style.zIndex = "1000"; // Sit above the header background
    helpIconBox.title = "Open ELI Help Center";

    // 2. Attach Listener
    helpIconBox.onclick = function (e) {
      // CRITICAL: Stop the event from bubbling to the header
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation(); // The nuclear option

      console.log("🏛️ Help Icon Clicked");
      showHelpCenter();
      return false;
    };

    // 3. Bind to the internal span too (Double safety)
    var iconSpan = helpIconBox.querySelector(".brand-icon");
    if (iconSpan) {
      iconSpan.style.pointerEvents = "none"; // Let clicks pass through to the box
    }
  }
  // --- 9. AUTO-RESIZE OBSERVER ---
  if (window.ResizeObserver) {
    var resizeObserver = new ResizeObserver(function (entries) {
      for (var entry of entries) {
        var width = entry.contentRect.width;
        var chkLaptop = document.getElementById("chkLaptopMode");
        var pinChk = document.getElementById("chkPinToolbar");
        var isCompact = document.body.classList.contains("compact-mode");

        // Threshold: 480px (Matches CSS breakpoint)
        if (width <= 480) {
          if (!isCompact) {
            document.body.classList.add("compact-mode");
            if (chkLaptop) chkLaptop.checked = true;
            showToast("💻 Auto-Switched to Laptop Mode");

            // Force Pin (UI Logic)
            if (pinChk) {
              if (!pinChk.checked) {
                pinChk.checked = true;
                if (typeof pinChk.onchange === "function") pinChk.onchange();
              }
              pinChk.disabled = true;
              pinChk.parentElement.style.opacity = "0.5";
              pinChk.parentElement.title = "Toolbar must be pinned in Laptop Mode";
            }
          }
        } else {
          // If Manual Mode enabled, IGNORE screen size checker (Sticky)
          if (!userSettings.manualLaptopMode) {
            if (isCompact) {
              document.body.classList.remove("compact-mode");
              if (chkLaptop) chkLaptop.checked = false;
              showToast("🖥️ Standard Mode Restored");

              // Release Pin
              if (pinChk) {
                pinChk.disabled = false;
                pinChk.parentElement.style.opacity = "1";
                pinChk.parentElement.title = "";
              }
            }
          }
        }
      }
    });
    resizeObserver.observe(document.body);
  }
  // START: PASTE THIS OVER THE OLD APP STORE LOGIC IN init()

  // 1. Initialize Workflow Playlists (Default)
  if (!userSettings.workflowMaps) {
    // Use JSON.parse/stringify to create a clean copy, not a reference
    userSettings.workflowMaps = JSON.parse(JSON.stringify(DEFAULT_WORKFLOW_MAPS));
  }
  // 2. Initialize Installed Apps Registry (If missing)
  if (!userSettings.installedStoreApps) {
    userSettings.installedStoreApps = [];
  }
  // 3. Register ONLY Installed Apps into Main System

  ELI_APP_STORE_CATALOG.forEach(function (app) {
    var isCore = app.category === "Core System Tools";
    var isInstalled = userSettings.installedStoreApps.indexOf(app.key) !== -1;

    // Use the helper function (ensure registerAppRuntime is defined globally below)
    if (isCore || isInstalled) {
      registerAppRuntime(app);
    }
  });

  // 4. Inject "App Store" Button into Settings Menu
  var settingsBody = document.querySelector("#settingsMenu .settings-body");
  if (settingsBody && !document.getElementById("btnOpenAppStore")) {
    var configHeader = Array.from(document.querySelectorAll('.settings-group-label')).find(el => el.innerText.includes("CONFIGURATION"));
    if (configHeader && configHeader.parentElement) {
      var storeBtn = document.createElement("div");
      storeBtn.id = "btnOpenAppStore";
      storeBtn.className = "settings-action-btn";
      storeBtn.innerHTML = "<span>🧩</span> <span><strong>ELI App Store</strong></span>";
      storeBtn.style.color = "#1565c0";
      storeBtn.style.backgroundColor = "#f0f7ff";
      storeBtn.style.borderTop = "1px solid #e3f2fd";
      storeBtn.style.borderBottom = "1px solid #e3f2fd";

      storeBtn.onclick = function () {
        document.getElementById("settingsMenu").classList.add("hidden");
        showAppStore();
      };

      // Insert after the Configuration header
      configHeader.parentElement.after(storeBtn);
    }
  }

  // 10. Initial UI State
  updateFooterName();
  runAutoBaseline();
  checkOnboarding();
  injectHistoryButton();
  // --- FIX: Connect the "Settings Bar" Store Button ---
  var headerStoreBtn = document.getElementById("btnAppStoreHeader");
  if (headerStoreBtn) {
    // Remove inline click to be safe and bind here
    headerStoreBtn.onclick = function (e) {
      e.stopPropagation(); // Stop menu from closing immediately if bubbling
      var menu = document.getElementById("settingsMenu");
      if (menu) menu.classList.add("hidden");
      showAppStore();
    };
  }
}
// ==========================================
// END: init
// ==========================================
// ==========================================
// START: runInteractiveErrorCheck
// ==========================================
function runInteractiveErrorCheck(onDoneCallback, isFastMode) {
  // Capture original setting to restore later
  var originalTrackState = userSettings.disableTrackChanges;

  // Wrapper to restore settings when leaving
  var safeCallback = function () {
    if (isFastMode) {
      // Restore original Redline setting
      userSettings.disableTrackChanges = originalTrackState;
    }
    if (typeof onDoneCallback === "function") onDoneCallback();
  };

  var p1 = validateApiAndGetText(false, "error").then(function (text) {
    var mode = userSettings.sensitivity || "Balanced";
    var instructions = "";

    // --- FAST MODE LOGIC ---
    if (isFastMode) {
      userSettings.disableTrackChanges = true;
      instructions = `TASK: CRITICAL PROOFREADING ONLY.
        1. Identify spelling mistakes.
        2. Identify broken cross-references (e.g. 'Error! Source not found').
        3. Identify severe grammar errors that make text unreadable.
        
        IGNORE: Style, passive voice, definitions, formatting, punctuation, and politeness suggestions.
        OUTPUT: Only report actual errors.`;
    } else {
      // --- STANDARD LOGIC ---
      var mode = userSettings.sensitivity || "Balanced";
      instructions = "TASK 1: GRAMMAR & SPELLING\nIdentify critical grammatical errors, spelling mistakes, and awkward phrasing.\n";

      if (userSettings.checkBrokenRefs) instructions += "\nTASK 2: BROKEN CROSS-REFERENCES\nIdentify broken internal references.\n";
      else instructions += "\nExclusion Rule: Do NOT report broken cross-references.\n";

      if (mode === "Picky") instructions += "\nTASK 3: DEFINITION X-RAY\nScan for Capitalized words not defined in the text.\n";

      if (userSettings.dialect === "UK") instructions += "\nIMPORTANT: Enforce UK/British spelling. Flag US spellings.\n";
      else instructions += "\nIMPORTANT: Enforce US spelling. Flag UK spellings.\n";

      if (userSettings.strictMode) instructions += "\nSTRICT MODE: Flag passive voice and weak verbs.\n";

      instructions += getFeedbackInstructions();
    }
    // ----------------------
    if (userSettings.dialect === "UK")
      instructions +=
        "\nIMPORTANT: Enforce UK/British spelling (Colour, Centre). Flag US spellings.\n";
    else
      instructions += "\nIMPORTANT: Enforce US spelling. Flag UK spellings.\n";

    if (userSettings.strictMode)
      instructions += "\nSTRICT MODE: Flag passive voice and weak verbs.\n";

    instructions += getFeedbackInstructions();

    return callGemini(API_KEY, PROMPTS.ERROR_CHECK(instructions, text)).then(
      function (jsonString) {
        var rawData = tryParseGeminiJSON(jsonString);
        return validateErrorData(rawData); // <--- Validation happens here
      },
    );
  });

  // --- CHANGE: Conditional Execution ---

  // --- OPTIMIZATION: SKIP FOOTER IN FAST MODE ---
  // If we are in Fast Mode, force this to false regardless of settings.
  var shouldCheck = (!isFastMode && userSettings.checkFooterRef);

  // If true, run scan. If false, return empty promise immediately.
  var p2 = shouldCheck ? checkFooterConsistency() : Promise.resolve([]);
  // ----------------------------------------------

  Promise.all([p1, p2])
    .then(function (results) {
      var aiErrors = results[0] || [];
      var footerErrors = results[1] || [];
      var allErrors = footerErrors.concat(aiErrors);

      if (allErrors.length === 0) {
        showOutput(
          "✅ No errors found in Body" +
          (userSettings.checkFooterRef ? " or Footers." : "."),
        );

        // --- INSERT THIS BLOCK ---
        if (typeof onDoneCallback === "function") {
          var btn = document.createElement("button");
          btn.className = "btn-action";
          // Check wizard flag for smart label
          btn.innerText = (window.eliWizardMode) ? "💾 Next: Save and/or Send ➔" : "⬅ Done / Back to Hub";
          btn.style.marginTop = "15px";
          btn.onclick = onDoneCallback;
          document.getElementById("output").appendChild(btn);
        }
        // ------------------------

      } else {
        renderCards(allErrors, "error", onDoneCallback);
        // --- NEW: INJECT BOTTOM BUTTON ---
        if (typeof onDoneCallback === "function") {
          var btnBottom = document.createElement("button");
          btnBottom.className = "btn-action";
          btnBottom.innerText = (window.eliWizardMode) ? "💾 Next: Save and/or Send ➔" : "⬅ Done / Back to Hub";
          // Distinct Style for Bottom Button
          btnBottom.style.cssText = "width:100%; margin-top:20px; margin-bottom:20px; background:#fff3e0; color:#e65100; border:1px solid #ffe0b2; font-weight:bold;";
          btnBottom.onclick = safeCallback;
          document.getElementById("output").appendChild(btnBottom);
        }
      }
    })
    .catch(handleError)
    .finally(function () {
      setLoading(false);
    });
}
// ==========================================
// END: runInteractiveErrorCheck
// ==========================================
// ==========================================
// START: runAiAssist
// ==========================================
function runAiAssist() {
  validateApiAndGetText(false, "legal")
    .then(function (text) {
      // Direct call to AI (No Guardrails)
      return callGemini(API_KEY, PROMPTS.LEGAL_ASSIST(text));
    })
    .then(function (result) {
      processGeminiResponse(result, "legal");
    })
    .catch(handleError)
    .finally(function () {
      setLoading(false);
    });
}
// ==========================================
// END: runAiAssist
// ==========================================
// ==========================================
// START: runSummary
// ==========================================
function runSummary() {
  validateApiAndGetText(false, "summary")
    .then(function (text) {
      // 1. LOAD PERSISTED SETTINGS
      var saved = userSettings.summaryPreferences || {};
      var audience = saved.audience || "General Audience";

      // 2. LOAD PERSISTED KEYS (The Fix)
      var keysList = saved.keys || [];

      // Check if we have custom keys in local storage that aren't in userSettings yet
      var storedKeys = localStorage.getItem("colgate_custom_summary_keys");
      if (storedKeys && keysList.length === 0) {
        try {
          var localCustom = JSON.parse(storedKeys);
          // If the user hasn't "Saved Settings" yet, use these as a fallback
          keysList = [
            "Parties",
            "Effective Date",
            "Term Length",
            "Liability Cap",
            "Governing Law",
            "Financials",
          ];
        } catch (e) { }
      }

      // 3. BUILD THE EXTRACTION STRING
      var keysString = null;
      if (keysList && keysList.length > 0) {
        keysString = "";
        for (var i = 0; i < keysList.length; i++) {
          var k = keysList[i];
          var safeKey = k.replace(/ /g, "_");
          keysString +=
            '"' +
            safeKey +
            '": { "value": "extracted info...", "quote": "exact text..." },\n';
        }
      }

      // 4. GENERATE
      generateSummaryWithAudience(text, audience, keysString);
    })
    .catch(handleError);
}
// ==========================================
// END: runSummary
// ==========================================
var currentSummaryAudience = "General Audience";

// 3. GEAR BUTTON: Configuration Modal
// ==========================================
// START: configureSummary
// ==========================================
function configureSummary() {
  validateApiAndGetText(false, null)
    .then(function (text) {
      setLoading(false);
      showSummaryAudienceSelection(text);
    })
    .catch(handleError);
}
// ==========================================
// END: configureSummary
// ==========================================
// ==========================================
// START: generateSummaryWithAudience
// ==========================================
function generateSummaryWithAudience(text, audienceDesc, customKeysString) {
  setLoading(true, "summary");

  // If specific keys weren't passed (e.g. from One-Click run), we leave it null.
  // The PROMPTS.SUMMARY function handles the default fallback.
  var finalKeys = customKeysString || null;

  callGemini(API_KEY, PROMPTS.SUMMARY(audienceDesc, text, finalKeys))
    .then(function (result) {
      processGeminiResponse(result, "summary");
    })
    .catch(handleError)
    .finally(function () {
      setLoading(false);
    });
}
// ==========================================
// END: generateSummaryWithAudience
// ==========================================
/// ==========================================
// START: showSummaryAudienceSelection (PREMIUM UNIFIED DESIGN)
// ==========================================
function showSummaryAudienceSelection(text) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- DATA LOADING ---
  var savedPrefs = userSettings.summaryPreferences || { audience: "lawyer", customInput: "", keys: [] };
  var localCustomKeys = [];
  try { localCustomKeys = JSON.parse(localStorage.getItem("colgate_custom_summary_keys") || "[]"); } catch (e) { }

  var hardcodedDefaults = [
    "Parties", "Effective Date", "Term Length", "Liability Cap", "Governing Law", "Financials",
    "Termination Notice", "Indemnity", "Jurisdiction", "Renewal", "Payment Terms", "Late Fees", "Total Contract Value", "Price Increases"
  ];
  var combined = hardcodedDefaults.concat(localCustomKeys);
  var allKeys = combined.filter((v, i, a) => a.indexOf(v) === i); // Dedupe

  var selectedKeys = (savedPrefs.keys && savedPrefs.keys.length > 0) ? savedPrefs.keys.slice() : ["Parties", "Effective Date", "Term Length", "Liability Cap", "Governing Law", "Financials"];

  // --- UI RENDER ---
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding: 0; border: 1px solid #d1d5db; border-left: 5px solid #1565c0; background: #fff; border-radius:0px;";

  // PREMIUM HEADER
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 12px; position: relative; overflow: hidden;`;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
          SYSTEM CONFIGURATION
      </div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
          <span>📋</span> Summary Settings
      </div>
  `;
  card.appendChild(header);

  // BODY
  var body = document.createElement("div");
  body.style.padding = "20px";

  body.innerHTML = `
        <label class="tool-label">Target Audience</label>
        <select id="sumAudienceSel" class="audience-select" style="margin-bottom: 15px; border-radius:0px;">
            <option value="lawyer">Legal Counsel (Standard)</option>
            <option value="non-lawyer">Business Stakeholder (Simple)</option>
            <option value="finance">Finance Team (Commercial Terms)</option>
            <option value="custom">Custom / Specific...</option>
        </select>
        <input type="text" id="sumCustomInput" class="audience-input hidden" style="margin-top: -10px; margin-bottom: 20px; border-radius:0px;" placeholder="E.g. The Marketing Team...">
        
        <div class="tool-label" style="display:flex; justify-content:space-between; border-bottom:1px solid #eee; padding-bottom:4px; margin-bottom:8px;">
            <span>Key Terms to Extract</span>
            <span id="termCounter" style="font-weight:700; color:#1565c0;">Loading...</span>
        </div>
        <div id="termListContainer" style="background:#f9fafb; border:1px solid #e0e0e0; padding:10px; height:180px; overflow-y:auto; margin-bottom:15px; border-radius:0px;"></div>
        
        <div style="display:flex; gap:8px; margin-bottom:10px;">
            <input type="text" id="sumAddTerm" class="tool-input" style="margin-bottom:0; font-size:11px; height:28px; border-radius:0px;" placeholder="+ Add custom term...">
            <button id="btnSumAdd" class="btn-action" style="height:28px; border-radius:0px;">Add</button>
        </div>
  `;
  card.appendChild(body);

  // FOOTER
  var footer = document.createElement("div");
  footer.style.cssText = "padding:15px 20px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px;";

  var btnCancel = document.createElement("button");
  btnCancel.className = "btn-platinum-config";
  btnCancel.style.flex = "1";
  btnCancel.innerText = "CANCEL";

  var btnGen = document.createElement("button");
  btnGen.className = "btn-platinum-config primary";
  btnGen.style.cssText = "flex:1; font-weight:800;";
  btnGen.innerText = "SAVE & GENERATE";

  footer.appendChild(btnCancel);
  footer.appendChild(btnGen);
  card.appendChild(footer);
  out.appendChild(card);

  // --- LOGIC ---
  var select = document.getElementById("sumAudienceSel");
  var customInput = document.getElementById("sumCustomInput");
  var termContainer = document.getElementById("termListContainer");
  var counter = document.getElementById("termCounter");
  var addInput = document.getElementById("sumAddTerm");

  // Restore State
  if (savedPrefs.audience) {
    if (["lawyer", "non-lawyer", "finance"].indexOf(savedPrefs.audience) !== -1) select.value = savedPrefs.audience;
    else { select.value = "custom"; customInput.classList.remove("hidden"); customInput.value = savedPrefs.audience; }
  }

  // Render List Helper
  function renderList() {
    termContainer.innerHTML = "";
    allKeys.forEach(function (key) {
      var isChecked = selectedKeys.indexOf(key) !== -1;
      var row = document.createElement("label");
      row.style.cssText = "display:flex; align-items:center; padding:6px 0; font-size:12px; color:#333; cursor:pointer; border-bottom:1px solid #eee; margin:0;";

      var chk = document.createElement("input");
      chk.type = "checkbox";
      chk.style.marginRight = "10px";
      chk.checked = isChecked;

      chk.onchange = function () {
        if (this.checked) {
          if (selectedKeys.length >= 6) { this.checked = false; showToast("⚠️ Max 6 terms allowed."); }
          else selectedKeys.push(key);
        } else {
          var idx = selectedKeys.indexOf(key);
          if (idx > -1) selectedKeys.splice(idx, 1);
        }
        updateCount();
      };

      row.appendChild(chk);
      row.appendChild(document.createTextNode(key));
      termContainer.appendChild(row);
    });
    updateCount();
  }

  function updateCount() { counter.innerText = selectedKeys.length + "/6 Selected"; }

  select.onchange = function () {
    if (select.value === "custom") customInput.classList.remove("hidden");
    else customInput.classList.add("hidden");

    // Auto-select presets
    if (select.value === "finance") selectedKeys = ["Parties", "Payment Terms", "Late Fees", "Total Contract Value", "Renewal", "Price Increases"];
    else if (select.value !== "custom") selectedKeys = ["Parties", "Effective Date", "Term Length", "Liability Cap", "Governing Law", "Financials"];

    renderList();
  };

  // Add Button
  document.getElementById("btnSumAdd").onclick = function () {
    var val = addInput.value.trim();
    if (val && allKeys.indexOf(val) === -1) {
      localCustomKeys.push(val);
      localStorage.setItem("colgate_custom_summary_keys", JSON.stringify(localCustomKeys));
      allKeys.push(val);
      if (selectedKeys.length < 6) selectedKeys.push(val);
      renderList();
      addInput.value = "";
    }
  };

  // Save/Generate Button
  btnGen.onclick = function () {
    var audienceVal = select.value === "custom" ? customInput.value : select.value;
    userSettings.summaryPreferences = { audience: audienceVal, customInput: customInput.value, keys: selectedKeys };
    if (select.value === "custom") currentSummaryAudience = customInput.value;
    else currentSummaryAudience = select.options[select.selectedIndex].text;

    saveCurrentSettings();
    showToast("✅ Settings Saved");
    generateSummaryWithAudience(text, currentSummaryAudience); // Trigger run immediately
  };

  btnCancel.onclick = showHome;
  renderList();
}
// ==========================================
// END: showSummaryAudienceSelection
// ==========================================
// REPLACE: runTemplateScan (Deep Scan + Prefix Support)
// ==========================================
function runTemplateScan() {
  setLoading(true, "template");

  Word.run(async function (context) {
    // 1. COLLECT ALL TEXT (Body + Headers + Footers)
    // We use the same robust "Collection Strategy" as the Email Scanner
    var fullText = "";
    var rangesToLoad = [];

    // A. Body
    var body = context.document.body;
    body.load("text");
    rangesToLoad.push(body);

    // B. Headers & Footers (Every Section, Every Type)
    var sections = context.document.sections;
    sections.load("items");
    await context.sync();

    sections.items.forEach(function (section) {
      ['Primary', 'FirstPage', 'EvenPages'].forEach(function (type) {
        try { rangesToLoad.push(section.getHeader(type)); } catch (e) { }
        try { rangesToLoad.push(section.getFooter(type)); } catch (e) { }
      });
    });

    // C. Load Text for All Ranges
    // We strictly load only the 'text' property to be efficient
    rangesToLoad.forEach(function (r) { r.load("text"); });

    await context.sync();

    // D. Concatenate
    rangesToLoad.forEach(function (r) {
      fullText += r.text + "\n";
    });

    return fullText;
  })
    .then(function (text) {
      // 2. REGEX (Now supports $! prefix)
      // Captures: 1=(Prefix), 2=(Name), 3=(Suffix)
      var bracketRegex = /(\$!|\$|!)?\[\[(.*?)\]\]/g;

      var curlyRegex = /{{([^}]+)}}(\*)?/g;

      var matches = new Set();
      var match;

      // Find all bracketed items (e.g. [Name], $![Fee])
      while ((match = bracketRegex.exec(text)) !== null) {
        // match[0] is the full string "$![Fee]"
        if (match[0].length > 2) matches.add(match[0]);
      }

      // Find all curly items (e.g. {{Date}})
      while ((match = curlyRegex.exec(text)) !== null) {
        if (match[0].length > 4) matches.add(match[0]);
      }

      setLoading(false);
      renderTemplateForm(Array.from(matches));
    })
    .catch(function (e) {
      setLoading(false);
      showOutput("Error scanning document: " + e.message, true);
    });
}

// ==========================================
function renderTemplateForm(placeholders) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- HEADER (Standard) ---
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 12px;`;
  header.innerHTML = `<div style="font-family:'Georgia', serif; font-size:16px; color:#fff;"><span>⚡</span> Auto-Fill Templates</div>`;
  out.appendChild(header);

  if (!placeholders || placeholders.length === 0) {
    // RESTORED: Detailed instructions style
    out.innerHTML += "<div class='text-block' style='background:#f5f5f5; border-color:#ddd; color:#555'><b>No placeholders detected.</b><br><br>1. Open an ELI template that has [[Brackets]].</div>";

    var btnHomeEmpty = document.createElement("button");
    btnHomeEmpty.className = "btn-action";
    btnHomeEmpty.innerText = "🏠 Home";
    btnHomeEmpty.style.width = "100%"; // Restored full width
    btnHomeEmpty.style.marginTop = "20px";
    btnHomeEmpty.onclick = showHome;
    out.appendChild(btnHomeEmpty);
    return;
  }

  // --- PREMIUM UPLOAD BOX (CONDITIONAL) ---

  var sourceData = window.AGREEMENT_DATA || AGREEMENT_DATA;
  // FIX: Check global variable explicitly if window.AGREEMENT_DATA is undefined
  var hasAgreement = (typeof AGREEMENT_DATA !== 'undefined' && AGREEMENT_DATA && AGREEMENT_DATA.name) ||
    (window.AGREEMENT_DATA && window.AGREEMENT_DATA.name);

  if (hasAgreement) {
    // Use the name from whichever source was valid
    var docName = window.AGREEMENT_DATA ? window.AGREEMENT_DATA.name : AGREEMENT_DATA.name;
    // CASE A: File Loaded in Footer
    var loadedBox = document.createElement("div");
    loadedBox.className = "tool-box slide-in";
    loadedBox.style.cssText = "padding:20px; border:1px solid #90caf9; background:#f0f7ff; border-left:5px solid #1565c0; margin-bottom:20px; border-radius:0px;";

    loadedBox.innerHTML = `
          <div style="display:flex; justify-content:space-between; align-items:center;">
              <div>
                  <div style="font-size:11px; font-weight:700; color:#1565c0; text-transform:uppercase;">Using Loaded Source</div>
                  <div style="font-size:14px; font-weight:600; color:#333; margin-top:4px;">📄 ${docName}</div>
              </div>
              <button id="btnFooterAutoFill" class="btn-action" style="width:auto; padding:8px 16px; font-size:11px;">⚡ Auto-Fill Fields</button>
          </div>
          <div id="footerFillStatus" style="font-size:11px; color:#555; margin-top:8px; display:none;"></div>
      `;
    out.appendChild(loadedBox);

    // Bind Logic
    var btnFill = loadedBox.querySelector("#btnFooterAutoFill");
    var statusDiv = loadedBox.querySelector("#footerFillStatus");

    // REPLACED HANDLER (Short & Fast)
    btnFill.onclick = function () {
      btnFill.innerHTML = "⏳ Filling...";
      statusDiv.style.display = "block";
      statusDiv.innerHTML = "<div class='spinner' style='width:12px;height:12px;border-width:2px;display:inline-block;'></div> Reading RAM data...";
      btnFill.disabled = true;

      // 1. GET TEXT DIRECTLY FROM RAM (No DB call needed)
      if (!sourceData || (!sourceData.text && !sourceData.isBase64)) {
        statusDiv.innerHTML = "❌ Error: No readable text in memory.";
        btnFill.disabled = false;
        btnFill.innerHTML = "⚡ Auto-Fill Fields";
        return;
      }

      var textToAnalyze = sourceData.text || "";

      // 2. RUN AI
      statusDiv.innerHTML = "<div class='spinner' style='width:12px;height:12px;border-width:2px;display:inline-block;'></div> Analyzing data...";

      var cleanKeys = placeholders.map(function (ph) {
        var clean = ph.replace(/[\[\]{}]/g, "").replace(/^(\$!|\$|!)/, "");
        if (clean.includes("||")) return clean.split("||")[0].trim();
        return clean;
      });

      var prompt = PROMPTS.TEMPLATE_MAP(
        cleanKeys,
        textToAnalyze.substring(0, 500000)
      );

      callGemini(API_KEY, prompt).then(function (result) {
        try {
          var mapping = tryParseGeminiJSON(result);
          var inputs = out.querySelectorAll(".template-input");
          var count = 0;
          inputs.forEach(function (inp) {
            var val = mapping[inp.dataset.target] || mapping[inp.dataset.aiKey];
            if (val) {
              if (typeof smartFillInput === 'function') {
                smartFillInput(inp, val);
              } else {
                inp.value = val;
                inp.style.borderColor = "#2e7d32";
                inp.style.backgroundColor = "#fff";
              }
              count++;
            } else {
              inp.style.borderColor = "#c62828";
              inp.style.backgroundColor = "#ffebee";
            }
          });
          // --- SURGICAL EDIT START ---
          var missing = inputs.length - count;
          statusDiv.innerHTML = "✅ Filled " + count + " fields. <span style='color:#c62828'>Missing " + missing + " fields.</span>";
          // --- SURGICAL EDIT END ---

          btnFill.disabled = false;
          btnFill.innerHTML = "⚡ Auto-Fill Fields";
        } catch (err) {
          console.error(err);
          statusDiv.innerHTML = "❌ Error parsing AI response.";
          btnFill.disabled = false;
          btnFill.innerHTML = "⚡ Auto-Fill Fields";
        }
      }).catch(function (err) {
        statusDiv.innerHTML = "❌ API Error: " + err.message;
        btnFill.disabled = false;
        btnFill.innerHTML = "⚡ Auto-Fill Fields";
      });
    };

  } else {
    // CASE B: Manual Upload (Original Logic)
    var uploadBox = document.createElement("div");
    uploadBox.className = "tool-box slide-in";
    uploadBox.style.padding = "0"; // Removed padding for full-width style
    uploadBox.style.border = "none"; // Removed standard border
    uploadBox.style.marginBottom = "20px";
    uploadBox.style.borderRadius = "0px";

    // New Drop Zone Container
    var dropZone = document.createElement("div");
    dropZone.className = "file-upload-zone";
    dropZone.style.cssText = `
          border: 2px dashed #90caf9; 
          background-color: #f0f7ff; 
          padding: 20px; 
          text-align: center; 
          cursor: pointer; 
          transition: all 0.2s ease;
          border-radius: 4px;
      `;

    // Hover Effects
    dropZone.onmouseover = function () { this.style.backgroundColor = "#e3f2fd"; this.style.borderColor = "#1565c0"; };
    dropZone.onmouseout = function () { this.style.backgroundColor = "#f0f7ff"; this.style.borderColor = "#90caf9"; };

    // New Content with Icon & Subtext
    dropZone.innerHTML = `
          <div style="font-size: 24px; margin-bottom: 8px; color: #1565c0;">📂</div>
          <div style="font-size: 12px; font-weight: 700; color: #1565c0; text-transform: uppercase; margin-bottom: 4px;">
              Import Data Source
          </div>
          <div style="font-size: 10px; color: #546e7a; line-height: 1.4;">
              Click to upload a filled document or data file.<br>
              <span style="opacity: 0.7;">(Supports .docx, .txt, .csv, .json)</span>
          </div>
      `;

    var fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.style.display = "none";

    dropZone.onclick = function () { fileInput.click(); };
    fileInput.onchange = function (e) {
      // Added "Reading File..." spinner state
      dropZone.innerHTML = `<div class="spinner" style="width:16px;height:16px;border-width:2px;display:inline-block;border-color:#1565c0 transparent #1565c0 transparent;"></div><div style="font-size:11px;font-weight:bold;color:#1565c0;margin-top:6px;">Reading File...</div>`;
      handleTemplateFileUpload(e, placeholders, out);
    };

    uploadBox.appendChild(dropZone);
    uploadBox.appendChild(fileInput);
    out.appendChild(uploadBox);
  }

  // --- MANUAL INPUTS ---
  var formContainer = document.createElement("div");

  placeholders.forEach(function (ph) {
    // CLEANING LOGIC: Removes "$", "!" from the display label (Removed *)
    var rawText = ph.replace(/[\[\]{}]/g, "")  // Remove brackets
      .replace(/^(\$!|\$|!)/, ""); // Removes $! OR $ OR !

    var labelText = rawText;
    var aiKey = rawText;

    // --- NEW: PIPE SUPPORT [Key || Label] ---
    if (rawText.includes("||")) {
      var parts = rawText.split("||");
      if (parts.length > 1) {
        aiKey = parts[0].trim();
        labelText = parts[1].trim();
      }
    }

    var fieldWrapper = document.createElement("div");
    fieldWrapper.style.marginBottom = "2px";

    // --- UPDATED LABEL (Flex container) ---
    var label = document.createElement("label");
    label.className = "tool-label";
    label.style.marginBottom = "2px";
    label.style.display = "flex";
    label.style.justifyContent = "space-between";
    label.style.alignItems = "center";

    // 1. Label Text
    var textSpan = document.createElement("span");
    textSpan.innerText = labelText;
    label.appendChild(textSpan);

    // 2. Locate Button (📍)
    var locBtn = document.createElement("span");
    locBtn.innerHTML = "Locate📍";
    locBtn.title = "Locate in Document";
    locBtn.style.cssText = "cursor:pointer; font-size:12px; opacity:0.6; transition:all 0.2s; padding:2px 6px; border-radius:4px; background:rgba(0,0,0,0.03);";

    // Hover Effects
    locBtn.onmouseover = function () { this.style.opacity = "1"; this.style.background = "#e3f2fd"; this.style.color = "#1565c0"; };
    locBtn.onmouseout = function () { this.style.opacity = "0.6"; this.style.background = "rgba(0,0,0,0.03)"; this.style.color = ""; };

    label.appendChild(locBtn);

    // --- NEW: SPECIAL PAYMENT INDICATOR ---
    if (ph.startsWith("$![")) {
      var payBadge = document.createElement("span");
      payBadge.innerText = "Payment terms helper will open on completion.";
      payBadge.style.cssText = "font-size:10px; color:#2e7d32; font-style:italic; margin-right:auto; margin-left:8px; display:block;";
      label.insertBefore(payBadge, locBtn);
    }

    label.appendChild(locBtn);
    fieldWrapper.appendChild(label);

    var input = document.createElement("textarea");
    // --- BIND LOCATE CLICK ---
    locBtn.onclick = function () {
      locatePlaceholderOrValue(ph, input.value, locBtn);
    };
    input.className = "tool-input template-input auto-expand";

    // Store the raw tag (with !) so the logic can find it later
    input.dataset.target = ph;
    input.dataset.aiKey = aiKey; // Store clean key for AI matching
    input.placeholder = "Enter " + labelText;

    input.rows = 1;
    input.style.cssText = "width:100%; resize:none; overflow:hidden; min-height:28px;";
    input.addEventListener("input", function () {
      this.style.height = 'auto';
      this.style.height = (this.scrollHeight) + 'px';

      // FIX: Clear hidden "smart" data if user manually types
      // This ensures we rely on the manual input value instead of stale auto-fill data
      if (this.dataset.fullTerms) {
        delete this.dataset.fullTerms;
        // Visual Reset
        this.style.backgroundColor = "";
        this.style.fontWeight = "";
        this.title = "";
      }
    });

    fieldWrapper.appendChild(input);
    formContainer.appendChild(fieldWrapper);
  });

  var btnRow = document.createElement("div");
  btnRow.style.marginTop = "20px";
  var applyBtn = document.createElement("button");
  applyBtn.className = "btn-action";
  applyBtn.innerHTML = "FILL ALL FIELDS";
  applyBtn.style.width = "100%";
  applyBtn.onclick = function () { applyTemplateValues(formContainer); };

  btnRow.appendChild(applyBtn);
  formContainer.appendChild(btnRow);
  out.appendChild(formContainer);
  // --- RESTORED: SMART FILL TOOL (Paste Blob) ---
  var smartDivider = document.createElement("div");
  smartDivider.style.cssText = "margin: 30px 0 15px 0; border-top: 1px dashed #ddd; text-align:center; height: 10px; cursor: pointer;";
  smartDivider.title = "Click to show/hide tools";
  smartDivider.innerHTML = "<span id='smartLabel' style='background:#fff; padding:0 10px; color:#999; font-size:10px; position:relative; top:-10px; font-weight:700;'>SMART TOOLS ▼</span>";
  out.appendChild(smartDivider);

  var magicBox = document.createElement("div");
  magicBox.className = "tool-box slide-in";
  magicBox.style.marginBottom = "10px";
  magicBox.style.borderLeft = "4px solid #1565c0";
  magicBox.style.display = "none";
  magicBox.id = "smartToolsContainer";

  smartDivider.onclick = function () {
    var container = document.getElementById("smartToolsContainer");
    var label = document.getElementById("smartLabel");
    if (container.style.display === "none") {
      container.style.display = "block";
      label.innerText = "SMART TOOLS ▲";
      label.style.color = "#1565c0";
    } else {
      container.style.display = "none";
      label.innerText = "SMART TOOLS ▼";
      label.style.color = "#999";
    }
  };

  magicBox.innerHTML = `
        <div class="tool-label" style="color:#1565c0;">✨ Smart Fill (Paste Text)</div>
        <div id="magicPasteArea" class="" style="margin-top:10px;">
            <textarea id="pasteInput" class="tool-input" placeholder="Paste messy text here (e.g. email signature, slack message)..." style="height:80px; font-size:11px;"></textarea>
        </div>
    `;

  var btnMagic = document.createElement("button");
  btnMagic.id = "btnMagicFill";
  btnMagic.className = "btn-action";
  btnMagic.style.cssText = `
        width:100%; margin-top:8px;
        background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
        border: 1px solid #b0bec5;
        color: #546e7a;
        font-weight: 800;
        border-radius: 0px;
        padding: 8px;
        box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
    `;
  btnMagic.innerText = "✨ Magic Fill";
  btnMagic.onclick = function () {
    var text = document.getElementById("pasteInput").value;
    runSmartFill(text, this);
  };

  magicBox.querySelector("#magicPasteArea").appendChild(btnMagic);
  out.appendChild(magicBox);
}
// ==========================================
// START: handleTemplateFileUpload
// ==========================================
function handleTemplateFileUpload(event, placeholders, container) {
  var file = event.target.files[0];
  if (!file) return;

  var btnZone = container.querySelector(".file-upload-zone");

  // --- Helper: What to do once we have the text ---
  // ==========================================
  // START: runAiExtraction
  // ==========================================
  function runAiExtraction(contentText) {
    // 1. Animation Logic
    var msgs = LOADING_MSGS.template;
    var msgIdx = 0;
    btnZone.innerHTML =
      "<div class='spinner' style='width:16px;height:16px;border-width:2px;'></div> " +
      msgs[0];

    var tmplInterval = setInterval(function () {
      msgIdx = (msgIdx + 1) % msgs.length;
      if (msgIdx === 0) msgIdx = 1;
      btnZone.innerHTML =
        "<div class='spinner' style='width:16px;height:16px;border-width:2px;'></div> " +
        msgs[msgIdx];
    }, 1500);

    // 2. Call AI
    // NEW: Clean placeholders for AI (Handle || pipe)
    var cleanKeys = placeholders.map(function (ph) {
      var clean = ph.replace(/[\[\]{}]/g, "")
        .replace(/^(\$!|\$|!)/, "");
      if (clean.includes("||")) return clean.split("||")[0].trim();
      return clean;
    });

    var prompt = PROMPTS.TEMPLATE_MAP(
      cleanKeys,
      contentText.substring(0, 8000),
    ); // Increased limit slightly
    callGemini(API_KEY, prompt)
      .then(function (result) {
        clearInterval(tmplInterval);
        try {
          var mapping = tryParseGeminiJSON(result);
          // NOTE: Use the class selector we fixed earlier (.template-input)
          var inputs = container.querySelectorAll(".template-input");
          var count = 0;
          inputs.forEach(function (inp) {
            // MATCH: Try Exact Target OR Clean AI Key
            var val = mapping[inp.dataset.target] || mapping[inp.dataset.aiKey];
            if (val) {
              inp.value = val;
              inp.style.borderColor = "#2e7d32";

              // Remove error styling if previously set
              inp.style.backgroundColor = "#fff";

              // Trigger auto-grow for the new Textareas
              inp.style.height = "auto";
              inp.style.height = inp.scrollHeight + "px";

              count++;
            } else {
              // HIGHLIGHT MISSING FIELDS IN RED
              inp.style.borderColor = "#c62828";
              inp.style.backgroundColor = "#ffebee";
            }
          });
          var missing = inputs.length - count;
          btnZone.innerHTML = "✅ Filled " + count + " fields. Missing " + missing + " fields.";
        } catch (err) {
          console.error(err);
          btnZone.innerHTML = "❌ Error parsing AI response.";
        }
      })
      .catch(function (err) {
        clearInterval(tmplInterval);
        btnZone.innerHTML = "❌ API Error: " + err.message;
      });
  }
  // ==========================================
  // END: runAiExtraction
  // ==========================================

  if (file.name.endsWith(".docx")) {
    // Check for Mammoth Library
    if (typeof mammoth === "undefined") {
      showOutput("❌ Error: Mammoth.js library is missing.", true);
      return;
    }

    btnZone.innerHTML =
      "<div class='spinner' style='width:16px;height:16px;border-width:2px;'></div> Unzipping DOCX...";

    var reader = new FileReader();
    reader.onload = function (e) {
      // Extract text from DOCX binary
      mammoth
        .extractRawText({ arrayBuffer: e.target.result })
        .then(function (result) {
          runAiExtraction(result.value); // Pass extracted text to AI
        })
        .catch(function (err) {
          console.error(err);
          btnZone.innerHTML = "❌ Error reading DOCX file.";
        });
    };
    reader.readAsArrayBuffer(file);
  } else {
    // Handle .txt, .json, .csv
    var reader = new FileReader();
    reader.onload = function (e) {
      runAiExtraction(e.target.result);
    };
    reader.readAsText(file);
  }
}
// ==========================================
// END: handleTemplateFileUpload
// ==========================================
// ==========================================
// REPLACE: runSmartFill (Updated to use smartFillInput)
// ==========================================
function runSmartFill(pastedText, btnElement) {
  var formContainer = document.querySelector(".output-area");
  // Fallback selector
  if (!formContainer) formContainer = document.getElementById("output");

  var inputs = formContainer.querySelectorAll(".template-input");
  var placeholders = [];

  inputs.forEach(function (inp) {
    if (inp.dataset.target) placeholders.push(inp.dataset.target);
  });

  if (placeholders.length === 0) { showToast("⚠️ No fields to fill."); return; }
  if (!pastedText || pastedText.trim().length === 0) { showToast("⚠️ Paste text first."); return; }

  var originalText = btnElement.innerHTML;
  btnElement.innerHTML = '<div class="spinner" style="width:12px;height:12px;border-width:2px;"></div>';
  btnElement.disabled = true;

  var prompt = PROMPTS.TEMPLATE_MAP(placeholders, pastedText);

  callGemini(API_KEY, prompt).then(function (result) {
    var mapping = tryParseGeminiJSON(result);
    var count = 0;
    inputs.forEach(function (inp) {
      if (mapping[inp.dataset.target]) {
        // USE NEW HELPER HERE
        smartFillInput(inp, mapping[inp.dataset.target]);
        count++;
      }
    });
    showToast(`✨ Filled ${count} fields.`);
  })
    .catch(handleError)
    .finally(function () {
      btnElement.innerHTML = originalText;
      btnElement.disabled = false;
    });
}
// ==========================================
// REPLACE: applyTemplateValues (Literal Cleanup Mode)
// ==========================================
function applyTemplateValues(container) {
  var inputs = container.querySelectorAll(".template-input");
  var standardReplacements = [];

  // 1. We will store the EXACT text of every special tag we find (e.g. "$![Fee]")
  // This allows us to search for them explicitly later, bypassing wildcard bugs.
  var tagsToKill = [];

  var feeTrigger = null;
  var paymentCandidates = []; // Fix: Defined here to collect candidates from loop
  var missingFields = [];

  // --- A. GATHER INPUTS ---
  inputs.forEach(function (input) {
    var val = input.value;

    // Check Missing
    if (!val || val.trim() === "") {
      input.style.border = "2px solid #c62828";
      input.style.backgroundColor = "#ffebee";
      var label = input.dataset.target.replace(/[\[\]]/g, "") || "Unknown Field";
      missingFields.push("• " + label);
    } else {
      input.style.border = "1px solid #ced4da";
      input.style.backgroundColor = "#fff";
    }

    // Parse Targets
    var targets = [];
    try { targets = JSON.parse(input.dataset.targets || "[]"); } catch (e) { }
    if (targets.length === 0 && input.dataset.target) targets = [input.dataset.target];

    targets.forEach(function (t) {
      var cleanT = t.trim();

      // DETECT SPECIAL TAGS
      // FINAL STRATEGY: Strict Separation based on PREFIX.
      // - $![...] -> FEES. STRICTLY Standard Text Replacement. NEVER triggers Modal.
      //              Must be SHORT (<50 chars). Long text is blocked (prevents hijacking).
      // - ![...]  -> TERMS. CANDIDATE for Modal.

      if (cleanT.indexOf("$![") === 0) {
        // FEE LOGIC
        // We allow "100" or "100 USD". We BLOCK "100 on signing..."
        if (val && val.length < 50) {
          standardReplacements.push({
            target: t,
            val: val,
            isCurrency: true
          });
          // CAPTURE THE FEE AMOUNT for the Modal!
          // If we find a valid fee, we save it.
          window.detectedFeeAmount = val;
        }
        // If > 50, we simply ignore it. It's likely a mis-mapped Term.
        // We do NOT add to paymentCandidates.

        // We add to kill list to ensure cleanup if it wasn't replaced (e.g. if we blocked it)
        // But if we blocked it, we want it empty? Yes.
        tagsToKill.push(t);

      } else if (cleanT.indexOf("![") === 0) {

        // TERM LOGIC -> Candidate for Modal
        // Even if EMPTY, we add it as a candidate. Why?
        // Because if we have a Fee ($![...]), we might want to trigger the modal 
        // to generate the schedule for the Terms field, even if the user didn't type anything in Terms.
        var candidateVal = (input.value && input.value.trim() !== "") ? input.value : (input.dataset.fullTerms || "");
        paymentCandidates.push({
          target: t,
          val: candidateVal,
          element: input
        });

      } else {
        // Standard Field
        standardReplacements.push({
          target: t,
          val: val,
          isCurrency: cleanT.indexOf("$[") === 0
        });
      }
    });
  });

  // --- NEW: Process Payment Candidates ---
  // We collected potential payment triggers. Now pick the BEST one, and treat others as standard fills.
  if (paymentCandidates.length > 0) {

    paymentCandidates.sort(function (a, b) {
      var aManual = (a.element.value && a.element.value.trim() !== "");
      var bManual = (b.element.value && b.element.value.trim() !== "");
      if (aManual && !bManual) return -1; // a wins
      if (!aManual && bManual) return 1;  // b wins
      return (b.val || "").length - (a.val || "").length; // Longer wins (usually Terms description)
    });


    if (!paymentCandidates[0].val && !paymentCandidates[0].element.value && !window.detectedFeeAmount) {
      feeTrigger = null;
    } else {
      feeTrigger = paymentCandidates[0];
    }

    if (feeTrigger) {
      // Losers -> Standard Replacements
      for (var i = 0; i < paymentCandidates.length; i++) {
        var cand = paymentCandidates[i];
        if (cand === feeTrigger) {
          tagsToKill.push(cand.target); // Winner handled by Modal
          continue;
        }

        standardReplacements.push({
          target: cand.target,
          val: cand.val, // Use its own value
          isCurrency: cand.target.indexOf("$![") === 0
        });
        // Also add to kill list
        tagsToKill.push(cand.target);
      }
    }
  }

  // --- B. EXECUTION HELPER (Standard) ---
  var runBatch = function (listToFill) {
    if (listToFill.length === 0) return Promise.resolve(0);
    listToFill.sort(function (a, b) { return b.target.length - a.target.length; });

    return Word.run(async function (context) {
      var sections = context.document.sections;
      sections.load("items");
      await context.sync();

      var count = 0;
      for (var i = 0; i < listToFill.length; i++) {
        var item = listToFill[i];
        if (!item.val) continue;

        var finalVal = String(item.val);
        if (item.isCurrency && !finalVal.trim().startsWith("$")) finalVal = "$" + finalVal;
        var cleanVal = finalVal.replace(/\n/g, "\u000b");

        // Literal Search (Safest)
        var searchOpts = { matchCase: false, ignoreSpace: true, ignorePunctuation: false };

        var bodyResults = context.document.body.search(item.target, searchOpts);
        bodyResults.load("items");
        await context.sync();
        for (var k = 0; k < bodyResults.items.length; k++) {
          var inserted = bodyResults.items[k].insertText(cleanVal, "Replace");
          if (userSettings.highlightFill) {
            inserted.font.highlightColor = "Yellow";
          }
          count++;
        }

        // Headers/Footers
        for (var s = 0; s < sections.items.length; s++) {
          var section = sections.items[s];
          var types = ["Primary", "FirstPage", "EvenPages"];
          for (var t = 0; t < types.length; t++) {
            try {
              var ranges = [section.getHeader(types[t]), section.getFooter(types[t])];
              for (var r = 0; r < ranges.length; r++) {
                var res = ranges[r].search(item.target, searchOpts);
                res.load("items");
                await context.sync();
                for (var h = 0; h < res.items.length; h++) {
                  var inserted = res.items[h].insertText(cleanVal, "Replace");
                  if (userSettings.highlightFill) {
                    inserted.font.highlightColor = "Yellow";
                  }
                  count++;
                }
              }
            } catch (e) { }
          }
        }
      }
      return count;
    });
  };

  // --- C. EXACT MATCH CLEANUP (New Strategy) ---
  var cleanupSpecificTags = function (tagsList) {
    if (!tagsList || tagsList.length === 0) return Promise.resolve(0);

    // Unique list to avoid duplicate searches
    var uniqueTags = tagsList.filter(function (v, i, a) { return a.indexOf(v) === i; });

    return Word.run(async function (context) {
      var deletedCount = 0;

      // Loop through every tag we know exists in the form (e.g. "$![Fee]")
      for (var i = 0; i < uniqueTags.length; i++) {
        var tag = uniqueTags[i];

        // search() with NO wildcards finds the exact string literal
        var results = context.document.body.search(tag, { matchCase: false, matchWildcards: false });
        results.load("items");
        await context.sync();

        // Delete found instances
        for (var k = results.items.length - 1; k >= 0; k--) {
          results.items[k].delete();
          deletedCount++;
        }
      }
      if (deletedCount > 0) console.log("Cleaned " + deletedCount + " tags.");
      return deletedCount;
    }).catch(function (e) {
      console.warn("Cleanup warning:", e);
    });
  };

  // --- D. EXECUTION FLOW ---
  var execute = function () {
    var btn = container.querySelector(".btn-action") || container.querySelector(".btn-platinum-config");
    if (btn) {
      btn.dataset.originalText = btn.innerText;
      btn.innerText = "Applying...";
      btn.disabled = true;
    }

    // 1. Fill Standard Fields
    runBatch(standardReplacements)
      .then(function (count) {

        if (feeTrigger && ((feeTrigger.val && feeTrigger.val.trim() !== "") || window.detectedFeeAmount)) {



          // Pass the FULL TERMS (from dataset) if available, otherwise the value.
          var scheduleContext = feeTrigger.element.dataset.fullTerms || feeTrigger.val;
          var modalVal = window.detectedFeeAmount ? window.detectedFeeAmount : feeTrigger.val;

          showPaymentScheduleModal(modalVal, scheduleContext, function (finalString) {
            // Replace the trigger with the generated schedule
            var paymentUpdate = [{ target: feeTrigger.target, val: finalString }];

            runBatch(paymentUpdate).then(function (c2) {
              // 3. Cleanup ALL special tags (removes any $! tags that weren't the trigger)
              cleanupSpecificTags(tagsToKill).then(function () {
                finishUp(count + c2, btn);
              });
            });
          });

        } else {
          // 3. Cleanup immediately if no modal needed
          cleanupSpecificTags(tagsToKill).then(function () {
            finishUp(count, btn);
          });
        }
      })
      .catch(function (e) {
        console.error(e);
        if (btn) { btn.innerText = "Error"; btn.disabled = false; }
        showOutput("Error: " + e.message, true);
      });
  };

  // --- E. MISSING CHECK ---
  if (missingFields.length > 0) {
    var overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

    var dialog = document.createElement("div");
    dialog.className = "tool-box slide-in";
    dialog.style.cssText = "background:white; width:300px; padding:20px; border-left:5px solid #f57f17; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

    dialog.innerHTML = `
            <div style="font-size:24px; margin-bottom:10px;">⚠️</div>
            <div style="font-weight:700; color:#e65100; margin-bottom:10px; font-size:16px;">Missing Fields</div>
            <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px; max-height:150px; overflow-y:auto;">
                The following fields are empty:<br><br>
                <strong>${missingFields.join("<br>")}</strong>
            </div>
            <div style="display:flex; gap:10px;">
                <button id="btnFixMissing" class="btn-action" style="flex:1; border-color:#999; color:#333;">I'll Fix Them</button>
                <button id="btnSkipMissing" class="btn-platinum-config" style="flex:1; color:#d32f2f;">Skip & Fill</button>
            </div>
        `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    document.getElementById("btnFixMissing").onclick = function () {
      overlay.remove();
      var firstRed = container.querySelector("[style*='border: 2px solid rgb(198, 40, 40)']");
      if (firstRed) firstRed.scrollIntoView({ behavior: "smooth", block: "center" });
    };

    document.getElementById("btnSkipMissing").onclick = function () {
      overlay.remove();
      execute();
    };
  } else {
    execute();
  }
  function finishUp(totalCount, btn) {
    if (userSettings.highlightFill) {
      showVerificationModal(totalCount, btn);
    } else {
      if (btn) {
        btn.innerText = "✅ Done (" + totalCount + ")";
        setTimeout(() => {
          btn.innerText = btn.dataset.originalText || "FILL ALL FIELDS";
          btn.disabled = false;
        }, 2000);
      }
      showToast("✅ Filling Complete. Scanning...");

      // --- CRITICAL CHANGE: HAND OFF TO SCANNER ---
      // The scanner will handle the "Next Step" logic (Wizard vs Hub)
      if (typeof runPreSaveScan === "function") {
        runPreSaveScan();
      } else {
        showAutoFillHub(); // Fallback
      }
      // ------------------------------------------
    }
  }

  function showVerificationModal(count, btn) {
    var overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

    var dialog = document.createElement("div");
    dialog.className = "tool-box slide-in";
    dialog.style.cssText = "background:white; width:300px; padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #ffa000; box-shadow:0 20px 50px rgba(0,0,0,0.3);";

    // --- CHECK FOR LOADED FILE ---
    var hasFile = (typeof AGREEMENT_DATA !== 'undefined' && AGREEMENT_DATA);
    var viewBtnHtml = "";

    if (hasFile) {
      viewBtnHtml = `
            <button id="btnVerifyViewSource" class="btn-action" style="
                width:100%; 
                background: #fff; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                font-weight: 700; 
                text-transform: uppercase; 
                border-radius: 0px; 
                padding: 10px; 
                margin-bottom: 10px;
                display: flex; align-items: center; justify-content: center; gap: 6px;
            ">
                <span>📄</span> View Source Data
            </button>
        `;
    }

    dialog.innerHTML = `
        <div style="background: linear-gradient(135deg, #ffa000 0%, #ff6f00 100%); padding: 15px; border-bottom: 3px solid #ffca28;">
            <div style="font-size:10px; font-weight:800; color:#fff8e1; text-transform:uppercase;">VERIFICATION</div>
            <div style="font-weight:600; color:white; font-size:16px;">Review Changes</div>
        </div>
        <div style="padding:20px;">
            <div style="font-size:13px; color:#546e7a; margin-bottom:15px; line-height:1.5;">
                <strong>${count} fields</strong> have been filled and highlighted in yellow. <br><br>
                Please review the document against your source data. When satisfied, click below to finalize.
            </div>
            
            ${viewBtnHtml}

            <button id="btnLooksGood" class="btn-action" style="
                width:100%; 
                background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
                color: #2e7d32; 
                border: 1px solid #b0bec5;
                font-weight: 800; 
                text-transform: uppercase;
                border-radius: 0px; 
                padding: 12px;
                box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.05);
                transition: all 0.2s;
            ">
                ✅ Looks Good (Remove Highlights)
            </button>
        </div>
    `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    // --- BIND VIEW BUTTON ---
    if (hasFile) {
      var btnView = document.getElementById("btnVerifyViewSource");
      btnView.onclick = function () {
        // Open the viewer but keep this modal open so user can return
        if (window.openDocumentViewer) window.openDocumentViewer();
      };
    }

    var looksGoodBtn = document.getElementById("btnLooksGood");

    // Platinum Hover Effect
    looksGoodBtn.onmouseover = function () {
      this.style.background = "#e8f5e9";
      this.style.borderColor = "#2e7d32";
      this.style.boxShadow = "0 2px 4px rgba(46, 125, 50, 0.15)";
      this.style.transform = "translateY(-1px)";
    };
    looksGoodBtn.onmouseout = function () {
      this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      this.style.borderColor = "#b0bec5";
      this.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.05)";
      this.style.transform = "translateY(0)";
    };

    looksGoodBtn.onclick = function () {
      overlay.remove();
      // CLEANUP HIGHLIGHTS
      Word.run(async (context) => {
        const body = context.document.body;
        body.font.highlightColor = null; // Clears highlights

        var sections = context.document.sections;
        sections.load("items");
        await context.sync();

        sections.items.forEach(sec => {
          ['Primary', 'FirstPage', 'EvenPages'].forEach(type => {
            try {
              sec.getHeader(type).font.highlightColor = null;
              sec.getFooter(type).font.highlightColor = null;
            } catch (e) { }
          });
        });

        await context.sync();
      })
        .then(() => {

          // --- NEW SNAPSHOT LOGIC HERE ---
          // We use a generic label because we don't know the final filename yet
          createSnapshot("Auto-Fill Verified [Original Version]", null, function () {

            showToast("✅ Verified & Snapshot Taken");

            if (btn) {
              btn.innerText = "✅ Verified";
              setTimeout(() => {
                btn.innerText = btn.dataset.originalText || "FILL ALL FIELDS";
                btn.disabled = false;
              }, 2000);
            }
            if (typeof runPreSaveScan === "function") {
              runPreSaveScan();
            } else {
              showSaveAsManager("Auto-Fill Verified [ORIGINAL VERSION]");
            }
          });
        })
        .catch(e => {
          console.error("Cleanup error: " + e);
          showToast("⚠️ Error removing highlights");
        });
    };
  }
}
var pendingJustifyText = "";
var pendingRewriteText = "";
// ==========================================
// START: showChatInterface
// ==========================================
function showChatInterface() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. Chat Container (Scrollable Area)
  var chatContainer = document.createElement("div");
  chatContainer.id = "chatContainer";
  chatContainer.className = "chat-container";
  // Add padding so text isn't hidden behind the footer
  chatContainer.style.paddingBottom = "130px";

  // Welcome Message
  var welcome = document.createElement("div");
  welcome.className = "chat-bubble chat-ai";
  var currentLang = localStorage.getItem("eli_language") || "en";
  var welcomeText = (typeof dictionary !== "undefined" && dictionary[currentLang])
    ? dictionary[currentLang]["msgChatWelcome"]
    : "<strong>👋 ELI</strong>: Ask a question or click a tool below to start.";
  welcome.innerHTML = welcomeText;

  chatContainer.appendChild(welcome);
  out.appendChild(chatContainer);

  // 2. Fixed Footer Container (Anchored to Bottom)
  var footerContainer = document.createElement("div");
  footerContainer.id = "chatFooter";
  footerContainer.style.cssText =
    "position:fixed; bottom:24px; left:0; width:100%; background:#f8f9fa; border-top:1px solid #d1d5db; z-index:1000; display:flex; flex-direction:column; box-shadow: 0 -4px 10px rgba(0,0,0,0.05);";

  // --- ROW A: SQUARED TOOL BAR ---
  var toolBar = document.createElement("div");
  toolBar.style.cssText =
    "display:flex; flex-direction:row; gap:1px; padding:8px 12px; background:#fff; border-bottom:1px solid #eee;";

  function createPremiumBtn(id, icon, label) {
    var btn = document.createElement("button");
    btn.id = id;
    btn.className = "tool-btn-premium"; // Uses new SQUARED CSS
    btn.innerHTML = `<span style="font-size:14px; margin-bottom:1px;">${icon}</span> ${label}`;
    return btn;
  }

  var btnDraft = createPremiumBtn("btnToolDraft", "⚖️", "Draft");
  var btnRewrite = createPremiumBtn("btnToolRewrite", "✏️", "Rewrite");
  var btnJustify = createPremiumBtn("btnToolJustify", "🛡️", "Defend");
  var btnTranslate = createPremiumBtn("btnToolTranslate", "🌐", "Translate");

  toolBar.append(btnDraft, btnRewrite, btnJustify, btnTranslate);
  footerContainer.appendChild(toolBar);

  // --- ROW B: INPUT + TOGGLE (Inline) ---
  var inputRow = document.createElement("div");
  inputRow.style.cssText =
    "display:flex; align-items:center; gap:8px; padding:8px 12px; background:#f1f3f4;";

  // 1. The Slider (Compact & Inline)
  var sliderContainer = document.createElement("div");
  sliderContainer.style.cssText =
    "display:flex; align-items:center; gap:6px; padding-right:8px; border-right:1px solid #d1d5db; margin-right:4px;";

  sliderContainer.innerHTML = `
        <label class="switch-toggle" style="transform:scale(0.8); margin:0;" title="Strict: Facts from document only | Broad: General knowledge from outside document">
            <input type="checkbox" id="scopeSlider">
            <span class="slider round"></span>
        </label>
        <div id="scopeLabel" style="font-size:9px; font-weight:800; color:#666; text-transform:uppercase; min-width:35px; letter-spacing:0.5px;">STRICT</div>
    `;
  inputRow.appendChild(sliderContainer);

  // 2. Text Input
  var input = document.createElement("textarea");
  input.id = "chatInput";
  input.className = "tool-input auto-expand";
  input.placeholder = "Ask a question...";
  input.rows = 1;
  input.style.cssText =
    "flex:1; height:34px; min-height:34px; max-height:100px; padding:8px 10px; font-size:13px; border-radius:0px; border:1px solid #ced4da; resize:none; line-height:1.4; margin:0; outline:none;";

  // Auto-expand logic
  input.addEventListener('input', function () {
    this.style.height = 'auto';
    this.style.height = (this.scrollHeight) + 'px';
  });

  inputRow.appendChild(input);

  // 3. Send Button (Squared)
  var sendBtn = document.createElement("button");
  sendBtn.className = "btn-action";
  sendBtn.innerText = "➤";
  sendBtn.style.cssText =
    "width:34px !important; height:34px !important; min-width:34px !important; background:#1565c0; color:white; border:none; border-radius:0px; cursor:pointer; flex:none !important; display:flex; align-items:center; justify-content:center; font-size:14px; margin:0; box-shadow:0 2px 5px rgba(0,0,0,0.2);";
  inputRow.appendChild(sendBtn);

  footerContainer.appendChild(inputRow);
  out.appendChild(footerContainer);

  // --- LOGIC: Slider Toggle ---
  var slider = document.getElementById("scopeSlider");
  var label = document.getElementById("scopeLabel");

  if (typeof currentChatScope !== "undefined" && currentChatScope === "broad") {
    slider.checked = true;
    label.innerText = "BROAD";
    label.style.color = "#1565c0";
  }

  slider.onchange = function () {
    if (this.checked) {
      label.innerText = "BROAD";
      label.style.color = "#1565c0";
      if (typeof window.setChatScope === "function") window.setChatScope("broad");
    } else {
      label.innerText = "STRICT";
      label.style.color = "#666";
      if (typeof window.setChatScope === "function") window.setChatScope("strict");
    }
  };

  // --- LOGIC: Toolbar Buttons ---
  function resetTools() {
    input.className = "tool-input auto-expand";
    input.classList.remove("input-mode-draft", "input-mode-rewrite", "input-mode-justify", "input-mode-translate");
    input.placeholder = "Ask a question...";
    [btnDraft, btnRewrite, btnJustify, btnTranslate].forEach(b => b.classList.remove("active-mode"));
  }

  function setMode(btn, modeClass, placeholder, chatMsg) {
    resetTools();
    input.classList.add(modeClass);
    input.placeholder = placeholder;
    input.focus();
    btn.classList.add("active-mode");
    if (chatMsg) addChatBubble(chatMsg, "ai");
  }

  btnDraft.onclick = () => setMode(btnDraft, "input-mode-draft", "Describe the clause...", "<strong>⚖️ Drafting Mode:</strong> Describe the clause you want me to create below.");

  btnRewrite.onclick = () => {
    validateApiAndGetText(true).then((t) => {
      if (!t || t.trim().length === 0) {
        addChatBubble("<strong>✏️ Selection Needed:</strong> Please highlight text first.", "ai");
        return;
      }
      if (typeof pendingRewriteText !== "undefined") pendingRewriteText = t;
      setMode(btnRewrite, "input-mode-rewrite", "How should I rewrite this?", '<strong>🎯 Rewrite Locked:</strong><br><em>"' + t.substring(0, 70) + '..."</em>');
    });
  };

  btnJustify.onclick = () => {
    validateApiAndGetText(true).then((t) => {
      if (!t || t.trim().length === 0) {
        addChatBubble("<strong>🛡️ Selection Needed:</strong> Highlight text to defend.", "ai");
        return;
      }
      if (typeof pendingJustifyText !== "undefined") pendingJustifyText = t;
      setMode(btnJustify, "input-mode-justify", "Explain your change...", '<strong>🛡️ Defense Locked:</strong><br>Ready to justify: <em>"' + t.substring(0, 70) + '..."</em>');
    });
  };

  btnTranslate.onclick = () => {
    validateApiAndGetText(true).then((t) => {
      if (!t || t.trim().length === 0) {
        addChatBubble("<strong>🌐 Selection Needed:</strong> Highlight text to translate.", "ai");
        return;
      }
      if (typeof pendingTranslateText !== "undefined") pendingTranslateText = t;
      setMode(btnTranslate, "input-mode-translate", "Target language?", "<strong>🌐 Translation Locked:</strong> Selection ready. Enter language below.");
    });
  };

  var submit = () => {
    if (!input.value) return;
    handleChatSubmit(input);
    resetTools();
    input.style.height = '34px';
  };

  sendBtn.onclick = submit;
  input.onkeydown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  };
}
// ==========================================
// START: resetModes
// ==========================================
function resetModes(input, btn1, btn2, btn3) {
  input.classList.remove("input-mode-draft");
  input.classList.remove("input-mode-rewrite");
  input.classList.remove("input-mode-justify");
  input.placeholder = "Ask a question or click a button above...";
  btn1.style.background = "";
  btn2.style.background = "";
  if (btn3) btn3.style.background = "";
}
// ==========================================
// END: resetModes
// ==========================================
// ==========================================
// START: handleChatSubmit
// ==========================================
function handleChatSubmit(inputElement) {
  var val = inputElement.value;
  if (!val) return;

  if (inputElement.classList.contains("input-mode-draft")) {
    runDraftInChat(val);
  } else if (inputElement.classList.contains("input-mode-rewrite")) {
    runRewriteInChat(val);
  } else if (inputElement.classList.contains("input-mode-justify")) {
    runJustifyInChat(val);
  } else if (inputElement.classList.contains("input-mode-translate")) {
    runTranslateInChat(val);
  } else {
    runChatQuery(val);
  }
  inputElement.value = "";
}
// ==========================================
// END: handleChatSubmit
// ==========================================

var currentChatScope = "strict";

// 2. Helper Function for the Toggle Click
window.setChatScope = function (mode) {
  currentChatScope = mode;

  // Update UI Visuals
  var btnStrict = document.getElementById("btnScopeStrict");
  var btnBroad = document.getElementById("btnScopeBroad");

  if (btnStrict && btnBroad) {
    if (mode === "strict") {
      btnStrict.className = "scope-btn active";
      btnBroad.className = "scope-btn";
      showToast("🛡️ Strict Mode: Facts Only");
    } else {
      btnStrict.className = "scope-btn";
      btnBroad.className = "scope-btn active";
      showToast("🌎 Broad Mode: Knowledge Unlocked");
    }
  }
};
// ==========================================
// START: runChatQuery (HYBRID CONTEXT)
// ==========================================
function runChatQuery(question) {
  addChatBubble(question, "user");
  var loaderId = addSkeletonBubble();

  // 1. Get Context (ALWAYS fetch Full Doc + Selection)
  Word.run(async function (context) {
    // A. Get Selection
    var selection = context.document.getSelection();
    selection.load("text");

    // B. Get Full Document (via HTML for cleaning)
    var body = context.document.body;
    var bodyHtml = body.getHtml();

    await context.sync();

    // Check if selection is "real" (more than just a cursor blink)
    var hasSelection = selection.text && selection.text.trim().length > 1;
    var cleanDoc = getCleanTextAndContext(bodyHtml.value);

    return {
      fullText: cleanDoc.text,
      selectionText: hasSelection ? selection.text : null
    };
  })
    .then(function (data) {
      // 2. Prepare Context String
      var limit = 500000; // Large token limit for full context
      var fullText = data.fullText.substring(0, limit);

      var finalText = "";

      if (data.selectionText) {
        // CASE A: HYBRID MODE (Selection + Full Doc)
        // We clearly label both parts so the AI knows what to focus on vs what is context
        finalText = `TARGET SELECTION (USER HIGHLIGHTED):\n"""\n${data.selectionText}\n"""\n\nFULL DOCUMENT CONTEXT:\n"""\n${fullText}\n"""`;
      } else {
        // CASE B: FULL DOC MODE (No selection)
        finalText = fullText;
      }

      var recentHistory = chatHistory.slice(-3);

      // 3. Call AI
      return callGemini(
        API_KEY,
        PROMPTS.QA(recentHistory, question, finalText, currentChatScope),
      ).then(function (result) {
        return {
          result: result,
          isHybrid: !!data.selectionText, // Flag for UI
        };
      });
    })
    // CHANGE 'package' TO 'pkg' (or any other name)
    .then(function (pkg) {
      removeBubble(loaderId);
      try {
        // Make sure to update the reference inside the function too:
        var data = tryParseGeminiJSON(pkg.result); // <--- Was package.result

        chatHistory.push({ role: "user", content: question });
        chatHistory.push({ role: "model", content: data.answer });

        var answerHtml = data.answer;

        // Update the reference here as well:
        if (pkg.isHybrid) { // <--- Was package.isHybrid
          answerHtml += `<div style="margin-top:8px; padding-top:6px; border-top:1px dashed #e0e0e0; font-size:9px; color:#1565c0; text-align:right;">🔍 Analyzed Selection + Document Context</div>`;
        }

        addChatBubble(answerHtml, "ai", data.citations);
      } catch (e) {
        addChatBubble("Error parsing answer.", "ai");
      }
    })
    .catch(function (e) {
      removeBubble(loaderId);
      console.error(e);
      if (e.message !== "Request cancelled.")
        addChatBubble("Error: " + e.message, "ai");
    });
}
// ==========================================
// END: runChatQuery
// ==========================================
// ==========================================
// START: runDraftInChat
// ==========================================
function runDraftInChat(topic) {
  addChatBubble("Drafting clause for: " + topic, "user");
  var loaderId = addSkeletonBubble();

  // CHANGED: Removed "chat"
  validateApiAndGetText(true)
    .then(function (text) {
      return callGemini(
        API_KEY,
        PROMPTS.DRAFT(topic, text.substring(0, 500000)),
      );
    })
    .then(function (jsonString) {
      removeBubble(loaderId);
      try {
        var data = tryParseGeminiJSON(jsonString);
        renderClauseBubble(data);
      } catch (e) {
        addChatBubble("Error generating draft.", "ai");
      }
    })
    .catch(function (e) {
      removeBubble(loaderId);
      addChatBubble("Error: " + e.message, "ai");
    });
}
// ==========================================
// START: addChatBubble (NO JUMP FIX)
// ==========================================
function addChatBubble(text, type, citations) {
  var container = document.getElementById("chatContainer");
  if (!container) return;

  var bubble = document.createElement("div");
  bubble.className = "chat-bubble chat-" + type;

  // Formatting Logic
  var formattedText = text;
  if (!/<[a-z][\s\S]*>/i.test(text)) {
    formattedText = formattedText.replace(/\*\*(.*?)\*\*/g, "<b>$1</b>");
    formattedText = formattedText.replace(/\*(.*?)\*/g, "<i>$1</i>");
    formattedText = formattedText.replace(/\n/g, "<br>");
  }
  bubble.innerHTML = formattedText;

  // Citations
  if (citations && citations.length > 0) {
    var citeContainer = document.createElement("div");
    citeContainer.style.marginTop = "10px";
    citeContainer.style.paddingTop = "8px";
    citeContainer.style.borderTop = "1px dashed #ccc";
    var label = document.createElement("div");
    label.innerText = "SOURCES:";
    label.style.fontSize = "9px";
    label.style.fontWeight = "bold";
    label.style.color = "#666";
    citeContainer.appendChild(label);

    citations.forEach(function (quote) {
      if (!quote) return;
      var chip = document.createElement("a");
      chip.className = "citation-chip";
      chip.innerHTML = "📍 " + quote.substring(0, 15) + "...";
      chip.onclick = function () { highlightCitation(quote, chip); };
      citeContainer.appendChild(chip);
    });
    bubble.appendChild(citeContainer);
  }

  container.appendChild(bubble);

  // --- THE FIX: Scroll container ONLY, not the window ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) {
    // Small timeout ensures the DOM has rendered the new height
    setTimeout(() => {
      scrollParent.scrollTop = scrollParent.scrollHeight;
    }, 10);
  }
}
// ==========================================
// START: addSkeletonBubble (NO JUMP FIX)
// ==========================================
function addSkeletonBubble() {
  var container = document.getElementById("chatContainer");
  var id = "loader-" + Date.now();

  var bubble = document.createElement("div");
  bubble.id = id;
  bubble.className = "chat-bubble chat-ai";

  bubble.innerHTML = `
    <div class="typing-indicator">
      <div class="typing-dot"></div>
      <div class="typing-dot"></div>
      <div class="typing-dot"></div>
    </div>
  `;

  container.appendChild(bubble);

  // --- THE FIX: Scroll container ONLY ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) {
    setTimeout(() => {
      scrollParent.scrollTop = scrollParent.scrollHeight;
    }, 10);
  }

  return id;
}

// ==========================================
// START: removeBubble
// ==========================================
function removeBubble(id) {
  var el = document.getElementById(id);
  if (el) el.remove();
}
// ==========================================
// END: removeBubble
// ==========================================

// ==========================================
// START: highlightCitation
// ==========================================
function highlightCitation(textToFind, chipElement) {
  chipElement.style.opacity = "0.5";
  Word.run(function (context) {
    return findSmartMatch(context, textToFind).then(function (range) {
      if (range) {
        range.select();
        chipElement.innerHTML = "✅ Found";
        chipElement.style.backgroundColor = "#dcedc8";
      } else {
        chipElement.innerHTML = "❌ Not found";
        chipElement.style.backgroundColor = "#ffcdd2";
      }
      chipElement.style.opacity = "1";
      return context.sync();
    });
  }).catch((e) => console.error(e));
}
// ==========================================
// END: highlightCitation
// ==========================================

// ==========================================
// START: processGeminiResponse
// ==========================================
function processGeminiResponse(result, type) {
  try {
    var data = tryParseGeminiJSON(result);

    // 1. LEGAL FILTERING LOGIC
    if (type === "legal" && Array.isArray(data)) {
      // A. Filter by Sensitivity Setting
      if (userSettings.riskSensitivity === "High") {
        // Show ONLY High
        data = data.filter((item) =>
          (item.severity || "").toLowerCase().includes("high"),
        );
      } else if (userSettings.riskSensitivity === "Medium") {
        // Show High AND Medium (Filter out Low)
        data = data.filter(
          (item) => !(item.severity || "").toLowerCase().includes("low"),
        );
      }

      // B. Filter by Blacklist (Training)
      if (legalBlacklist.length > 0) {
        data = data.filter(function (item) {
          // Check if this issue title is in the blacklist
          return !legalBlacklist.includes(item.issue);
        });
      }
    }

    if (type === "summary") {
      renderSummary(data);
    } else {
      if (!Array.isArray(data) || data.length === 0) {
        showOutput("No issues found! (Filters active)");
      } else {
        renderCards(data, type);

        if (type === "legal" && userSettings.showRiskScore) {
          renderRiskScoreCard(data);
        }
      }
    }
  } catch (e) {
    showOutput("Error parsing AI response: " + e.message, true);
  }
}
// ==========================================
// END: processGeminiResponse
// ==========================================
// ==========================================
// START: renderSummary
// ==========================================
function renderSummary(data) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  // Clear previous

  // Create the main card container
  var card = document.createElement("div");
  card.className = "error-card slide-in";
  // Reset default card padding/border for a cleaner look
  card.style.padding = "0";
  card.style.border = "none";
  card.style.boxShadow = "none";
  card.style.background = "transparent";

  // 1. HEADER: Premium Banner Style
  ;

  // Design: Deep Navy Gradient + Gold Accent Line + Subtle "Noise" Texture overlay
  var html = `
        <div style="background: #fff; border: 1px solid #cfd8dc; border-radius: 0px; overflow: hidden; margin-bottom: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.08);">
            
            <div style="
                padding: 12px 15px; 
                background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); 
                position: relative;
                overflow: hidden;
                border-bottom: 3px solid #c9a227; /* Colgate Gold Accent */
            ">
                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>

                <div style="position: relative; z-index: 2;">
                    <div style="
                        font-family: 'Inter', sans-serif;
                        font-size: 8px; 
                        font-weight: 700; 
                        color: #8fa6c7; /* Muted Blue-Grey */
                        text-transform: uppercase; 
                        letter-spacing: 2px; 
                        margin-bottom: 2px;
                    ">
                        LEGAL DOCUMENT ABSTRACT
                    </div>
                    
                    <div style="
                        font-family: 'Georgia', serif;
                        font-size: 16px; 
                        font-weight: 400; 
                        color: #ffffff; 
                        letter-spacing: 0.5px;
                        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
                    ">
                        ${data.doc_type || "Document Analysis"}
                    </div>
                </div>
            </div>
    `;
  // 2. KEY TERMS: Strict 2-Column Grid
  html += `<div style="padding: 15px; background: #fff;">`;
  // Grid: 1fr 1fr forces two equal columns
  html += `<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 15px;">`;
  for (var key in data.key_terms) {
    var item = data.key_terms[key];
    var displayValue =
      typeof item === "object" && item.value ? item.value : item;
    var sourceQuote = typeof item === "object" && item.quote ? item.quote : "";
    var cleanKey = key.replace(/_/g, " ");
    // --- CUSTOM FIX: "vs." -> "AND" ---
    if (cleanKey === "Parties" && displayValue) {
      displayValue = displayValue.replace(
        /\s+vs\.?\s+/gi,
        " <span style='color:#90a4ae; font-weight:700; font-size:0.8em;'> AND </span> ",
      );
    }

    // Interactive Attributes
    var divAttrs = "";
    var cursorStyle = "cursor: text;";

    if (sourceQuote) {
      var safeQuote = sourceQuote.replace(/"/g, "&quot;");
      divAttrs = `title="👆 Click to locate" data-quote="${safeQuote}" onclick="locateSummaryItem(this, this.getAttribute('data-quote'))"`;
      cursorStyle =
        "cursor: pointer; text-decoration-line: underline; text-decoration-style: dotted; text-decoration-color: #90caf9;";
    } else {
      divAttrs = `title="Copy" onclick="navigator.clipboard.writeText(this.innerText); showToast('Copied');"`;
      cursorStyle = "cursor: copy;";
    }

    // Render Card (Fixed 2-column layout + Light Blue Background)
    html += `
            <div style="background: #f0f7ff;
                border: 1px solid #dae0e5; border-radius: 0px; padding: 10px; min-width: 0; display: flex; flex-direction: column; justify-content: flex-start;
                transition: all 0.2s ease;" 
                 onmouseover="this.style.backgroundColor='#fff'; this.style.borderColor='#2196f3';" 
                 onmouseout="this.style.backgroundColor='#f0f7ff'; this.style.borderColor='#dae0e5';">
                
                <div style="font-size: 10px;
                    font-weight: 700; text-transform: uppercase; color: #78909c; margin-bottom: 4px; white-space: nowrap; overflow: hidden;
                    text-overflow: ellipsis;">
                    ${cleanKey}
                </div>
                
                <div ${divAttrs} style="font-size: 12px;
                    font-weight: 600; color: #263238; line-height: 1.35; word-wrap: break-word; ${cursorStyle}">
                     ${displayValue}
                </div>
            </div>`;
  }
  html += `</div>`; // End Grid

  // 3. EXECUTIVE SUMMARY (UPDATED: Enterprise Yellow Background)
  html += `
        <div style="margin-bottom: 10px;">
            <div style="font-size: 10px; font-weight: 700; color: #f57f17; text-transform: uppercase; margin-bottom: 6px;">Executive Summary</div>
            <div style="font-size: 12px; line-height: 1.5; color: #37474f; background: #fffde7; border-left: 3px solid #fbc02d; padding: 10px; border-radius: 0 4px 4px 0; border:1px solid #fff59d; border-left-width:3px;">
                ${data.executive_summary}
            </div>
        </div>
    `;

  // 4. RED FLAGS (If any)
  if (data.red_flags && data.red_flags.length > 0) {
    html += `<div style="border-top: 1px dashed #e0e0e0;
            padding-top: 12px; margin-top: 12px;">`;
    html += `<div style="font-size: 10px; font-weight: 700; color: #c62828;
            text-transform: uppercase; margin-bottom: 8px; display:flex; align-items:center;">
                    <span style="margin-right:4px;">⚠️</span> Key Risks
                 </div>`;
    html += `<ul style="margin: 0;
            padding-left: 0; list-style: none;">`;

    // [Search for "data.red_flags.forEach" inside renderSummary around line 643]
    // [Search for "data.red_flags.forEach" inside renderSummary]
    data.red_flags.forEach(function (item) {
      var title = typeof item === "object" && item.risk ? item.risk : item;
      var detail =
        typeof item === "object" && item.explanation ? item.explanation : "";
      var quote = typeof item === "object" && item.quote ? item.quote : "";

      var sourceBtn = "";
      if (quote) {
        // Safe Encoding for HTML Attributes
        var safeQuote = quote.replace(/"/g, "&quot;");
        var safeTitle = title.replace(/"/g, "&quot;").replace(/'/g, "\\'");
        var safeDetail = detail
          .replace(/"/g, "&quot;")
          .replace(/'/g, "\\'")
          .replace(/\n/g, " ");

        // --- FIX IS HERE: Added ', true' as the 5th argument 👇 ---
        // This forces the highlighter to grab the FULL paragraph, not just the search snippet.
        sourceBtn = `<span style="font-size:9px; color:#1976d2; cursor:pointer; margin-left:6px; font-weight:600; background:#e3f2fd; padding:1px 4px; border-radius:3px;" 
            onclick="locateSummaryItem(this, '${safeQuote}', '${safeTitle}', '${safeDetail}', true)">📍 Locate</span>`;
      }

      html += `
            <li style="margin-bottom: 8px; padding-left: 8px; border-left: 2px solid #ffcdd2;">
                <div style="font-size: 12px; font-weight: 700; color: #b71c1c; display:flex; align-items:center; flex-wrap:wrap;">
                        ${title} ${sourceBtn}
                </div>
                <div style="font-size: 11px; color: #546e7a; margin-top: 2px; line-height: 1.3;">
                    ${detail}
                </div>
            </li>`;
    });
    html += `</ul></div>`;
  }

  // 5. SEMANTIC NAVIGATOR (Quick Jumps)
  if (data.navigation_markers && data.navigation_markers.length > 0) {
    html += `<div style="margin-top: 12px;
            padding-top: 12px; border-top: 1px dashed #e0e0e0;">`;
    html += `<div style="font-size: 9px; font-weight: 700;
            color: #90a4ae; margin-bottom: 6px;">JUMP TO SECTION</div>`;
    html += `<div style="display: flex; flex-wrap: wrap;
            gap: 5px;">`;

    // [Search for "data.navigation_markers.forEach" inside renderSummary around line 1802]
    data.navigation_markers.forEach(function (nav) {
      if (
        !nav.quote ||
        nav.quote.length < 3 ||
        nav.quote.toLowerCase().includes("not found")
      )
        return;

      // FIX: Escape single quotes (') to prevent SyntaxError in onclick
      var safeQuote = nav.quote
        .replace(/\\/g, "\\\\")
        .replace(/'/g, "\\'")
        .replace(/"/g, "&quot;");
      var safeLabel = nav.label
        .replace(/\\/g, "\\\\")
        .replace(/'/g, "\\'")
        .replace(/"/g, "&quot;");

      html += `
                <span onclick="locateSummaryItem(this, '${safeQuote}', '${safeLabel}')" 
                      style="font-size: 10px; background: #f5f5f5; color: #607d8b; padding: 3px 8px; border-radius: 10px; cursor: pointer; border: 1px solid #e0e0e0; transition: all 0.1s;"
                      onmouseover="this.style.background='#e0e0e0'"
                      onmouseout="this.style.background='#f5f5f5'">
                    📍 ${nav.label}
                </span>`;
    });
    html += `</div></div>`;
  }

  html += `</div></div>`;
  // Close Inner Body and Card

  // 6. ACTION BUTTONS (Email + Export)
  var btnRow = document.createElement("div");
  btnRow.style.cssText = "display: flex; gap: 10px; margin-top: 20px;";

  var emailSectionId = "emailSection_" + Date.now();

  // Helper for Platinum Style
  function stylePlatinumBtn(btn, icon, label, baseColor) {
    var txtColor = baseColor || "#546e7a";
    btn.className = "btn-action";
    btn.innerHTML = `<span style="font-size:14px; opacity:0.8; filter:grayscale(100%); transition:all 0.2s;">${icon}</span> <span style="margin-left:6px;">${label}</span>`;
    btn.style.cssText = `
          flex: 1; 
          background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
          border: 1px solid #b0bec5; 
          border-top: 1px solid #cfd8dc;
          color: ${txtColor}; 
          font-size: 11px; 
          font-weight: 700; 
          text-transform: uppercase;
          border-radius: 0px;
          padding: 10px 5px;
          box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
          transition: all 0.2s ease;
          display: flex; align-items: center; justify-content: center;
      `;

    btn.onmouseover = function () {
      this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)";
      this.style.borderColor = baseColor || "#90caf9";
      this.style.color = baseColor || "#1565c0";
      this.style.transform = "translateY(-1px)";
      this.style.boxShadow = "0 3px 6px rgba(0,0,0,0.1), inset 0 1px 0 #ffffff";
      var ico = this.querySelector("span");
      if (ico) { ico.style.filter = "none"; ico.style.opacity = "1"; }
    };

    btn.onmouseout = function () {
      this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      this.style.borderColor = "#b0bec5";
      this.style.color = txtColor;
      this.style.transform = "translateY(0)";
      this.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08)";
      var ico = this.querySelector("span");
      if (ico) { ico.style.filter = "grayscale(100%)"; ico.style.opacity = "0.8"; }
    };
  }

  // Email Button
  var btnEmail = document.createElement("button");
  stylePlatinumBtn(btnEmail, "📧", "Draft Email", "#1565c0");
  btnEmail.onclick = function () {
    document.getElementById(emailSectionId).classList.toggle("hidden");
  };

  // Export Button
  var btnExport = document.createElement("button");
  stylePlatinumBtn(btnExport, "📊", "Export CSV", "#2e7d32");
  btnExport.onclick = function () {
    exportDealAbstract(data);
  };

  // Adobe Button
  var btnSign = document.createElement("button");
  stylePlatinumBtn(btnSign, "✍️", "Adobe Tags", "#d32f2f");
  btnSign.title = "Inject invisible signature tags";
  btnSign.onclick = function () {
    autoTagAdobe();
  };

  btnRow.appendChild(btnEmail);
  btnRow.appendChild(btnExport);
  btnRow.appendChild(btnSign);

  // Add buttons to card HTML manually since we closed the string above
  // Note: We need to append the row to the card element, not the string

  card.innerHTML = html;
  card.appendChild(btnRow); // <--- Button Row Added

  // 7. HIDDEN EMAIL DRAFT AREA
  var emailContainer = document.createElement("div");
  emailContainer.id = emailSectionId;
  emailContainer.className = "hidden slide-in";
  emailContainer.style.cssText =
    "margin-top:10px; border:1px solid #e0e0e0; border-radius:6px; background:#fff; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.05);";

  var emailContent =
    data.email_draft ||
    "⚠️ The AI could not generate the email draft this time.";
  emailContainer.innerHTML = `
        <div style="background:#f5f5f5; padding:8px 12px; font-size:11px; font-weight:700; color:#607d8b; border-bottom:1px solid #eee;">
             EMAIL PREVIEW
        </div>
        <textarea id="summary-body" style="width:100%; height:180px; padding:12px; border:none; font-family:inherit; font-size:12px; color:#37474f; resize:vertical; outline:none; background:#fff;">${emailContent}</textarea>
        <div style="padding:8px; border-top:1px solid #eee; background:#fafafa;">
            <button id="btnCopyEmail" class="btn-action" style="width:100%; background:#2e7d32; color:white; border:none;">Copy to Clipboard</button>
        </div>
    `;

  card.appendChild(emailContainer);
  out.appendChild(card);

  // Re-attach button logic for the email copy
  setTimeout(function () {
    var btn = emailContainer.querySelector("#btnCopyEmail"); // Scope to container
    if (btn) {
      btn.onclick = function () {
        var txt = emailContainer.querySelector("#summary-body");
        txt.select();
        document.execCommand("copy");
        btn.innerText = "Copied! ✓";
        showToast("Email copied");
        setTimeout(() => {
          btn.innerText = "Copy to Clipboard";
        }, 2000);
      };
    }
  }, 100);
  out.classList.remove("hidden");
}
// ==========================================
// END: renderSummary
// ==========================================
// ==========================================
// START: renderCards (Updated with AI Drafter Button)
// ==========================================
function renderCards(items, type, customBackAction) {
  var out = document.getElementById("output");
  var title = type === "error" ? "Suggested Fixes" : "Legal Risks";
  var subLabel = type === "error" ? "EDITS FOUND" : "RISKS DETECTED";

  // 1. PREMIUM HEADER
  // --- 1. PREMIUM HEADER (MATCHES SUMMARY) ---
  out.innerHTML = `
        <div style="background: #fff; border: 1px solid #cfd8dc; border-radius: 0px; overflow: hidden; margin-bottom: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
            
            <div style="
                padding: 12px 15px; 
                background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); 
                position: relative;
                overflow: hidden;
                border-bottom: 3px solid #c9a227; /* Colgate Gold */
                display: flex; justify-content: space-between; align-items: center;
            ">
                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>

                <div style="position: relative; z-index: 2;">
                    <div style="
                        font-family: 'Inter', sans-serif;
                        font-size: 8px; 
                        font-weight: 700; 
                        color: #8fa6c7; 
                        text-transform: uppercase; 
                        letter-spacing: 2px; 
                        margin-bottom: 2px;
                    ">
                        ${subLabel}
                    </div>
                    
                    <div style="
                        font-family: 'Georgia', serif;
                        font-size: 16px; 
                        font-weight: 400; 
                        color: #ffffff; 
                        letter-spacing: 0.5px;
                        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
                    ">
                        ${title}
                    </div>
                </div>

                <div style="position: relative; z-index: 2; background: rgba(255,255,255,0.15); padding: 4px 10px; border-radius: 0px; border: 1px solid rgba(255,255,255,0.3); color: #fff; font-size: 12px; font-weight: 700;">
                    ${items.length}
                </div>
            </div>
        </div>
    `;
  if (typeof customBackAction === "function") {
    var btnBack = document.createElement("button");
    btnBack.className = "btn-action";
    // Smart Label: Check if we are in Wizard mode
    btnBack.innerText = (window.eliWizardMode) ? "💾 Next: Save & Send ➔" : "⬅ Done / Back to Hub";
    btnBack.style.cssText = "width:100%; margin-bottom:15px; background:#fff3e0; color:#e65100; border:1px solid #ffe0b2; font-weight:bold;";
    btnBack.onclick = customBackAction;
    out.appendChild(btnBack);
  }

  items.forEach(function (item, index) {
    var card = document.createElement("div");
    card.className = "error-card slide-in";
    card.style.animationDelay = index * 0.05 + "s";

    // --- COLORS & ICONS ---
    var tabColor = "#1565c0"; // Default Blue
    var tabBg = "#e3f2fd";
    var icon = "📝";

    if (type === "error" && (item.type === "ghost" || (item.reason && item.reason.includes("Ghost")))) {
      icon = "👻";
    } else if (type === "legal") {
      icon = "⚖️";
    }

    if (type === "legal" && item.severity) {
      var sev = item.severity.toLowerCase();
      if (sev === "high") { tabColor = "#c62828"; tabBg = "#ffebee"; }
      else if (sev === "medium") { tabColor = "#e65100"; tabBg = "#fff3e0"; }
      else { tabColor = "#2e7d32"; tabBg = "#e8f5e9"; }
    }

    // --- CARD CONTAINER ---
    card.style.cssText = `
            position: relative;
            background: #fff;
            border: 1px solid ${tabColor}; 
            border-top: 3px solid ${tabColor}; 
            border-radius: 0px; 
            margin-top: 30px;
            margin-bottom: 35px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            padding: 0;
            overflow: visible;
            box-sizing: border-box;
            width: 100%;
        `;

    var originalText = type === "error" ? item.original : item.quote;
    var displayText = type === "error" ? item.correction : item.issue;
    var reasonText = type === "error" ? item.reason : item.explanation;
    var headerLabel = type === "error" ? "Suggestion" : "Risk";
    var labelText = headerLabel + " " + (index + 1);

    // --- 1. INDEX TAB ---
    var html = `
            <div style="
                position: absolute;
                top: -25px;
                left: -1px; 
                
                /* PREMIUM UPGRADE: Gradient & Depth */
                background: linear-gradient(to bottom, #ffffff 0%, ${tabBg} 100%);
                color: ${tabColor};
                
                border: 1px solid ${tabColor};
                border-bottom: 1px solid ${tabBg}; /* Blends into card body */
                
                padding: 6px 14px;
                border-radius: 0px;
                
                font-size: 10px;
                font-weight: 800;
                text-transform: uppercase;
                letter-spacing: 1px;
                
                /* Highlights */
                box-shadow: inset 0 1px 0 #ffffff, 0 -2px 5px rgba(0,0,0,0.05);
                text-shadow: 0 1px 0 rgba(255,255,255,1);
                
                z-index: 2;
                line-height: 1;
                white-space: nowrap;
            ">
                ${icon} ${labelText}
            </div>
        `;

    // --- 2. CONTENT ---
    html += `<div style="padding: 16px; box-sizing: border-box;">`;

    if (type === "error") {
      html += `<div class='text-label' style="color:#90a4ae;">ORIGINAL</div>`;
      html += `<div class='text-block' style='font-family: "Georgia", serif; color: #546e7a; background: #f8f9fa; border-left: 3px solid #cfd8dc; padding: 8px 12px; border-radius: 0; margin-bottom: 12px; font-size: 13px; word-wrap: break-word;'>${originalText}</div>`;
      html += `<div class='text-label' style="color:#2e7d32;">FIX</div>`;
      var diffHtml = getRedlineHtml(originalText, displayText);
      html += `<div class='fix-block' style='font-family: "Georgia", serif; background: #f1f8e9; border-left: 3px solid #66bb6a; color: #1b5e20; padding: 8px 12px; border-radius: 0; font-size: 13px; line-height: 1.5; margin-bottom: 0; word-wrap: break-word;'>${diffHtml}</div>`;
    } else {
      // --- LEGAL ISSUE CONTENT ---
      html += `<div class='text-label' style="color:#d32f2f;">RISK DETECTED</div>`;
      html += `<div class='text-block' style='color:#b71c1c; background:#fef2f2; border-left: 3px solid #ef9a9a; word-wrap: break-word;'>${displayText}</div>`;

      // --- NEW: SUGGESTION BOX WITH AI DRAFTER BUTTON ---
      html += `<div id='suggest-${index}' class='suggestion-box hidden' style="margin-top:12px; background:#e3f2fd; border-left:4px solid #1565c0; padding:12px;">
                        <div style="font-size:11px; font-weight:700; color:#1565c0; margin-bottom:6px; text-transform:uppercase;">Suggestion</div>
                        <div style="font-size:13px; color:#333; margin-bottom:10px;">${item.fix_suggestion}</div>
                        
                        <button id="btnAiDraft_${index}" class="btn-action btn-ai-trigger" style="width:100%; font-weight:700; display:flex; align-items:center; justify-content:center; gap:6px; padding:10px;">
    <span style="font-size:14px;">⚡</span> AI Drafter: Write this Clause
</button>
                     </div>`;
    }

    html += `</div>`;

    // --- 3. EXPLANATION FOOTER ---
    html += `
            <div style="
                background: #fcfcfc; 
                border-top: 1px solid #f0f0f0; 
                padding: 10px 16px;
                display: flex; gap: 10px; align-items: flex-start;
                box-sizing: border-box;
            ">
                <div style="font-size: 16px; opacity: 0.8; margin-top: 0px;">💡</div>
                <div style="font-size: 13px; color: #263238; line-height: 1.5; font-weight: 500;">
                    ${reasonText}
                </div>
            </div>
        `;

    card.innerHTML = html;

    // --- BIND THE NEW AI DRAFTER BUTTON (LEGAL ONLY) ---
    if (type === 'legal') {
      setTimeout(function () {
        var draftBtn = card.querySelector(`#btnAiDraft_${index}`);
        if (draftBtn) {
          draftBtn.onclick = function () {
            // Pass the suggestion text to the drafter
            runDraftFromFix(item.fix_suggestion || item.issue);
          };
        }
      }, 0);
    }

    // --- ACTIONS AREA (Rest of your existing logic) ---
    var actionsDiv = document.createElement("div");
    actionsDiv.className = "card-actions";
    actionsDiv.style.padding = "0 16px 16px 16px";
    actionsDiv.style.boxSizing = "border-box";

    var buttonRow = document.createElement("div");
    buttonRow.style.cssText = "display: flex; gap: 8px; width: 100%; box-sizing: border-box; margin-bottom: 12px;";

    // PLATINUM BUTTON HELPER
    function createPlatinumBtn(label, iconStr, textColor) {
      var b = document.createElement("button");
      b.className = "btn-action";
      b.innerHTML = `<span style="font-size:12px; color:${textColor}; text-shadow:0 1px 0 rgba(255,255,255,0.8);">${iconStr}</span> <span style="color:${textColor}; text-shadow:0 1px 0 rgba(255,255,255,0.8);">${label}</span>`;
      b.style.cssText = `
                flex: 1; min-width: 0; height: 30px; font-size: 10px; font-weight: 700; text-transform: uppercase; border-radius: 0px; cursor: pointer; white-space: nowrap;
                background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
                border: 1px solid #b0bec5; border-top: 1px solid #cfd8dc;
                box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
                display: flex; align-items: center; justify-content: center; gap: 6px; padding: 0 4px; transition: all 0.2s ease; z-index: 1;
            `;
      b.onmouseover = function () {
        this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)";
        this.style.borderColor = "#90caf9";
        this.style.transform = "translateY(-1px)";
        this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.1), inset 0 1px 0 #ffffff";
        this.style.zIndex = "10";
      };
      b.onmouseout = function () {
        this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
        this.style.borderColor = "#b0bec5";
        this.style.transform = "translateY(0)";
        this.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08)";
        this.style.zIndex = "1";
      };
      return b;
    }

    // 1. LOCATE BUTTON
    var btnLocate = createPlatinumBtn("Locate", "👁", "#455a64");
    btnLocate.onclick = function () {
      if (type === "legal") {
        locateSummaryItem(btnLocate, originalText, displayText, reasonText, true);
      } else {
        locateText(originalText, btnLocate, false);
      }
    };
    buttonRow.appendChild(btnLocate);

    if (type === "error") {
      // 2. APPLY BUTTON (Green Text)
      var btnApply = createPlatinumBtn("Apply", "✅", "#2e7d32");

      // --- SCRIPT 40/31 LIST LOGIC ---
      btnApply.onclick = function () {
        var issueText = (item.reason || "") + " " + (item.correction || "");
        var isNumbering = /numbering|sequence|duplicate section|list item/i.test(issueText) ||
          /^\d+(\.|\))$/.test(item.correction.trim());

        if (isNumbering) {
          var overlay = document.createElement("div");
          overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

          var dialog = document.createElement("div");
          dialog.className = "tool-box slide-in";
          dialog.style.cssText = "background:white; width:280px; border-left:5px solid #f57f17; padding:20px; box-shadow:0 10px 25px rgba(0,0,0,0.2); border-radius:4px;";
          dialog.innerHTML = `
                    <div style="font-size:24px; margin-bottom:10px;">🔢</div>
                    <div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Manual Fix Required</div>
                    <div style="font-size:12px; color:#555; line-height:1.5; margin-bottom:15px;">
                        This appears to be a <strong>list numbering error</strong>.<br><br>
                        Microsoft Word's auto-numbering is complex.
                        Applying this fix via AI may break your list formatting.<br><br>
                        Please make this specific edit manually.
                    </div>
                    <button id="btnOkManual" class="btn-action" style="width:100%; font-weight:bold;">OK, I'll fix it</button>
                `;
          overlay.appendChild(dialog);
          document.body.appendChild(overlay);

          document.getElementById("btnOkManual").onclick = function () {
            overlay.remove();
          };
          return;
        }

        applyFix(originalText, displayText, btnApply).then(function (success) {
          if (success) {
            btnLocate.disabled = true;
            btnLocate.innerHTML = "Resolved";
            btnLocate.style.opacity = "0.5";
            btnLocate.style.cursor = "default";
            btnLocate.onclick = null;
            updateProgressBar();
          }
        });
      };
      buttonRow.appendChild(btnApply);

      // 3. IGNORE BUTTON
      var btnIgnore = createPlatinumBtn("Ignore", "🚫", "#c62828");
      btnIgnore.onclick = function () {
        card.style.transition = "all 0.4s ease";
        card.style.transform = "translateX(100%)";
        card.style.opacity = "0";
        setTimeout(function () { card.remove(); updateProgressBar(); }, 400);
        showToast("Suggestion Ignored");
      };
      buttonRow.appendChild(btnIgnore);
    } else {
      // Legal Buttons
      var btnSuggest = createPlatinumBtn("Suggest Fix", "💡", "#1565c0");
      btnSuggest.onclick = function () {
        var box = document.getElementById("suggest-" + index);
        box.classList.toggle("hidden");
        // Update text based on visibility
        this.innerHTML = box.classList.contains("hidden") ?
          `<span style="font-size:12px; color:#1565c0;">💡</span> <span style="color:#1565c0;">Suggest Fix</span>` :
          `<span style="font-size:12px; color:#1565c0;">▲</span> <span style="color:#1565c0;">Hide Fix</span>`;
      };
      buttonRow.appendChild(btnSuggest);

      // ... inside renderCards function ...

      // --- UPDATED IGNORE LOGIC ---
      var btnIgnoreLegal = createPlatinumBtn("Ignore", "🚫", "#c62828");
      btnIgnoreLegal.onclick = function () {
        // 1. Train AI (Add to blacklist so it doesn't appear next time)
        if (typeof legalBlacklist !== "undefined") legalBlacklist.push(displayText);

        // 2. Animate Out (Slide Right & Fade) <--- NEW ANIMATION
        card.style.transition = "all 0.4s ease";
        card.style.transform = "translateX(100%)";
        card.style.opacity = "0";

        // 3. Remove from DOM <--- NEW REMOVAL
        setTimeout(function () {
          card.remove();
          updateProgressBar();
        }, 400);

        showToast("Risk Ignored");
      };
      buttonRow.appendChild(btnIgnoreLegal);
    }

    actionsDiv.appendChild(buttonRow);

    // --- FOOTER ROW (Error Only) ---
    if (type === "error") {
      var footerRow = document.createElement("div");
      footerRow.style.cssText = "border-top: 1px dashed #eee; padding-top: 10px; display: flex; justify-content: space-between; align-items: center; width: 100%; box-sizing: border-box; overflow: hidden;";

      var checkLabel = document.createElement("label");
      checkLabel.style.fontSize = "11px";
      checkLabel.style.color = "#78909c";
      checkLabel.style.display = "flex";
      checkLabel.style.alignItems = "center";
      checkLabel.style.cursor = "pointer";
      checkLabel.style.flex = "1";
      checkLabel.style.whiteSpace = "nowrap";
      checkLabel.innerHTML = `<input type="checkbox" class="email-check" data-original="${originalText}" data-fix="${displayText}" style="margin-right:6px;"> Add to Email`;

      var bugBtn = document.createElement("div");
      bugBtn.innerHTML = "Flag Incorrect";
      bugBtn.style.cssText = "font-size: 10px; color: #b0bec5; cursor: pointer; text-decoration: underline; margin-left: 10px; flex-shrink: 0;";
      bugBtn.onclick = function () {
        var parent = footerRow.parentElement;
        var existingFb = parent.querySelector('.feedback-container');
        if (existingFb) { existingFb.remove(); } else {
          var fbContainer = document.createElement("div");
          fbContainer.className = "feedback-container slide-in";
          fbContainer.style.cssText = "display: flex; gap: 5px; width: 100%; margin-top: 8px; box-sizing: border-box;";
          fbContainer.innerHTML = `<input type="text" placeholder="Why is this wrong?" style="flex:1; border:1px solid #ccc; font-size:11px; padding:4px; border-radius:0px;"><button style="background:#78909c; color:white; border:none; font-size:10px; padding:4px 8px; cursor:pointer; border-radius:0px;">Send</button>`;
          var send = fbContainer.querySelector("button");
          var inp = fbContainer.querySelector("input");
          send.onclick = function () {
            if (inp.value.trim()) {
              userFeedback.push({ original: originalText, reason: inp.value });
              fbContainer.innerHTML = "<span style='color:#2e7d32; font-size:10px; font-weight:bold;'>Thanks!</span>";
              setTimeout(() => fbContainer.remove(), 1500);
            }
          };
          parent.appendChild(fbContainer);
          inp.focus();
        }
      };
      footerRow.appendChild(checkLabel);
      footerRow.appendChild(bugBtn);
      actionsDiv.appendChild(footerRow);
    }

    // --- NEGOTIATION ROW (Legal Only) ---
    if (type === "legal") {
      var negoRow = document.createElement("div");
      negoRow.style.marginTop = "8px";
      negoRow.style.paddingTop = "8px";
      negoRow.style.borderTop = "1px dashed #eee";
      negoRow.innerHTML = `
                <label style="display:flex; align-items:center; cursor:pointer; font-size:11px; font-weight:600; color:#1565c0;">
                    <input type="checkbox" class="negotiation-check" data-clause="${item.quote.replace(/"/g, '&quot;')}" data-issue="${item.issue.replace(/"/g, '&quot;')}" style="margin-right:6px;"> Include in Negotiation Email
                </label>
            `;
      actionsDiv.appendChild(negoRow);
    }

    card.appendChild(actionsDiv);
    out.appendChild(card);
  });
  // BOTTOM BUTTONS
  if (type === "error") {
    var emailSection = document.createElement("div");
    emailSection.id = "emailSection";
    emailSection.style.marginTop = "30px";
    emailSection.style.borderTop = "1px dashed #cfd8dc";
    emailSection.style.paddingTop = "20px";
    var btnEmail = document.createElement("button");
    btnEmail.className = "btn-action";
    btnEmail.style.width = "100%";
    btnEmail.innerHTML = "📧 Draft Email Summary";
    btnEmail.style.borderRadius = "0px";
    btnEmail.onclick = function (e) { generateEmailDraft(e); };
    emailSection.appendChild(btnEmail);
    out.appendChild(emailSection);
    // ... inside renderCards, at the bottom ...

    // ... inside renderCards, near the bottom ...

  } else if (type === "legal") {
    var negoSection = document.createElement("div");
    negoSection.style.marginTop = "30px";
    negoSection.style.borderTop = "1px dashed #cfd8dc";
    negoSection.style.paddingTop = "20px";

    // --- UPDATED: TRUE PLATINUM (SILVER) BUTTON ---
    var btnNego = document.createElement("button");
    btnNego.className = "btn-platinum-config";
    btnNego.innerHTML = "📧 DRAFT NEGOTIATION EMAIL";

    // Apply Silver/Metallic CSS
    btnNego.style.cssText = `
        width: 100%; 
        background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%); /* Silver Gradient */
        color: #455a64; /* Dark Grey Text */
        border: 1px solid #b0bec5; 
        border-bottom: 3px solid #90a4ae; /* Thick Metallic Bottom */
        font-weight: 800; 
        border-radius: 0px; 
        text-transform: uppercase; 
        padding: 12px; 
        letter-spacing: 0.5px;
        cursor: pointer;
        transition: all 0.2s;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    `;

    // Hover: Light Blue Tint
    btnNego.onmouseover = function () {
      this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)";
      this.style.borderColor = "#90caf9";
    };
    btnNego.onmouseout = function () {
      this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      this.style.borderColor = "#b0bec5";
    };

    btnNego.onclick = function () { generateNegotiationEmail(); };

    negoSection.appendChild(btnNego);
    out.appendChild(negoSection);
  }

  out.classList.remove("hidden");
  updateProgressBar();
}
// ==========================================
// END: renderCards
// ==========================================
// / START: getFeedbackInstructions
// ==========================================
function getFeedbackInstructions() {
  var feedbackInstructions = "";
  if (userFeedback.length > 0) {
    feedbackInstructions =
      "\nIMPORTANT - PREVIOUS USER FEEDBACK (DO NOT REPORT THESE):\n";
    userFeedback.forEach(function (item) {
      feedbackInstructions +=
        '- Ignore text similar to: "' +
        item.original +
        "\". User's Reason: " +
        item.reason +
        "\n";
    });
    feedbackInstructions +=
      "INSTRUCTION: Strictly follow the user's feedback above.\n";
  }
  return feedbackInstructions;
}
// ==========================================
// END: getFeedbackInstructions
// ==========================================
// ==========================================
// START: generateEmailDraft (Fix: No Scroll Jump)
// ==========================================
function generateEmailDraft(event) {
  var checkedBoxes = document.querySelectorAll(".email-check:checked");
  if (checkedBoxes.length === 0) {
    var btn = event ? event.target || event.srcElement : null;
    if (btn) {
      var oldText = btn.innerHTML;
      btn.innerHTML = "⚠️ Select items first!";
      setTimeout(function () {
        btn.innerHTML = oldText;
      }, 2000);
    }
    return;
  }

  var emailBody = "Dear [Third Party],\n\nAttached please find [Name of Document] between the parties. I have made the following changes:\n\n";
  checkedBoxes.forEach(function (box) {
    emailBody += "• " + box.dataset.fix + '\n  (Replaced: "' + box.dataset.original + '")\n\n';
  });
  emailBody += "Thanks,\nAndrew";

  var emailSection = document.getElementById("emailSection");
  var existingDraft = document.getElementById("emailDraftContainer");
  if (existingDraft) existingDraft.remove();

  var draftContainer = document.createElement("div");
  draftContainer.id = "emailDraftContainer";
  draftContainer.className = "email-inline-container slide-in";

  var header = document.createElement("div");
  header.innerHTML = "<strong>Email Draft:</strong>";
  header.style.marginBottom = "8px";

  var textarea = document.createElement("textarea");
  textarea.className = "email-textarea";
  textarea.value = emailBody;

  var copyBtn = document.createElement("button");
  copyBtn.className = "btn btn-red"; // Or use btn-action if you prefer consistency
  copyBtn.innerText = "Copy to Clipboard";
  copyBtn.style.marginTop = "10px";

  copyBtn.onclick = function () {
    textarea.select();
    document.execCommand("copy");
    copyBtn.innerText = "Copied! ✓";
    setTimeout(function () {
      copyBtn.innerText = "Copy to Clipboard";
    }, 2000);
  };

  var cancelBtn = document.createElement("button");
  cancelBtn.className = "btn-action";
  cancelBtn.innerText = "Close Draft";
  cancelBtn.style.marginTop = "10px";
  cancelBtn.onclick = function () {
    draftContainer.remove();
  };

  draftContainer.appendChild(header);
  draftContainer.appendChild(textarea);
  draftContainer.appendChild(copyBtn);
  draftContainer.appendChild(cancelBtn);
  emailSection.appendChild(draftContainer);

  // --- FIX: GENTLE SCROLL INSTEAD OF JUMP ---
  // REPLACED: draftContainer.scrollIntoView({ behavior: "smooth" });
  var out = document.getElementById("output");
  setTimeout(() => {
    out.scrollTop = out.scrollHeight; // Scroll inner container only
  }, 10);
}
// ==========================================
// END: generateEmailDraft
// ==========================================
// ==========================================
// START: locateText (Auto-Toggles Track Changes)
// ==========================================
function locateText(textToFind, btnElement, grabParagraph) {
  var originalLabel = btnElement.innerText;
  btnElement.innerText = "Scanning...";
  btnElement.style.opacity = "0.7";

  // Regex to identify lists/headers
  var listRegex = /^(\([a-zA-Z0-9]+\)|[a-zA-Z0-9]+\.|[ivxIVX]+\.)\s+(.+)/;
  var cleanSearch = textToFind;

  // SMART CLEANING
  if (listRegex.test(cleanSearch)) {
    var potentialMatch = cleanSearch.replace(listRegex, "$2").trim();
    if (potentialMatch.length > 3) cleanSearch = potentialMatch;
  }
  if (cleanSearch.length > 250) cleanSearch = cleanSearch.substring(0, 250);

  Word.run(async function (context) {
    // 1. CAPTURE & TOGGLE TRACKING OFF
    var doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    var wasTracking = doc.changeTrackingMode !== "Off";
    if (wasTracking) {
      doc.changeTrackingMode = "Off"; // Temporarily disable
      await context.sync();
    }

    // 2. PERFORM SEARCH
    var foundRange = null;
    try {
      foundRange = await findSmartMatch(context, cleanSearch);
    } catch (e) { console.log("Smart match failed"); }

    if (!foundRange) {
      // Footer Fallback
      var sections = context.document.sections;
      sections.load("items");
      await context.sync();

      // (Simplified footer loop for brevity/speed)
      for (var i = 0; i < sections.items.length; i++) {
        if (foundRange) break;
        var section = sections.items[i];
        // Check Primary Footer only for speed in fallback
        try {
          var footer = section.getFooter("Primary");
          var footerResults = footer.search(cleanSearch, { matchCase: false, ignoreSpace: true });
          footerResults.load("items");
          await context.sync();
          if (footerResults.items.length > 0) foundRange = footerResults.items[0];
        } catch (e) { }
      }
    }

    // 3. RESTORE TRACKING
    if (wasTracking) {
      doc.changeTrackingMode = "TrackAll"; // Restore
      await context.sync();
    }

    // 4. HANDLE UI
    if (foundRange) {
      if (grabParagraph) {
        var para = foundRange.paragraphs.getFirst();
        para.getRange().select();
      } else {
        foundRange.select();
      }
      btnElement.innerHTML = "<span>✅</span> Found";
      btnElement.style.backgroundColor = "#e8f5e9";
      btnElement.style.color = "#2e7d32";
      btnElement.style.opacity = "1";
    } else {
      btnElement.innerHTML = "<span>⚠️</span> Not Found";
      btnElement.style.backgroundColor = "#fff3e0";
      btnElement.style.color = "#e65100";
      btnElement.style.opacity = "1";
      setTimeout(function () {
        btnElement.innerText = originalLabel;
        btnElement.style.backgroundColor = "";
        btnElement.style.color = "";
      }, 3000);
    }
    return context.sync();
  }).catch(function (e) {
    console.error(e);
    btnElement.innerText = "Error";
  });
}
// ==========================================
// END: locateText
// ==========================================
// ==========================================
// REPLACE: findSmartMatch (Deep Scan: Body -> Headers -> Footers)
// ==========================================
async function findSmartMatch(context, textToFind) {
  var cleanText = textToFind.replace(/[\r\n\t]+/g, " ").trim();

  // Safety check
  if (cleanText.length < 1) return null;

  var searchOptions = {
    matchCase: false,
    ignoreSpace: true,
    ignorePunctuation: true,
  };

  // Truncate search string to avoid Word API limits (255 chars)
  var exactSearchString = cleanText.substring(0, 250);

  // --- 1. SEARCH BODY (Fastest) ---
  var body = context.document.body;
  var bodyResults = body.search(exactSearchString, searchOptions);
  bodyResults.load("items");
  await context.sync();

  if (bodyResults.items.length > 0) return bodyResults.items[0];

  // --- 2. SEARCH HEADERS & FOOTERS (Deep Scan) ---
  // If not found in body, we must iterate through sections
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  for (var i = 0; i < sections.items.length; i++) {
    var section = sections.items[i];
    var types = ["Primary", "FirstPage", "EvenPages"];

    for (var t = 0; t < types.length; t++) {
      // A. Check Footer (Most likely culprit for your issue)
      try {
        var footer = section.getFooter(types[t]);
        var fRes = footer.search(exactSearchString, searchOptions);
        fRes.load("items");
        await context.sync();
        if (fRes.items.length > 0) return fRes.items[0];
      } catch (e) {
        // Ignore (e.g. if FirstPage footer doesn't exist)
      }

      // B. Check Header
      try {
        var header = section.getHeader(types[t]);
        var hRes = header.search(exactSearchString, searchOptions);
        hRes.load("items");
        await context.sync();
        if (hRes.items.length > 0) return hRes.items[0];
      } catch (e) { }
    }
  }

  // --- 3. OPTIONAL: ANCHOR SEARCH (Fallback for long text) ---
  // If exact search failed, try searching by start/end anchors (only in Body for speed)
  if (exactSearchString.length > 50) {
    var startAnchor = exactSearchString.substring(0, 20).trim();
    var endAnchor = exactSearchString.substring(exactSearchString.length - 20).trim();

    var startResults = body.search(startAnchor, searchOptions);
    var endResults = body.search(endAnchor, searchOptions);
    startResults.load("items");
    endResults.load("items");
    await context.sync();

    if (startResults.items.length > 0 && endResults.items.length > 0) {
      return startResults.items[0].expandTo(endResults.items[0]);
    }
  }

  // AI Fallback (Last resort)
  if (cleanText.length > 15) {
    return await findMatchByAI(context, cleanText);
  }

  return null;
}// [Search for "function applyFix" around line 763]
// ==========================================
// END: findSmartMatch
// ==========================================
// ==========================================
// START: applyFix
// ==========================================
function applyFix(originalText, correctionText, btnElement, replaceFullPara) {
  // 1. Define Patterns to strip noise
  var listRegex = /^(\([a-zA-Z0-9]+\)|[a-zA-Z0-9]+\.|[ivxIVX]+\.)\s+/;
  var headerRegex = /^(\d+(\.\d+)*\.?)?\s*([A-Z][A-Za-z\s\/]+)(\.|:)(\s+)/;

  // ==========================================
  // START: cleanString
  // ==========================================
  function cleanString(str) {
    var isListNum = /^[\d\w]+[\.\)]\s*$/.test(str.trim());
    if (!isListNum) {
      if (listRegex.test(str)) str = str.replace(listRegex, "").trim();
      var match = str.match(headerRegex);
      if (match && match[0].length < 60) {
        str = str.substring(match[0].length).trim();
      }
    }
    return str;
  }
  // ==========================================
  // END: cleanString
  // ==========================================

  var cleanOriginal = cleanString(originalText);
  var cleanCorrection = correctionText.includes("<")
    ? correctionText
    : cleanString(correctionText);

  return Word.run(function (context) {
    return findSmartMatch(context, cleanOriginal).then(function (range) {
      if (range) {
        // --- NEW: FULL PARAGRAPH REPLACEMENT LOGIC ---
        // [Search for "if (replaceFullPara) {" inside applyFix around line 847 in script final 3c.txt]

        // --- NEW: FULL PARAGRAPH REPLACEMENT LOGIC (With Font Matching) ---
        if (replaceFullPara) {
          var para = range.paragraphs.getFirst();

          // 1. CAPTURE ORIGINAL FONT
          // We load the font of the target range so we can re-apply it later
          range.font.load("name, size, color");

          return context.sync().then(function () {
            var insertedRange;

            // 2. INSERT NEW CONTENT
            if (
              cleanCorrection.includes("<") &&
              cleanCorrection.includes(">")
            ) {
              insertedRange = para.insertHtml(cleanCorrection, "Replace");
            } else {
              insertedRange = para.insertText(cleanCorrection, "Replace");
            }

            // 3. RE-APPLY ORIGINAL FONT
            // This ensures the new text (and its heading) matches the document's style
            if (range.font.name) insertedRange.font.name = range.font.name;
            if (range.font.size) insertedRange.font.size = range.font.size;
            if (range.font.color) insertedRange.font.color = range.font.color;

            // 4. UI FEEDBACK
            btnElement.innerText = "Fixed ✓";
            btnElement.disabled = true;
            btnElement.classList.add("btn-disabled");

            var logDetail =
              cleanCorrection.replace(/<[^>]*>/g, "").substring(0, 50) + "...";
            addAuditLog(
              "Playbook Fallback Applied",
              "Replaced clause with: '" + logDetail + "'",
            );

            return context.sync().then(function () {
              return true;
            });
          });
        }
        // ---------------------------------------------
        // ---------------------------------------------

        var hasHtml =
          cleanCorrection.includes("<") && cleanCorrection.includes(">");
        var diff = !hasHtml
          ? getStringDiff(cleanOriginal, cleanCorrection)
          : null;
        var useSpecificSearch = diff && diff.badPart.length < 250;
        var specificSearch = null;

        if (useSpecificSearch) {
          specificSearch = range.search(diff.badPart, {
            matchCase: false,
            ignorePunctuation: false,
            matchWholeWord: false,
          });
          specificSearch.load("items");
        }

        return context.sync().then(function () {
          try {
            if (
              typeof userSettings !== "undefined" &&
              userSettings.disableTrackChanges
            ) {
              context.document.changeTrackingMode = "Off";
            } else {
              context.document.changeTrackingMode = "TrackAll";
            }
          } catch (e) { }

          var targetRange = null;
          if (useSpecificSearch && specificSearch.items.length === 1) {
            specificSearch.items[0].insertText(diff.goodPart, "Replace");
            targetRange = specificSearch.items[0];
          } else {
            if (hasHtml) {
              range.insertHtml(cleanCorrection, "Replace");
            } else {
              range.insertText(cleanCorrection, "Replace");
            }
            targetRange = range;
          }

          if (targetRange) {
            try {
              targetRange.select();
            } catch (e) { }
          }

          btnElement.innerText = "Fixed ✓";
          btnElement.disabled = true;
          btnElement.classList.add("btn-disabled");

          var logDetail = cleanCorrection
            .replace(/<[^>]*>/g, "")
            .substring(0, 60);
          if (cleanCorrection.length > 60) logDetail += "...";
          addAuditLog("Applied Fix", "Updated text: '" + logDetail + "'");

          return context.sync().then(function () {
            return true;
          });
        });
        // ... inside applyFix ...
      } else {
        // 1. UPDATE BUTTON VISUALS (Red Error State)
        btnElement.innerText = "⚠️ Match Failed";
        btnElement.style.backgroundColor = "#f8d7da";
        btnElement.style.color = "#721c24";

        // 2. CREATE OVERLAY (Dimmed Background)
        var overlay = document.createElement("div");
        overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

        // 3. CREATE DIALOG BOX
        var dialog = document.createElement("div");
        dialog.className = "tool-box slide-in";
        dialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #c62828; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

        dialog.innerHTML = `
        <div style="font-size:24px; margin-bottom:10px;">⚠️</div>
        <div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Manual Fix Required</div>
        <div style="font-size:12px; color:#555; line-height:1.5; margin-bottom:15px;">
            ELI couldn't find the exact text match to apply this fix automatically.<br><br>
            This often happens if the text was edited recently or is inside a restricted area (like a table or header).<br><br>
            <strong>Please make this edit manually.</strong>
        </div>
        <button id="btnOkMatchFail" class="btn-action" style="width:100%; font-weight:bold;">OK, I'll fix it</button>
    `;

        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        // 4. BIND CLOSE BUTTON
        document.getElementById("btnOkMatchFail").onclick = function () {
          overlay.remove();
        };

        return false;
      }
    });
  }).catch(function (error) {
    console.error(error);
    btnElement.innerText = "Error";
    return false;
  });
} // ADD this helper function to the bottom of your script

// ==========================================
// START: getStringDiff
// ==========================================
function getStringDiff(oldStr, newStr) {
  // Helper: robust whitespace check
  // ==========================================
  // START: isSpace
  // ==========================================
  function isSpace(char) {
    return char && /\s/.test(char);
  }
  // ==========================================
  // END: isSpace
  // ==========================================

  // 1. Find common Prefix
  var start = 0;
  while (
    start < oldStr.length &&
    start < newStr.length &&
    oldStr[start] === newStr[start]
  ) {
    start++;
  }

  // 2. Find common Suffix
  var endOld = oldStr.length - 1;
  var endNew = newStr.length - 1;
  while (
    endOld >= start &&
    endNew >= start &&
    oldStr[endOld] === newStr[endNew]
  ) {
    endOld--;
    endNew--;
  }

  // 3. Initial Bridge
  var chunk = oldStr.substring(start, endOld + 1);
  var isInsertion = start > endOld;
  var isTiny = chunk.trim().length < 3;
  var expand = isInsertion || isTiny;

  // ==========================================
  // START: getCurrent
  // ==========================================
  function getCurrent() {
    return oldStr.substring(start, endOld + 1);
  }
  // ==========================================
  // END: getCurrent
  // ==========================================

  // ==========================================
  // START: isAmbiguous
  // ==========================================
  function isAmbiguous() {
    var s = getCurrent();
    if (!s) return true;
    return oldStr.split(s).length - 1 > 1;
  }
  // ==========================================
  // END: isAmbiguous
  // ==========================================

  // 4. THE UNIQUENESS LOOP (Now with robust space detection)
  var maxSteps = 20;
  while ((expand || isAmbiguous()) && maxSteps > 0) {
    expand = false;
    maxSteps--;
    var expanded = false;

    // Step Left
    if (start > 0) {
      start--;
      // Eat chars until we hit ANY whitespace
      while (start > 0 && !isSpace(oldStr[start])) {
        start--;
      }
      expanded = true;
    }

    // Step Right
    if (endOld < oldStr.length - 1) {
      endOld++;
      endNew++;
      while (endOld < oldStr.length - 1 && !isSpace(oldStr[endOld])) {
        endOld++;
        endNew++;
      }
      expanded = true;
    }

    if (!expanded) break;
  }

  var badPart = oldStr.substring(start, endOld + 1);
  var goodPart = newStr.substring(start, endNew + 1);

  if (!badPart || badPart.trim() === "")
    return { badPart: oldStr, goodPart: newStr };

  return { badPart: badPart, goodPart: goodPart };
} // --- NEW: HTML PARSER (Removes Track Changes, Footers, & COMMENTS) ---
// ==========================================
// END: getStringDiff
// START: getCleanTextAndContext
// ==========================================
function getCleanTextAndContext(html) {
  if (!html) return { text: "", context: "" };

  var parser = new DOMParser();
  var doc = parser.parseFromString(html, "text/html");

  // 1. Remove Deletions (<del>) to Simulate "Accept Changes"
  // We remove the node entirely, meaning the text is GONE (Accepted Deletion).
  // Insertions (<ins>) are just treated as text, meaning they STAY (Accepted Insertion).
  var delNodes = doc.querySelectorAll(
    "del, s, strike, span[style*='line-through'], span[style*='text-decoration:line-through'], span[style*='text-decoration: line-through']",
  );

  var deletedPhrases = [];
  delNodes.forEach(function (node) {
    var txt = node.textContent.trim();
    if (txt.length > 3) deletedPhrases.push(txt);
    node.remove(); // <--- This deletes the text from our memory copy
  });

  // 2. Remove Comments & Hidden Anchors
  var commentNodes = doc.querySelectorAll(
    "div[style*='mso-element:comment'], .CommentText, .msocommentanchor, a[name*='_msocom_']",
  );
  commentNodes.forEach(function (node) {
    node.remove();
  });

  // 3. Handle line breaks
  var brs = doc.querySelectorAll("br");
  brs.forEach(function (br) {
    br.replaceWith("\n");
  });

  // 4. Extract Clean Text
  var tempDiv = document.createElement("div");
  tempDiv.innerHTML = doc.body.innerHTML;
  var cleanText = tempDiv.innerText || tempDiv.textContent || "";

  // Cleanup whitespace
  cleanText = cleanText.replace(/\n\s*\n/g, "\n").trim();
  // Cleanup footer/page noise
  cleanText = cleanText.replace(/Page\s+\d+\s+(of\s+\d+)?/gi, " ");

  // 5. Build Context (Optional - used for AI hints)
  var contextMsg = "";
  if (deletedPhrases.length > 0) {
    var unique = deletedPhrases
      .filter((v, i, a) => a.indexOf(v) === i)
      .slice(0, 5);
    contextMsg =
      "NOTE: The user recently deleted these phrases: [" +
      unique.join(", ") +
      "]";
  }

  return { text: cleanText, context: contextMsg };
}
// ==========================================
// END: getCleanTextAndContext
// ==========================================
// ==========================================
// START: renderRewriteBubble (FIXED)
// ==========================================
function renderRewriteBubble(data, originalText) {
  var container = document.getElementById("chatContainer");
  var bubble = document.createElement("div");
  bubble.className = "chat-bubble chat-ai";
  bubble.style.borderLeft = "4px solid #1976d2";

  var commentBtnId = "btnComment_" + Date.now();

  bubble.innerHTML = `
        <div style="font-weight:800; color:#1565c0; font-size:10px; text-transform:uppercase; margin-bottom:6px; letter-spacing:0.5px;">SUGGESTED REWRITE</div>
        <div style="font-family:'Georgia', serif; font-size:13px; background:#e3f2fd; padding:12px; border:1px solid #90caf9; margin-bottom:12px; color:#0d47a1; line-height:1.6;">${data.rewrite}</div>
        <div style="font-size:11px; color:#546e7a; margin-bottom:12px;"><strong>Reasoning:</strong> ${data.summary}</div>
        
        <div style="display:flex; flex-direction:column; gap:8px;">
            <button class="btn-platinum-action btn-apply">
                <span style="font-size:14px;">📝</span> Apply Redline
            </button>
            <button id="${commentBtnId}" class="btn-platinum-action btn-comment hidden">
                <span style="font-size:14px;">💬</span> Attach Rationale
            </button>
        </div>
    `;
  container.appendChild(bubble);

  // Apply Platinum Styles dynamically
  var btns = bubble.querySelectorAll(".btn-platinum-action");
  btns.forEach(b => {
    b.style.cssText = `
          width: 100%;
          background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
          border: 1px solid #b0bec5;
          border-bottom: 2px solid #90a4ae;
          border-radius: 0px;
          color: #1565c0; /* Blue Text */
          font-family: 'Inter', sans-serif;
          font-size: 10px;
          font-weight: 800;
          text-transform: uppercase;
          padding: 8px;
          cursor: pointer;
          display: flex; align-items: center; justify-content: center; gap: 6px;
      `;
    b.onmouseover = function () { this.style.background = "#e3f2fd"; };
    b.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)"; };
  });

  // --- JUMP FIX ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) setTimeout(() => { scrollParent.scrollTop = scrollParent.scrollHeight; }, 10);

  // 1. APPLY REDLINE BUTTON
  var btnApply = bubble.querySelector(".btn-apply");
  var btnComment = bubble.querySelector("#" + commentBtnId);

  btnApply.onclick = function () {
    btnApply.innerText = "Applying...";
    Word.run(function (context) {
      var selection = context.document.getSelection();
      try { context.document.changeTrackingMode = "TrackAll"; } catch (e) { }
      var insertedRange = selection.insertText(data.rewrite, "Replace");
      insertedRange.font.bold = false;
      insertedRange.font.italic = false;
      insertedRange.select();
      return context.sync();
    })
      .then(() => {
        btnApply.innerHTML = "✅ Applied";
        btnApply.disabled = true;
        btnComment.classList.remove("hidden"); // Reveal comment button
        showToast("✅ Redline applied");
      })
      .catch(function (e) { console.error(e); btnApply.innerText = "Error"; });
  };

  // 2. ATTACH RATIONALE BUTTON
  btnComment.onclick = function () {
    btnComment.innerText = "Attaching...";
    Word.run(function (context) {
      var selection = context.document.getSelection();
      selection.insertComment(data.summary);
      return context.sync();
    })
      .then(() => {
        btnComment.innerHTML = "✅ Attached";
        btnComment.disabled = true;
        showToast("✅ Rationale Attached");
      });
  };
}
// ==========================================
// START: tryParseGeminiJSON
// ==========================================
// ==========================================
// HELPER: Validate Error Data (Prevents Crashes)
// ==========================================
function validateErrorData(data) {
  if (!Array.isArray(data)) return [];
  // Filter out items that don't have the minimum required fields
  // This prevents "undefined" errors in the renderCards function
  return data.filter(function (item) {
    return item && item.original && item.correction;
  });
}
// ==========================================
// START: tryParseGeminiJSON (Robust)
// ==========================================
function tryParseGeminiJSON(jsonString) {
  if (!jsonString) return {};

  // 1. Aggressive Cleanup (Your existing code)
  var clean = jsonString
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .replace(/^[^{[]*/, "") // Remove text before first { or [
    .replace(/[^}\]]*$/, "") // Remove text after last } or ]
    .trim();

  // 2. Try Standard Parse (Fast Path)
  try {
    return JSON.parse(clean);
  } catch (e) {
    console.warn("Standard Parse Failed. Attempting Recovery...");

    try {
      var recoveredArray = [];
      // This Regex looks for balanced { ... } blocks
      var objectRegex = /\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}/g;

      var match;
      while ((match = objectRegex.exec(clean)) !== null) {
        try {
          // Parse each block individually
          var validObj = JSON.parse(match[0]);
          recoveredArray.push(validObj);
        } catch (err) {
          // If one block is bad, skip it but save the rest
          console.warn("Skipping bad block:", match[0]);
        }
      }

      if (recoveredArray.length > 0) {
        console.log("✅ Recovered " + recoveredArray.length + " items via Regex.");
        return recoveredArray;
      }
    } catch (e2) {
      console.error("Regex Recovery Failed:", e2);
    }
    // ---------------------------------------------------------

    // 4. Secondary Recovery: Fix common syntax errors (Your existing code)
    try {
      var fixed = clean.replace(/(?<!["}])\n/g, "\\n");
      if (fixed.match(/^\[\w+\]/)) throw new Error("Plain text list detected");
      return JSON.parse(fixed);
    } catch (e3) {

      // 5. Emergency Fallback (Your existing code)
      if (jsonString.includes("1.") || jsonString.includes("- ")) {
        return {
          summary: "AI returned plain text analysis.",
          changes: [{
            clause: "General",
            change: jsonString.substring(0, 500),
            impact: "Review Manually"
          }]
        };
      }

      // Return empty structure to prevent app crash
      if (clean.startsWith("[")) return [];
      return {};
    }
  }
}
// ==========================================
// START: callGemini (SECURE + FULL FEATURE)
// ==========================================
function callGemini(apiKey, input, modelOverride) {
  // 1. Setup Controller
  currentController = new AbortController();
  var signal = currentController.signal;

  // 2. Determine Model (Use override OR setting OR default)
  var targetModel = modelOverride || userSettings.model || "gemini-2.5-flash-preview-09-2025";

  // 3. Format Payload
  // If input is a string, wrap it in the standard Gemini format
  // If input is already an array (multimodal), use it as is
  var contentsPayload = Array.isArray(input) ? [{ parts: input }] : [{ parts: [{ text: input }] }];

  // 4. Call YOUR Vercel Backend
  return fetch(BACKEND_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: contentsPayload, // Send full structure
      model: targetModel         // Send selected model
    }),
    signal: signal,
  })
    .then(function (res) {
      if (res.status === 503 || res.status === 429) throw new Error("OVERLOADED");
      if (!res.ok) return res.text().then(t => { throw new Error("Backend Error: " + t) });
      return res.json();
    })
    .then(function (data) {
      currentController = null;
      if (data.candidates && data.candidates.length > 0)
        return data.candidates[0].content.parts[0].text;
      return "No response from AI.";
    })
    .catch(function (err) {
      if (err.name === "AbortError") console.log("Cancelled");
      else throw err;
    });
}
// ==========================================
// END: callGemini
// ==========================================
// ==========================================
// START: setLoading (Merged: Dental Visuals + Original Logic)
// ==========================================
function setLoading(isLoading, type) {
  var loader = document.getElementById("loading");
  var out = document.getElementById("output");

  // --- 1. VISUAL INJECTION (Only runs once) ---
  // Checks if the animation exists. If not, it sets up the dental HTML structure.
  if (loader && !loader.querySelector(".mouth-container")) {
    loader.innerHTML = `
        <div class="mouth-container">
            <div class="teeth-row upper">
                <div class="tooth"></div><div class="tooth"></div>
                <div class="tooth"></div><div class="tooth"></div>
                <div class="tooth"></div>
            </div>
            <div class="teeth-row lower">
                <div class="tooth"></div><div class="tooth"></div>
                <div class="tooth"></div><div class="tooth"></div>
                <div class="tooth"></div>
            </div>
            <div class="brush-anim">🪥</div>
            <div class="foam f1"></div><div class="foam f2"></div><div class="foam f3"></div>
        </div>
        <p class="loading-text" style="margin-top: 15px; font-weight: 700; color: #1565c0; font-size: 13px;">Processing...</p>
      `;
  }
  // ------------------------------------------------------------

  var textP = loader.querySelector("p");

  // 2. SELECT BUTTONS TO LOCK (Toolbar + Arrows)
  var buttonsToLock = document.querySelectorAll(".btn-tile, .nav-arrow");

  // Safety: If already loading, prevent re-entry
  if (isLoading && !loader.classList.contains("hidden")) return;

  if (loadingInterval) {
    clearInterval(loadingInterval);
    loadingInterval = null;
  }

  // 3. Ensure Cancel Button Exists (ORIGINAL LOGIC)
  var cancelBtn = document.getElementById("btnCancelLoad");
  if (!cancelBtn) {
    cancelBtn = document.createElement("button");
    cancelBtn.id = "btnCancelLoad";
    cancelBtn.innerText = "Cancel Request";
    cancelBtn.className = "btn-action";
    cancelBtn.style.display = "block";
    cancelBtn.style.margin = "15px auto 0 auto";
    cancelBtn.style.width = "auto";

    // Optional: Visual tweak to match the compact theme, but keeps your logic intact
    cancelBtn.style.background = "transparent";
    cancelBtn.style.border = "1px solid #cfd8dc";
    cancelBtn.style.fontSize = "10px";
    cancelBtn.style.padding = "4px 10px";

    loader.appendChild(cancelBtn);
  }

  cancelBtn.onclick = function () {
    if (currentController) {
      currentController.abort();
      currentController = null;
    }
    setLoading(false);
    showOutput("<div class='placeholder' style='color:#c62828'>🛑 Request Cancelled.</div>");
  };

  if (isLoading) {
    // --- LOCK INTERFACE (Invisible) ---
    buttonsToLock.forEach(function (btn) {
      btn.style.pointerEvents = "none"; // Physically blocks clicks
      btn.style.cursor = "wait";        // Changes mouse cursor to hourglass/spinner
    });

    // Start Messages
    var messages = LOADING_MSGS[type] || LOADING_MSGS["default"];
    var i = 0;
    textP.innerText = messages[0];

    loadingInterval = setInterval(function () {
      i = (i + 1) % messages.length;
      if (i === 0 && messages.length > 2) i = 2; // Original skip logic
      textP.innerText = messages[i];
    }, 1500);

    if (loader) loader.classList.remove("hidden");
    if (out) out.classList.add("hidden");

  } else {
    // --- UNLOCK INTERFACE ---
    buttonsToLock.forEach(function (btn) {
      btn.style.pointerEvents = "auto"; // Re-enable clicks
      btn.style.cursor = "pointer";     // Restore hand cursor
    });

    if (loader) loader.classList.add("hidden");
    if (out) out.classList.remove("hidden");

    setTimeout(function () {
      if (textP) textP.innerText = "Analyzing...";
    }, 200);
  }
}
// ==========================================
// END: setLoading
// ==========================================
// ==========================================
// START: showOutput
// ==========================================
function showOutput(text, isError) {
  var out = document.getElementById("output");
  if (!out) return;
  out.innerHTML = "";
  if (isError) {
    out.style.color = "red";
    out.innerText = text;
  } else {
    out.style.color = "#212529";
    out.innerHTML = text.replace(/\n/g, "<br>");
  }
  out.classList.remove("hidden");
}
// ==========================================
// END: showOutput
// ==========================================

// ==========================================
// START: handleError (Robust Version)
// ==========================================
function handleError(error) {
  // 1. SAFETY: Handle null/undefined immediately
  if (!error) return;

  // 2. NORMALIZE: Convert string errors to objects
  // (Fixes the crash when error is just "cancelled" or "Invalid API Key")
  if (typeof error === "string") {
    // Special case for API Key
    if (error === "Invalid API Key") {
      showOutput("⚠️ Error: Please put your Google Gemini API Key at the top of the script.", true);
      return;
    }
    // Wrap other strings
    error = { message: error };
  }

  // 3. Extract Message Safely
  var msg = error.message ? error.message : JSON.stringify(error);

  // 4. IGNORE: User Cancellations (Silent)
  if (msg.includes("cancelled") || msg.includes("AbortError")) {
    console.log("Action stopped by user.");
    return;
  }

  if (error.name === "AbortError") return;

  // 5. HANDLER: Overload
  if (msg === "OVERLOADED") {
    renderOverloadCard();
    return;
  }

  // 6. HANDLER: Script Error (CORS / External)
  if (msg.toLowerCase().includes("script error")) {
    msg = "External Script Error. (This often happens if a library like Mammoth.js failed to load or due to browser security settings).";
  }

  // 7. SHOW OUTPUT
  showOutput("Error: " + msg, true);
}
// ==========================================
// END: handleError
// ==========================================

// --- NEW: Friendly Overload Message ---
// ==========================================
// START: renderOverloadCard
// ==========================================
function renderOverloadCard() {
  var out = document.getElementById("output");
  if (!out) return;
  out.innerHTML = "";
  out.classList.remove("hidden");

  var card = document.createElement("div");
  card.className = "error-card slide-in";
  // Friendly Orange/Amber Styling
  card.style.borderLeft = "5px solid #ff9800";
  card.style.backgroundColor = "#fff3e0";
  card.style.padding = "15px";

  var html = "";
  html +=
    "<div style='display:flex; align-items:center; gap:12px; margin-bottom:10px;'>";
  html += "  <span style='font-size:24px;'>🐢</span>";
  html += "  <div>";
  html +=
    "    <strong style='color:#e65100; font-size:16px; display:block;'>Gemini is Busy</strong>";
  html +=
    "    <span style='color:#bf360c; font-size:12px;'>Google's servers are overloaded.</span>";
  html += "  </div>";
  html += "</div>";
  html += "<div style='color:#5d4037; font-size:13px; line-height:1.5;'>";
  html +=
    "  Don't worry, your document is fine. The AI just needs a moment to catch its breath.<br><br>";
  html += "  <strong>Please wait 30 seconds and try again.</strong>";
  html += "</div>";

  card.innerHTML = html;
  out.appendChild(card);
}
// ==========================================
// END: renderOverloadCard
// ==========================================
// ==========================================
// START: showConfigModal (PREMIUM UNIFIED DESIGN)
// ==========================================
function showConfigModal(type) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  var titleMap = { error: "Error Check", legal: "Legal Risk", email: "Email Templates", compare: "Comparison Settings", filename: "File Naming" };
  var title = titleMap[type] || "Settings";

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.padding = "0"; // Reset padding for full-bleed header
  card.style.borderRadius = "0px";
  card.style.border = "1px solid #d1d5db";
  card.style.borderLeft = "5px solid #1565c0";
  card.style.background = "#fff";

  // --- PREMIUM HEADER ---
  var header = document.createElement("div");
  header.style.cssText = `
      background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); 
      padding: 8px 15px; 
      border-bottom: 3px solid #c9a227; 
      box-shadow: 0 2px 5px rgba(0,0,0,0.1); 
      margin-bottom: 12px; 
      position: relative;
      overflow: hidden;
  `;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
          SYSTEM CONFIGURATION
      </div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
          <span>⚙️</span> ${title}
      </div>
  `;
  card.appendChild(header);

  // --- CONTENT CONTAINER ---
  var contentDiv = document.createElement("div");
  contentDiv.style.padding = "10px 15px";

  var html = "";

  // -------------------------
  // A. EMAIL SETTINGS
  // -------------------------
  if (type === "email") {
    var defaults = { myName: "", thirdParty: { subject: "", body: "" }, requester: { subject: "", body: "" } };
    var tmpl = userSettings.emailTemplates || defaults;
    if (!tmpl.thirdParty) tmpl.thirdParty = defaults.thirdParty;
    if (!tmpl.requester) tmpl.requester = defaults.requester;

    html += `
        <div style="margin-bottom:10px;">
            <label class="tool-label">My Name</label>
            <input type="text" id="txtMyName" class="tool-input" value="${tmpl.myName}" style="width:100%;">
            <label class="checkbox-wrapper" style="margin-top:8px;">
                <input type="checkbox" id="chkUseFullName" ${tmpl.useFullName ? "checked" : ""} style="margin-right:6px;">
                <span style="font-size:11px; color:#555;">Use Full Name in Signature</span>
            </label>
        </div>
        <div style="background:#f8f9fa; padding:15px; border:1px solid #e0e0e0; margin-bottom:15px;">
            <div style="font-weight:700; color:#37474f; font-size:11px; margin-bottom:8px; border-bottom:1px solid #ccc; padding-bottom:4px;">TO: THIRD PARTY (External)</div>
            <label class="tool-label">Subject</label>
            <input type="text" id="txtSubjTP" class="tool-input" value="${tmpl.thirdParty.subject}" style="width:100%; margin-bottom:8px;">
            <label class="tool-label">Body</label>
            <textarea id="txtBodyTP" class="tool-input" style="height:100px; font-family:sans-serif; line-height:1.4;">${tmpl.thirdParty.body}</textarea>
        </div>
        <div style="background:#e3f2fd; padding:15px; border:1px solid #90caf9; margin-bottom:15px;">
            <div style="font-weight:700; color:#1565c0; font-size:11px; margin-bottom:8px; border-bottom:1px solid #90caf9; padding-bottom:4px;">TO: REQUESTER (Internal)</div>
            <label class="tool-label">Subject</label>
            <input type="text" id="txtSubjReq" class="tool-input" value="${tmpl.requester.subject}" style="width:100%; margin-bottom:8px;">
            <label class="tool-label">Body</label>
            <textarea id="txtBodyReq" class="tool-input" style="height:100px; font-family:sans-serif; line-height:1.4;">${tmpl.requester.body}</textarea>
        </div>
        <div style="font-size:9px; color:#666; background:#f5f5f5; padding:8px; border:1px solid #e0e0e0; border-radius:4px; line-height:1.5;">
            <strong>Available Variables:</strong><br>
            [Entity], [TP Entity], [Type], [CP Ref], [Date], [TP Contact], [Requester], [My Name]
        </div>
    `;
  }
  // -------------------------
  // C. FILENAME SETTINGS (NEW)
  // -------------------------
  else if (type === "filename") {
    var defaults = {
      colgate: "[Ref] - CP-[Party] -- [Type] (DRAFT [Date])",
      hills: "[Ref] - HPN-[Party] -- [Type] (DRAFT [Date])"
    };
    var current = userSettings.fileNamePatterns || defaults;

    html += `
        <div style="font-size:12px; color:#37474f; font-weight:700; margin-bottom:10px; text-transform:uppercase; border-bottom:1px solid #eee; padding-bottom:5px;">
            File Naming Convention
        </div>
        <div style="font-size:11px; color:#666; margin-bottom:15px; line-height:1.4; background:#f9fafb; padding:10px; border:1px solid #eee;">
            Define how filenames are generated automatically.<br>
            <strong>Available Tokens:</strong><br>
            <span style="font-family:monospace; color:#1565c0;">[Ref]</span>, 
            <span style="font-family:monospace; color:#1565c0;">[Party]</span>, 
            <span style="font-family:monospace; color:#1565c0;">[Type]</span>, 
            <span style="font-family:monospace; color:#1565c0;">[Date]</span>
        </div>

        <label class="tool-label">Colgate Pattern</label>
        <input type="text" id="patColgate" class="tool-input" value="${current.colgate}" style="font-family:monospace; font-size:11px; margin-bottom:15px;">

        <label class="tool-label">Hill's Pattern</label>
        <input type="text" id="patHills" class="tool-input" value="${current.hills}" style="font-family:monospace; font-size:11px;">
        
        <!-- Verification Section Hidden -->
      `;
  }
  // -------------------------
  // B. ERROR CHECK SETTINGS
  // -------------------------
  else if (type === "error") {
    html += `
        <div style="margin-bottom:15px;">
            <label class="tool-label">Regional Dialect</label>
            <select id="selDialect" class="audience-select" style="border-radius:0px;">
                <option value="US" ${userSettings.dialect === "US" ? "selected" : ""}>🇺🇸 US English (Color, Z)</option>
                <option value="UK" ${userSettings.dialect === "UK" ? "selected" : ""}>🇬🇧 UK/Global English (Colour, S)</option>
            </select>
        </div>
        <label class="checkbox-wrapper" style="margin-bottom:10px; background:#fff3e0; padding:10px; border:1px solid #ffe0b2; border-radius:0px;">
            <input type="checkbox" id="chkRefs" ${userSettings.checkBrokenRefs ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Check Broken References</strong> (Finds 'Error! Source not found')</span>
        </label>
        <label class="checkbox-wrapper" style="margin-bottom:15px; background:#e8f5e9; padding:10px; border:1px solid #c8e6c9; border-radius:0px;">
            <input type="checkbox" id="chkFooters" ${userSettings.checkFooterRef ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Check Footer Consistency</strong> (Match 'CP Ref:' across sections)</span>
        </label>
        <div style="margin-bottom:15px;">
            <label class="tool-label">Sensitivity</label>
            <select id="selSensitivity" class="audience-select" style="border-radius:0px;">
                <option value="Essential" ${userSettings.sensitivity === "Essential" ? "selected" : ""}>Essential Only (Critical Errors)</option>
                <option value="Balanced" ${userSettings.sensitivity === "Balanced" ? "selected" : ""}>Balanced (Standard)</option>
                <option value="Picky" ${userSettings.sensitivity === "Picky" ? "selected" : ""}>Picky (All Suggestions)</option>
            </select>
        </div>
        <label class="checkbox-wrapper" style="margin-bottom:15px; padding:8px;">
            <input type="checkbox" id="chkStrict" ${userSettings.strictMode ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Strict Mode</strong> (Flag passive voice & weak verbs)</span>
        </label>
        <label class="checkbox-wrapper" style="margin-bottom:15px; background:#f5f5f5; padding:10px; border:1px solid #ddd; border-radius:0px;">
            <input type="checkbox" id="chkNoTrack" ${userSettings.disableTrackChanges ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Disable Track Changes</strong> (Apply fixes directly without redlines)</span>
        </label>
    `;
  }
  // -------------------------
  // C. LEGAL RISK SETTINGS
  // -------------------------
  // C. COMPARE SETTINGS (UPDATED WITH JSON IMPORT)
  // -------------------------
  else if (type === "compare") {
    html += `
        <label class="checkbox-wrapper" style="margin-bottom:15px; background:#f5f5f5; padding:10px; border:1px solid #e0e0e0; border-radius:0px;">
            <input type="checkbox" id="chkCompareRisk" ${userSettings.showCompareRisk ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Show Risk Badges</strong> (Low, Medium, Critical)</span>
        </label>
        
        <div style="margin-top:20px; border-top:1px solid #eee; padding-top:15px;">
            <div style="font-size:12px; font-weight:700; color:#1565c0; text-transform:uppercase; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                <span>📚 Standard Templates</span>
            </div>
            
            <div id="tmplList" style="max-height:150px; overflow-y:auto; border:1px solid #e0e0e0; background:#fff; margin-bottom:10px;">
                </div>

            <div style="background:#f0f7ff; padding:10px; border:1px solid #bbdefb; margin-bottom:10px;">
                <div style="font-size:10px; font-weight:700; color:#1565c0; margin-bottom:5px;">ADD SINGLE TEMPLATE</div>
                <input type="text" id="newTmplName" class="tool-input" placeholder="Template Name (e.g. NDA 2025)" style="margin-bottom:5px; height:28px;">
                <div style="display:flex; gap:5px;">
                    <input type="file" id="newTmplFile" accept=".docx,.html,.txt" style="font-size:10px; width:100%;">
                    <button id="btnUploadTmpl" class="btn-action" style="padding:4px 8px; width:auto; font-size:10px;">Upload</button>
                </div>
                <div id="tmplUploadStatus" style="font-size:10px; color:#666; margin-top:4px;"></div>
            </div>

            <div style="background:#f9f9f9; padding:10px; border:1px dashed #ccc;">
                <div style="font-size:10px; font-weight:700; color:#546e7a; margin-bottom:5px;">BULK IMPORT (JSON)</div>
                <div style="display:flex; gap:5px; align-items:center;">
                    <input type="file" id="bulkJsonFile" accept=".json" style="font-size:10px; width:100%;">
                    <button id="btnBulkUpload" class="btn-action" style="padding:4px 8px; width:auto; font-size:10px;">Import</button>
                </div>
                <div style="font-size:9px; color:#999; margin-top:4px;">Structure: Type > Subtype > Templates</div>
            </div>
        </div>
    `;

    // --- Inject JS Logic for Template List after rendering ---
    // --- Inject JS Logic for Template List after rendering ---
    setTimeout(function () {
      var listDiv = document.getElementById("tmplList");
      // 1. Declare variable in the outer scope of this closure
      var savedTemplates = [];

      // 2. Load from Database (Async)
      // This fixes the "savedTemplates is not defined" and the "Split Brain" storage issue
      loadFeatureData(STORE_TEMPLATES).then(function (data) {
        savedTemplates = data || [];
        renderTmplList();
      });

      function renderTmplList() {
        listDiv.innerHTML = "";
        if (savedTemplates.length === 0) {
          listDiv.innerHTML = "<div style='padding:10px; font-size:11px; color:#999; text-align:center;'>No templates saved.</div>";
          return;
        }
        savedTemplates.forEach((t, idx) => {
          var row = document.createElement("div");
          row.style.cssText = "display:flex; justify-content:space-between; align-items:center; padding:6px 10px; border-bottom:1px solid #eee; font-size:11px;";
          row.innerHTML = `
                    <span style="font-weight:600; color:#37474f;">${t.name}</span>
                    <span class="del-tmpl" data-idx="${idx}" style="cursor:pointer; color:#c62828; font-weight:bold;">🗑️</span>
                `;
          listDiv.appendChild(row);
        });
        // Bind Deletes
        document.querySelectorAll(".del-tmpl").forEach(btn => {
          btn.onclick = function () {
            savedTemplates.splice(this.dataset.idx, 1);
            saveFeatureData(STORE_TEMPLATES, savedTemplates);
            renderTmplList();
          };
        });
      }

      // Bind Single Upload
      document.getElementById("btnUploadTmpl").onclick = function () {
        var name = document.getElementById("newTmplName").value.trim();
        var fileInp = document.getElementById("newTmplFile");
        var status = document.getElementById("tmplUploadStatus");

        if (!name) { status.innerText = "⚠️ Enter a name."; return; }
        if (!fileInp.files[0]) { status.innerText = "⚠️ Select a file."; return; }

        status.innerText = "⏳ Reading...";
        handleFileSelect(fileInp.files[0], function (text) {
          // savedTemplates is now guaranteed to be defined from the scope above
          savedTemplates.push({ name: name, content: text });
          saveFeatureData(STORE_TEMPLATES, savedTemplates);
          renderTmplList();
          status.innerText = "✅ Saved!";
          document.getElementById("newTmplName").value = "";
          fileInp.value = "";
        });
      };

      // Bind Bulk JSON Import
      document.getElementById("btnBulkUpload").onclick = function () {
        var fileInp = document.getElementById("bulkJsonFile");
        var btn = this;

        if (!fileInp.files[0]) {
          showToast("⚠️ Select JSON file");
          return;
        }

        btn.innerText = "⏳ Reading...";
        var reader = new FileReader();

        reader.onload = function (e) {
          try {
            var json = JSON.parse(e.target.result);
            if (!Array.isArray(json)) throw new Error("Format Error: Root must be an array");

            var validItems = [];
            var skippedCount = 0;

            json.forEach(function (item) {
              if (item && typeof item === 'object' && item.name && (item.content || item.text)) {
                validItems.push({
                  name: String(item.name).trim(),
                  content: String(item.content || item.text || ""),
                  type: String(item.type || "General"),
                  library: String(item.library || "Imported")
                });
              } else {
                skippedCount++;
              }
            });

            if (validItems.length === 0) throw new Error("No valid templates found.");

            // Update Memory & Storage
            savedTemplates = savedTemplates.concat(validItems);
            saveFeatureData(STORE_TEMPLATES, savedTemplates);

            renderTmplList();
            showToast("✅ Imported " + validItems.length + " templates.");
            btn.innerText = "Import";
            fileInp.value = "";

          } catch (err) {
            console.error("Bulk Import Error:", err);
            showToast("❌ " + err.message);
            btn.innerText = "Import";
          }
        };
        reader.readAsText(fileInp.files[0]);
      };

    }, 100);
  }
  else {
    html += `
        <div style="margin-bottom:15px;">
            <label class="tool-label">Risk Threshold (Sensitivity)</label>
            <select id="selRisk" class="audience-select" style="border-radius:0px;">
                <option value="All" ${userSettings.riskSensitivity === "All" ? "selected" : ""}>Show Everything (High Sensitivity)</option>
                <option value="Medium" ${userSettings.riskSensitivity === "Medium" ? "selected" : ""}>Ignore Low Risks (Standard)</option>
                <option value="High" ${userSettings.riskSensitivity === "High" ? "selected" : ""}>Critical/High Only (Low Sensitivity)</option>
            </select>
        </div>
        <label class="checkbox-wrapper" style="margin-bottom:15px; background:#f9fafb; padding:10px; border:1px solid #e0e0e0; border-radius:0px;">
            <input type="checkbox" id="chkScore" ${userSettings.showRiskScore ? "checked" : ""} style="margin-right:8px;">
            <span><strong>Show Risk Score Card</strong> (Visual Gauge)</span>
        </label>
        <div style="margin-top:10px; padding-top:10px; border-top:1px dashed #eee;">
             <button id="btnResetTraining" class="btn-bug" style="color:#c62828;">Reset "Ignored" Risks (${legalBlacklist.length})</button>
        </div>
    `;
  }

  // --- ADVANCED PROMPT ---
  if (type !== "email" && type !== "compare" && type !== "filename") {
    var currentPrompt = userSettings.customPrompts[type] || DEFAULT_PERSONAS[type];
    html += `
        <div style="margin-top:20px; padding-top:15px; border-top:1px dashed #d1d5db;">
            <label class="tool-label" style="color:#b71c1c; cursor:pointer;" onclick="document.getElementById('advPromptBox').classList.toggle('hidden')">🧠 AI System Instructions (Advanced)</label>
            <div id="advPromptBox" class="hidden" style="margin-top:10px;">
                <textarea id="txtCustomPrompt" class="tool-input" style="height:100px; font-family:monospace; font-size:12px; line-height:1.4;">${currentPrompt}</textarea>
                <button id="btnResetPrompt" class="btn-action" style="margin-top:8px; font-size:10px; padding:4px 8px;">Reset to Default</button>
            </div>
        </div>
    `;
  }

  contentDiv.innerHTML = html;
  card.appendChild(contentDiv);


  // --- SAVE BUTTON (PLATINUM PRIMARY) ---
  var footerDiv = document.createElement("div");
  footerDiv.style.cssText = "padding:15px 20px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px;";

  var btnCancel = document.createElement("button");
  btnCancel.className = "btn-platinum-config"; // Standard styling
  btnCancel.style.flex = "1";
  btnCancel.innerText = "CANCEL";
  btnCancel.onclick = showHome;

  var btnSave = document.createElement("button");
  btnSave.id = "btnSaveConfig";
  btnSave.className = "btn-platinum-config primary"; // Platinum (Primary variant)
  btnSave.style.cssText = "flex: 1; font-weight:800;";
  btnSave.innerText = "SAVE SETTINGS";

  footerDiv.appendChild(btnCancel);
  footerDiv.appendChild(btnSave);
  card.appendChild(footerDiv);
  out.appendChild(card);

  // --- BIND EVENTS ---
  var btnResetPrompt = document.getElementById("btnResetPrompt");
  if (btnResetPrompt) {
    btnResetPrompt.onclick = function () { document.getElementById("txtCustomPrompt").value = DEFAULT_PERSONAS[type]; };
  }

  var btnResetTrain = document.getElementById("btnResetTraining");
  if (btnResetTrain) {
    btnResetTrain.onclick = function () {
      legalBlacklist = [];
      this.innerText = "Reset Done (0)";
      this.style.color = "#2e7d32";
    };
  }

  // Save Logic
  btnSave.onclick = function () {
    if (type === "email") {
      userSettings.emailTemplates = {
        myName: document.getElementById("txtMyName").value.trim(),
        useFullName: document.getElementById("chkUseFullName").checked,
        thirdParty: { subject: document.getElementById("txtSubjTP").value, body: document.getElementById("txtBodyTP").value },
        requester: { subject: document.getElementById("txtSubjReq").value, body: document.getElementById("txtBodyReq").value }
      };
    } else if (type === "filename") {
      userSettings.fileNamePatterns = {
        colgate: document.getElementById("patColgate").value.trim(),
        hills: document.getElementById("patHills").value.trim()
      };
      // userSettings.highlightFill = document.getElementById("chkHighlightFill").checked; // Hidden in Main Settings
    } else if (type === "error") {
      userSettings.dialect = document.getElementById("selDialect").value;
      userSettings.checkBrokenRefs = document.getElementById("chkRefs").checked;
      userSettings.checkFooterRef = document.getElementById("chkFooters").checked;
      userSettings.strictMode = document.getElementById("chkStrict").checked;
      userSettings.disableTrackChanges = document.getElementById("chkNoTrack").checked;
      userSettings.sensitivity = document.getElementById("selSensitivity").value;
      var elPrompt = document.getElementById("txtCustomPrompt");
      if (elPrompt) userSettings.customPrompts[type] = elPrompt.value;
    } else if (type === "compare") {
      userSettings.showCompareRisk = document.getElementById("chkCompareRisk").checked;
      var elPrompt = document.getElementById("txtCustomPrompt");
      if (elPrompt) userSettings.customPrompts[type] = elPrompt.value;
    } else {
      userSettings.riskSensitivity = document.getElementById("selRisk").value;
      userSettings.showRiskScore = document.getElementById("chkScore").checked;
      var elPrompt = document.getElementById("txtCustomPrompt");
      if (elPrompt) userSettings.customPrompts[type] = elPrompt.value;
    }

    saveCurrentSettings();
    btnSave.innerText = "SAVED!";
    btnSave.style.background = "#2e7d32";
    btnSave.style.borderColor = "#1b5e20";
    setTimeout(() => { showHome(); }, 600);
  };
}
// ==========================================
// END: showConfigModal
// ==========================================
// ==========================================
// START: renderRiskScoreCard
// ==========================================
function renderRiskScoreCard(items) {
  // 1. Calculate Score
  // Start at 100 (Perfect). Deduct leniently.
  var score = 100;
  var highCount = 0;
  var medCount = 0;

  items.forEach(function (item) {
    var sev = (item.severity || "").toLowerCase();
    if (sev.includes("high")) {
      score -= 10; // Was 15 (Now more forgiving)
      highCount++;
    } else if (sev.includes("medium")) {
      score -= 4; // Was 5
      medCount++;
    } else {
      score -= 1; // Minimal penalty for minor issues
    }
  });

  // Bonus: If there are NO high risks, boost score slightly to bias towards green
  if (highCount === 0 && score < 95) {
    score += 5;
  }

  // Clamp score between 1 and 100
  score = Math.max(1, Math.min(100, Math.round(score)));

  // 2. Determine Color & Label
  var color = "#12b76a"; // Green
  var label = "Low Risk";
  var bg = "#ecfdf3";

  if (score < 50) {
    // Only critical if below 50 (was 60)
    color = "#b42318"; // Red
    label = "Critical Risk";
    bg = "#fef3f2";
  } else if (score < 80) {
    // Warning range (was 85)
    color = "#f79009"; // Orange
    label = "Moderate Risk";
    bg = "#fffaeb";
  }

  // 3. Render HTML
  var out = document.getElementById("output");
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.borderLeft = "5px solid " + color;
  card.style.background = bg;

  var html = `
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div style="font-size:11px; font-weight:700; color:${color}; text-transform:uppercase; letter-spacing:1px;">Colgate Safety Score</div>
                <div style="font-size:32px; font-weight:800; color:${color}; line-height:1;">${score}<span style="font-size:14px; opacity:0.6;">/100</span></div>
            </div>
            <div style="text-align:right;">
                <div style="font-size:12px; font-weight:600; color:#333;">${label}</div>
                <div style="font-size:10px; color:#666;">${highCount} High Alerts found</div>
            </div>
        </div>
        <div style="margin-top:8px; height:4px; width:100%; background:rgba(0,0,0,0.05); border-radius:2px;">
            <div style="width:${score}%; height:100%; background:${color}; border-radius:2px; transition: width 1s ease;"></div>
        </div>
    `;

  card.innerHTML = html;
  // Insert at the VERY TOP of the output
  out.insertBefore(card, out.firstChild);
}
// ==========================================
// END: renderRiskScoreCard
// ==========================================

// ==========================================
// REPLACE: validateApiAndGetText (Prevents Empty Doc Hallucinations)
// ==========================================
function validateApiAndGetText(isInline, loadingType) {
  if (!isInline && loadingType) setLoading(true, loadingType);
  if (!API_KEY || API_KEY.includes("PASTE_YOUR_KEY")) {
    setLoading(false);
    showOutput(
      "⚠️ Error: Please put your Google Gemini API Key at the top of the script.",
      true,
    );
    return Promise.reject("Invalid API Key");
  }

  return Word.run(async function (context) {
    var selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    // A. Strict Selection Check
    // If text length is 5 or less, it's likely just a cursor blink, not a real selection.
    var hasRealSelection = selection.text && selection.text.trim().length > 5;

    if (hasRealSelection) {
      if (loadingType) {
        setLoading(false);
        try {
          var choice = await askUserSelectionScope();
          setLoading(true, loadingType);
          if (choice === "selection") return selection.text;
        } catch (e) {
          throw new Error("Request cancelled.");
        }
      } else {
        return selection.text;
      }
    }

    // B. Chat Tool Logic: If no real selection, return empty string.
    // This allows the Rewrite/Defend/Translate buttons to show the "Selection Needed" bubble.
    if (isInline) {
      return "";
    }

    // Default Fallback: Get Whole Document (Only for Dashboard tools like Error Check)
    var bodyHtml = context.document.body.getHtml();
    await context.sync();

    var cleanDoc = getCleanTextAndContext(bodyHtml.value);

    // --- NEW: BLANK DOCUMENT GUARDRAIL ---
    // If the document has less than 50 characters of real text, stop here.
    if (!cleanDoc.text || cleanDoc.text.trim().length < 50) {
      if (loadingType) setLoading(false);
      throw new Error("Document is empty or too short to analyze.");
    }
    // -------------------------------------

    return cleanDoc.text;
  });
}
// ==========================================
// END: validateApiAndGetText
// ==========================================
// ==========================================
// START: renderMainFooter (CONNECTED)
// ==========================================
// GLOBAL: Agreement Data Memory
var AGREEMENT_DATA = null;
function renderMainFooter() {
  if (document.getElementById("main-footer")) return;

  var footer = document.createElement("div");
  footer.id = "main-footer";
  footer.style.cssText =
    "position:fixed; bottom:0; left:0; width:100%; background:#f8f9fa; border-top:1px solid #d1d5db; padding:0 12px; height:24px; z-index:9999; display:flex; justify-content:space-between; align-items:center; box-sizing:border-box;";

  // --- LOADER LOGIC ---
  var mode = userSettings.defaultDataSource || "onit";
  var sourceId = userSettings.dataSourceId || "Onit";

  // Check Persistence
  var savedInfo = null;
  try {
    if (typeof Office !== 'undefined' && Office && Office.context && Office.context.document && Office.context.document.settings) {
      var raw = Office.context.document.settings.get("AgreementDataDetails");
      if (raw) savedInfo = raw;
    }
  } catch (e) { console.log("Data Load Error", e); }

  var isLoaded = !!savedInfo;
  // CHANGE: Updated Text
  var rawLabel = isLoaded ? savedInfo.sourceName : (mode === 'local' ? "DATA" : "Load " + sourceId);

  // 2. FIX: Apply truncation logic on reload
  var displayLabel = rawLabel;
  if (isLoaded && displayLabel.length > 18) {
    displayLabel = displayLabel.substring(0, 15) + "...";
  }
  // RE-HYDRATION FIX:
  if (savedInfo && !AGREEMENT_DATA) {
    mode = savedInfo.type || mode;
    console.log("Re-hydrating Agreement Data from DB...");
    if (typeof PROMPTS !== 'undefined' && PROMPTS.loadDocFromDB) {
      // ✅ FIX: Use Global ID
      var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';
      PROMPTS.loadDocFromDB(uniqueKey).then(function (doc) {
        if (doc) {
          AGREEMENT_DATA = doc;
          // Set to Lightning once fully loaded
          var btn = document.querySelector("#footerBtnViewFile span");
          if (btn) btn.innerText = "⚡";
        }
      });
    }
  } else if (savedInfo && savedInfo.type) {
    mode = savedInfo.type;
  }

  var html = `
        <div style="display:flex; align-items:center; gap:0px; min-width:120px;">
           <div id="footerDataToggle" title="Switch Source: Local vs Data Source" style="display:${isLoaded ? 'none' : 'flex'}; border:1px solid #ced4da; border-radius:3px; background:#fff; cursor:pointer;" onclick="window.toggleFooterDataSource()">
               <div id="togBtnLocal" style="padding:1px 5px; font-size:8px; font-weight:700; color:${mode === 'local' ? '#fff' : '#78909c'}; background:${mode === 'local' ? '#546e7a' : 'transparent'};">FILE</div>
               <div id="togBtnOnit"  style="padding:1px 5px; font-size:8px; font-weight:700; color:${mode === 'onit' ? '#fff' : '#78909c'}; background:${mode === 'onit' ? '#546e7a' : 'transparent'};">SRC</div>
           </div>
           
           <div class="footer-divider-left" style="display:${isLoaded ? 'none' : 'block'}; width:1px; height:14px; background:#cfd8dc; margin:0 8px;"></div>
   
           <div id="footerBtnViewFile" title="View Processed File" onclick="window.openDocumentViewer()" style="display:${isLoaded ? 'flex' : 'none'}; margin-left:6px; margin-right:4px; cursor:pointer; align-items:center; justify-content:center; padding:0 4px; height:16px; border-radius:3px; background:#e0f7fa; border:1px solid #4dd0e1;">
                <span style="font-size:8px; font-weight:800; color:#006064;">${isLoaded ? '⏳' : 'VIEW'}</span>
           </div>
           
           <div class="footer-divider-middle" style="display:${isLoaded ? 'block' : 'none'}; width:1px; height:12px; background:#b0bec5; margin:0 8px;"></div>

           <div id="footerLoaderLocal" style="display:${mode === 'local' ? 'flex' : 'none'}; align-items:center;">
<label for="footerFileLoader" style="font-size:11px; font-weight:700; color:${isLoaded && mode === 'local' ? '#2e7d32' : '#546e7a'}; font-family:'Inter', sans-serif; text-transform:uppercase; letter-spacing:0.5px; cursor:${isLoaded ? 'default' : 'pointer'}; pointer-events:${isLoaded ? 'none' : 'auto'}; display:flex; align-items:center; margin:0; white-space:nowrap;">                   <span id="txtFooterFile">${mode === 'local' ? displayLabel : "DATA"}</span>
                   <span id="iconFooterFile" style="font-size:18px; margin-left:4px; transform:translateY(-2px); display:${isLoaded && mode === 'local' ? 'none' : 'inline'};">📂</span>
                </label>
                <div id="btnDeleteFile" title="Clear File & Cache" onclick="window.clearLoadedFile()" style="display:${isLoaded && mode === 'local' ? 'block' : 'none'}; cursor:pointer; margin-left:4px; font-size:10px; color:#c62828; font-weight:800; padding:0 3px;">✕</div>
                <input type="file" id="footerFileLoader" style="display:none;" onchange="window.handleFooterFileLoad(this)">
           </div>
   
           <div id="footerLoaderOnit" style="display:${mode === 'onit' ? 'flex' : 'none'}; align-items:center;">
                <div onclick="window.handleFooterSourceLoad()" style="font-size:11px; font-weight:700; color:${isLoaded && mode === 'onit' ? '#2e7d32' : '#546e7a'}; font-family:'Inter', sans-serif; text-transform:uppercase; letter-spacing:0.5px; cursor:pointer; display:flex; align-items:center; white-space:nowrap;">
                   <span id="txtFooterOnit">${mode === 'onit' ? displayLabel : "ONIT"}</span>
                   <span style="font-size:18px; margin-left:4px; transform:translateY(-2px);">☁️</span>
                </div>
           </div>
        </div>

        <div id="footerCenter" style="position:absolute; left:50%; transform:translateX(-50%); display:flex; justify-content:center;">
            <div style="display:flex; background:#e0e0e0; border-radius:0px; padding:1px; border:1px solid #cfd8dc; align-items:center;">
                <div class="footer-seg-btn" id="footSegDraft" 
                     style="padding:0 6px; font-size:9px; line-height:16px; font-weight:600; color:#78909c; cursor:pointer; border-radius:0px; transition:all 0.2s;">
                    Draft
                </div>
                <div class="footer-seg-btn" id="footSegAll" 
                     style="padding:0 6px; font-size:9px; line-height:16px; font-weight:700; color:#455a64; background:#fff; cursor:pointer; border-radius:0px; transition:all 0.2s;">
                    All
                </div>
                <div class="footer-seg-btn" id="footSegReview" 
                     style="padding:0 6px; font-size:9px; line-height:16px; font-weight:600; color:#78909c; cursor:pointer; border-radius:0px; transition:all 0.2s;">
                    Review
                </div>
            </div>
        </div>

        <div id="footerRight" style="display:flex; align-items:center; gap:0px; min-width:60px; justify-content:flex-end;">
            <span id="footerBtnSnapshot" title="Take Snapshot" style="cursor:pointer; font-size:24px; filter: sepia(1) hue-rotate(190deg) saturate(5); transform: translateY(-4px); display:inline-block;">📷</span>
            <div class="footer-divider-right" style="width:1px; height:12px; background:#b0bec5; margin:0 8px;"></div>
            <span id="footerBtnHistory" title="View Version History" style="font-size:11px; color:#546e7a; font-weight:700; font-family:'Inter', sans-serif; text-transform:uppercase; letter-spacing:0.5px; cursor:pointer;">Versions</span>
            <span id="footerVerCount" style="font-size:11px; color:#90a4ae; font-weight:600; font-family:'Inter', sans-serif; margin-left:2px;">(0)</span>
        </div>
    `;

  footer.innerHTML = html;
  document.body.appendChild(footer);
  document.body.style.paddingBottom = "30px";

  // Bind Clicks
  document.getElementById("footSegDraft").onclick = function () { setWorkflowMode('draft'); };
  document.getElementById("footSegAll").onclick = function () { setWorkflowMode('all'); };
  document.getElementById("footSegReview").onclick = function () { setWorkflowMode('review'); };

  // Bind Version History (Footer)
  var fSnap = document.getElementById("footerBtnSnapshot");
  var fHist = document.getElementById("footerBtnHistory");
  if (fSnap) fSnap.onclick = function () { showVersionHistoryInterface(); };
  if (fHist) fHist.onclick = function () { showVersionHistoryInterface(true); };

  // Update Count
  if (typeof getHistoryFromDoc === 'function') {
    getHistoryFromDoc().then(function (list) {
      var cnt = list ? list.length : 0;
      var el = document.getElementById("footerVerCount");
      if (el) el.innerText = "(" + cnt + ")";
    });
  }
}
// ==========================================
// REPLACE: clearLoadedFile (Targeted Wipe)
// ==========================================
window.clearLoadedFile = async function () {
  console.log("🗑️ Starting Nuclear Wipe...");

  // 1. INSTANT MEMORY WIPE
  window.AGREEMENT_DATA = null;
  AGREEMENT_DATA = null;
  if (typeof GLOBAL_FILENAME !== 'undefined') GLOBAL_FILENAME = "";
  if (typeof GLOBAL_FILE_BLOB !== 'undefined') GLOBAL_FILE_BLOB = null;

  // 2. WIPE DATABASE (Using Unique ID)
  if (typeof PROMPTS !== 'undefined') {
    try {
      // --- FIX: Use the Unique Document ID ---
      // This ensures we only delete the file attached to THIS document
      var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';

      // A. Overwrite with null (Safety overwrite)
      await PROMPTS.saveDocToDB(uniqueKey, null);

      // B. Delete the keys entirely
      if (PROMPTS.deleteDocFromDB) {
        await PROMPTS.deleteDocFromDB(uniqueKey);
      }

      // C. Wipe Cache
      if (PROMPTS.saveAICacheToDB) {
        await PROMPTS.saveAICacheToDB(uniqueKey, null);
      }
      console.log("✅ Database Records Destroyed for: " + uniqueKey);
    } catch (e) {
      console.error("DB Wipe Error:", e);
    }
  }

  // 3. WIPE PERSISTENCE (Office Settings)
  // This clears the "State" so the UI knows to show "DATA" instead of the filename next time
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.remove("AgreementDataDetails");
      Office.context.document.settings.saveAsync();
    }
  } catch (e) { console.warn("Settings wipe error", e); }

  // 4. RESET UI
  var label = document.querySelector("#footerLoaderLocal label");
  if (label) label.style.color = "#546e7a";

  var span = document.getElementById("txtFooterFile");
  if (span) span.innerText = "DATA";

  var icon = document.getElementById("iconFooterFile");
  if (icon) icon.style.display = "inline";

  // Hide buttons
  var viewBtn = document.getElementById("footerBtnViewFile");
  if (viewBtn) viewBtn.style.display = "none";

  var delBtn = document.getElementById("btnDeleteFile");
  if (delBtn) delBtn.style.display = "none";

  // Restore Toggles
  var toggle = document.getElementById("footerDataToggle");
  if (toggle) toggle.style.display = "flex";

  var divider = document.querySelector(".footer-divider-left");
  if (divider) divider.style.display = "block";

  var divMid = document.querySelector(".footer-divider-middle");
  if (divMid) divMid.style.display = "none";

  // Clear Input Value
  var fileInput = document.getElementById("footerFileLoader");
  if (fileInput) fileInput.value = "";
  // --- UNLOCK UPLOAD ---
  var lbl = document.querySelector("label[for='footerFileLoader']");
  if (lbl) { lbl.style.pointerEvents = "auto"; lbl.style.cursor = "pointer"; }
  var inp = document.getElementById("footerFileLoader");
  if (inp) inp.disabled = false;
  showToast("🗑️ OARS Data Removed");
};
// --- FOOTER DATA LOADER HELPERS ---

// Toggle between File and Source
window.toggleFooterDataSource = function () {
  var localBox = document.getElementById("footerLoaderLocal");
  var onitBox = document.getElementById("footerLoaderOnit");
  var btnLocal = document.getElementById("togBtnLocal");
  var btnOnit = document.getElementById("togBtnOnit");

  // Check current state (if local is flex, we go onit)
  var isLocal = localBox.style.display !== 'none';

  if (isLocal) {
    // Switch to Onit
    localBox.style.display = 'none';
    onitBox.style.display = 'flex';

    btnLocal.style.background = 'transparent';
    btnLocal.style.color = '#78909c';

    btnOnit.style.background = '#546e7a';
    btnOnit.style.color = '#fff';
  } else {
    // Switch to Local
    onitBox.style.display = 'none';
    localBox.style.display = 'flex';

    btnOnit.style.background = 'transparent';
    btnOnit.style.color = '#78909c';

    btnLocal.style.background = '#546e7a';
    btnLocal.style.color = '#fff';
  }
};

// 1. Add 'async' here
window.handleFooterFileLoad = async function (input) {
  if (input.files && input.files[0]) {
    var file = input.files[0];

    // --- UI UPDATES (Run these immediately for feedback) ---
    //showToast("🔄 Swapping File: Wiping old data...");

    // UI Update: Name
    var displayName = file.name.length > 18 ? file.name.substring(0, 15) + "..." : file.name;
    var txt = document.getElementById("txtFooterFile");
    if (txt) {
      txt.innerText = displayName;
      txt.title = file.name;
      txt.parentElement.style.color = "#2e7d32";
    }

    // UI Update: Icons
    var icon = document.getElementById("iconFooterFile");
    if (icon) icon.style.display = "none";

    var toggle = document.getElementById("footerDataToggle");
    if (toggle) toggle.style.display = "none";

    var divider = document.querySelector(".footer-divider-left");
    if (divider) divider.style.display = "none";

    var btnView = document.getElementById("footerBtnViewFile");
    if (btnView) {
      btnView.style.display = "flex";
      btnView.querySelector("span").innerText = "⏳";
    }

    var divMid = document.querySelector(".footer-divider-middle");
    if (divMid) divMid.style.display = "block";

    var btnDel = document.getElementById("btnDeleteFile");
    if (btnDel) btnDel.style.display = "block";
    // --- LOCK UPLOAD ---
    var lbl = document.querySelector("label[for='footerFileLoader']");
    if (lbl) { lbl.style.pointerEvents = "none"; lbl.style.cursor = "default"; }
    input.disabled = true;

    // --- LOGIC: CRITICAL WIPE FIX ---

    // 1. Clear RAM immediately
    AGREEMENT_DATA = null;

    var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';

    // 2. Wipe Database & Cache (WAIT for it to finish)
    if (typeof PROMPTS !== 'undefined') {
      try {
        console.log("Waiting for DB Wipe...");
        // We use 'await' here to ensure the delete finishes 
        // BEFORE we start reading/saving the new file.
        await Promise.all([
          PROMPTS.deleteDocFromDB(uniqueKey),
          PROMPTS.saveAICacheToDB(uniqueKey, null)
        ]);
        console.log("DB Wipe Complete. Starting Load.");
      } catch (e) {
        console.log("Wipe failed, but proceeding:", e);
      }
    }

    // --- LOAD LOGIC (Now safe to run) ---

    if (file.name.toLowerCase().endsWith(".docx")) {
      var reader = new FileReader();
      reader.onload = function (e) {
        var arrayBuffer = e.target.result;
        mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
          .then(function (result) {
            var cleanText = result.value.replace(/<[^>]*>/g, ' ');
            AGREEMENT_DATA = { text: cleanText, name: file.name };

            // Overwrite DB Record
            finishFooterLoad(file, uniqueKey);
          })
          .catch(function (err) { console.error("Mammoth Error", err); });
      };
      reader.readAsArrayBuffer(file);
    }
    else if (file.name.toLowerCase().endsWith(".pdf")) {
      var reader = new FileReader();
      reader.onload = function (e) {
        AGREEMENT_DATA = {
          isBase64: true,
          mime: "application/pdf",
          data: e.target.result.split(',')[1],
          name: file.name
        };
        finishFooterLoad(file, uniqueKey);
      };
      reader.readAsDataURL(file);
    }
    else {
      var reader = new FileReader();
      reader.onload = function (e) {
        AGREEMENT_DATA = { text: e.target.result, name: file.name };
        finishFooterLoad(file, uniqueKey);
      };
      reader.readAsText(file);
    }

    // Reset input
    input.value = "";
  }
};
// Helper to finish the UI/DB work after AGREEMENT_DATA is set
function finishFooterLoad(file, optionalBuffer) {
  console.log("Agreement Data Loaded: " + file.name);

  // 1. CRITICAL: Immediate Memory Wipe
  window.CURRENT_AI_RESPONSE = null;

  // 4. Persist State (Metadata)
  persistAgreementState(file.name, "local");

  // 5. Persist Data (IndexedDB)
  var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';

  // Save the DOC using the unique key...
  PROMPTS.saveDocToDB(uniqueKey, file).then(() => {

    // ...and Clear the CACHE using the SAME unique key
    // ✅ FIX: Changed 'currentDoc' to uniqueKey
    PROMPTS.saveAICacheToDB(uniqueKey, null).then(() => {
      console.log("Old AI Cache Purged.");

      // Double check memory wipe here just to be safe
      window.CURRENT_AI_RESPONSE = null;

      showToast("💾 Agreement Data Loaded");

      // Load indicator
      var statusSpan = document.querySelector("#footerBtnViewFile span");
      if (statusSpan) statusSpan.innerText = "⏳";

      // Trigger Background Analysis
      PROMPTS.runBackgroundAIAnalysis(file);
    });
  }).catch(err => console.error("DB Save Fail", err));
}

window.handleFooterSourceLoad = function () {
  var sourceId = userSettings.dataSourceId || "Onit";
  var txt = document.getElementById("txtFooterOnit");

  // Simulate Loading
  if (txt) txt.innerText = "Loading...";

  setTimeout(function () {
    if (txt) {
      txt.innerText = sourceId + " (Connected)";
      txt.parentElement.style.color = "#2e7d32";
      // CHANGE: Removed icon change logic. Icon stays ☁️.
    }

    AGREEMENT_DATA = { source: sourceId, status: "Connected" };

    // Persist State
    persistAgreementState(sourceId + " (Connected)", "onit");

  }, 800);
};
// Persist Helper
function persistAgreementState(name, type) {
  try {
    if (typeof Office !== 'undefined' && Office && Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.set("AgreementDataDetails", { sourceName: name, type: type });
      Office.context.document.settings.saveAsync(function (res) {
        console.log("Agreement State Persisted", res);
      });
    }
  } catch (e) { }
}

// ==========================================
function checkSelectionState() {
  return;
}
// ==========================================
// END: checkSelectionState
// ==========================================
// ==========================================
// START: updatePlaceholderUI
// ==========================================
function updatePlaceholderUI(hasSelection) {
  var ph = document.querySelector(".placeholder");
  if (!ph) return;

  var currentLang = localStorage.getItem("eli_language") || "en";

  if (hasSelection) {
    // STATE: PARTIAL SELECTION (Orange Warning)
    if (!ph.classList.contains("warning-mode")) {
      ph.classList.add("warning-mode");
      ph.style.borderTop = "4px solid #f57f17";
      ph.style.background = "#fff3e0";

      // --- FIX: Use data-key tags for the warning text ---
      ph.innerHTML = `
    <div style = "font-size:28px; margin-bottom:10px;" >⚠️</div >
                <strong style="color:#e65100;" data-key="msgWarnTitle">Partial Selection Active</strong><br>
                <span style="font-size:12px; color:#bf360c;" data-key="msgWarnDesc">
                    You are analyzing <strong>only</strong> the highlighted text.<br>
                    (Click anywhere else to reset to Full Document)
                </span>
            `;

      // --- FORCE TRANSLATION UPDATE IMMEDIATELY ---
      if (typeof setLanguage === "function") setLanguage(currentLang);
    }
  } else {
    // STATE: NO SELECTION (Default/Ready)
    if (ph.classList.contains("warning-mode")) {
      ph.classList.remove("warning-mode");

      // Restore Yellow Theme
      ph.style.background = "linear-gradient(180deg, #fffde7 0%, #fff9c4 100%)";
      ph.style.border = "1px solid #fff59d";
      ph.style.borderTop = "none";
      ph.style.outline = "2px dashed #cfd8dc";
      ph.style.outlineOffset = "6px";

      // --- FIX: Use data-key tags for the standard text ---
      ph.innerHTML = `
        <span data-key="msgSelectText">Select text in your document or click a button above.</span><br>
        <span style="font-size: 11px; opacity: 0.8;" data-key="msgScanLimit">If no text is selected, entire document will be scanned.</span>
      `;

      // --- FORCE TRANSLATION UPDATE IMMEDIATELY ---
      if (typeof setLanguage === "function") setLanguage(currentLang);
    }
  }
}
// ==========================================
// END: updatePlaceholderUI

// START: showHome (ADDED STEP 5)
// ==========================================
function showHome() {
  setActiveButton("");
  if (currentController) { currentController.abort(); currentController = null; }
  setLoading(false);

  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // Reset UI
  var toolbar = document.getElementById("mainToolbar");
  if (toolbar) toolbar.classList.remove("collapsed");
  var expander = document.getElementById("toolbarExpander");
  if (expander) expander.classList.remove("visible");

  // --- WORKFLOW CONTENT LOGIC ---
  var contentHTML = "";

  if (currentWorkflowMode === 'draft') {
    contentHTML = `
        <div style="font-weight:800; color:#e65100; margin-bottom:8px; text-transform:uppercase; letter-spacing:1px; border-bottom:1px solid #ffe0b2; padding-bottom:4px;">
            📝 DRAFTING WORKFLOW
            <div style="font-size:10px; font-weight:700; color:#bf360c; margin-top:2px; letter-spacing:0;">(LOAD AN ELI TEMPLATE AND AGREEMENT DATA IN FOOTER)</div>
        </div>
        <div style="text-align:left; font-size:12px; color:#5d4037; line-height:1.8;">
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #ffe0b2;">1</span> 
                <span><strong>Build Draft:</strong> Does it all: Use to build a draft, populate template fields, error check and save and send draft.</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #ffe0b2;">2</span> 
                <span><strong>Error Check:</strong> Fix grammar, typos, undefined terms, section headings, etc.</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #ffe0b2;">3</span> 
                <span><strong>Ask ELI:</strong> Refine specific clauses.</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #ffe0b2;">4</span> 
                <span><strong>Email:</strong> Send to counterparty.</span>
            </div>
        </div>
      `;
  } else if (currentWorkflowMode === 'review') {
    contentHTML = `
        <div style="font-weight:800; color:#1565c0; margin-bottom:8px; text-transform:uppercase; letter-spacing:1px; border-bottom:1px solid #bbdefb; padding-bottom:4px;">
            🔍 REVIEW WORKFLOW
        </div>
        <div style="text-align:left; font-size:12px; color:#37474f; line-height:1.8;">
            <div style="display:flex; align-items:flex-start; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #bbdefb; height:fit-content;">1</span> 
                <div>
                    <strong>Choose Your Review Method (OR):</strong>
                    <ul style="margin:4px 0 4px 15px; padding:0; list-style-type:disc; color:#546e7a;">
                        <li><strong>Playbook:</strong> Check draft against Colgate playbook and review/insert fallbacks.</li>
                        <li><strong>Compare:</strong> Analyze redlines or compare to original.</li>
                        <li><strong>Legal Issues:</strong> Find legal risks.</li>
                    </ul>
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #bbdefb;">2</span> 
                <span><strong>Ask ELI:</strong> Negotiate or fix clauses.</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-weight:bold; background:#fff; padding:0 5px; border-radius:3px; border:1px solid #bbdefb;">3</span> 
                <span><strong>Email:</strong> Send to counterparty.</span>
            </div>
        </div>
      `;
  } else {
    // Default View
    contentHTML = `
        <span data-key="msgSelectText">Click a button above to scan entire document.</span><br>
        <span style="font-size: 13px; opacity: 0.8;" data-key="msgScanLimit">If text is selected, only that text will be scanned.</span>
      `;
  }

  // Render Placeholder
  var placeholder = document.createElement("div");
  placeholder.className = "placeholder";
  if (currentWorkflowMode !== 'all') placeholder.style.padding = "20px";

  placeholder.innerHTML = contentHTML;
  out.appendChild(placeholder);

  // Restore Language
  var currentLang = localStorage.getItem("eli_language") || "en";
  if (typeof setLanguage === "function") setLanguage(currentLang);

  checkSelectionState();
}

// ==========================================
// REPLACE: showCompareInterface (Clean Object Support)
// ==========================================
function showCompareInterface() {
  var out = document.getElementById("output");
  var savedTemplates = []; // <--- 1. GLOBAL SCOPE DECLARATION
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- PREMIUM HEADER ---
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 12px; position: relative; overflow: hidden;`;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">DOCUMENT ANALYSIS</div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;"><span>↔️</span> Smart Compare</div>
  `;
  out.appendChild(header);

  // --- 1. REDLINE SUMMARY ---
  var redlineCard = document.createElement("div");
  redlineCard.className = "tool-box slide-in";
  redlineCard.style.cssText = "background: #fff; border: 1px solid #e0e0e0; border-radius: 0px; padding: 0; margin-bottom: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);";
  redlineCard.innerHTML = `
        <div style="padding: 8px 12px; border-bottom: 1px solid #f0f0f0; background: linear-gradient(to right, #fff5f5, #fff);">
            <div style="display:flex; align-items:center; gap:8px;">
                <div style="background:#ffebee; color:#c62828; width:24px; height:24px; display:flex; align-items:center; justify-content:center; border-radius:50%; font-size:12px; border:1px solid #ffcdd2;">📝</div>
                <div><div style="font-size:11px; font-weight:800; color:#b71c1c;">TRACK CHANGES ANALYSIS</div><div style="font-size:9px; color:#ef5350; font-weight:600; text-transform:uppercase;">Option 1: Existing Redlines</div></div>
            </div>
        </div>
        <div style="padding: 10px;">
            <div style="font-size:10px; color:#546e7a; margin-bottom:8px;">Analyze existing redlines in this document.</div>
            <button id="btnSumRedlines" class="btn-action" style="width:100%; padding:8px; font-size:11px;">⚡ Analyze Track Changes</button>
        </div>
    `;
  out.appendChild(redlineCard);

  // --- 2. COMPARE OPTIONS CONTAINER ---
  var rowContainer = document.createElement("div");
  rowContainer.style.cssText = "display: flex; gap: 10px; margin-bottom: 10px;";

  // --- OPTION 2A: TEMPLATE LIBRARY ---
  var templateCard = document.createElement("div");
  templateCard.className = "tool-box slide-in";
  templateCard.style.cssText = "flex: 1; background: #fff; border: 1px solid #e0e0e0; border-radius: 0px; padding: 0; margin: 0; box-shadow: 0 4px 10px rgba(0,0,0,0.05); min-width: 0;";
  // FIX: Append immediately to reserve the left slot (Option 2A)
  rowContainer.appendChild(templateCard); // <--- ADD THIS LINE


  // Load Templates
  // WRAPPER START
  loadFeatureData(STORE_TEMPLATES).then(function (data) {
    savedTemplates = data || []; // Update the global variable    // Migration Logic
    var oldLocal = localStorage.getItem("colgate_standard_templates");
    if (oldLocal) {
      try {
        var oldArr = JSON.parse(oldLocal);
        savedTemplates = savedTemplates.concat(oldArr);
        saveFeatureData(STORE_TEMPLATES, savedTemplates);
        localStorage.removeItem("colgate_standard_templates");
      } catch (e) { }
    }

    // 1. DATA PROCESSING (Group by Library -> Type)
    var libraryData = {}; // { LibName: { TypeName: [ {index, label} ] } }

    savedTemplates.forEach((t, idx) => {
      var lib = t.library || "Standard Library";
      var type = t.type || "General";
      var name = t.name;

      // Legacy Fallback (Parsing "Lib - Type - Name" string)
      if (!t.library && t.name.includes(" - ")) {
        var parts = t.name.split(" - ");
        if (parts.length >= 3) { lib = parts[0]; type = parts[1]; name = parts.slice(2).join(" - "); }
        else if (parts.length === 2) { type = parts[0]; name = parts[1]; }
      }

      if (!libraryData[lib]) libraryData[lib] = {};
      if (!libraryData[lib][type]) libraryData[lib][type] = [];
      libraryData[lib][type].push({ index: idx, label: name });
    });

    var libraryList = Object.keys(libraryData);
    var dropStyle = `width: 100%; border-radius: 0px; margin-bottom: 8px; font-size: 11px; padding: 6px; border: 1px solid #ffe0b2; background: #fff8e1; color: #e65100; font-weight: 600; outline: none;`;

    // 2. RENDER CARD CONTENT (Templates)
    var cardContent = "";
    if (savedTemplates.length > 0) {
      cardContent = `
        <label style="font-size:9px; font-weight:700; color:#f57f17; display:block; margin-bottom:4px; text-transform:uppercase;">AGREEMENT TYPE</label>
        <select id="selTmplType" class="audience-select" style="${dropStyle}"></select>

        <label style="font-size:9px; font-weight:700; color:#f57f17; display:block; margin-bottom:4px; text-transform:uppercase;">TEMPLATE VARIANT</label>
        <select id="selTmplSub" class="audience-select" style="${dropStyle} background:#fff; border-color:#e0e0e0; color:#333;"></select>
        
        <button id="btnRunTemplateCompare" class="btn-action" style="width:100%; padding:8px; font-size:10px; margin-top:4px;">⚡ Compare</button>
      `;
    } else {
      cardContent = `<div style="font-size:9px; color:#c62828; margin-bottom:8px; background:#ffebee; padding:6px; border:1px solid #ffcdd2; line-height:1.3;"><strong>No Templates.</strong><br>Import JSON in Settings.</div>`;
    }

    // 3. RENDER HEADER (Library Selector if multiple)
    var headerInner = "";
    if (libraryList.length > 1) {
      headerInner = `<select id="selLibraryRef" style="font-family:inherit; font-weight:400; font-size:9px; margin-left:20px; border:none; background:transparent; opacity:0.8; cursor:pointer; outline:none; max-width:140px;">
                        ${libraryList.map(l => `<option value="${l}">${l}</option>`).join('')}
                     </select>`;
    } else {
      headerInner = `<div style="font-weight:400; opacity:0.8; font-size:9px; margin-left:20px;">(${libraryList[0] || "Standard Library"})</div>`;
    }

    templateCard.innerHTML = `
        <div style="padding: 10px 8px; border-bottom: 1px solid #f0f0f0; background: #fff8e1;">
            <div style="font-size:10px; font-weight:800; color:#f57f17; display:flex; flex-direction:column; gap:2px;">
                <div style="display:flex; align-items:center; gap:6px;"><span>📚</span> OPTION 2A</div>
                ${headerInner}
            </div>
        </div>
        <div style="padding: 10px;">${cardContent}</div>
    `;
    if (savedTemplates.length > 0) {
      var libSel = document.getElementById("selLibraryRef");
      var typeSel = document.getElementById("selTmplType");
      var subSel = document.getElementById("selTmplSub");
      var btnTempl = document.getElementById("btnRunTemplateCompare");

      // DATA BINDING HELPERS
      function updateTypes() {
        var currentLib = libSel ? libSel.value : (libraryList[0] || "");
        var types = [];
        if (currentLib && libraryData[currentLib]) {
          types = Object.keys(libraryData[currentLib]);
        }
        typeSel.innerHTML = types.map(k => `<option value="${k}">${k}</option>`).join('');
        updateSubtypes(); // Chain reaction
      }

      function updateSubtypes() {
        var currentLib = libSel ? libSel.value : (libraryList[0] || "");
        var cat = typeSel.value;
        var items = [];
        if (libraryData[currentLib] && libraryData[currentLib][cat]) {
          items = libraryData[currentLib][cat];
        }
        subSel.innerHTML = items.map(i => `<option value="${i.index}">${i.label}</option>`).join('');
      }

      // INITIALIZE & ATTACH LISTENERS
      updateTypes();

      if (libSel) libSel.onchange = updateTypes;
      typeSel.onchange = updateSubtypes;

      btnTempl.onclick = function () {
        var template = savedTemplates[subSel.value];
        if (template && template.content) {
          runCompareAnalysis(template.content, "template");
        }
      };
    }

  });

  // --- OPTION 2B: UPLOAD FILE ---
  var deepCompareCard = document.createElement("div");
  deepCompareCard.className = "tool-box slide-in";
  deepCompareCard.style.cssText = "flex: 1; background: #fff; border: 1px solid #e0e0e0; border-radius: 0px; padding: 0; margin: 0; box-shadow: 0 4px 10px rgba(0,0,0,0.05); min-width: 0;";
  var blueDropStyle = `width: 100%; border-radius: 0px; margin-bottom: 8px; font-size: 11px; padding: 6px; border: 1px solid #bbdefb; background: #e3f2fd; color: #1565c0; font-weight: 600; outline: none;`;

  deepCompareCard.innerHTML = `
        <div style="padding: 10px 8px; border-bottom: 1px solid #f0f0f0; background: #e3f2fd;">
            <div style="font-size:10px; font-weight:800; color:#1565c0; display:flex; align-items:center; gap:6px;"><span>📂</span> OPTION 2B</div>
        </div>
        <div style="padding: 10px;">
            <label style="font-size:9px; font-weight:700; color:#1565c0; display:block; margin-bottom:4px; text-transform:uppercase;">Context</label>
            <select id="compareContextSelect" class="audience-select" style="${blueDropStyle}">
                <option value="old" selected>File is OLD (Original)</option>
                <option value="new">File is NEW (Draft)</option>
            </select>
            <div class="file-upload-zone" id="compareDropZone" style="padding: 12px 4px; border: 2px dashed #90caf9; background:#fff; text-align: center; cursor: pointer; border-radius: 0px; transition: all 0.2s; margin-top:4px;" onmouseover="this.style.background='#e3f2fd'; this.style.borderColor='#1565c0';" onmouseout="this.style.background='#fff'; this.style.borderColor='#90caf9';">
                <div style="font-size:10px; font-weight:700; color:#1565c0;">Select File</div>
                <div style="font-size:9px; color:#90caf9;">(Word / PDF)</div>
            </div>
            <input type="file" id="compareFileInput" style="display:none" accept=".docx,.pdf,.txt">
            <div id="compareStatus" style="margin-top:6px; font-size:9px; color:#666; text-align:center; display:none;"><div class="spinner" style="width:10px;height:10px;border-width:2px;display:inline-block;"></div> Reading...</div>
        </div>
    `;
  rowContainer.appendChild(deepCompareCard);
  out.appendChild(rowContainer);

  // --- BIND EVENTS ---
  var btnSum = document.getElementById("btnSumRedlines");
  if (btnSum) btnSum.onclick = function () { runRedlineSummary(this); };



  var compDrop = document.getElementById("compareDropZone");
  var compInput = document.getElementById("compareFileInput");
  var compStatus = document.getElementById("compareStatus");
  var compSelect = document.getElementById("compareContextSelect");

  if (compDrop && compInput) {
    compDrop.onclick = function () { compInput.value = null; compInput.click(); };
    compInput.onchange = function (e) {
      compStatus.style.display = "block";
      handleFileSelect(e.target.files[0], function (text) {
        compStatus.style.display = "none";
        var mode = compSelect ? compSelect.value : "old";
        runCompareAnalysis(text, mode);
      });
    };
  }

  var btnHome = document.createElement("button");
  btnHome.className = "btn-action";
  btnHome.innerHTML = "🏠 HOME";
  btnHome.style.cssText = `width: 100%; margin-top: 20px; border-radius:0px; padding:10px;`;
  btnHome.onclick = showHome;
  out.appendChild(btnHome);


}
// ==========================================
// END: showCompareInterface
// ==========================================
// ==========================================
// HELPER: Handle File Select (Extracts Text + HTML)
// ==========================================
function handleFileSelect(file, callback) {
  if (!file) return;
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    var pdfReader = new FileReader();
    pdfReader.onload = function (e) {
      // Extract Base64 string
      var b64 = e.target.result.split(',')[1];
      // Pass special object to callback instead of text
      if (callback) callback({ isBase64: true, mime: "application/pdf", data: b64 }, null);
    };
    pdfReader.readAsDataURL(file);
    return; // 🛑 EXIT HERE so we don't run the DOCX logic below
  }
  var reader = new FileReader();
  reader.onload = function (loadEvent) {
    var arrayBuffer = loadEvent.target.result;

    // 1. Extract Raw Text (For AI/Comparison)
    mammoth.extractRawText({ arrayBuffer: arrayBuffer })
      .then(function (resultText) {
        var text = resultText.value;

        // 2. Extract HTML (For 'View' Window)
        mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
          .then(function (resultHtml) {
            var html = resultHtml.value;

            // Pass BOTH back to the main function
            if (callback) callback(text, html);
          })
          .catch(function (err) {
            console.error("Mammoth HTML failed:", err);
            // Fallback: Return text, but null HTML
            if (callback) callback(text, null);
          });
      })
      .catch(function (err) {
        console.error("Mammoth Text failed:", err);
        showToast("❌ Read Failed");
      });
  };

  reader.readAsArrayBuffer(file);
}
// ==========================================
// END: handleFileSelect
// ==========================================
// ==========================================
// START: runCompareAnalysis
// ==========================================
function runCompareAnalysis(uploadedText, mode) {
  setLoading(true, "default");

  // 1. Get Current Open Document Text
  validateApiAndGetText(false, "default")
    .then(function (openDocText) {
      // 2. Determine Order
      var originalText, newText;

      if (mode === "new") {
        // Upload is NEW. Open Doc is OLD.
        originalText = openDocText;
        newText = uploadedText;
      } else {
        // Upload is OLD or TEMPLATE. Open Doc is NEW.
        originalText = uploadedText;
        newText = openDocText;
      }

      // 3. Call Gemini with the mode
      return callGemini(API_KEY, PROMPTS.COMPARE(originalText, newText, mode));
    })
    .then(function (result) {
      try {
        var data = tryParseGeminiJSON(result);
        renderCompareResults(data, mode);
      } catch (e) {
        showOutput("Error parsing comparison: " + e.message, true);
      }
    })
    .catch(handleError)
    .finally(function () {
      setLoading(false);
    });
}
// ==========================================
// END: runCompareAnalysis
// ==========================================
// ==========================================
// START: renderCompareResults
// ==========================================
function renderCompareResults(items, mode) {
  // --- SAFETY PATCH START ---
  if (!Array.isArray(items)) {
    // Check if it's the specific "AI returned plain text" error object
    if (items.changes && Array.isArray(items.changes)) {
      items = items.changes; // Recover the nested data
    } else {
      items = []; // Prevent crash
    }
  }
  var out = document.getElementById("output");
  out.innerHTML = "<div class='results-header'>Comparison Results</div>";

  if (!items || items.length === 0) {
    out.innerHTML +=
      "<div class='placeholder'>✅ No significant legal differences found.</div>";
    return;
  }

  // --- DYNAMIC LABELS ---
  var labelOld = "OLD (Deleted/Changed)";
  var labelNew = "NEW (Added)";
  var bgOld = "#ffebee"; // Redish
  var bgNew = "#e8f5e9"; // Greenish

  if (mode === "template") {
    labelOld = "TEMPLATE (Standard)";
    labelNew = "DRAFT (Variation)";
    bgOld = "#e3f2fd"; // Blueish for Template
    bgNew = "#ffebee"; // Orangeish for Draft
  }

  items.forEach(function (item, idx) {
    var card = document.createElement("div");
    card.className = "response-card slide-in";
    card.style.animationDelay = idx * 0.1 + "s";

    var color = "#333";
    var icon = "📝";
    var border = "4px solid #ccc";

    if (item.change_type === "Critical") {
      color = "#c62828";
      icon = "🚨";
      border = "4px solid #c62828";
    } else if (item.change_type === "Moderate") {
      color = "#f57f17";
      icon = "⚠️";
      border = "4px solid #f57f17";
    }

    // UPDATE: Full Border + Left Accent
    card.style.border = "1px solid #d1d5db";
    card.style.borderLeft = border;

    // UPDATE: Locate Button Logic
    var locateBtn = function (text) {
      var safe = text.replace(/"/g, "&quot;").replace(/'/g, "\\'");
      return `<button class="btn-action" style="padding:1px 6px; font-size:9px; margin-left:6px; display:inline-block; width:auto; border-radius:2px; background:#fff; border:1px solid #999; color:#333; cursor:pointer;" onclick="locateText('${safe}', this, true)">📍 Locate</button>`;
    };

    var leftBtn = (mode === "new") ? locateBtn(item.original_snippet) : "";
    var rightBtn = (mode === "old" || mode === "template") ? locateBtn(item.new_snippet) : "";

    // Explicitly check boolean (default FALSE)
    var showBadges = (userSettings.showCompareRisk === true);

    var html = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                <strong style="color:#212529; font-size:14px;">${item.section}</strong>
                ${showBadges ? `<span class="badge" style="background:${color}20; color:${color};">${icon} ${item.change_type.toUpperCase()}</span>` : ''}
            </div>
            
            <div style="font-size:13px; color:#444; margin-bottom:12px; font-weight:500;">
                ${item.description}
            </div>
            
            <div style="display:flex; gap:10px; font-size:11px;">
                <div style="flex:1; background:${bgOld}; padding:8px; border-radius:4px; border:1px solid rgba(0,0,0,0.1);">
                    <div style="font-weight:bold; color:#333; margin-bottom:4px; display:flex; align-items:center; justify-content:space-between;">
                        <span>${labelOld}</span>
                        ${leftBtn}
                    </div>
                    <div style="color:#555;">"${item.original_snippet}"</div>
                </div>
                
                <div style="flex:1; background:${bgNew}; padding:8px; border-radius:4px; border:1px solid rgba(0,0,0,0.1);">
                    <div style="font-weight:bold; color:#333; margin-bottom:4px; display:flex; align-items:center; justify-content:space-between;">
                        <span>${labelNew}</span>
                        ${rightBtn}
                    </div>
                    <div style="color:#555;">"${item.new_snippet}"</div>
                </div>
            </div>
        `;

    card.innerHTML = html;
    out.appendChild(card);
  });

  var btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "⬅ Back to Compare Upload";
  btnBack.style.marginTop = "15px";
  btnBack.onclick = showCompareInterface;
  out.appendChild(btnBack);
}
// ==========================================
// END: renderCompareResults
// ==========================================
// ==========================================
// START: generateNegotiationEmail (Fix: No Scroll Jump)
// ==========================================
function generateNegotiationEmail() {
  // 1. Gather Selected Items
  var checks = document.querySelectorAll(".negotiation-check:checked");
  if (checks.length === 0) {
    showToast("⚠️ Select at least one issue first");
    return;
  }

  var issues = [];
  checks.forEach(function (c) {
    issues.push({
      clause: c.dataset.clause,
      issue: c.dataset.issue,
    });
  });

  // 2. Show Loading UI
  var out = document.getElementById("output");
  var emailContainer = document.createElement("div");
  emailContainer.id = "negotiationResult";
  emailContainer.className = "email-inline-container slide-in";
  emailContainer.style.cssText =
    "background:#fff; border:1px solid #d1d5db; border-radius:8px; box-shadow:0 10px 25px rgba(0,0,0,0.1); margin-top:20px; padding:0; overflow:hidden;";

  emailContainer.innerHTML = `
        <div style="background:#f8f9fa; padding:15px; border-bottom:1px solid #eee; display:flex; align-items:center; gap:10px;">
            <div class="spinner" style="width:16px;height:16px;border-width:2px; border-color:#1565c0 transparent #1565c0 transparent;"></div>
            <span style="font-size:13px; font-weight:600; color:#495057;">Drafting negotiation strategy...</span>
        </div>
    `;

  var old = document.getElementById("negotiationResult");
  if (old) old.remove();
  out.appendChild(emailContainer);

  // --- FIX: SCROLL CONTAINER GENTLY INSTEAD OF JUMPING ---
  // REPLACED: emailContainer.scrollIntoView({behavior: "smooth" });
  setTimeout(() => {
    out.scrollTop = out.scrollHeight; // Scrolls the content div only
  }, 10);

  // 3. Call AI
  callGemini(API_KEY, PROMPTS.NEGOTIATE_BATCH(issues))
    .then(function (text) {
      var cleanText = text.trim();

      // ... (Cleaning Logic) ...
      try {
        if (cleanText.startsWith("{")) {
          var json = JSON.parse(cleanText);
          cleanText = json.email_body || json.email || json.text || cleanText;
        } else if (cleanText.startsWith('"') && cleanText.endsWith('"')) {
          cleanText = cleanText.slice(1, -1);
        }
      } catch (e) { }

      cleanText = cleanText.replace(/\\"/g, '"');
      cleanText = cleanText.replace(/\\n/g, "\n\n");
      cleanText = cleanText.replace(/\*\*(.*?)\*\*/g, "<b>$1</b>");

      // Fix Spacing: Normalize newlines (Max 1 blank line)
      cleanText = cleanText.replace(/(\r\n|\n|\r)/g, "\n");
      cleanText = cleanText.replace(/\n{2,}/g, "\n\n");

      // Bold numbered items without extra breaks
      cleanText = cleanText.replace(/(\d+\.)/g, "<b>$1</b>");
      // Convert to HTML
      cleanText = cleanText.replace(/\n/g, "<br>");

      var html = `
            <div style="background:#1565c0; color:white; padding:12px 16px; display:flex; justify-content:space-between; align-items:center;">
              <div style="font-weight:600; font-size:13px; display:flex; align-items:center; gap:8px;">
                <span>📧</span> Draft Email
              </div>
              <button id="btnCloseNego" style="background:transparent; border:none; color:white; font-size:16px; cursor:pointer; opacity:0.8;">✕</button>
            </div>

            <div style="background:#f8f9fa; padding:12px 16px; border-bottom:1px solid #e9ecef;">
              <span style="font-size:11px; color:#666; font-style:italic;">Copy text below to your email client.</span>
            </div>

            <div id="negoOutput" contenteditable="true"
              style="width:100%; height:300px; padding:16px; border:none; font-family:'Segoe UI', sans-serif; font-size:14px; line-height:1.3; color:#212529; overflow-y:auto; outline:none; background:white;">
              ${cleanText}
            </div>

            <div style="padding:12px 16px; background:#fff; border-top:1px solid #e9ecef; display:flex; justify-content:flex-end; gap:10px;">
              <button id="btnDiscardNego" class="btn-action" style="padding:8px 16px; color:#d32f2f; border-color:#ffcdd2; background:white;">Discard</button>
              <button id="btnCopyRich" class="btn-action" style="padding:8px 16px; background:#1565c0; color:white; border:none; font-weight:600; box-shadow:0 2px 4px rgba(21,101,192,0.2);">
                Copy Email
              </button>
            </div>
            `;
      emailContainer.innerHTML = html;
      emailContainer.style.padding = "0";

      // Bind Close/Discard
      var closer = function () { document.getElementById('negotiationResult').remove(); };
      document.getElementById("btnCloseNego").onclick = closer;
      document.getElementById("btnDiscardNego").onclick = closer;

      // Handle Rich Text Copy
      document.getElementById("btnCopyRich").onclick = function () {
        var node = document.getElementById("negoOutput");
        var range = document.createRange();
        range.selectNode(node);
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);
        document.execCommand("copy");
        window.getSelection().removeAllRanges();

        this.innerText = "Copied! ✅";
        setTimeout(() => { this.innerText = "Copy Email"; }, 2000);
      };
    })
    .catch(function (e) {
      emailContainer.innerHTML = `<div style="padding:20px; color:#c62828;">Error: ${e.message}</div>`;
    });
}
// ==========================================
// END: generateNegotiationEmail
// START: runJustifyInChat
// ==========================================
function runJustifyInChat(userReason) {
  addChatBubble("🛡️ Justifying: " + userReason, "user");
  var loaderId = addSkeletonBubble();

  callGemini(API_KEY, PROMPTS.JUSTIFY(userReason, pendingJustifyText))
    .then(function (jsonString) {
      removeBubble(loaderId);
      try {
        var data = tryParseGeminiJSON(jsonString);
        renderJustifyBubble(data);
      } catch (e) {
        addChatBubble("Error generating justification.", "ai");
      }
    })
    .catch(function (e) {
      removeBubble(loaderId);
      addChatBubble("Error: " + e.message, "ai");
    });
}
// ==========================================
// END: runJustifyInChat
// ==========================================
// ==========================================
// START: renderJustifyBubble (FIXED)
// ==========================================
function renderJustifyBubble(data) {
  var container = document.getElementById("chatContainer");
  var bubble = document.createElement("div");
  bubble.className = "chat-bubble chat-ai";
  bubble.style.borderLeft = "4px solid #7b1fa2";

  var cleanRationale = (data.rationale || "").replace(/^"|"$/g, "").trim();

  var html = `
            <div style="font-weight:800; color:#7b1fa2; font-size:10px; text-transform:uppercase; margin-bottom:6px; letter-spacing:0.5px;">NEGOTIATION RATIONALE</div>
            <div style="font-family:'Segoe UI', sans-serif; font-size:13px; background:#f3e5f5; padding:12px; border:1px solid #e1bee7; margin-bottom:12px; color:#4a148c; font-style:italic; line-height:1.5;">
              "${cleanRationale}"
            </div>

            <div style="display:flex; gap:8px;">
              <button class="btn-plat-purple btn-copy">📄 Copy</button>
              <button class="btn-plat-purple btn-comment">💬 Attach</button>
            </div>
            `;
  bubble.innerHTML = html;
  container.appendChild(bubble);

  // Apply Styles
  var btns = bubble.querySelectorAll(".btn-plat-purple");
  btns.forEach(b => {
    b.style.cssText = `
          flex: 1;
          background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
          border: 1px solid #b0bec5;
          border-bottom: 2px solid #90a4ae;
          border-radius: 0px;
          color: #7b1fa2; /* Purple Text */
          font-family: 'Inter', sans-serif;
          font-size: 10px;
          font-weight: 800;
          text-transform: uppercase;
          padding: 8px;
          cursor: pointer;
          display: flex; align-items: center; justify-content: center; gap: 6px;
      `;
    b.onmouseover = function () { this.style.background = "#f3e5f5"; };
    b.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)"; };
  });

  // --- JUMP FIX ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) setTimeout(() => { scrollParent.scrollTop = scrollParent.scrollHeight; }, 10);

  // Bind Buttons
  var btnCopy = bubble.querySelector(".btn-copy");
  btnCopy.onclick = function () {
    navigator.clipboard.writeText(cleanRationale);
    this.innerHTML = "✅ Copied";
    setTimeout(() => { btnCopy.innerHTML = "📄 Copy"; }, 2000);
  };

  var btnComment = bubble.querySelector(".btn-comment");
  btnComment.onclick = function () {
    if (typeof pendingJustifyText !== "undefined" && pendingJustifyText) {
      addWordComment(pendingJustifyText, cleanRationale, btnComment);
    } else {
      showToast("⚠️ Error: Source text reference lost.");
    }
  };
}
// ----------------------------------------------------------------------------
// HELPER: Normalize AI Arrays (Handles "string" vs ["string"] vs [{obj}])
// ----------------------------------------------------------------------------
function normalizeAIArray(input) {
  if (!input) return [];

  // Case A: Single String ("john@doe.com")
  if (typeof input === 'string') {
    // Basic cleanup
    var clean = input.replace(/[\[\]"]/g, '').trim();
    if (clean.includes(",")) {
      // "a@b.com, c@d.com" -> Split
      return clean.split(",").map(s => ({ email: s.trim() }));
    }
    return [{ email: clean }];
  }

  // Case B: Array (could be strings OR objects)
  if (Array.isArray(input)) {
    return input.map(item => {
      // If string, wrap it
      if (typeof item === 'string') return { email: item };
      // If object, keep it
      return item;
    });
  }

  return [];
}
// ==========================================
// START: renderClauseBubble (FIXED)
// ==========================================
function renderClauseBubble(data) {
  var container = document.getElementById("chatContainer");
  var bubble = document.createElement("div");
  bubble.className = "chat-bubble chat-ai";
  // Green accent for "Drafting"
  bubble.style.borderLeft = "4px solid #2e7d32";

  // 1. Render the Text Content
  var html = `
            <div style="font-weight:800; color:#2e7d32; font-size:10px; text-transform:uppercase; margin-bottom:6px; letter-spacing:0.5px;">
              ⚖️ DRAFT CLAUSE: ${data.title || "New Clause"}
            </div>
            <div style="font-family:'Georgia', serif; font-size:13px; background:#f9fafb; padding:12px; border:1px solid #e0e0e0; margin-bottom:12px; color:#37474f; line-height:1.6;">
              ${data.text.replace(/\n/g, "<br>")}
            </div>
            <div style="font-size:11px; color:#607d8b; margin-bottom:12px; font-style:italic;">
              <strong>Note:</strong> ${data.explanation}
            </div>
            `;
  bubble.innerHTML = html;

  // 2. Create Buttons (PLATINUM SQUARED STYLE)
  var btnRow = document.createElement("div");
  btnRow.style.display = "flex";
  btnRow.style.gap = "8px";

  // SHARED BUTTON STYLE
  var btnStyle = `
            flex: 1;
            background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
            border: 1px solid #b0bec5;
            border-bottom: 2px solid #90a4ae;
            border-radius: 0px;
            color: #455a64;
            font-family: 'Inter', sans-serif;
            font-size: 10px;
            font-weight: 800;
            text-transform: uppercase;
            padding: 8px 12px;
            cursor: pointer;
            display: flex; align-items: center; justify-content: center; gap: 6px;
            transition: all 0.2s;
            `;

  // --- INSERT BUTTON ---
  var btnInsert = document.createElement("button");
  btnInsert.style.cssText = btnStyle;
  // Green Text for Primary Action
  btnInsert.style.color = "#2e7d32";
  btnInsert.style.borderBottomColor = "#81c784";
  btnInsert.innerHTML = "<span style='font-size:14px;'>📍</span> Insert";

  btnInsert.onmouseover = function () { this.style.background = "#e8f5e9"; };
  btnInsert.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)"; };

  btnInsert.onclick = function () {
    // Show the "Scroll to Location" Modal
    var overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

    var dialog = document.createElement("div");
    dialog.className = "tool-box slide-in";
    dialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #2e7d32; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:0px;";

    dialog.innerHTML = `
            <div style="font-size:24px; margin-bottom:10px;">📍</div>
            <div style="font-weight:800; color:#2e7d32; margin-bottom:10px; font-size:12px; text-transform:uppercase;">Ready to Insert</div>
            <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">
              1. Click where you want this clause to go.<br>
                2. Click the button below.
            </div>
            <button id="btnConfirmInsert" class="btn-action" style="${btnStyle} width:100%; color:#2e7d32; margin-bottom:8px;">Insert Here</button>
            <button id="btnCancelInsert" class="btn-action" style="${btnStyle} width:100%;">Cancel</button>
            `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    document.getElementById("btnCancelInsert").onclick = function () { overlay.remove(); };

    document.getElementById("btnConfirmInsert").onclick = function () {
      overlay.remove();
      btnInsert.innerText = "Inserting...";
      Word.run(function (context) {
        var selection = context.document.getSelection();
        var insertedRange = selection.insertText(data.text, "Replace");
        return context.sync();
      })
        .then(function () {
          btnInsert.innerHTML = "✅ Done";
          btnInsert.disabled = true;
          showToast("✅ Clause Inserted");
        })
        .catch(function (e) {
          console.error(e);
          btnInsert.innerText = "Error";
        });
    };
  };

  // --- COPY BUTTON ---
  var btnCopy = document.createElement("button");
  btnCopy.style.cssText = btnStyle;
  btnCopy.innerHTML = "<span style='font-size:14px;'>📄</span> Copy";

  btnCopy.onclick = function () {
    navigator.clipboard.writeText(data.text);
    this.innerHTML = "✅ Copied";
    setTimeout(() => { this.innerHTML = "<span style='font-size:14px;'>📄</span> Copy"; }, 2000);
  };

  // Assemble
  btnRow.appendChild(btnInsert);
  btnRow.appendChild(btnCopy);
  bubble.appendChild(btnRow);

  container.appendChild(bubble);

  // --- JUMP FIX ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) setTimeout(() => { scrollParent.scrollTop = scrollParent.scrollHeight; }, 10);
}
// ==========================================
// START: askUserSelectionScope (Platinum V3)
// ==========================================
function askUserSelectionScope() {
  return new Promise((resolve, reject) => {
    // 1. Create Overlay
    var overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(33, 33, 33, 0.6); backdrop-filter:blur(3px); z-index:20000; display:flex; align-items:center; justify-content:center; animation: fadeIn 0.2s ease;";

    // 2. Create Modal Box (Squared & Premium)
    var box = document.createElement("div");
    box.className = "slide-in";
    box.style.cssText = `
            background: #ffffff;
            width: 320px;
            padding: 0;
            box-shadow: 0 20px 50px rgba(0,0,0,0.3);
            border-top: 4px solid #1565c0; /* Enterprise Blue Accent */
            position: relative;
            border-radius: 0px; /* STRICT SQUARED */
            font-family: 'Inter', sans-serif;
            `;

    // 3. Content Header
    box.innerHTML = `
            <div style="padding: 25px 25px 15px 25px; text-align: center;">
              <div style="font-size: 32px; margin-bottom: 12px;">🧐</div>
              <div style="font-size: 14px; font-weight: 800; color: #1565c0; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">
                Partial Selection
              </div>
              <div style="font-size: 12px; color: #546e7a; line-height: 1.5; font-weight: 500;">
                You highlighted specific text.<br>What would you like ELI to analyze?
              </div>
            </div>
            <div id="modalBtnContainer" style="padding: 10px 25px 25px 25px; display: flex; flex-direction: column; gap: 10px;">
            </div>
            `;

    // 4. Helper: Create PLATINUM Buttons
    function createPlatinumModalBtn(label, icon, callback) {
      var btn = document.createElement("button");

      // --- THE PLATINUM AESTHETIC ---
      btn.style.cssText = `
            width: 100%;
            padding: 12px;

            /* Platinum Gradient */
            background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);

            /* Borders & Radius */
            border: 1px solid #b0bec5;
            border-top: 1px solid #cfd8dc; /* light top edge */
            border-radius: 0px; /* Squared */

            /* Shadows & Typography */
            box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
            color: #546e7a;
            font-family: 'Inter', sans-serif;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            text-shadow: 0 1px 0 rgba(255,255,255,0.8);

            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.25, 0.8, 0.25, 1);
            display: flex; align-items: center; justify-content: center; gap: 8px;
            `;

      btn.innerHTML = `<span style="font-size:14px; filter: grayscale(100%); opacity:0.8;">${icon}</span> ${label}`;

      // --- HOVER STATE (Blue Lift) ---
      btn.onmouseover = function () {
        this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)";
        this.style.borderColor = "#90caf9";
        this.style.color = "#1565c0";
        this.style.transform = "translateY(-1px)";
        this.style.boxShadow = "0 4px 8px rgba(21, 101, 192, 0.15), inset 0 1px 0 #ffffff";

        // Colorize Icon on Hover
        var iconSpan = this.querySelector("span");
        if (iconSpan) { iconSpan.style.filter = "none"; iconSpan.style.opacity = "1"; }
      };

      btn.onmouseout = function () {
        this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
        this.style.borderColor = "#b0bec5";
        this.style.color = "#546e7a";
        this.style.transform = "translateY(0)";
        this.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08)";

        // Reset Icon
        var iconSpan = this.querySelector("span");
        if (iconSpan) { iconSpan.style.filter = "grayscale(100%)"; iconSpan.style.opacity = "0.8"; }
      };

      btn.onclick = function () {
        overlay.remove();
        callback();
      };
      return btn;
    }

    // 5. Create & Append Buttons
    var container = box.querySelector("#modalBtnContainer");

    // Option A: Selection Only
    var btnSel = createPlatinumModalBtn("Target Selection Only", "🎯", function () {
      resolve("selection");
    });

    // Option B: Full Document
    var btnFull = createPlatinumModalBtn("Analyze Whole Document", "📄", function () {
      resolve("full");
    });

    container.appendChild(btnSel);
    container.appendChild(btnFull);

    // 6. Close 'X' Button
    var closeX = document.createElement("div");
    closeX.innerHTML = "✕";
    closeX.style.cssText = `
            position: absolute; top: 8px; right: 8px;
            cursor: pointer; font-size: 14px; color: #b0bec5;
            width: 24px; height: 24px; text-align: center; line-height: 24px;
            transition: color 0.2s;
            `;
    closeX.onmouseover = function () { this.style.color = "#d32f2f"; };
    closeX.onmouseout = function () { this.style.color = "#b0bec5"; };

    closeX.onclick = function () {
      overlay.remove();
      reject("cancelled");
    };

    box.appendChild(closeX);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
  });
}
// ==========================================
// ==========================================
function showToast(message) {
  // 1. Create the element
  var toast = document.createElement("div");
  toast.innerText = message;

  // 2. Style it (Floating pill at bottom center)
  toast.style.cssText =
    "position:fixed; bottom:30px; left:50%; transform:translateX(-50%); background:#323232; color:white; padding:10px 20px; border-radius:4px; font-size:13px; z-index:2147483647; box-shadow:0 4px 10px rgba(0,0,0,0.3); opacity:0; transition:opacity 0.3s ease; pointer-events:none; white-space:nowrap; max-width: 90%; box-sizing: border-box;";

  document.body.appendChild(toast);

  // 3. Animate In
  setTimeout(function () {
    toast.style.opacity = "1";
  }, 10);

  // 4. Animate Out & Remove
  setTimeout(function () {
    toast.style.opacity = "0";
    setTimeout(function () {
      toast.remove();
    }, 300);
  }, 3000);
}
// ==========================================
// END: showToast
// ==========================================
// ==========================================
// START: locateSummaryItem (Stable Layout + Manual Fix Dialog)
// ==========================================
function locateSummaryItem(el, textToFind, riskTitle, riskDetail, forceParagraph) {
  if (!textToFind) return;
  var originalHtml = el.innerHTML; // Capture full HTML (icon + text) for restore

  // 1. UI Feedback (Lock width to prevent jumping)
  el.style.width = el.offsetWidth + "px";
  el.style.justifyContent = "center";
  el.innerHTML = "Scanning...";
  el.style.opacity = "0.7";

  // --- MULTI-SECTION RISK CHECK (Keep existing logic) ---
  var sectionRegex = /(?:Section|Clause|Article)\s+(\d+(?:\.\d+)*|[IVX]+)/gi;
  var foundSections = (riskDetail || "").match(sectionRegex) || [];
  var titleSections = (riskTitle || "").match(sectionRegex) || [];
  foundSections = foundSections.concat(titleSections);
  var rootSet = new Set();
  foundSections.forEach(function (s) {
    var val = s.replace(/Section|Clause|Article/gi, "").trim();
    var numberMatch = val.match(/^(\d+)/);
    if (numberMatch) rootSet.add(numberMatch[1]);
    else rootSet.add(val.toLowerCase());
  });

  if (Array.from(rootSet).length > 1) {
    // (Simplified Multi-Section Dialog - kept brief for stability)
    var uniqueDisplayNames = foundSections.map(s => s.trim()).filter((v, i, a) => a.indexOf(v) === i);
    var listHtml = `<ul style="margin:10px 0; padding-left:20px;">${uniqueDisplayNames.map(s => `<li style="text-transform:capitalize;"><strong>${s}</strong></li>`).join("")}</ul>`;

    var overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.4); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";
    var dialog = document.createElement("div");
    dialog.className = "tool-box slide-in";
    dialog.style.cssText = "background:white; width:320px; border-left:5px solid #d32f2f; padding:20px; box-shadow:0 10px 25px rgba(0,0,0,0.2); border-radius:4px;";
    dialog.innerHTML = `<div style="font-size:24px; margin-bottom:10px;">⚠️</div><div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Multi-Section Risk</div><div style="font-size:12px; color:#555; line-height:1.5; margin-bottom:15px;">This risk spans multiple areas:<br>${listHtml}</div><button id="btnOkLocate" class="btn-action" style="width:100%; font-weight:bold;">OK, Review Manually</button>`;
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    document.getElementById("btnOkLocate").onclick = function () { overlay.remove(); };

    // Reset Button immediately
    el.innerHTML = originalHtml;
    el.style.width = "";
    el.style.opacity = "1";
    return;
  }

  // --- FRAGMENT PREPARATION ---
  var fragments = textToFind.split(/[\r\n]+|\.{3,}/g);
  fragments = fragments.filter(function (f) { return f.trim().length > 20; });
  fragments.sort(function (a, b) { return b.length - a.length; });
  var quoteTarget = fragments.length > 0 ? fragments[0].trim() : textToFind.trim();

  Word.run(async function (context) {
    // A. CAPTURE & TOGGLE TRACKING OFF (Fixes Search)
    var doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    var wasTracking = doc.changeTrackingMode !== "Off";
    if (wasTracking) {
      doc.changeTrackingMode = "Off";
      await context.sync();
    }

    // B. PERFORM SEARCH LAYERS
    var range = await findSmartMatch(context, quoteTarget);
    var resultType = "none";

    if (range) {
      if (forceParagraph) range.paragraphs.getFirst().getRange().select();
      else range.select();
      resultType = "quote";
    } else {
      // Layer 1.5: Fuzzy Scan
      var fuzzyRange = await findByFuzzyScan(context, quoteTarget);
      if (fuzzyRange) {
        fuzzyRange.select();
        resultType = "quote";
      } else if (riskTitle && riskTitle.length > 3) {
        // Layer 2: Title Scan
        var range2 = await findSmartMatch(context, riskTitle);
        if (range2) {
          range2.paragraphs.getFirst().getRange().select();
          resultType = "label";
        } else {
          // Layer 3: Brute Paragraph Scan
          var range3 = await findByParagraphScan(context, riskTitle);
          if (range3) {
            range3.select();
            resultType = "header";
          }
        }
      }
    }

    // C. RESTORE TRACKING
    if (wasTracking) {
      doc.changeTrackingMode = "TrackAll";
      await context.sync();
    }

    return resultType;
  })
    .then(function (result) {
      if (result === "quote" || result === "label" || result === "header") {
        // SUCCESS
        el.innerHTML = "<span>✅</span> Found"; // Span ensures icon alignment matches original
        el.style.backgroundColor = "#e8f5e9";
        el.style.borderColor = "#2e7d32";
        el.style.color = "#2e7d32";
        showToast("📍 Found");
      } else {
        // FAILURE: Show "Manual Fix" Dialog
        el.innerHTML = "⚠️ Not Found";
        el.style.backgroundColor = "#fff3e0";
        el.style.borderColor = "#f57f17";
        el.style.color = "#e65100";

        var overlay = document.createElement("div");
        overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

        var dialog = document.createElement("div");
        dialog.className = "tool-box slide-in";
        dialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #c62828; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

        dialog.innerHTML = `
            <div style="font-size:24px; margin-bottom:10px;">⚠️</div>
            <div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Cannot Locate Text</div>
            <div style="font-size:12px; color:#555; line-height:1.5; margin-bottom:15px;">
              ELI couldn't find the section: <strong>"${(riskTitle || "").substring(0, 30)}..."</strong><br><br>
                It may have been renamed or deleted.<br>
                  <strong>Please scroll to it manually.</strong>
                </div>
                <button id="btnOkSumFail" class="btn-action" style="width:100%; font-weight:bold;">OK</button>
                `;

        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        document.getElementById("btnOkSumFail").onclick = function () {
          overlay.remove();
        };
      }

      // 3. RESET BUTTON VISUALS (After Delay)
      setTimeout(function () {
        el.innerHTML = originalHtml; // Restore original HTML (Icon + Text)
        el.style.backgroundColor = "";
        el.style.borderColor = "";
        el.style.color = "";
        el.style.width = ""; // Release width lock
        el.style.opacity = "1";
      }, 2000);
    })
    .catch(function (e) {
      console.error("Locate Error:", e);
      el.innerHTML = originalHtml;
      el.style.width = "";
      el.style.opacity = "1";
    });
}
// ==========================================
// END: locateSummaryItem

// START: updateProgressBar
// ==========================================
function updateProgressBar() {
  // 1. Remove the bar if it exists
  var bar = document.getElementById("reviewProgressBar");
  if (bar) bar.remove();

  // 2. Ensure the main footer is visible (in case the bar hid it)
  var mainFooter = document.getElementById("main-footer");
  if (mainFooter) mainFooter.style.display = "flex";
}
// ==========================================
// END: updateProgressBar
// ==========================================
// ==========================================
// START: findByParagraphScan
// ==========================================
function findByParagraphScan(context, textToFind) {
  // 1. Create a "Fingerprint" of the target text
  // Remove ALL punctuation, spaces, and make lowercase.
  // Example: "Term (3) Years." -> "term3years"
  var cleanTarget = textToFind.toLowerCase().replace(/[^a-z0-9]/g, "");

  // Safety check: Don't scan for tiny strings
  if (cleanTarget.length < 10) return Promise.resolve(null);

  // 2. Load all paragraphs (Batching for performance)
  var paragraphs = context.document.body.paragraphs;
  paragraphs.load("text");

  return context.sync().then(function () {
    // 3. Iterate through paragraphs in JavaScript
    for (var i = 0; i < paragraphs.items.length; i++) {
      var para = paragraphs.items[i];

      // Create fingerprint of the document paragraph
      var paraText = para.text.toLowerCase().replace(/[^a-z0-9]/g, "");

      // 4. Check for Match
      // We check if the paragraph *contains* our target fingerprint
      if (paraText.includes(cleanTarget)) {
        return para; // Return the whole paragraph object
      }
    }
    return null; // Not found
  });
}
// ==========================================
// END: findByParagraphScan
// ==========================================
// ==========================================
// START: findByFuzzyScan
// ==========================================
function findByFuzzyScan(context, textToFind) {
  // 1. Tokenize: Split search text into significant words (4+ letters)
  var targetWords = textToFind.toLowerCase().match(/\b\w{4,}\b/g);

  // Safety: If text is too short to tokenize, give up
  if (!targetWords || targetWords.length < 3) return Promise.resolve(null);

  // 2. Load all paragraphs in the document
  var paragraphs = context.document.body.paragraphs;
  paragraphs.load("text");

  return context.sync().then(function () {
    var bestPara = null;
    var bestScore = 0;

    // 3. Iterate and Score
    for (var i = 0; i < paragraphs.items.length; i++) {
      var para = paragraphs.items[i];
      var paraText = para.text.toLowerCase();

      // Skip short paragraphs
      if (paraText.length < textToFind.length * 0.5) continue;

      var matches = 0;
      // Count how many target words exist in this paragraph
      for (var j = 0; j < targetWords.length; j++) {
        if (paraText.includes(targetWords[j])) {
          matches++;
        }
      }

      // Calculate score (Percentage of words matched)
      var score = matches / targetWords.length;

      // 4. Threshold: Must match at least 60% of words to win
      if (score > 0.6 && score > bestScore) {
        bestScore = score;
        bestPara = para;
      }
    }

    return bestPara; // Returns the winner (or null)
  });
}
// ==========================================
// END: findByFuzzyScan
// ==========================================
// ==========================================
// START: exportDealAbstract
// ==========================================
function exportDealAbstract(summaryData) {
  if (!summaryData || !summaryData.key_terms) {
    showToast("⚠️ No summary data to export. Run Summary first.");
    return;
  }

  // 1. Flatten the Data for CSV
  var rows = [["Field", "Value", "Source Clause"]];

  // Add Metadata
  rows.push(["Document Type", summaryData.doc_type || "Unknown", ""]);
  rows.push(["Risk Level", summaryData.risk_level || "Unknown", ""]);

  // Add Key Terms
  for (var key in summaryData.key_terms) {
    var item = summaryData.key_terms[key];
    var val = typeof item === "object" && item.value ? item.value : item;
    var quote = typeof item === "object" && item.quote ? item.quote : "";

    // Clean up text for CSV (remove newlines/commas)
    val = val.replace(/[\r\n]+/g, " ").replace(/"/g, '""');
    quote = quote.replace(/[\r\n]+/g, " ").replace(/"/g, '""');

    rows.push([key, `"${val}"`, `"${quote}"`]); // Quote to handle commas safely
  }

  // 2. Convert to CSV String
  var csvContent = rows.map((e) => e.join(",")).join("\n");

  // 3. Trigger Download
  var blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  var url = URL.createObjectURL(blob);
  var link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", "Deal_Abstract_" + Date.now() + ".csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  showToast("📊 Abstract Exported!");
}
// ==========================================
// END: exportDealAbstract
// ==========================================
// ==========================================
// START: autoTagAdobe
// ==========================================
function autoTagAdobe() {
  Word.run(function (context) {
    var body = context.document.body;

    // 1. Search for Signature Block Indicators
    // We look for "By:" as the anchor point
    var searchResults = body.search("By:", { matchCase: true });
    searchResults.load("items");

    return context.sync().then(function () {
      var count = 0;

      // 2. Iterate through found locations
      searchResults.items.forEach(function (range, index) {
        // Logic: Alternate between Signer 1 and Signer 2
        // Even index (0, 2) = Signer 1 (Company)
        // Odd index (1, 3) = Signer 2 (Counterparty)
        var signerNum = (index % 2) + 1;

        // Adobe Sign Syntax: {{Sig_es_:signerN:signature}}
        var tag = `{{Sig_es_:signer${signerNum}:signature}}`;

        // 3. Insert Tag
        // Insert "After" the "By:" so it sits on the signature line
        var inserted = range.insertText(" " + tag, "After");

        // 4. "Hide" it (White text, tiny font)
        // Adobe reads the text even if it's invisible to the human eye.
        inserted.font.color = "white";
        inserted.font.size = 1;

        count++;
      });

      if (count > 0) {
        showToast("✅ Added " + count + " hidden Adobe Sign tags.");
      } else {
        showToast("⚠️ No 'By:' lines found to tag.");
      }
    });
  }).catch(function (e) {
    console.error(e);
    showOutput("Error tagging for Adobe: " + e.message, true);
  });
}
// ==========================================
// END: autoTagAdobe
// ==========================================
var pendingTranslateText = "";

// ==========================================
// START: runTranslateInChat
// ==========================================
function runTranslateInChat(language) {
  addChatBubble("🌐 Translate to " + language, "user");
  var loaderId = addSkeletonBubble();

  callGemini(API_KEY, PROMPTS.TRANSLATE(language, pendingTranslateText))
    .then((jsonString) => {
      removeBubble(loaderId);
      var data = tryParseGeminiJSON(jsonString);
      renderTranslateBubble(data, pendingTranslateText);
    })
    .catch((e) => {
      removeBubble(loaderId);
      addChatBubble("Error during translation: " + e.message, "ai");
    });
}
// ==========================================
// END: runTranslateInChat
// ==========================================

// ==========================================
// START: renderTranslateBubble (FIXED)
// ==========================================
function renderTranslateBubble(data, originalText) {
  var container = document.getElementById("chatContainer");
  var bubble = document.createElement("div");
  bubble.className = "chat-bubble chat-ai";
  bubble.style.borderLeft = "4px solid #7b1fa2";

  bubble.innerHTML = `
                <div style="font-weight:800; color:#4a148c; font-size:10px; text-transform:uppercase; margin-bottom:6px; letter-spacing:0.5px;">TRANSLATION (${data.language})</div>
                <div style="font-size:13px; background:#f3e5f5; padding:12px; border:1px solid #e1bee7; margin-bottom:12px; color:#4a148c; line-height:1.5;">${data.translation}</div>
                <button class="btn-plat-trans" style="width:100%;">Replace Selection</button>
                `;
  container.appendChild(bubble);

  // Style Button
  var btn = bubble.querySelector(".btn-plat-trans");
  btn.style.cssText = `
                width: 100%;
                background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
                border: 1px solid #b0bec5;
                border-bottom: 2px solid #90a4ae;
                border-radius: 0px;
                color: #4a148c; /* Deep Purple */
                font-family: 'Inter', sans-serif;
                font-size: 10px;
                font-weight: 800;
                text-transform: uppercase;
                padding: 10px;
                cursor: pointer;
                `;
  btn.onmouseover = function () { this.style.background = "#f3e5f5"; };
  btn.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)"; };

  // --- JUMP FIX ---
  var scrollParent = document.getElementById("output");
  if (scrollParent) setTimeout(() => { scrollParent.scrollTop = scrollParent.scrollHeight; }, 10);

  btn.onclick = function () {
    btn.innerText = "Applying...";
    Word.run(function (context) {
      var selection = context.document.getSelection();
      var insertedRange = selection.insertText(data.translation, "Replace");
      insertedRange.font.bold = false;
      insertedRange.font.italic = false;
      return context.sync();
    }).then(() => {
      btn.innerText = "✅ Replaced";
      btn.disabled = true;
      showToast("✅ Text Updated");
    });
  };
}
// --- HELPER: AI DEEP SEARCH (The "Harvey" Layer) ---
// [Search for function findMatchByAI around line 2394 in script final 3c.txt]
// [Search for function findMatchByAI around line 2397 in script final 3c.txt]
// ==========================================
// START: findMatchByAI
// ==========================================
async function findMatchByAI(context, textToFind) {
  // 1. Safety Check
  if (!textToFind || textToFind.length < 3) return null;

  // 2. LOAD DOCUMENT TEXT (Critical Fix)
  // We need to feed the text to the AI so it can find the header.
  var body = context.document.body;
  body.load("text");
  await context.sync();
  var docText = body.text;

  // 3. Determine Search Mode
  // If text is short (< 60 chars), assume it's a "Label" (e.g. "Indemnity") that needs semantic matching.
  // If text is long, assume it's a "Quote" that needs exact matching.
  var isLabelSearch = textToFind.length < 60;

  var instruction = "";
  if (isLabelSearch) {
    instruction = `
        MODE: CONCEPTUAL HEADER SEARCH
        The user is looking for the section related to: "${textToFind}".
        TASK: Scan the document text below. Find the ACTUAL SECTION HEADER or NUMBERED LIST ITEM that corresponds to this concept.
        (Example: If user seeks "Indemnity" but doc says "12. Hold Harmless", select "12. Hold Harmless").
        (Example: If user seeks "Term" but doc says "2.1 Duration", select "2.1 Duration").
      `;
  } else {
    instruction = `
        MODE: EXACT TEXT LOCATION
        TARGET TEXT: "${textToFind.substring(0, 500)}"
        TASK: Find this exact text snippet in the document below.
        CRITICAL: If the text starts with numbering (e.g. "1.1"), include it in the anchor.
      `;
  }

  // 4. Construct Prompt
  var prompt = `
                ${instruction}

                OUTPUT FORMAT: Return a RAW JSON object with the exact start and end strings found in the document text.
                {"found": true, "start_anchor": "exact first 10-15 chars of the found header/text", "end_anchor": "exact last 10-15 chars of the found header/text" }

                DOCUMENT TEXT:
                """
                ${docText.substring(0, 500000)}
                """
                `;

  // 5. Call AI
  try {
    var jsonStr = await callGemini(API_KEY, prompt);
    var result = tryParseGeminiJSON(jsonStr);

    if (!result.found) return null;

    // 6. Search for the Anchors returned by AI
    // The AI tells us "I found '12. Hold Harmless'". Now we use Word Search to highlight it.
    var startResults = body.search(result.start_anchor, {
      matchCase: false,
      ignoreSpace: true,
    });
    var endResults = body.search(result.end_anchor, {
      matchCase: false,
      ignoreSpace: true,
    });
    startResults.load("items");
    endResults.load("items");
    await context.sync();

    if (startResults.items.length > 0 && endResults.items.length > 0) {
      var startRange = startResults.items[0];
      // Find the closest end anchor to the start anchor to avoid spanning huge gaps
      var endRange = endResults.items[0];

      // Create the range covering the header
      var preciseRange = startRange.expandTo(endRange);
      return preciseRange;
    }
  } catch (e) {
    console.warn("AI Deep Search failed:", e);
  }
  return null;
} // --- NEW FEATURE: AI AUDIT TRAIL ---
// ==========================================
// END: findMatchByAI
// ==========================================
// ==========================================
// START: addAuditLog
// ==========================================
function addAuditLog(actionType, details) {
  // 1. Check if feature is enabled
  if (!userSettings.enableAuditTrail) return;

  Word.run(function (context) {
    var body = context.document.body;

    // 2. Prepare the Log Entry
    var timestamp = new Date().toLocaleString();
    // Optional: Add extra newlines "\n" if you want even more spacing
    var logEntry = `[AI AUDIT] ${timestamp} | ACTION: ${actionType} | DETAILS: ${details}`;

    // 3. Insert as a NEW Paragraph at the End (Forces Hard Return)
    var paragraph = body.insertParagraph(logEntry, "End");

    // 4. Style it (Larger & Distinct)
    paragraph.font.size = 10; // Increased from 8
    paragraph.font.color = "#555555"; // Darker Grey for readability
    paragraph.font.name = "Courier New"; // Monospace alignment

    return context.sync();
  }).catch(function (e) {
    console.warn("Audit Log failed:", e);
  });
}
// ==========================================
// START: showInterfaceConfig (FIXED DROPDOWN DEFAULT)
// ==========================================
function showInterfaceConfig() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. Local State (Deep Copy)
  var localMaps = JSON.parse(JSON.stringify(userSettings.workflowMaps || {
    "default": [],
    "draft": [],
    "review": []
  }));

  // --- FIX: Ensure activeEditMode matches a valid dropdown option ---
  var activeEditMode = currentWorkflowMode;
  if (activeEditMode !== "draft" && activeEditMode !== "review") {
    activeEditMode = "default"; // Fallback for 'all' or undefined
  }

  // Definition of Tools (Metadata)
  var toolMeta = {
    "error": { label: "Error Check", icon: "📝" },
    "legal": { label: "Legal Assist", icon: "⚖️" },
    "summary": { label: "Summary", icon: "📋" },
    "chat": { label: "Ask ELI", icon: "💬" },
    "playbook": { label: "Playbook", icon: "📖" },
    "compare": { label: "Compare", icon: "↔️" },
    "template": { label: "Build Draft", icon: "⚡" },
    "email": { label: "Email Tool", icon: "✉️" }
  };

  // Ensure store apps are included
  TOOL_DEFINITIONS.forEach(function (t) {
    if (!toolMeta[t.key]) toolMeta[t.key] = { label: t.label, icon: t.icon };
  });

  // 2. Container
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding:0; overflow:hidden; border-left:5px solid #1565c0; border-radius:0px; background:#fff; position:relative;";

  // 3. Header
  var header = document.createElement("div");
  header.style.cssText = "background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; display: flex; justify-content: space-between; align-items: center;";
  header.innerHTML = '<div style="color:white;">' +
    '<div style="font-family:\'Inter\',sans-serif; font-size:9px; font-weight:700; color:#8fa6c7; text-transform:uppercase; letter-spacing:2px;">INTERFACE</div>' +
    '<div style="font-family:\'Georgia\',serif; font-size:16px;">🎨 Customize Toolbar</div>' +
    '</div>' +
    '<div id="btnCloseConfig" style="cursor:pointer; color:#fff; font-weight:bold; font-size:16px; opacity:0.8;">✕</div>';
  card.appendChild(header);

  // 4. Workflow Selector
  var selectorRow = document.createElement("div");
  selectorRow.style.cssText = "background:#f1f8e9; padding:12px 15px; border-bottom:1px solid #dcedc8; display:flex; align-items:center; gap:10px;";

  // Note: We use "default" as the value for the "All Tools" option
  var selHtml = '<label style="font-size:10px; font-weight:700; color:#33691e; text-transform:uppercase;">EDITING PROFILE:</label>' +
    '<select id="cfgModeSel" class="audience-select" style="flex:1; margin-bottom:0; height:28px; padding:2px 8px; border:1px solid #a5d6a7; color:#1b5e20; font-weight:600;">' +
    '<option value="default">Default (All Tools)</option>' +
    '<option value="draft">Drafting Mode</option>' +
    '<option value="review">Review Mode</option>' +
    '</select>';

  selectorRow.innerHTML = selHtml;
  card.appendChild(selectorRow);

  // 5. Content Area
  var content = document.createElement("div");
  content.id = "configListContent";
  content.style.padding = "8px 10px 10px 10px";
  card.appendChild(content);

  // --- HELPER: Create Mini Buttons ---
  function createMiniBtn(text, onClick, type) {
    var b = document.createElement("button");
    b.className = "btn-platinum-config";
    b.innerHTML = text;
    b.style.padding = "2px 8px";
    b.style.minWidth = "24px";
    if (type === "delete") { b.classList.add("delete"); b.style.color = "#c62828"; }
    if (type === "add") { b.style.color = "#2e7d32"; b.style.fontWeight = "900"; }
    b.onclick = onClick;
    return b;
  }

  // --- RENDER LOGIC ---
  function renderList() {
    var playlist = localMaps[activeEditMode] || [];
    content.innerHTML = "";

    var page1 = playlist.slice(0, 4);
    var page2 = playlist.slice(4, 8);

    // Determine hidden items
    var allKeys = Object.keys(toolMeta);
    var hidden = allKeys.filter(function (k) { return playlist.indexOf(k) === -1; });

    // Inner Function: Render Section
    function renderSection(title, items, pageNum) {
      var secHeader = document.createElement("div");
      secHeader.className = "config-header";
      var countInfo = (pageNum > 0) ? items.length + "/4" : items.length + " Available";

      // --- COLLAPSIBLE LOGIC ---
      var isCollapsible = (pageNum === 0);
      var containerDiv = document.createElement("div");

      if (isCollapsible) {
        // Default state: Hidden
        containerDiv.style.display = "none";

        secHeader.style.cursor = "pointer";
        secHeader.style.display = "flex";
        secHeader.style.justifyContent = "space-between";
        secHeader.style.alignItems = "center";
        secHeader.setAttribute("title", "Click to expand/collapse");

        secHeader.innerHTML = '<div style="display:flex; align-items:center; gap:6px;">' +
          '<span id="arrow_' + pageNum + '" style="font-size:10px; color:#546e7a;">▶</span> ' +
          '<span>' + title + '</span>' +
          '</div>' +
          '<span style="opacity:0.5; font-size:9px;">' + countInfo + '</span>';

        secHeader.onclick = function () {
          var isHidden = containerDiv.style.display === "none";
          containerDiv.style.display = isHidden ? "block" : "none";
          var arrow = this.querySelector("#arrow_" + pageNum);
          if (arrow) arrow.innerText = isHidden ? "▼" : "▶";
        };
      } else {
        // Standard Header
        secHeader.innerHTML = '<span>' + title + '</span> <span style="opacity:0.5; font-size:9px;">' + countInfo + '</span>';
      }

      content.appendChild(secHeader);

      if (items.length === 0) {
        containerDiv.innerHTML = '<div style="font-size:10px; color:#ccc; font-style:italic; padding:4px 12px; margin-bottom:8px;">(Empty slot)</div>';
      }

      items.forEach(function (key, index) {
        var tool = toolMeta[key];
        if (!tool) return;

        var row = document.createElement("div");
        row.className = "config-row" + (pageNum === 0 ? " hidden-item" : "");

        // Left: Icon + Label
        var left = document.createElement("div");
        left.style.display = "flex"; left.style.alignItems = "center"; left.style.gap = "8px";
        left.innerHTML = '<span style="font-size:16px;">' + tool.icon + '</span> <span style="font-size:11px; font-weight:600; color:#37474f;">' + tool.label + '</span>';

        // Right: Controls
        var right = document.createElement("div");
        right.style.display = "flex"; right.style.gap = "4px";

        if (pageNum > 0) {
          // Swap Up
          var globalIdx = (pageNum === 1) ? index : index + 4;

          if (globalIdx > 0) {
            var btnUp = createMiniBtn("▲", function () {
              var temp = playlist[globalIdx];
              playlist[globalIdx] = playlist[globalIdx - 1];
              playlist[globalIdx - 1] = temp;
              renderList();
            });
            right.appendChild(btnUp);
          }
          // Swap Down
          if (globalIdx < playlist.length - 1) {
            var btnDown = createMiniBtn("▼", function () {
              var temp = playlist[globalIdx];
              playlist[globalIdx] = playlist[globalIdx + 1];
              playlist[globalIdx + 1] = temp;
              renderList();
            });
            right.appendChild(btnDown);
          }
          // Remove
          var btnHide = createMiniBtn("✕", function () {
            playlist.splice(globalIdx, 1);
            renderList();
          }, "delete");
          right.appendChild(btnHide);

        } else {
          // Restore
          var btnAdd = createMiniBtn("✚", function () {
            if (playlist.length >= 8) {
              showToast("⚠️ Toolbar is full (8/8). Remove one first.");
            } else {
              playlist.push(key);
              renderList();
            }
          }, "add");
          right.appendChild(btnAdd);
        }

        row.appendChild(left);
        row.appendChild(right);
        containerDiv.appendChild(row);
      });

      content.appendChild(containerDiv);
    }

    renderSection("Page 1 (Primary)", page1, 1);
    renderSection("Page 2 (Secondary)", page2, 2);
    if (hidden.length > 0) renderSection("Unused Tools", hidden, 0);
  }

  // --- FOOTER (Actions) ---
  var footer = document.createElement("div");
  footer.style.cssText = "padding:10px 16px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px; flex-direction:column;";

  var btnRow = document.createElement("div");
  btnRow.style.cssText = "display:flex; gap:10px;";

  var btnReset = document.createElement("button");
  btnReset.className = "btn-platinum-config";
  btnReset.innerHTML = "↺ Reset This Mode";
  btnReset.style.flex = "1";

  // Custom Reset Dialog
  btnReset.onclick = function () {
    var overlay = document.createElement("div");
    overlay.style.cssText = "position:absolute; top:0; left:0; width:100%; height:100%; background:rgba(255,255,255,0.95); z-index:100; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; padding:20px; box-sizing:border-box;";

    overlay.innerHTML =
      '<div style="font-size:24px; margin-bottom:10px;">⚠️</div>' +
      '<div style="font-weight:800; font-size:14px; color:#c62828; margin-bottom:5px;">RESET ' + activeEditMode.toUpperCase() + '?</div>' +
      '<div style="font-size:11px; color:#546e7a; margin-bottom:15px;">This will return this specific profile to its default layout.</div>' +
      '<div style="display:flex; gap:10px; width:100%; justify-content:center;">' +
      '<button id="btnConfReset" class="btn-platinum-config delete" style="width:auto; padding:8px 16px; color:#c62828;">YES, RESET</button>' +
      '<button id="btnCancReset" class="btn-platinum-config" style="width:auto; padding:8px 16px;">CANCEL</button>' +
      '</div>';

    card.appendChild(overlay);

    document.getElementById("btnCancReset").onclick = function () { overlay.remove(); };

    document.getElementById("btnConfReset").onclick = function () {
      // Use Global Constant
      var defArr = DEFAULT_WORKFLOW_MAPS[activeEditMode];
      localMaps[activeEditMode] = defArr ? defArr.slice() : [];

      renderList();
      overlay.remove();
      showToast("↺ Reset applied (Click Save to commit)");
    };
  };

  var btnSave = document.createElement("button");
  btnSave.className = "btn-platinum-config primary";
  btnSave.style.flex = "2";
  btnSave.innerText = "SAVE CHANGES";
  btnSave.onclick = function () {
    userSettings.workflowMaps = JSON.parse(JSON.stringify(localMaps));

    if (activeEditMode === currentWorkflowMode || (activeEditMode === 'default' && currentWorkflowMode !== 'draft' && currentWorkflowMode !== 'review')) {
      // Update live toolOrder if we just edited the currently active profile
      userSettings.toolOrder = userSettings.workflowMaps[activeEditMode];

      var newPages = {};
      userSettings.toolOrder.forEach(function (key, index) {
        newPages[key] = (index < 4) ? 1 : 2;
        userSettings.visibleButtons[key] = true;
      });

      var allBtns = Object.keys(userSettings.visibleButtons);
      allBtns.forEach(function (k) {
        if (userSettings.toolOrder.indexOf(k) === -1) userSettings.visibleButtons[k] = false;
      });

      userSettings.buttonPages = newPages;
      renderToolbar();
    }

    saveCurrentSettings();
    showToast("✅ Configuration Saved");
    showHome();
  };

  btnRow.appendChild(btnReset);
  btnRow.appendChild(btnSave);
  footer.appendChild(btnRow);
  card.appendChild(footer);
  out.appendChild(card);

  // --- BINDINGS ---
  var sel = document.getElementById("cfgModeSel");

  // Set initial value (This is the fix!)
  sel.value = activeEditMode;

  // Handle Switch
  sel.onchange = function () {
    activeEditMode = this.value;
    renderList();
    content.style.opacity = "0.5";
    setTimeout(function () { content.style.opacity = "1"; }, 200);
  };

  // Initial Render Call
  renderList();

  document.getElementById("btnCloseConfig").onclick = showHome;
}
// ==========================================
// END: showInterfaceConfig
// ==========================================
// ==========================================
// START: setActiveButton
// ==========================================
function setActiveButton(activeId) {
  var ids = [
    "btnError",
    "btnAssist",
    "btnSummary",
    "btnPlaybook",
    "btnQA",
    "btnTemplate",
    "btnCompare",
    "btnEmailTool",
  ];

  ids.forEach(function (id) {
    var btn = document.getElementById(id);
    if (!btn) return;

    // Cleanup visuals
    if (btn.dataset.originalText) btn.innerText = btn.dataset.originalText;

    // --- CRITICAL FIX: REMOVE THIS LINE ---
    // btn.style.transform = "none";  <-- DELETE OR COMMENT OUT THIS LINE
    // Let CSS handle the transform so hover works!
    btn.style.transform = ""; // Clear inline style to let CSS take over

    btn.style.filter = "none";
    btn.style.opacity = "1";

    if (id === activeId) {
      // --- ACTIVE STATE ---
      btn.classList.add("active-tool");
      // Force styles for the active state
      btn.style.backgroundColor = "#e3f2fd";
      btn.style.color = "#1565c0";
      btn.style.fontWeight = "700";
      btn.style.border = "1px solid #1565c0";
      btn.style.boxShadow = "inset 0 -4px 0 0 #1565c0";
    } else {
      // --- INACTIVE STATE ---
      btn.classList.remove("active-tool");
      // Clear all inline overrides so CSS hover works
      btn.style.backgroundColor = "";
      btn.style.color = "";
      btn.style.fontWeight = "";
      btn.style.border = "";
      btn.style.boxShadow = "";
    }
  });
}
// ==========================================
// END: setActiveButton
// ==========================================
// START: getRedlineHtml
// ==========================================
function getRedlineHtml(oldStr, newStr) {
  if (!oldStr || !newStr) return newStr;

  // 1. Find Common Prefix
  var start = 0;
  while (
    start < oldStr.length &&
    start < newStr.length &&
    oldStr[start] === newStr[start]
  ) {
    start++;
  }
  // Backtrack to the last space to avoid cutting words in half (e.g. "app|le" vs "app|lication")
  while (start > 0 && oldStr[start] !== " " && newStr[start] !== " ") {
    start--;
  }

  // 2. Find Common Suffix
  var endOld = oldStr.length;
  var endNew = newStr.length;
  while (
    endOld > start &&
    endNew > start &&
    oldStr[endOld - 1] === newStr[endNew - 1]
  ) {
    endOld--;
    endNew--;
  }
  // Forward track to next space to avoid cutting words
  while (
    endOld < oldStr.length &&
    oldStr[endOld] !== " " &&
    newStr[endNew] !== " "
  ) {
    endOld++;
    endNew++;
  }

  // 3. Extract Parts
  var prefix = oldStr.substring(0, start);
  var suffix = oldStr.substring(endOld);

  var deleted = oldStr.substring(start, endOld);
  var inserted = newStr.substring(start, endNew);

  // 4. Build HTML
  var html = prefix;

  // RED LINE (Deleted Text)
  if (deleted.trim().length > 0) {
    html += `<span style="color:#c62828; text-decoration:line-through; margin-right:4px; opacity:0.7;">${deleted}</span>`;
  }

  // GREEN TEXT (New Fix)
  if (inserted.trim().length > 0) {
    html += `<span style="color:#2e7d32; font-weight:800; background:#e8f5e9; padding:0 2px; border-radius:2px;">${inserted}</span>`;
  }

  html += suffix;

  return html;
}
// ==========================================
// END: getRedlineHtml
// ==========================================
// ==========================================
// START: showPlaybookSourceManager
// ==========================================
function showPlaybookSourceManager() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. CONTAINER (Premium Aesthetic)
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.padding = "0";
  card.style.borderRadius = "0px";
  card.style.border = "1px solid #d1d5db";
  card.style.borderLeft = "5px solid #1565c0";
  card.style.background = "#fff";

  // 2. HEADER
  var header = document.createElement("div");
  header.style.cssText = `
                background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
                padding: 12px 15px;
                border-bottom: 3px solid #c9a227;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                margin-bottom: 0;
                position: relative;
                overflow: hidden;
                `;
  header.innerHTML = `
                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
                <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
                  SYSTEM CONFIGURATION
                </div>
                <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
                  <span>📚</span> Playbook Configuration
                </div>
                `;
  card.appendChild(header);

  // 3. BODY
  var contentDiv = document.createElement("div");
  contentDiv.style.padding = "20px";
  contentDiv.innerHTML = `
                <div style="margin-bottom:15px;">
                  <label class="tool-label">Source Strategy</label>
                  <select id="pbSourceSel" class="audience-select" style="font-weight:600;">
                    <option value="hardcoded">Standard (Built-in Rules)</option>
                    <option value="local">Local Database (JSON File)</option>
                    <option value="online">Online Resource (URL)</option>
                  </select>
                </div>

                <div style="font-size:12px; color:#37474f; font-weight:700; margin-bottom:10px; text-transform:uppercase; border-bottom:1px solid #eee; padding-bottom:5px;">
                  Workflows & Policies
                </div>

                <label class="checkbox-wrapper" style="margin-bottom:10px; background:#e3f2fd; padding:10px; border-radius:4px; border:1px solid #bbdefb;">
                  <input type="checkbox" id="chkEnableComments" style="transform:scale(1.2); margin-right:8px;">
                    <span><strong>Enable Justification Workflow</strong><br><span style="font-size:10px; font-weight:400; color:#555;">Review AI reasoning before inserting clauses.</span></span>
                </label>

                <label class="checkbox-wrapper" style="margin-bottom:15px; background:#fff3e0; padding:10px; border-radius:4px; border:1px solid #ffe0b2;">
                  <input type="checkbox" id="chkHideMatches" style="transform:scale(1.2); margin-right:8px;">
                    <span><strong>Hide Exact Matches</strong><br><span style="font-size:10px; font-weight:400; color:#555;">Focus only on risks/deviations to save time.</span></span>
                </label>

                <div style="margin-bottom:15px;">
                  <label class="tool-label">Default Approver Email</label>
                  <input type="text" id="pbGlobalEmail" class="tool-input" placeholder="manager@colgate.com">
                </div>

                <div id="pbLocalArea" class="hidden" style="background:#f5f5f5; padding:15px; border:1px solid #ddd; margin-bottom:15px; border-radius:4px;">
                  <div class="tool-label">📂 Upload JSON File</div>
                  <input type="file" id="pbFile" accept=".json" style="font-size:11px; width:100%;">
                    <div id="pbStatus" style="font-size:11px; color:#2e7d32; margin-top:8px; font-weight:bold;"></div>
                </div>

                <div id="pbOnlineArea" class="hidden" style="margin-bottom:15px;">
                  <label class="tool-label">Resource URL</label>
                  <input type="text" id="pbUrl" class="tool-input" placeholder="https://api.yoursite.com/playbook.json">
                </div>
                `;
  card.appendChild(contentDiv);

  // 4. FOOTER BUTTONS
  var footerDiv = document.createElement("div");
  footerDiv.style.cssText = "padding:15px 20px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px;";

  var btnCancel = document.createElement("button");
  btnCancel.className = "btn-platinum-config";
  btnCancel.style.flex = "1";
  btnCancel.innerText = "CANCEL";
  btnCancel.id = "btnCancelPb";

  var btnSave = document.createElement("button");
  btnSave.className = "btn-platinum-config primary";
  btnSave.style.flex = "2";
  btnSave.innerText = "SAVE SETTINGS";
  btnSave.id = "btnSavePb";

  footerDiv.appendChild(btnCancel);
  footerDiv.appendChild(btnSave);
  card.appendChild(footerDiv);
  out.appendChild(card);

  // 5. RESTORE BINDINGS
  var sel = document.getElementById("pbSourceSel");
  sel.value = userSettings.playbookSource || "hardcoded";

  var chkComments = document.getElementById("chkEnableComments");
  chkComments.checked = userSettings.enableJustificationComments || false;

  var chkHide = document.getElementById("chkHideMatches");
  chkHide.checked = userSettings.hideMatches || false;

  if (document.getElementById("pbUrl"))
    document.getElementById("pbUrl").value = userSettings.playbookUrl || "";
  if (document.getElementById("pbGlobalEmail"))
    document.getElementById("pbGlobalEmail").value = userSettings.approvalEmail || "";

  // Toggle Logic
  function toggle() {
    var loc = document.getElementById("pbLocalArea");
    var onl = document.getElementById("pbOnlineArea");
    if (loc) loc.classList.add("hidden");
    if (onl) onl.classList.add("hidden");

    if (sel.value === "local" && loc) loc.classList.remove("hidden");
    if (sel.value === "online" && onl) onl.classList.remove("hidden");
  }
  sel.onchange = toggle;
  toggle();

  // File Handler
  // File Handler (Fixed: Writes directly to DB, bypassing LocalStorage)
  var fileInp = document.getElementById("pbFile");
  if (fileInp) {
    fileInp.onchange = function (e) {
      var file = e.target.files[0];
      if (!file) return;

      var reader = new FileReader();
      reader.onload = function (evt) {
        try {
          var raw = evt.target.result;
          var json = JSON.parse(raw);

          if (!Array.isArray(json)) {
            showToast("⚠️ Invalid Playbook: Expected valid JSON Array");
            document.getElementById("pbStatus").innerHTML = "<span style='color:orange'>⚠️ Invalid Format (Not an Array)</span>";
            return;
          }

          // FIX: Save directly to IndexedDB (ELI_FeatureDB)
          saveFeatureData(STORE_PLAYBOOKS, json).then(function () {
            var successMsg = "✅ Loaded " + json.length + " clauses into DB.";
            document.getElementById("pbStatus").innerText = successMsg;
            showToast(successMsg);

            // Clear the file input so it can be re-selected if needed
            fileInp.value = "";

            if (typeof refreshPlaybookData === "function") refreshPlaybookData();
          });

        } catch (err) {
          console.error("JSON Load Error:", err);
          showToast("⚠️ Bad JSON File. Please check format.");
          document.getElementById("pbStatus").innerHTML = "<span style='color:red'>⚠️ Failed to Parse JSON</span>";
        }
      };
      reader.readAsText(file);
    };
  }

  // Save Link
  document.getElementById("btnSavePb").onclick = function () {
    userSettings.playbookSource = sel.value;
    userSettings.playbookUrl = document.getElementById("pbUrl") ? document.getElementById("pbUrl").value : "";
    userSettings.approvalEmail = document.getElementById("pbGlobalEmail").value;
    userSettings.enableJustificationComments = chkComments.checked;
    userSettings.hideMatches = chkHide.checked;

    saveCurrentSettings();
    showToast("✅ Settings Saved");
    showHome();
  };

  document.getElementById("btnCancelPb").onclick = showHome;
}
// ==========================================
// END: showPlaybookSourceManager
// ==========================================
// ==========================================
// START: refreshPlaybookData
// ==========================================
function refreshPlaybookData() {
  var source = userSettings.playbookSource || "hardcoded";
  if (source === "local") {
    // 1. Load from DB
    loadFeatureData(STORE_PLAYBOOKS).then(function (data) {
      loadedPlaybook = data;

      // 2. Migration: Check LocalStorage one last time
      var localRaw = localStorage.getItem("colgate_local_playbook");
      if (localRaw) {
        try {
          var oldData = JSON.parse(localRaw);
          loadedPlaybook = oldData;
          saveFeatureData(STORE_PLAYBOOKS, oldData); // Move to DB
          localStorage.removeItem("colgate_local_playbook"); // Delete Old
          console.log("📦 Playbook Migrated.");
        } catch (e) { }
      }

      // 3. Trigger UI Refresh (Since this is async)
      var scrollBox = document.getElementById("playbook-scroll-area");
      var typeSel = document.getElementById("agreementType");
      if (scrollBox && typeSel) {
        renderPlaybookChecklist(typeSel.value, scrollBox);
      }
    });
  } else if (source === "online") {
    loadedPlaybook = STANDARD_PLAYBOOK; // Placeholder
  } else {
    loadedPlaybook = STANDARD_PLAYBOOK;
  }
}
// ==========================================
// END: refreshPlaybookData
// ==========================================
// START: showPlaybookInterface (Platinum, Compressed & Standard Home)
// ==========================================
function showPlaybookInterface(preSelectedType) {
  refreshPlaybookData();
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. STATE
  if (typeof window.currentPlaybookType === "undefined") window.currentPlaybookType = "NDA";
  if (preSelectedType && typeof preSelectedType === "string") window.currentPlaybookType = preSelectedType;

  // 2. PREMIUM HEADER (Squared & Compact)
  var header = document.createElement("div");
  header.style.cssText = `
                background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
                padding: 12px 15px;
                border-bottom: 3px solid #c9a227;
                margin-bottom: 12px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                position: relative;
                overflow: hidden;
                `;
  header.innerHTML = `
                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
                <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
                  LEGAL STRATEGY
                </div>
                <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
                  <span>📖</span> Playbook Manager
                </div>
                `;
  out.appendChild(header);

  // 3. APPROVAL STATUS BAR (Squared & Compressed)
  var pending = approvalQueue.filter(function (x) { return x.status === "Pending"; }).length;

  var bellBtn = document.createElement("button");
  bellBtn.className = "btn-action";
  // Custom Override for Square/Compact
  var baseStatusStyle = "width: 100%; padding: 6px; cursor: pointer; border-radius: 0px; font-size: 11px; display: flex; justify-content: center; align-items: center; gap: 8px; margin-bottom: 12px; border: 1px solid;";

  if (pending > 0) {
    bellBtn.innerHTML = `<span>🔔</span> <strong>${pending} PENDING APPROVAL(S)</strong>`;
    bellBtn.style.cssText = baseStatusStyle + "background: #ffebee; color: #c62828; border-color: #ffcdd2;";
  } else {
    bellBtn.innerHTML = `<span style="font-size:12px;">✅</span> <strong>ALL APPROVALS CLEARED</strong>`;
    bellBtn.style.cssText = baseStatusStyle + "background: #e8f5e9; color: #2e7d32; border-color: #c8e6c9;";
  }
  bellBtn.onclick = showApprovalDashboard;
  out.appendChild(bellBtn);

  // 4. COMPLIANCE SCAN CARD (Squared)
  var scanCard = document.createElement("div");
  scanCard.className = "tool-box slide-in";
  // Compact box styles: 0px radius, tighter padding (10px)
  scanCard.style.cssText = "border-left: 5px solid #d32f2f; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.05); padding: 10px; margin-bottom: 12px; border-radius: 0px; border: 1px solid #e0e0e0; border-left-width: 5px;";

  // Title
  var scanHead = document.createElement("div");
  scanHead.innerHTML = "🔍 COMPLIANCE SCAN";
  scanHead.style.cssText = "font-weight:800; color:#b71c1c; font-size:10px; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; padding-bottom:4px; border-bottom:1px solid #f5f5f5;";
  scanCard.appendChild(scanHead);

  // Filter Row (Compact)
  var filterRow = document.createElement("div");
  filterRow.style.cssText = "margin-bottom: 8px; display: flex; align-items: center; gap: 8px;";

  var typeLabel = document.createElement("div");
  typeLabel.innerText = "TYPE:";
  typeLabel.className = "tool-label";
  typeLabel.style.marginBottom = "0";
  typeLabel.style.fontSize = "10px";
  typeLabel.style.fontWeight = "700";
  typeLabel.style.color = "#546e7a";
  filterRow.appendChild(typeLabel);

  var typeSelect = document.createElement("select");
  typeSelect.id = "agreementType";
  typeSelect.className = "audience-select";
  // Compact Select: Height 22px
  typeSelect.style.cssText = "margin-bottom: 0; padding: 0 4px; height: 22px; font-size: 11px; border-radius: 0px; flex: 1; border: 1px solid #d1d5db; background-color: #f9fafb;";

  ["NDA", "Services", "JDA", "Other"].forEach(function (t) {
    var o = document.createElement("option"); o.value = t; o.text = t === "Other" ? "All Types" : t;
    typeSelect.appendChild(o);
  });
  typeSelect.value = window.currentPlaybookType;
  typeSelect.onchange = function () { window.currentPlaybookType = this.value; render(); };
  filterRow.appendChild(typeSelect);
  scanCard.appendChild(filterRow);

  // List Header (Grey Bar)
  var listHeader = document.createElement("div");
  listHeader.style.cssText = "display:flex; justify-content:space-between; align-items:center; margin-bottom:0px; background:#eceff1; padding:4px 8px; border:1px solid #e0e0e0; border-bottom:none;";
  listHeader.innerHTML = `<span style="font-size:9px; font-weight:700; color:#455a64;">ACTIVE CLAUSES</span>`;

  // Select All Checkbox
  var checkAllLabel = document.createElement("label");
  checkAllLabel.style.cssText = "font-size: 9px; color: #1565c0; cursor: pointer; font-weight: 700; display:flex; align-items:center;";
  var checkAllInput = document.createElement("input");
  checkAllInput.type = "checkbox";
  checkAllInput.style.marginRight = "4px";
  checkAllInput.onclick = function (e) {
    document.querySelectorAll(".playbook-check").forEach(b => b.checked = e.target.checked);
    updatePlaybookButtonState();
  };
  checkAllLabel.appendChild(checkAllInput);
  checkAllLabel.appendChild(document.createTextNode("SELECT ALL"));
  listHeader.appendChild(checkAllLabel);
  scanCard.appendChild(listHeader);

  // Scroll Box (Squared & Compact)
  var scrollBox = document.createElement("div");
  scrollBox.id = "playbook-scroll-area";
  scrollBox.style.cssText = "height: 150px; overflow-y: auto; border: 1px solid #e0e0e0; background: #fff; border-radius: 0px; margin-bottom: 8px;";
  scanCard.appendChild(scrollBox);

  function render() { renderPlaybookChecklist(typeSelect.value, scrollBox); }
  render();

  // Buttons (Updated to TRUE Platinum Aesthetic)
  var btnStyle = `
                width: 100%; border-radius: 0px; font-size: 11px; font-weight: 800; padding: 10px; text-transform: uppercase; cursor: pointer; transition: all 0.2s;
                background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
                border: 1px solid #b0bec5;
                border-top: 1px solid #cfd8dc;
                box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
                `;

  var btnAnalyze = document.createElement("button");
  btnAnalyze.id = "btnPlaybookAnalyze"; // Added ID
  btnAnalyze.style.cssText = btnStyle + "color: #1565c0; margin-bottom: 6px; opacity: 0.5; cursor: not-allowed;"; // Default Disabled
  btnAnalyze.innerHTML = "⚡ RUN ANALYSIS";
  btnAnalyze.disabled = true; // Default Disabled

  btnAnalyze.onmouseover = function () {
    if (!this.disabled) {
      this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)";
      this.style.borderColor = "#90caf9";
    }
  };
  btnAnalyze.onmouseout = function () {
    if (!this.disabled) {
      this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      this.style.borderColor = "#b0bec5";
    }
  };
  btnAnalyze.onclick = runPlaybookAnalysis;
  scanCard.appendChild(btnAnalyze);

  // Update State Immediately
  setTimeout(updatePlaybookButtonState, 100);

  var btnEmailNew = document.createElement("button");
  btnEmailNew.style.cssText = btnStyle + "color: #546e7a;";
  btnEmailNew.innerHTML = "✉️ DRAFT EMAIL";
  btnEmailNew.onmouseover = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)"; this.style.color = "#1565c0"; };
  btnEmailNew.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)"; this.style.color = "#546e7a"; };
  btnEmailNew.onclick = runAiEmailScanner;
  scanCard.appendChild(btnEmailNew);

  out.appendChild(scanCard);

  // 5. CLAUSE LIBRARY (Squared & Compact)
  var libCard = document.createElement("div");
  libCard.className = "tool-box slide-in";
  libCard.style.cssText = "border-left: 5px solid #1565c0; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.05); padding: 10px; margin-bottom: 10px; border-radius: 0px; border: 1px solid #e0e0e0; border-left-width: 5px;";

  var libHead = document.createElement("div");
  libHead.innerHTML = "📝 CLAUSE LIBRARY";
  libHead.style.cssText = "font-weight:800; color:#0d47a1; font-size:10px; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; padding-bottom:4px; border-bottom:1px solid #f5f5f5;";
  libCard.appendChild(libHead);

  var clauseSelect = document.createElement("select");
  clauseSelect.className = "audience-select";
  clauseSelect.style.cssText = "width: 100%; margin-bottom: 6px; height: 26px; font-size: 11px; border-radius: 0px; padding: 2px; border:1px solid #cfd8dc;";

  if (loadedPlaybook.length === 0) {
    var opt = document.createElement("option"); opt.innerText = "No clauses loaded"; clauseSelect.appendChild(opt);
  } else {
    loadedPlaybook.forEach(function (c, idx) {
      var opt = document.createElement("option"); opt.value = idx;
      var lock = c.approvalRequired ? "🔒 " : "";
      opt.innerText = lock + c.name;
      clauseSelect.appendChild(opt);
    });
  }
  libCard.appendChild(clauseSelect);

  var btnInsert = document.createElement("button");
  btnInsert.style.cssText = btnStyle + "color: #1565c0; margin-top: 0;";
  btnInsert.innerHTML = "INSERT SELECTED";
  btnInsert.onmouseover = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)"; };
  btnInsert.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)"; };
  btnInsert.onclick = function () {
    var idx = clauseSelect.value;
    var selectedClause = loadedPlaybook[idx];
    if (selectedClause) {
      checkApprovalAndExecute(selectedClause, selectedClause.standard, "Insert Standard");
    }
  };
  libCard.appendChild(btnInsert);
  out.appendChild(libCard);

  // 6. HOME BUTTON (Standard Style + Squared)
  var btnHome = document.createElement("button");
  btnHome.className = "btn-action";
  btnHome.innerHTML = "🏠 HOME";
  btnHome.style.cssText = `
                width: 100%; margin-top: 20px;
                background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
                border: 1px solid #b0bec5;
                color: #546e7a; font-weight: 800; border-radius: 0px;
                padding: 10px;
                box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
                `;
  btnHome.onmouseover = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #e3f2fd 100%)"; this.style.color = "#1565c0"; };
  btnHome.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)"; this.style.color = "#546e7a"; };
  btnHome.onclick = showHome;
  out.appendChild(btnHome);
}
// ==========================================
// END: showPlaybookInterface
// ==========================================
// ==========================================
// START: renderPlaybookChecklist (Compact)
// ==========================================
function renderPlaybookChecklist(filterType, container) {
  container.innerHTML = "";
  if (loadedPlaybook.length === 0) {
    container.innerHTML = "<div style='padding:15px; color:#999; text-align:center; font-size:11px;'>No clauses found.</div>";
    return;
  }

  loadedPlaybook.forEach(function (clause, idx) {
    var shouldShow = false;
    if (filterType === "Other") shouldShow = true;
    else if (
      !clause.types ||
      clause.types.includes("All") ||
      clause.types.includes(filterType)
    )
      shouldShow = true;

    if (shouldShow) {
      var row = document.createElement("div");
      row.className = "playbook-item slide-in";

      // COMPRESSED STYLING
      row.style.cssText = "display: flex; align-items: center; padding: 4px 8px; border-bottom: 1px solid #f0f0f0; min-height: 24px;";

      var chk = document.createElement("input");
      chk.type = "checkbox";
      chk.className = "playbook-check";
      chk.dataset.name = clause.name;
      chk.dataset.standard = clause.standard;
      chk.style.marginRight = "8px";
      chk.style.cursor = "pointer";
      chk.onchange = updatePlaybookButtonState;

      var nameLabel = document.createElement("label");
      nameLabel.style.cssText = "flex: 1; font-size: 11px; cursor: pointer; color: #333; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;";

      // Add Lock Icon if needed
      var prefix = clause.approvalRequired ? "🔒 " : "";
      nameLabel.innerText = prefix + clause.name;
      nameLabel.title = (clause.standard || "").substring(0, 100);

      nameLabel.onclick = function () {
        chk.click();
      };

      row.appendChild(chk);
      row.appendChild(nameLabel);
      container.appendChild(row);
    }
  });
}
// ==========================================
// END: renderPlaybookChecklist
// ==========================================
// ==========================================
// START: runPlaybookAnalysis
// ==========================================
function runPlaybookAnalysis() {
  var selected = [];
  var checkboxes = document.querySelectorAll(".playbook-check:checked");

  checkboxes.forEach(function (box) {
    selected.push({ name: box.dataset.name, standard: box.dataset.standard });
  });

  if (selected.length === 0) {
    showToast("⚠️ Please select at least one clause.");
    return;
  }

  var prompt = `You are a Senior Legal Editor for Colgate.
                TASK: Compare the attached document against a list of "Standard Playbook Clauses".
                INPUT STANDARDS: ${JSON.stringify(selected)}

                CRITICAL RULES:
                1. Check EVERY standard listed above.
                2. SPLIT ISSUES: If a single clause (e.g. "AI Policy") has multiple distinct failures, return separate objects for each failure.
                3. STATUS: Use "Mismatch" if text exists but is wrong. Use "Missing" ONLY if the concept is totally absent.
                4. DATA PERSISTENCE: If you split a clause into multiple issues, you MUST include the "doc_text" in EVERY object.

                5. BOUNDARY RULE (CRITICAL):
                - Extract ONLY the text relevant to the specific clause.
                - STOP immediately if you encounter a new Section Header or a clearly distinct topic.
                - Do NOT include the next clause even if it is in the same paragraph.

                6. CLEAN TEXT: Extract only the body text (ignore bold headers).

                Return a RAW JSON Array (no markdown):
                [{
                  "name": "Clause Name (must match input list)",
                "status": "Match" | "Mismatch" | "Missing",
                "doc_text": "The exact body text of THIS clause only.",
                "analysis": "Specific explanation of THIS issue only."
    }]

                Document Text to Check:`;

  validateApiAndGetText(false, "default")
    .then(function (text) {
      setLoading(true, "default");
      return callGemini(API_KEY, prompt + text);
    })
    .then(function (result) {
      try {
        var data = tryParseGeminiJSON(result);
        renderPlaybookResults(data); // Calls the defined renderer
      } catch (e) {
        showOutput("Error parsing results: " + e.message, true);
      }
    })
    .catch(handleError)
    .finally(function () {
      setLoading(false);
    });
}

// ==========================================
function renderPlaybookResults(items) {
  if (!Array.isArray(items)) {
    // If we got an error object or empty object, wrap it or show error
    if (items.summary || items.error) {
      showOutput("⚠️ AI Analysis Issue: " + (items.summary || "Could not parse results."), true);
      return;
    }
    items = []; // Default to empty array to prevent crash
  }
  var out = document.getElementById("output");
  out.innerHTML = "";
  var header = document.createElement("div");
  header.style.cssText = `
                background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
                padding: 12px 15px;
                border-bottom: 3px solid #c9a227;
                margin-bottom: 12px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                position: relative;
                overflow: hidden;
                `;
  header.innerHTML = `
                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
                <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
                  PLAYBOOK
                </div>
                <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
                  <span>⚡</span> Analysis Results
                </div>
                `;
  out.appendChild(header);

  // Force Sync Queue
  try {
    var storedQ = localStorage.getItem("colgate_approval_queue");
    approvalQueue = storedQ ? JSON.parse(storedQ) : [];
  } catch (e) {
    approvalQueue = [];
  }

  var allClauses = loadedPlaybook;

  items.forEach(function (item, index) {
    // --- 1. SETTING CHECK: Skip Matches if requested ---
    if (userSettings.hideMatches && item.status === "Match") return;

    var sourceClause = allClauses.find(function (c) {
      return c.name === item.name || item.name.includes(c.name);
    });
    var isLocked = sourceClause && sourceClause.approvalRequired;

    // --- APPROVAL STATUS CHECK ---
    var approvedRecord = approvalQueue.find(function (r) {
      return (
        sourceClause &&
        r.clause === sourceClause.name &&
        r.status === "Approved"
      );
    });

    // --- DESIGN THEME ---
    var cardBg = "#ffffff";
    var accentColor = "#757575";
    var icon = "❓";
    var statusLabel = item.status.toUpperCase();
    var labelBg = "#eee";
    var labelColor = "#333";

    if (item.status === "Match") {
      accentColor = "#2e7d32"; // Green
      icon = "✅";
      labelBg = "#2e7d32";
      labelColor = "#ffffff";
    } else if (item.status === "Mismatch") {
      accentColor = "#f57f17"; // Orange
      icon = "⚠️";
      labelBg = "#f57f17";
      labelColor = "#ffffff";
    } else if (item.status === "Missing") {
      accentColor = "#c62828"; // Red
      icon = "❌";
      labelBg = "#c62828";
      labelColor = "#ffffff";
    }

    // --- LOCKED HEADER ---
    var lockHeader = "";
    var borderStyle = "solid";
    var approvalBadge = "";

    if (approvedRecord) {
      var dateStr = approvedRecord.timestamp.split(",")[0];
      approvalBadge = `<div style="margin-top:4px; font-size:10px; color:#2e7d32; font-weight:700; display:flex; align-items:center; gap:4px;">🔓 Approval Acquired on ${dateStr}</div>`;
    } else if (isLocked) {
      borderStyle = "dashed";
      lockHeader = `
                <div style="background:${accentColor}; color:white; padding:4px 12px; font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; display:flex; align-items:center; justify-content:center; gap:6px;">
                  <span>🔒 RESTRICTED CLAUSE - APPROVAL REQUIRED</span>
                </div>`;
      cardBg = "#ffffff";
    }

    // --- 2. CONDITIONAL AI INSIGHT ---
    // Only generate this HTML if status is NOT "Match"
    var insightHtml = "";
    if (item.status !== "Match") {
      insightHtml = `
            <div style="margin-top: 15px;">
                <div style="font-size:11px; font-weight:700; text-transform:uppercase; color:#5c6bc0; margin-bottom:6px; letter-spacing:0.5px; display:flex; align-items:center; gap:4px;">
                    <span>🤖</span> AI INSIGHT
                </div>
                <div style="font-family: 'Segoe UI', system-ui, sans-serif; font-size: 15px; font-weight: 500; color: #1a237e; line-height: 1.6; background: linear-gradient(to right, #e8eaf6, #f5f5fa); padding: 16px; border-left: 5px solid #3949ab; border-radius: 0 4px 4px 0; box-shadow: 0 2px 6px rgba(0,0,0,0.04);">
                    ${item.analysis}
                </div>
            </div>`;
    }

    // --- CONTENT HTML ---
    var contentHtml = `
                ${lockHeader}
                <div style="padding:16px;">
                  <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:15px;">
                    <div>
                      <div style="font-size:16px; font-weight:800; color:#263238; letter-spacing:-0.3px; line-height:1.2;">${item.name}</div>
                      ${approvalBadge}
                    </div>
                    <div>
                      <span style="display:inline-block; background:${labelBg}; color:${labelColor}; padding:5px 10px; border-radius:0px; font-size:11px; font-weight:800; text-transform:uppercase; box-shadow:0 2px 4px rgba(0,0,0,0.15); border: 1px solid #000;">
                        ${icon} ${statusLabel}
                      </span>
                    </div>
                  </div>

                  ${item.status !== "Missing"
        ? `<div style="margin-bottom:15px; padding-top:15px; border-top: 3px solid #37474f;">
                    <div style="font-size:10px; font-weight:700; text-transform:uppercase; color:#37474f; margin-bottom:6px; letter-spacing:0.5px;">Current Document Text</div>
                    <div style="background:#f5f5f5; border-left:3px solid #bdbdbd; padding:12px; border-radius:0px; font-family:'Georgia', serif; font-size:13px; color:#424242; line-height:1.5;">"${item.doc_text || "Found implicitly."}"</div>
                </div>`
        : `<div style="margin-bottom:15px; padding-top:15px; border-top: 3px solid #c62828;">
                    <div style="font-size:10px; font-weight:700; text-transform:uppercase; color:#c62828;">⚠️ Clause Missing from Document</div>
                </div>`
      }

                  ${insightHtml}

                  <div id="action_wrapper_${index}" style="margin-top:20px; padding-top:15px; border-top:1px dashed #cfd8dc;"></div>
                </div>
                `;

    // --- CARD CONTAINER ---
    var card = document.createElement("div");
    card.className = "response-card slide-in";
    card.id = "card_dev_" + index;
    card.style.cssText = `
                background: ${cardBg};
                border: 1px ${borderStyle} ${accentColor};
                border-left: 5px solid ${accentColor};
                margin-bottom: 20px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.08);
                overflow: hidden;
                padding: 0;
                border-radius: 0px;
                `;

    var actionWrapper = document.getElementById(`action_wrapper_${index}`);

    // ... (Keep the existing action wrapper logic, e.g. "var actionBtnStyle = ...")

    card.innerHTML = contentHtml;
    out.appendChild(card);

    var actionWrapper = document.getElementById(`action_wrapper_${index}`);


    var actionBtnStyle = "flex: 1; background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%) !important; color: #1565c0 !important; border: 1px solid #b0bec5 !important; border-top: 1px solid #cfd8dc !important; font-weight: 800 !important; font-size: 11px !important; padding: 10px 14px !important; border-radius: 0px !important; cursor: pointer !important; transition: all 0.2s ease; text-transform: uppercase !important; letter-spacing: 0.5px !important; box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08) !important; text-shadow: 0 1px 0 rgba(255,255,255,0.8) !important; display: flex; align-items: center; justify-content: center; gap: 6px;";
    // ==========================================
    // START: hoverOn
    // ==========================================
    function hoverOn(btn) {
      btn.style.background = "#e3f2fd";
      btn.style.borderColor = "#90caf9";
      btn.style.color = "#1565c0";
      btn.style.boxShadow = "0 2px 5px rgba(0,0,0,0.1)";
    }
    // ==========================================
    // END: hoverOn
    // ==========================================
    // ==========================================
    // START: hoverOff
    // ==========================================
    function hoverOff(btn) {
      btn.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      btn.style.borderColor = "#b0bec5";
      btn.style.color = "#1565c0";
      btn.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08)";
    }
    // ==========================================
    // END: hoverOff
    // ==========================================

    // 1. LOCATE BUTTON
    if (item.status !== "Missing") {
      var btnLoc = document.createElement("button");
      btnLoc.innerHTML = "<span>👁️</span> Locate in Doc";
      // btnLoc.className = "btn-action"; // Removed to avoid conflict
      btnLoc.style.cssText =
        actionBtnStyle + " width: 100%; margin-bottom: 10px;";
      btnLoc.onmouseover = function () {
        hoverOn(this);
      };
      btnLoc.onmouseout = function () {
        hoverOff(this);
      };
      btnLoc.onclick = function () {
        locateText(item.doc_text, btnLoc, true);
      };
      actionWrapper.appendChild(btnLoc);
    }

    // 2. SECONDARY ROW (Fix / Ignore)
    var secondaryRow = document.createElement("div");
    secondaryRow.style.cssText = "display:flex; gap:10px;";

    // UPDATE: Check if ANY fallback options exist (including the new Standard one)
    // Since we are adding Standard to the list below, we always have at least one option if sourceClause exists.
    var showFallbacks = item.status !== "Match" && sourceClause;

    if (showFallbacks) {
      var btnExpand = document.createElement("button");
      btnExpand.innerHTML = "<span>⚡</span> Fix / Fallbacks";
      // btnExpand.className = "btn-action";
      btnExpand.style.cssText = actionBtnStyle;
      // Differentiate the Fix button slightly (Solid border)
      btnExpand.style.border = "1px solid #1565c0";
      btnExpand.onmouseover = function () {
        hoverOn(this);
      };
      btnExpand.onmouseout = function () {
        this.style.background = "#e3f2fd";
        this.style.borderColor = "#1565c0";
        this.style.boxShadow = "0 1px 2px rgba(0,0,0,0.05)";
      };
      btnExpand.onclick = function () {
        var fbDiv = document.getElementById("fb_container_" + index);
        if (fbDiv.style.display === "none") {
          fbDiv.style.display = "block";
          this.innerHTML = "<span>▲</span> Hide Options";
        } else {
          fbDiv.style.display = "none";
          this.innerHTML = "<span>⚡</span> Fix / Fallbacks";
        }
      };
      secondaryRow.appendChild(btnExpand);
    }

    // 3. IGNORE BUTTON (Updated to Match Uniform Style)
    var btnIgnore = document.createElement("button");
    btnIgnore.innerHTML = "<span>✕</span> Ignore";
    // btnIgnore.className = "btn-action";

    // Apply the standard 'actionBtnStyle' defined earlier in the function
    btnIgnore.style.cssText = actionBtnStyle;
    // Use the same hover effects as the other buttons
    btnIgnore.onmouseover = function () {
      hoverOn(this);
    };
    btnIgnore.onmouseout = function () {
      hoverOff(this);
    };

    btnIgnore.onclick = function () {
      card.style.opacity = "0";
      setTimeout(function () {
        card.remove();
      }, 300);
      showToast("Deviation Ignored");
    };
    secondaryRow.appendChild(btnIgnore);
    actionWrapper.appendChild(secondaryRow);
    // --- 3. HIDDEN FALLBACK CONTAINER ---
    if (showFallbacks) {
      var fbContainer = document.createElement("div");
      fbContainer.id = "fb_container_" + index;
      fbContainer.style.display = "none";
      fbContainer.style.marginTop = "15px";
      fbContainer.style.paddingTop = "10px";
      fbContainer.style.borderTop = "1px solid rgba(0,0,0,0.05)";
      fbContainer.innerHTML = `<div style="color:${accentColor}; margin-bottom:10px; font-size:10px; font-weight:800; text-transform:uppercase;">AVAILABLE REPLACEMENTS</div>`;

      // -----------------------------------------------------------------
      // NEW LOGIC: Combine Standard + Existing Fallbacks
      // -----------------------------------------------------------------
      var allOptions = [];

      // 1. Always add Standard Clause first
      allOptions.push({
        label: "Revert to Standard",
        text: sourceClause.standard,
        desc: "Original Colgate Policy",
        isStandard: true,
      });

      // 2. Add existing fallbacks if they exist
      if (sourceClause.fallbacks && sourceClause.fallbacks.length > 0) {
        allOptions = allOptions.concat(sourceClause.fallbacks);
      }

      allOptions.forEach(function (fb, i) {
        var num = i + 1; // Calculate the number

        var btn = document.createElement("button");
        btn.className = "btn-action";

        // --- 1. PREMIUM STYLING LOGIC ---
        var isStd = fb.isStandard;
        var isRestricted = isLocked && !approvedRecord && !isStd;

        // COLORS: Gold for Standard, Red for Locked, Gray/Blue for Regular
        var accentColor = isStd
          ? "#fbc02d"
          : isRestricted
            ? "#d32f2f"
            : "#90a4ae";
        var bgColor = isStd ? "#fffde7" : "#ffffff";
        var borderColor = isStd ? "#fdd835" : "#e0e0e0";
        var titleColor = isStd
          ? "#f9a825"
          : isRestricted
            ? "#c62828"
            : "#546e7a";

        // --- 2. CARD CONTAINER STYLE ---
        btn.style.cssText = `
                width: 100%;
                text-align: left;
                margin-bottom: 12px;
                background: ${bgColor};
                border: 1px solid ${borderColor};
                border-left: 4px solid ${accentColor};
                padding: 12px 16px;
                border-radius: 4px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.03);
                transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
                cursor: pointer;
                position: relative;
                `;

        // --- 3. INTERACTIVE HOVER EFFECTS ---
        btn.onmouseover = function () {
          this.style.transform = "translateY(-2px)";
          this.style.boxShadow = "0 8px 16px rgba(0,0,0,0.08)";
          if (!isStd && !isRestricted) {
            this.style.borderColor = "#2196f3";
            this.style.borderLeftColor = "#2196f3";
            // Make the number badge blue on hover too
            var badge = this.querySelector(".num-badge");
            if (badge) {
              badge.style.borderColor = "#2196f3";
              badge.style.color = "#2196f3";
            }
          }
        };
        btn.onmouseout = function () {
          this.style.transform = "translateY(0)";
          this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.03)";
          if (!isStd && !isRestricted) {
            this.style.borderColor = "#e0e0e0";
            this.style.borderLeftColor = "#90a4ae";
            var badge = this.querySelector(".num-badge");
            if (badge) {
              badge.style.borderColor = "#cfd8dc";
              badge.style.color = "#90a4ae";
            }
          }
        };

        // --- 4. CONTENT STRUCTURE ---
        var icon = isRestricted ? "🔒" : isStd ? "⭐" : "⚡";
        // Only show "Standard Policy" badge if it's the standard option
        var badgeLabel = isStd
          ? `<span style="background:#fbc02d; color:white; padding:2px 6px; border-radius:3px; font-size:9px; font-weight:800; margin-left:6px; text-transform:uppercase; letter-spacing:0.5px;">Standard</span>`
          : "";
        var descHtml = fb.desc
          ? `<div style="font-size:11px; color:#90a4ae; margin-top:2px; font-weight:400;">${fb.desc}</div>`
          : "";

        btn.innerHTML = `
                <div style="display:flex; align-items:flex-start; margin-bottom:10px;">

                  <div class="num-badge" style="
                            flex-shrink: 0;
                            width: 20px; height: 20px;
                            background: white;
                            color: ${accentColor};
                            border: 1px solid ${isStd ? accentColor : "#cfd8dc"};
                  border-radius: 50%;
                  font-size: 10px; font-weight: 700;
                  display: flex; align-items: center; justify-content: center;
                  margin-right: 10px;
                  margin-top: -1px;
                            transition: all 0.2s;">
                  ${num}
                </div>

                <div style="flex-grow:1;">
                  <div style="font-size:12px; font-weight:700; color:${titleColor}; text-transform:uppercase; letter-spacing:0.5px; line-height:1.2;">
                    ${fb.label} ${badgeLabel} ${icon}
                  </div>
                  ${descHtml}
                </div>
              </div>

                <div style="
                        padding-top: 8px;
                        border-top: 1px dashed ${isStd ? "#fff59d" : "#eceff1"};
                font-family: 'Georgia', serif;
                font-size: 13px;
                color: #37474f;
                        line-height: 1.5;">
                "${fb.text}"
            </div>
            `;

        // --- CLICK HANDLER ---
        btn.onclick = async function () {
          // 1. APPROVAL CHECK
          if (isLocked && !approvedRecord && !fb.isStandard) {
            var initialComment = fb.justification || "Standard term applied.";
            if (typeof checkApprovalAndExecute === "function") {
              checkApprovalAndExecute(
                sourceClause,
                fb.text,
                item.doc_text,
                initialComment,
                btn,
              );
            }
            return;
          }

          // --- DEFINITION: The Core Insertion Logic ---
          // ==========================================
          // START: executeInsertion
          // ==========================================
          var executeInsertion = function () {
            var target = item.status === "Missing" ? null : item.doc_text;
            performInsertion(fb.text, target, null, btn);

            // SHOW EDITOR (Justification/Comment)
            var siblings = fbContainer.querySelectorAll("button");
            // FIX: Use setProperty to override the "!important" CSS rule
            siblings.forEach(function (s) {
              if (s !== btn) {
                s.style.setProperty("display", "none", "important");
              }
            });
            var editArea = document.createElement("div");
            editArea.className = "slide-in";
            editArea.style.cssText =
              "background:#f8f9fa; border:1px solid #d1d5db; padding:15px; border-radius:0px; margin-top:5px; border-left: 4px solid #1565c0;";

            editArea.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
              <div style="font-size:11px; font-weight:700; color:#1565c0;">📝 JUSTIFICATION (FOR COMMENT):</div>
              <div id="source_badge_${index}"></div>
            </div>
            <div style="font-size:10px; color:#666; margin-bottom:8px;">This note will be added as a comment to the counterparty.</div>
            <textarea id="justification_txt_${index}" style="width:100%; height:60px; font-family:sans-serif; font-size:12px; padding:8px; border:1px solid #ccc; border-radius:0px; box-sizing:border-box;">Generating...</textarea>

            <div style="display:flex; gap:10px; margin-top:10px;">
              <button id="btn_apply_${index}" class="btn-action"
                style="flex:2; background:#2e7d32; color:white; border:none; font-weight:bold; border-radius:0px;"
                onmouseover="this.style.setProperty('background', '#1b5e20', 'important'); this.style.setProperty('color', '#ffffff', 'important');"
                onmouseout="this.style.setProperty('background', '#2e7d32', 'important'); this.style.setProperty('color', '#ffffff', 'important');">Attach Comment</button>
              <button id="btn_cancel_${index}" class="btn-action" style="flex:1; border-radius:0px;">Cancel</button>
            </div>
            `;

            fbContainer.appendChild(editArea);

            var txtBox = document.getElementById(`justification_txt_${index}`);
            var badge = document.getElementById(`source_badge_${index}`);

            // GENERATE JUSTIFICATION
            if (fb.justification && fb.justification.trim().length > 0) {
              txtBox.value = fb.justification;
              badge.innerHTML =
                "<span style='color:#2e7d32; font-size:9px; background:#e8f5e9; padding:3px 6px; font-weight:700; letter-spacing:0.5px;'>📘 FROM PLAYBOOK</span>";
            } else if (fb.isStandard) {
              txtBox.value = "Restoring standard Colgate policy terms.";
              badge.innerHTML =
                "<span style='color:#2e7d32; font-size:9px; background:#e8f5e9; padding:3px 6px; font-weight:700;'>STANDARD</span>";
            } else {
              try {
                var aiPrompt = `Write a single professional sentence justifying the change to this contract clause: "${fb.text}". Perspective: You are Colgate legal counsel. Rules: 1. Start with "Colgate requires..." or "Colgate standard policy is..." 2. Do NOT use the word "Justification" or "Rationale" at the start. 3. Return plain text only.`;
                callGemini(API_KEY, aiPrompt)
                  .then(function (aiRes) {
                    var cleanRes = aiRes
                      .replace(/[\{\}]/g, "")
                      .replace(/Justification:/i, "")
                      .replace(/Rationale:/i, "")
                      .replace(/"/g, "")
                      .trim();
                    txtBox.value = cleanRes;
                    badge.innerHTML =
                      "<span style='color:#1565c0; font-size:9px; background:#e3f2fd; padding:3px 6px; font-weight:700; letter-spacing:0.5px;'>✨ AI GENERATED</span>";
                  })
                  .catch(function (e) {
                    txtBox.value =
                      "Colgate requires standard terms for compliance.";
                    badge.innerHTML =
                      "<span style='color:#757575; font-size:9px; background:#eee; padding:3px 6px;'>⚠️ FALLBACK</span>";
                  });
              } catch (e) { }
            }

            document.getElementById(`btn_cancel_${index}`).onclick =
              function () {
                // Restore the other fallback buttons by removing the inline style
                siblings.forEach(function (s) {
                  if (s !== btn) {
                    s.style.display = ""; // This removes the inline "display: none !important"
                  }
                });
                editArea.remove();
              };

            document.getElementById(`btn_apply_${index}`).onclick =
              function () {
                var finalComment = txtBox.value;
                Word.run(function (context) {
                  var selection = context.document.getSelection();
                  selection.insertComment(finalComment);
                  return context.sync();
                }).then(function () {
                  editArea.innerHTML =
                    "<div style='color:#2e7d32; font-weight:bold; text-align:center; padding:10px;'>✅ Comment Attached</div>";
                });
              };
          };
          // ==========================================
          // END: executeInsertion
          // ==========================================

          // --- 2. CONDITIONAL WORKFLOW ---
          if (item.status === "Missing") {
            // A. MISSING: Show Modal first
            var overlay = document.createElement("div");
            overlay.style.cssText =
              "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";
            var dialog = document.createElement("div");
            dialog.className = "tool-box slide-in";
            dialog.style.cssText =
              "background:white; width:280px; padding:20px; border-left:5px solid #1565c0; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";
            dialog.innerHTML = `<div style="font-size:24px; margin-bottom:10px;">📍</div><div style="font-weight:700; color:#1565c0; margin-bottom:10px; font-size:14px;">Insert Missing Clause</div><div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">Please scroll to the location in the document you want to insert the clause.</div><button id="btnInsertHere" class="btn btn-blue" style="width:100%;">I am there - Insert Here</button>`;
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            document.getElementById("btnInsertHere").onclick = function () {
              overlay.remove();
              executeInsertion();
            };
          } else {
            // B. NORMAL: Run logic immediately
            executeInsertion();
          }
        }; // END CLICK HANDLER

        fbContainer.appendChild(btn);
      });
      actionWrapper.appendChild(fbContainer);
    }
  });

  var btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "⬅ Back to Checklist";
  btnBack.style.marginTop = "20px";
  btnBack.style.borderRadius = "0px";
  btnBack.onclick = function () {
    showPlaybookInterface();
  };
  out.appendChild(btnBack);
}
// ==========================================
// END: renderPlaybookResults
// ==========================================
function performInsertion(newText, targetText, commentText, btnElement, exactMatch) {
  // 1. UI Feedback immediately
  if (btnElement) {
    btnElement.dataset.originalText = btnElement.innerHTML;
    btnElement.innerHTML = "⏳ Applying...";
  }

  Word.run(async function (context) {
    let finalRange = null;
    let fontProps = null;
    let rangeToReplace = null;

    // A. FIND THE TARGET
    if (targetText && targetText.length > 5) {
      // Use findSmartMatch instead of basic search
      rangeToReplace = await findSmartMatch(context, targetText);
    }

    // B. FALLBACK: If not found, use Selection
    if (!rangeToReplace) {
      console.log(
        "Target text not found via SmartMatch. Falling back to cursor.",
      );
      rangeToReplace = context.document.getSelection();
    } else {
      // If we found it via search, expand to the full paragraph to ensure a clean swap
      // BUT SKIP THIS if exactMatch is requested (e.g. by Document Cleaner)
      if (!exactMatch) {
        var p = rangeToReplace.paragraphs.getFirst();
        rangeToReplace = p.getRange();
      }
    }

    // C. CAPTURE FORMATTING (So the new clause looks like the old one)
    rangeToReplace.font.load("name, size, color");
    await context.sync();

    fontProps = {
      name: rangeToReplace.font.name,
      size: rangeToReplace.font.size,
      color: rangeToReplace.font.color,
    };

    // D. PERFORM REPLACEMENT
    // We use the found range (or selection)
    if (newText.includes("<") && newText.includes(">")) {
      finalRange = rangeToReplace.insertHtml(newText, "Replace");
    } else {
      finalRange = rangeToReplace.insertText(newText, "Replace");
    }

    // E. RE-APPLY FORMATTING
    if (finalRange) {
      if (fontProps.name) finalRange.font.name = fontProps.name;
      if (fontProps.size) finalRange.font.size = fontProps.size;
      if (fontProps.color) finalRange.font.color = fontProps.color;

      finalRange.select();

      // F. ADD JUSTIFICATION COMMENT (If passed immediately)
      if (
        commentText &&
        typeof userSettings !== "undefined" &&
        userSettings.enableJustificationComments
      ) {
        finalRange.insertComment(commentText);
      }
    }

    await context.sync();
  })
    .then(function () {
      // UI Success
      if (btnElement) {
        btnElement.style.backgroundColor = "#e8f5e9";
        btnElement.style.borderColor = "#2e7d32";
        btnElement.innerHTML = `<div style="color:#2e7d32; font-weight:bold; text-align:center;">✅ APPLIED</div>`;
        btnElement.disabled = true;
      }
      showToast("✅ Clause Updated");
    })
    .catch(function (e) {
      console.error("Insertion Error:", e);
      // Reset button on error
      if (btnElement) {
        btnElement.innerHTML = "❌ Error";
        btnElement.style.backgroundColor = "#ffebee";
      }
      showOutput("Error inserting text: " + e.message, true);
    });
}
// ==========================================
// END: performInsertion
// ==========================================
function promptForApproval(name, text, defaultEmail) {
  // 1. Load Saved Requesters (ASYNC FROM DB)
  loadFeatureData(STORE_REQUESTERS).then(function (savedReqs) {
    if (!savedReqs) savedReqs = [];

    // 2. Create the Modal Overlay
    var overlay = document.createElement("div");
    overlay.id = "approvalOverlay";
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

    var dialog = document.createElement("div");
    dialog.className = "tool-box slide-in";
    dialog.style.cssText = "background:white; width:320px; padding:0; border-left:5px solid #d32f2f; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px; overflow:hidden;";

    // --- CONTENT ---
    dialog.innerHTML = `
        <div style="padding:15px; border-bottom:1px solid #eee; background:#fff5f5;">
            <div style="font-weight:800; color:#c62828; font-size:14px; margin-bottom:4px;">🔒 APPROVAL REQUIRED</div>
            <div style="font-size:11px; color:#b71c1c;">This clause requires authorization.</div>
        </div>
        <div style="padding:20px;">
            <label class="tool-label">Select Approver:</label>
            <select id="selApprover" class="audience-select" style="width:100%; margin-bottom:10px; border-color:#d32f2f;">
                <option value="">-- Select Saved Person --</option>
                ${savedReqs.map(r => `<option value="${r.email}">${r.name} (${r.email})</option>`).join("")}
                <option value="add_new" style="font-weight:bold; color:#1565c0;">➕ ADD NEW PERSON...</option>
            </select>
            <div id="manualApproverBox" class="hidden" style="background:#f9f9f9; padding:10px; border:1px dashed #ccc; margin-bottom:10px;">
                <input type="text" id="newApproverName" class="tool-input" placeholder="Name" style="margin-bottom:5px;">
                <input type="text" id="newApproverEmail" class="tool-input" placeholder="Email" style="margin-bottom:5px;">
                <button id="btnSaveNewApprover" class="btn-action" style="width:100%; font-size:10px;">Save & Select</button>
            </div>
            <label class="tool-label">Reason for Request:</label>
            <textarea id="approvalReason" class="tool-input" style="height:60px; font-family:inherit; margin-bottom:15px;" placeholder="E.g. Counterparty requires this change..."></textarea>
            <div style="display:flex; flex-direction:column; gap:8px;">
                <button id="btnSubmitAndEmail" style="width:100%; background: linear-gradient(to bottom, #d32f2f 0%, #b71c1c 100%); color: white; border: 1px solid #b71c1c; font-weight: 800; font-size: 11px; text-transform: uppercase; padding: 10px; border-radius: 0px; cursor: pointer; box-shadow: inset 0 1px 0 rgba(255,255,255,0.2), 0 1px 2px rgba(0,0,0,0.2);">
                    📝 Log & Draft Email
                </button>
                <button id="btnLogOnly" class="btn-action" style="width:100%; color:#d32f2f; border-color:#ef9a9a;">
                    💾 Log Request Only (No Email)
                </button>
                <button id="btnCancelApproval" class="btn-action" style="width:100%;">Cancel</button>
            </div>
        </div>
      `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    // --- LOGIC ---
    var sel = document.getElementById("selApprover");
    var manualBox = document.getElementById("manualApproverBox");
    var nameInp = document.getElementById("newApproverName");
    var emailInp = document.getElementById("newApproverEmail");

    if (defaultEmail) {
      for (var i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === defaultEmail) { sel.selectedIndex = i; break; }
      }
    }

    sel.onchange = function () {
      if (this.value === "add_new") manualBox.classList.remove("hidden");
      else manualBox.classList.add("hidden");
    };

    // --- DB SAVE LOGIC ---
    document.getElementById("btnSaveNewApprover").onclick = function () {
      var n = nameInp.value.trim();
      var e = emailInp.value.trim();
      if (!n || !e) { showToast("⚠️ Name and Email required"); return; }

      savedReqs.push({ id: String(Date.now()), name: n, email: e });

      // SAVE TO DB INSTEAD OF LOCALSTORAGE
      saveFeatureData(STORE_REQUESTERS, savedReqs).then(() => {
        var opts = `<option value="">-- Select Saved Person --</option>` +
          savedReqs.map(r => `<option value="${r.email}" ${r.email === e ? "selected" : ""}>${r.name} (${r.email})</option>`).join("") +
          `<option value="add_new" style="font-weight:bold; color:#1565c0;">➕ ADD NEW PERSON...</option>`;
        sel.innerHTML = opts;
        manualBox.classList.add("hidden");
        showToast("✅ Saved to DB");
      });
    };

    // (Rest of logic remains identical, just inside the Promise)
    function handleSubmit(isEmailMode) {
      var finalEmail = sel.value;
      var reason = document.getElementById("approvalReason").value;
      if (!finalEmail || finalEmail === "add_new") { sel.style.borderColor = "red"; showToast("⚠️ Select an Approver"); return; }
      if (!reason.trim()) { document.getElementById("approvalReason").style.borderColor = "red"; showToast("⚠️ Reason required"); return; }

      overlay.remove();
      approvalQueue.push({
        id: Date.now(), clause: name, text: text, reason: reason, approver: finalEmail, timestamp: new Date().toLocaleString(), status: "Pending",
      });
      saveApprovalQueue();
      if (isEmailMode) renderApprovalNextSteps(name, text, finalEmail, reason);
      else { showToast("✅ Request Logged"); showApprovalDashboard(); }
    }

    document.getElementById("btnSubmitAndEmail").onclick = function () { handleSubmit(true); };
    document.getElementById("btnLogOnly").onclick = function () { handleSubmit(false); };
    document.getElementById("btnCancelApproval").onclick = function () { overlay.remove(); };
  });
}
// ==========================================
// REPLACE: renderApprovalNextSteps (Editable Email)
// ==========================================
function renderApprovalNextSteps(name, text, email, reason) {
  // 1. Generate Default Content
  var defaultSubject = "APPROVAL REQUEST: " + name;
  var safeText = text.length > 800 ? text.substring(0, 800) + "... [TRUNCATED]" : text;
  var defaultBody = "Hi,\n\nI need approval to use the restricted clause: " +
    name +
    ".\n\nPROPOSED TEXT:\n" +
    safeText +
    "\n\nREASON:\n" +
    reason +
    "\n\nPlease reply with your approval.\n\nThanks.";

  var out = document.getElementById("output");
  out.innerHTML = "";

  var div = document.createElement("div");
  div.className = "tool-box slide-in";
  div.style.cssText = "border-left: 5px solid #d32f2f; padding: 0; background:white;";

  // --- HEADER ---
  var header = document.createElement("div");
  header.innerHTML = `
    <div style="background:#fff5f5; padding:15px; border-bottom:1px solid #ffcdd2;">
        <div style="font-size:14px; font-weight:800; color:#c62828;">📧 Draft Approval Email</div>
        <div style="font-size:11px; color:#b71c1c;">Review and edit before sending.</div>
    </div>
  `;
  div.appendChild(header);

  // --- EDITABLE FORM ---
  var bodyContent = document.createElement("div");
  bodyContent.style.padding = "15px";
  bodyContent.innerHTML = `
    <div style="margin-bottom:10px;">
        <label class="tool-label">To:</label>
        <input type="text" value="${email}" disabled class="tool-input" style="background:#f0f0f0; color:#555;">
    </div>

    <div style="margin-bottom:10px;">
        <label class="tool-label">Subject:</label>
        <input type="text" id="appEmailSubject" class="tool-input" value="${defaultSubject}">
    </div>

    <div style="margin-bottom:15px;">
        <label class="tool-label">Body:</label>
        <textarea id="appEmailBody" class="tool-input" style="height:150px; font-family:sans-serif; line-height:1.4;">${defaultBody}</textarea>
    </div>

    <button id="btnLaunchGmail" class="btn btn-blue" style="width:100%; background:linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%); padding:12px; font-size:13px; font-weight:800; border:1px solid #b0bec5; border-radius:0px; text-transform:uppercase; box-shadow:inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08); color:#1565c0; display:flex; align-items:center; justify-content:center; gap:8px;">
        <span>🚀</span> Open Gmail Draft
    </button>

    <button id="btnSkipEmail" class="btn-action" style="width:100%; margin-top:10px;">Skip (Dashboard)</button>
  `;
  div.appendChild(bodyContent);
  out.appendChild(div);

  // --- LOGIC ---
  document.getElementById("btnLaunchGmail").onclick = function () {
    // 1. Get current values from inputs
    var subj = document.getElementById("appEmailSubject").value;
    var body = document.getElementById("appEmailBody").value;

    // 2. Construct Gmail URL dynamically
    var gmailUrl = "https://mail.google.com/mail/?view=cm&fs=1" +
      "&to=" + encodeURIComponent(email) +
      "&su=" + encodeURIComponent(subj) +
      "&body=" + encodeURIComponent(body);

    // 3. Open
    if (Office.context.ui && Office.context.ui.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(gmailUrl);
    } else {
      window.open(gmailUrl, "_blank");
    }

    showToast("🚀 Opening Gmail...");
    // Auto-redirect to dashboard after launch
    setTimeout(showApprovalDashboard, 1000);
  };

  document.getElementById("btnSkipEmail").onclick = function () {
    showApprovalDashboard();
  };
}
// ==========================================
// END: renderApprovalNextSteps
// ==========================================
// ==========================================
// START: showApprovalDashboard
// ==========================================
function showApprovalDashboard() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  var header = document.createElement("div");
  header.style.cssText = `
              background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
              padding: 12px 15px;
              border-bottom: 3px solid #c9a227;
              margin-bottom: 12px;
              box-shadow: 0 2px 5px rgba(0,0,0,0.1);
              position: relative;
              overflow: hidden;
              `;
  header.innerHTML = `
              <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
              <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
                WORKFLOW
              </div>
              <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
                <span>🕒</span> Approval Dashboard
              </div>
              `;
  out.appendChild(header);

  // Ensure global variable exists
  if (typeof approvalQueue === "undefined") approvalQueue = [];

  // --- RENDER LIST ---
  if (approvalQueue.length === 0) {
    out.innerHTML += "<div class='placeholder'>No requests found.</div>";
  }

  approvalQueue.forEach(function (req) {
    var card = document.createElement("div");
    card.className =
      "approval-card approval-item-card " +
      (req.status === "Approved"
        ? "approved"
        : req.status === "Rejected"
          ? "rejected"
          : "");

    var statusColor =
      req.status === "Pending"
        ? "#f57f17"
        : req.status === "Approved"
          ? "#2e7d32"
          : "#c62828";

    // Use the saved timestamp
    var displayDate = req.timestamp;

    // --- NEW LOGIC: Append Date to Status if Approved ---
    var statusDisplay = req.status;
    if (req.status === "Approved") {
      // Split to get just the date part (e.g. "12/20/2025") from the timestamp
      var dateOnly = req.timestamp.split(",")[0];
      statusDisplay += ` <span style="font-weight:400; font-size:11px; opacity:0.8;">(${dateOnly})</span>`;
    }
    // ----------------------------------------------------

    // Premium Card Style
    card.style.cssText = `
              background: #fff;
              border: 1px solid #e0e0e0;
              border-left: 5px solid ${statusColor};
              border-radius: 0px;
              padding: 12px;
              margin-bottom: 15px;
              box-shadow: 0 2px 8px rgba(0,0,0,0.05);
              transition: transform 0.2s;
              `;
    card.onmouseover = function () { this.style.transform = "translateY(-2px)"; this.style.boxShadow = "0 5px 15px rgba(0,0,0,0.1)"; };
    card.onmouseout = function () { this.style.transform = "none"; this.style.boxShadow = "0 2px 8px rgba(0,0,0,0.05)"; };

    card.innerHTML = `
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; padding-bottom:8px; border-bottom:1px dashed #eee;">
                <span style="font-size:10px; font-weight:700; color:#90a4ae; text-transform:uppercase;">${displayDate}</span>
                <span style="font-size:10px; font-weight:700; color:${statusColor}; text-transform:uppercase; background:${statusColor}10; padding:2px 6px; border-radius:2px;">${statusDisplay}</span>
              </div>

              <div style="font-size:14px; font-weight:800; color:#37474f; margin-bottom:4px; letter-spacing:-0.2px;">${req.clause}</div>

              <div style="font-size:11px; color:#546e7a; display:flex; gap:6px; margin-bottom:8px;">
                <span>👤 To:</span> <strong>${req.approver}</strong>
              </div>

              <div style="background:#f5f5f5; padding:8px; border-radius:0px; border-left:2px solid #cfd8dc; font-size:11px; color:#455a64; font-style:italic;">
                "${req.reason}"
              </div>
              `;

    // ... (Rest of the Confirm button logic remains unchanged)
    // PENDING ACTION: "I Received Approval"
    if (req.status === "Pending") {
      var btnConfirm = document.createElement("button");
      btnConfirm.className = "btn-action";
      btnConfirm.style.marginTop = "10px";
      btnConfirm.style.width = "100%";
      btnConfirm.style.border = "1px solid #2e7d32";
      btnConfirm.style.color = "#2e7d32";
      btnConfirm.innerHTML = "✅ I Received Approval Email";

      btnConfirm.onclick = function () {
        // --- CUSTOM CONFIRMATION MODAL ---
        var overlay = document.createElement("div");
        overlay.style.cssText =
          "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

        var dialog = document.createElement("div");
        dialog.className = "tool-box slide-in";
        dialog.style.cssText =
          "background:white; width:280px; padding:20px; border-left:5px solid #2e7d32; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

        dialog.innerHTML = `
              <div style="font-size:24px; margin-bottom:10px;">🛡️</div>
              <div style="font-weight:700; color:#2e7d32; margin-bottom:10px; font-size:14px;">Confirm Approval</div>
              <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">
                Are you sure you received explicit email approval for <strong>${req.clause}</strong>?
              </div>
              <div style="display:flex; gap:10px;">
                <button id="btnYesApprove" style="flex:1; background: linear-gradient(to bottom, #43a047 0%, #2e7d32 100%); color: white; border: 1px solid #1b5e20; font-weight: 800; font-size: 11px; text-transform: uppercase; padding: 10px; border-radius: 0px; cursor: pointer; box-shadow: inset 0 1px 0 rgba(255,255,255,0.2), 0 1px 2px rgba(0,0,0,0.2); text-shadow: 0 -1px 0 rgba(0,0,0,0.2);">Yes, Unlock</button>
                <button id="btnNoApprove" class="btn-action" style="flex:1;">Cancel</button>
              </div>
              `;

        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        // Dialog Actions
        document.getElementById("btnNoApprove").onclick = function () {
          overlay.remove();
        };

        document.getElementById("btnYesApprove").onmouseover = function () { this.style.background = "#66bb6a"; this.style.boxShadow = "0 3px 6px rgba(0,0,0,0.2)"; };
        document.getElementById("btnYesApprove").onmouseout = function () { this.style.background = "linear-gradient(to bottom, #43a047 0%, #2e7d32 100%)"; this.style.boxShadow = "inset 0 1px 0 rgba(255,255,255,0.2), 0 1px 2px rgba(0,0,0,0.2)"; };

        document.getElementById("btnYesApprove").onclick = function () {
          overlay.remove();

          // Update State
          req.status = "Approved";
          req.timestamp = new Date().toLocaleString(); // Update time to NOW

          // Save & Refresh
          saveApprovalQueue();
          showApprovalDashboard();
          showToast("✅ Clause Unlocked");
        };
      };
      card.appendChild(btnConfirm);
    }
    out.appendChild(card);
  });

  // --- CLEAR HISTORY BUTTON (Fixed: HTML Modal + Deep Clean) ---
  if (approvalQueue.length > 0) {
    var btnClear = document.createElement("button");
    // Removed className="btn-action" to avoid conflict
    btnClear.style.cssText = `
              width: 100%;
              margin-top: 25px;
              background: linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%);
              color: #c62828;
              border: 1px solid #b0bec5;
              border-top: 1px solid #cfd8dc;
              font-weight: 800;
              font-size: 11px;
              padding: 10px 14px;
              border-radius: 0px;
              cursor: pointer;
              transition: all 0.2s ease;
              text-transform: uppercase;
              letter-spacing: 0.5px;
              box-shadow: inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08);
              text-shadow: 0 1px 0 rgba(255,255,255,0.8);
              display: flex; align-items: center; justify-content: center; gap: 6px;
              `;
    btnClear.onmouseover = function () {
      this.style.background = "#ffebee"; // Light Red Warning Hover
      this.style.borderColor = "#ef9a9a";
      this.style.color = "#b71c1c";
      this.style.boxShadow = "0 2px 5px rgba(0,0,0,0.1)";
    };
    btnClear.onmouseout = function () {
      this.style.background = "linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%)";
      this.style.borderColor = "#b0bec5";
      this.style.color = "#c62828";
      this.style.boxShadow = "inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08)";
    };
    btnClear.innerHTML = "🗑️ Clear All History";

    btnClear.onclick = function () {
      // --- CUSTOM DELETE MODAL (Replaces native confirm) ---
      var overlay = document.createElement("div");
      overlay.style.cssText =
        "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

      var dialog = document.createElement("div");
      dialog.className = "tool-box slide-in";
      dialog.style.cssText =
        "background:white; width:280px; padding:20px; border-left:5px solid #c62828; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

      dialog.innerHTML = `
              <div style="font-size:24px; margin-bottom:10px;">🗑️</div>
              <div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Clear History?</div>
              <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">
                This will permanently delete <strong>all</strong> pending and approved records from this document.
              </div>
              <div style="display:flex; gap:10px;">
                <button id="btnYesClear" style="flex:1; background: linear-gradient(to bottom, #d32f2f 0%, #b71c1c 100%); color: white; border: 1px solid #b71c1c; font-weight: 800; font-size: 11px; text-transform: uppercase; padding: 10px; border-radius: 0px; cursor: pointer; box-shadow: inset 0 1px 0 rgba(255,255,255,0.2), 0 1px 2px rgba(0,0,0,0.2); text-shadow: 0 -1px 0 rgba(0,0,0,0.2);">Yes, Delete All</button>
                <button id="btnNoClear" class="btn-action" style="flex:1;">Cancel</button>
              </div>
              `;

      overlay.appendChild(dialog);
      document.body.appendChild(overlay);

      document.getElementById("btnNoClear").onclick = function () {
        overlay.remove();
      };

      document.getElementById("btnYesClear").onclick = function () {
        overlay.remove();

        // 1. Wipe Memory
        approvalQueue = [];
        window.approvalQueue = []; // Explicit global wipe

        // 2. Wipe Document Settings (Critical Step)
        // We use our helper which sets the Doc Setting to []
        saveApprovalQueue();

        // 3. Clear LocalStorage (Redundant but safe)
        localStorage.removeItem("colgate_approval_queue");

        // 4. Visual Feedback
        showToast("History Cleared");
        showApprovalDashboard(); // Instant refresh
      };
    };
    out.appendChild(btnClear);
  }

  var btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "⬅ Back";
  btnBack.style.marginTop = "10px";
  btnBack.onclick = showPlaybookInterface;
  out.appendChild(btnBack);
}
// ==========================================
// END: showApprovalDashboard
// ==========================================
// ==========================================
// START: insertText
// ==========================================
function insertText(text) {
  showRedlinePrompt(text);
}

function loadApprovalQueue() {
  // Office.context.document.settings saves data INSIDE the .docx file
  var saved = Office.context.document.settings.get("colgate_approval_queue");
  if (saved) {
    approvalQueue = saved;
  } else {
    approvalQueue = [];
  }
}
// ==========================================
// END: loadApprovalQueue
// ==========================================

// Save approvals to the specific Word file
// Save approvals to the specific Word file (Persistent Storage)
// ==========================================
// START: saveApprovalQueue
// ==========================================
function saveApprovalQueue() {
  // 1. Save to LocalStorage (Fast UI sync)
  localStorage.setItem("colgate_approval_queue", JSON.stringify(approvalQueue));

  // 2. Save to Word Document (Persistent across sessions/users)
  Office.context.document.settings.set("colgate_approval_queue", approvalQueue);

  Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(
        "Failed to save approvals to doc: " + asyncResult.error.message,
      );
    } else {
      console.log("✅ Approvals synced to document settings.");
    }
  });
}
// ==========================================
// END: saveApprovalQueue
// ==========================================
// 7. ENFORCEMENT & EXECUTION
function checkApprovalAndExecute(
  clause,
  textToInsert,
  textToReplace,
  comment,
  btnElement,
) {
  if (!clause.approvalRequired) {
    performInsertion(textToInsert, textToReplace, comment, btnElement);
    return;
  }

  var approved = approvalQueue.find(function (r) {
    return r.clause === clause.name && r.status === "Approved";
  });

  if (approved) {
    performInsertion(textToInsert, textToReplace, comment, btnElement);
    showToast("✅ Approved Action Executed");
  } else {
    var manager =
      clause.approvalEmail && clause.approvalEmail.length > 3
        ? clause.approvalEmail
        : userSettings.approvalEmail;
    if (!manager || manager.length < 3) manager = "manager@colgate.com";
    promptForApproval(clause.name, textToInsert, manager);
  }
}
// --- ROBUST FOOTER CHECKER (Based on User's Working Snippet) ---
// ==========================================
// START: checkFooterConsistency
// ==========================================
function checkFooterConsistency() {
  return Word.run(async function (context) {
    var sections = context.document.sections;
    sections.load("items");
    await context.sync();

    var footerRanges = [];

    // 1. Iterate Sections & Queue Ranges (Exactly as before)
    sections.items.forEach(function (section, index) {
      // Primary
      try {
        var primary = section.getFooter
          ? section.getFooter("Primary")
          : section.footers.getItem("Primary");
        var r1 = primary.getRange();
        // CHANGE: Get HTML instead of loading text
        var h1 = r1.getHtml();
        footerRanges.push({ type: "Primary", section: index + 1, html: h1 });
      } catch (e) { }

      // FirstPage
      try {
        var first = section.getFooter
          ? section.getFooter("FirstPage")
          : section.footers.getItem("FirstPage");
        var r2 = first.getRange();
        // CHANGE: Get HTML instead of loading text
        var h2 = r2.getHtml();
        footerRanges.push({ type: "FirstPage", section: index + 1, html: h2 });
      } catch (e) { }

      // EvenPages
      try {
        var even = section.getFooter
          ? section.getFooter("EvenPages")
          : section.footers.getItem("EvenPages");
        var r3 = even.getRange();
        // CHANGE: Get HTML instead of loading text
        var h3 = r3.getHtml();
        footerRanges.push({ type: "EvenPages", section: index + 1, html: h3 });
      } catch (e) { }
    });

    // 2. Execute the queued loads
    await context.sync();

    // 3. Analyze Results (Logic identical to original, just accessing cleaned HTML)
    var errors = [];
    var masterRef = null;
    var refRegex = /(CP Ref:|A-)\s*([A-Za-z0-9\-]+)/i;

    // Helper to get clean text from the HTML result
    // ==========================================
    // START: getText
    // ==========================================
    function getText(item) {
      if (!item || !item.html) return "";
      return getCleanTextAndContext(item.html.value).text;
    }
    // ==========================================
    // END: getText
    // ==========================================

    // A. Find Master Reference (Priority: Section 1 FirstPage -> Section 1 Primary)
    var p1First = footerRanges.find(function (r) {
      return (
        r.section === 1 &&
        r.type === "FirstPage" &&
        getText(r).trim().length > 0
      );
    });
    var p1Prim = footerRanges.find(function (r) {
      return (
        r.section === 1 && r.type === "Primary" && getText(r).trim().length > 0
      );
    });

    var masterText = p1First ? getText(p1First) : p1Prim ? getText(p1Prim) : "";

    var masterMatch = masterText.match(refRegex);
    if (masterMatch) {
      masterRef = masterMatch[0].trim();
    } else {
      // If we can't find a reference on Page 1, we can't compare others.
      return [];
    }

    // B. Compare All Others
    footerRanges.forEach(function (item) {
      var text = getText(item); // Use helper to get clean text

      if (text && text.trim().length > 0) {
        var match = text.match(refRegex);
        if (match) {
          var found = match[0].trim();
          if (found !== masterRef) {
            errors.push({
              original: found,
              correction: masterRef,
              reason:
                "Footer Mismatch: Section " +
                item.section +
                " (" +
                item.type +
                ") has '" +
                found +
                "', but Page 1 is '" +
                masterRef +
                "'.",
            });
          }
        }
      }
    });

    return errors;
  }).catch(function (e) {
    return [
      {
        original: "Script Error",
        correction: "Footer Check",
        reason: "Logic Error: " + e.message,
      },
    ];
  });
}

function showResetConfirmationModal() {
  var overlay = document.createElement("div");
  overlay.style.cssText =
    "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var dialog = document.createElement("div");
  dialog.className = "tool-box slide-in";
  dialog.style.cssText =
    "background:white; width:320px; padding:25px; border-left:6px solid #1565c0; box-shadow:0 20px 50px rgba(0,0,0,0.3); border-radius:4px; text-align:center;";

  dialog.innerHTML = `
      <div style="font-size:32px; margin-bottom:15px;">♻️</div>
      <div style="font-weight:800; color:#1565c0; margin-bottom:10px; font-size:16px; text-transform:uppercase;">App Reset</div>
      
      <div style="font-size:12px; color:#546e7a; margin-bottom:20px; line-height:1.5;">
        Select a reset level.
      </div>

      <button id="btnStandardReset" class="btn-action" style="width:100%; margin-bottom:10px; border:1px solid #1565c0; color:#1565c0; font-weight:700;">
        🛠️ Reset UI Settings<br>
        <span style="font-weight:400; font-size:9px; opacity:0.8;">Fixes toolbars & keys. Keeps data safe.</span>
      </button>

      <button id="btnSecureWipe" class="btn-action" style="width:100%; margin-bottom:15px; border:1px solid #c62828; color:#c62828;">
        ☢️ Wipe Data (Admin)<br>
        <span style="font-weight:400; font-size:9px; opacity:0.8;">Select what to delete (Libraries, History, etc).</span>
      </button>
      <button id="btnCheckStorage" class="btn-action" style="width:100%; margin-bottom:15px; border:1px solid #455a64; color:#455a64;">
        📊 Check Storage Usage
      </button>

      <div style="border-top:1px solid #eee; margin-bottom:10px;"></div>
      <button id="btnCancelReset" class="btn-action" style="width:100%; color:#555;">Cancel</button>
  `;
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  document.getElementById("btnCheckStorage").onclick = checkStorageUsage;
  // --- 1. CANCEL ---
  document.getElementById("btnCancelReset").onclick = function () {
    overlay.remove();
  };

  // --- 2. STANDARD RESET (Safe UI Reset) ---
  document.getElementById("btnStandardReset").onclick = function () {
    var btn = this;
    btn.innerText = "Resetting...";

    // 1. Reset Workflows & Apps
    userSettings.workflowMaps = JSON.parse(JSON.stringify(DEFAULT_WORKFLOW_MAPS));
    userSettings.installedStoreApps = [];

    // 2. Reset Active View State
    userSettings.toolOrder = userSettings.workflowMaps["default"];
    userSettings.currentWorkflow = "default";

    // 3. Reset Button Visibility & Pages
    userSettings.visibleButtons = {
      error: true, legal: true, summary: true, playbook: true,
      chat: true, template: true, compare: true, email: true
    };
    userSettings.buttonPages = {
      error: 1, legal: 1, summary: 1, chat: 1,
      email: 2, template: 2, playbook: 2, compare: 2
    };

    // 4. Force Onboarding Restart
    userSettings.hasOnboarded = false;

    // 5. Save & Reload
    saveCurrentSettings();
    showToast("✅ Interface Reset (Reloading...)");
    setTimeout(function () { window.location.reload(); }, 500);
  };

  // --- 3. SECURE DATA WIPE (Granular Choices) ---
  document.getElementById("btnSecureWipe").onclick = function () {
    dialog.style.display = "none";

    var passDialog = document.createElement("div");
    passDialog.className = "tool-box slide-in";
    passDialog.style.cssText = "background:white; width:300px; padding:20px; border-left:5px solid #c62828; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px; text-align:left;";

    // Step 1: Password Input
    passDialog.innerHTML = `
        <div style="font-size:24px; margin-bottom:10px; text-align:center;">🔒</div>
        <div style="font-weight:700; color:#c62828; margin-bottom:10px; text-align:center;">Admin Access Required</div>
        <input type="password" id="txtAdminPass" class="tool-input" placeholder="Enter Password" style="width:100%; margin-bottom:10px;">
        <button id="btnVerifyPass" class="btn-action" style="width:100%; background:#c62828; color:white; font-weight:bold;">Verify</button>
        <button id="btnCancelPass" class="btn-action" style="width:100%; margin-top:5px;">Cancel</button>
    `;

    overlay.appendChild(passDialog);
    setTimeout(function () { document.getElementById("txtAdminPass").focus(); }, 100);

    document.getElementById("btnCancelPass").onclick = function () { overlay.remove(); };

    document.getElementById("btnVerifyPass").onclick = function () {
      var pwd = document.getElementById("txtAdminPass").value;
      if (pwd !== "moops") {
        showToast("⛔ Incorrect Password.");
        return;
      }

      // Step 2: Granular Selection
      passDialog.innerHTML = `
            <div style="font-weight:800; color:#c62828; font-size:14px; margin-bottom:10px; text-transform:uppercase; border-bottom:1px solid #ffcdd2; padding-bottom:5px;">
                ☢️ SELECT DATA TO WIPE
            </div>
            
            <label style="display:flex; align-items:center; padding:8px; border:1px solid #eee; margin-bottom:5px; cursor:pointer;">
                <input type="checkbox" id="wipeClauses" checked style="margin-right:10px;">
                <div>
                    <div style="font-weight:bold; font-size:12px; color:#333;">Clause Library</div>
                    <div style="font-size:10px; color:#666;">Deletes all saved standard/negotiated clauses.</div>
                </div>
            </label>

            <label style="display:flex; align-items:center; padding:8px; border:1px solid #eee; margin-bottom:5px; cursor:pointer;">
                <input type="checkbox" id="wipePlaybooks" checked style="margin-right:10px;">
                <div>
                    <div style="font-weight:bold; font-size:12px; color:#333;">Playbooks & Templates</div>
                    <div style="font-size:10px; color:#666;">Deletes Playbook rules and Document Templates.</div>
                </div>
            </label>

            <label style="display:flex; align-items:center; padding:8px; border:1px solid #eee; margin-bottom:5px; cursor:pointer;">
                <input type="checkbox" id="wipeDocs" style="margin-right:10px;">
                <div>
                    <div style="font-weight:bold; font-size:12px; color:#333;">Document Files</div>
                    <div style="font-size:10px; color:#666;">Deletes cached OARS files & Version History.</div>
                </div>
            </label>

            <label style="display:flex; align-items:center; padding:8px; border:1px solid #eee; margin-bottom:15px; cursor:pointer;">
                <input type="checkbox" id="wipeSettings" style="margin-right:10px;">
                <div>
                    <div style="font-weight:bold; font-size:12px; color:#333;">User Settings</div>
                    <div style="font-size:10px; color:#666;">Resets name, email templates, and preferences.</div>
                </div>
            </label>

            <button id="btnExecuteWipe" class="btn-action" style="width:100%; background:#d32f2f; color:white; font-weight:800; padding:12px; border:1px solid #b71c1c;">DELETE SELECTED</button>
            <button id="btnAbortWipe" class="btn-action" style="width:100%; margin-top:10px;">Cancel</button>
        `;

      document.getElementById("btnAbortWipe").onclick = function () { overlay.remove(); };

      document.getElementById("btnExecuteWipe").onclick = function () {
        var btn = this;
        btn.innerText = "Wiping...";
        btn.disabled = true;

        var doClauses = document.getElementById("wipeClauses").checked;
        var doPlaybooks = document.getElementById("wipePlaybooks").checked;
        var doDocs = document.getElementById("wipeDocs").checked;
        var doSettings = document.getElementById("wipeSettings").checked;

        var tasks = [];

        // 1. Wipe Clause Library
        if (doClauses) {
          tasks.push(new Promise(resolve => {
            indexedDB.deleteDatabase("ELI_ClauseDB");
            localStorage.removeItem("colgate_library");
            resolve();
          }));
        }

        // 2. Wipe Playbooks & Templates
        if (doPlaybooks) {
          tasks.push(new Promise(resolve => {
            indexedDB.deleteDatabase("ELI_FeatureDB");
            localStorage.removeItem("colgate_local_playbook");
            localStorage.removeItem("colgate_standard_templates");
            resolve();
          }));
        }

        // 3. Wipe Document Files & History
        if (doDocs) {
          tasks.push(new Promise(resolve => {
            if (PROMPTS && PROMPTS._DOC_DB_NAME) indexedDB.deleteDatabase(PROMPTS._DOC_DB_NAME);

            // Clear Word Settings
            if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
              var s = Office.context.document.settings;
              s.remove("colgate_approval_queue");
              s.remove("eli_version_history");
              s.remove("AgreementDataDetails");
              s.remove("eli_doc_id");
              s.saveAsync(() => resolve());
            } else {
              resolve();
            }
          }));
        }

        // 4. Wipe User Settings
        if (doSettings) {
          tasks.push(new Promise(resolve => {
            localStorage.removeItem("colgate_user_settings");
            localStorage.removeItem("eli_language");
            localStorage.removeItem("colgate_saved_requesters");
            resolve();
          }));
        }

        Promise.all(tasks).then(() => {
          showToast("✅ Selected Data Wiped");
          setTimeout(() => window.location.reload(), 1000);
        });
      };
    };
  };
}
// ==========================================
// END: showResetConfirmationModal
// ==========================================

function showPaymentScheduleModal(totalFeeValue, scheduleString, callback) {
  // 1. Parse Total (Robust)
  // 1. Parse Total (Robust)
  // Fix: Handle commas, spaces, currency symbols. 
  // We want to find the first valid number in the string.
  var cleanFee = String(totalFeeValue).replace(/,/g, ''); // Remove commas first
  var match = cleanFee.match(/(\d+(\.\d+)?)/); // Find first number sequence
  var totalFee = match ? parseFloat(match[0]) : 0;
  if (isNaN(totalFee)) {
    totalFee = 0;
  }
  var formattedTotal = totalFee.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // 2. Create Overlay
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; margin:0; padding:0; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px); gap: 15px;";

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "background:white; width:340px; padding:0; border-radius:0px; box-shadow:0 20px 50px rgba(0,0,0,0.3); border:1px solid #b0bec5; display:flex; flex-direction:column; max-height:90vh;";

  // 3. Header
  card.innerHTML = `
                              <div style="background:linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding:15px; border-bottom:3px solid #c9a227;">
                                <div style="font-size:10px; font-weight:800; color:#8fa6c7; text-transform:uppercase; letter-spacing:1px;">AUTO-FILL DETECTED</div>
                                <div style="font-family:'Georgia', serif; font-size:18px; color:white; display:flex; justify-content:space-between; align-items:center;">
                                  <span>💰 Payment Schedule</span>
                                  <div id="headerTotalDisplay" style="font-size:14px; font-weight:700; background:rgba(0,0,0,0.3); padding:4px 8px; border-radius:4px; border:1px solid rgba(255,255,255,0.2);">
                                    $${formattedTotal}
                                  </div>
                                </div>
                              </div>
                              <div style="padding:15px; overflow-y:auto; flex:1;">

                                <!-- Currency Selection -->
                                <div style="margin-bottom:15px;">
                                  <label class="tool-label">Currency</label>
                                  <select id="payCurrency" class="audience-select" style="border-radius:0px; margin-bottom:0;">
                                    <option value="USD" selected>USD ($)</option>
                                    <option value="EUR">EUR (€)</option>
                                    <option value="GBP">GBP (£)</option>
                                  </select>
                                </div>

                                <!-- Mode Selection -->
                                <div style="display:flex; gap:15px; margin-bottom:15px; padding-bottom:10px; border-bottom:1px solid #eee;">
                                  <label style="font-size:12px; font-weight:700; color:#37474f; display:flex; align-items:center; cursor:pointer;">
                                    <input type="radio" name="payMode" value="percent" style="margin-right:6px;"> Percentage (%)
                                  </label>
                                  <label style="font-size:12px; font-weight:700; color:#37474f; display:flex; align-items:center; cursor:pointer;">
                                    <input type="radio" name="payMode" value="fixed" checked style="margin-right:6px;"> Fixed Amount ($)
                                  </label>
                                </div>

                                <!-- Rows Container -->
                                <div id="payRowsContainer"></div>

                                <!-- Add Button -->
                                <button id="btnAddPayRow" class="btn-action" style="width:100%; margin-top:10px; border-style:dashed; color:#1565c0;">+ Add Payment Milestone</button>

                                <!-- Validation Error -->
                                <div id="payError" style="color:#c62828; font-size:11px; font-weight:700; margin-top:10px; display:none; text-align:center;"></div>
                                <div id="payDateWarning" style="color:#e65100; font-size:11px; font-weight:700; margin-top:5px; display:none; text-align:center; background:#fff3e0; padding:5px; border:1px solid #ffe0b2;"></div>
                              </div>

                              <!-- Footer -->
                              <div style="padding:15px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px;">
                                <button id="btnPayApply" class="btn-platinum-config primary" style="flex:1;">APPLY SCHEDULE</button>
                                <button id="btnPayCancel" class="btn-platinum-config" style="flex:1;">CANCEL</button>
                              </div>
                              `;

  overlay.appendChild(card);
  var showVerbatim = scheduleString && scheduleString.length > 5;
  if (showVerbatim) {
    var verbatimBox = document.createElement("div");
    verbatimBox.className = "tool-box slide-in";
    verbatimBox.style.cssText = "background: #fffde7; border: 1px solid #fbc02d; width: 250px; max-height: 90vh; display: flex; flex-direction: column; border-radius: 0px; box-shadow: 0 10px 25px rgba(0,0,0,0.2);";
    verbatimBox.innerHTML = `
                              <div style="background:#fff9c4; padding:10px; font-size:11px; font-weight:700; color:#f57f17; border-bottom:1px solid #fbc02d;">
                                🤖 AI Detected Text
                              </div>
                              <div style="padding:15px; font-size:12px; color:#5d4037; line-height:1.5; overflow-y:auto;">
                                ${scheduleString}
                              </div>
                              `;
    overlay.appendChild(verbatimBox);
  }

  document.body.appendChild(overlay);

  // --- LOGIC ---
  var container = document.getElementById("payRowsContainer");
  var modeRadios = document.getElementsByName("payMode");
  var currentMode = "fixed";
  var currencySymbol = "$";
  var currencyCode = "USD";

  var currencyMap = {
    "USD": "$", "EUR": "€", "GBP": "£"
  };

  function renderRow(prefillAmt, prefillType, prefillDate) {
    var div = document.createElement("div");
    div.className = "pay-row";
    div.style.cssText = "background:#fcfcfc; border:1px solid #e0e0e0; padding:10px; margin-bottom:8px; position:relative;";

    div.innerHTML = `
                              <div style="display:flex; gap:10px; margin-bottom:8px;">
                                <div style="flex:1;">
                                  <label class="tool-label" style="font-size:9px;">Amount (${currentMode === 'percent' ? '%' : currencySymbol})</label>
                                  <input type="number" class="tool-input pay-amount" style="margin-bottom:0; height:28px;" placeholder="0" value="${prefillAmt || ''}">
                                </div>
                                <div style="flex:1; display:flex; flex-direction:column; justify-content:flex-end;">
                                  <div class="pay-calc" style="font-size:11px; color:#546e7a; font-weight:600; background:#eceff1; padding:6px; text-align:center; border-radius:2px;">${currencySymbol}0.00</div>
                                </div>
                              </div>
                              <div style="display:flex; gap:10px;">
                                <div style="flex:1;">
                                  <select class="audience-select pay-trigger-type" style="margin-bottom:0; height:28px; padding:2px;">
                                    <option value="event" ${!prefillType || prefillType === 'event' ? 'selected' : ''}>On Event</option>
                                    <option value="date" ${prefillType === 'date' ? 'selected' : ''}>Specific Date</option>
                                    <option value="other" ${prefillType === 'other' ? 'selected' : ''}>Other</option>
                                  </select>
                                </div>
                                <div style="flex:1;">
                                  <input type="${prefillType === 'date' ? 'date' : 'text'}" class="tool-input pay-trigger-val" style="margin-bottom:0; height:28px; font-family:sans-serif;" value="${prefillDate || ''}" placeholder="${prefillType === 'date' ? '' : 'e.g. Execution'}">
                                </div>
                              </div>
                              <div class="btn-del-row" style="position:absolute; top:-6px; right:-6px; background:#c62828; color:white; width:16px; height:16px; border-radius:50%; font-size:10px; display:flex; align-items:center; justify-content:center; cursor:pointer; box-shadow:0 2px 4px rgba(0,0,0,0.2);">✕</div>
                              `;

    // Bind Events
    var amountInput = div.querySelector(".pay-amount");
    var calcDisplay = div.querySelector(".pay-calc");
    var typeSel = div.querySelector(".pay-trigger-type");
    var valInput = div.querySelector(".pay-trigger-val");
    var delBtn = div.querySelector(".btn-del-row");

    amountInput.oninput = function () {
      var val = parseFloat(this.value) || 0;
      if (currentMode === 'percent') {
        var calc = (val / 100) * totalFee;
        calcDisplay.innerText = currencySymbol + calc.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      } else {
        var calc = (val / totalFee) * 100;
        calcDisplay.innerText = calc.toFixed(0) + "%"; // No decimals for percent
      }
      validateTotal();
    };

    typeSel.onchange = function () {
      if (this.value === 'date') {
        valInput.type = "date";
        valInput.placeholder = "";
      } else {
        valInput.type = "text";
        valInput.placeholder = this.value === 'event' ? "e.g. Execution" : "Details...";
      }
    };
    valInput.onchange = validateDates; // Check dates on change

    delBtn.onclick = function () {
      if (container.children.length > 1) {
        div.remove();
        validateTotal();
      }
    };

    container.appendChild(div);
    // Trigger calc immediately if prefilled
    if (prefillAmt) amountInput.oninput();
  }

  function validateTotal() {
    var inputs = document.querySelectorAll(".pay-amount");
    var sum = 0;
    inputs.forEach(inp => sum += (parseFloat(inp.value) || 0));

    var err = document.getElementById("payError");
    var btn = document.getElementById("btnPayApply");

    var isValid = false;
    if (currentMode === 'percent') {
      if (Math.abs(sum - 100) < 0.1) isValid = true;
      else err.innerText = `Total: ${sum}% (Must be 100%)`;
    } else {
      if (Math.abs(sum - totalFee) < 1.0) isValid = true;
      else err.innerText = `Total: ${currencySymbol}${sum} (Must be ${currencySymbol}${totalFee})`;
    }

    if (isValid) {
      err.style.display = "none";
      btn.style.opacity = "1";
      btn.disabled = false;
    } else {
      err.style.display = "block";
      btn.style.opacity = "0.5";
      btn.disabled = true;
    }
  }

  function validateDates() {
    var rows = container.querySelectorAll(".pay-row");
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var warnings = [];
    var lastDate = null;

    rows.forEach((row, index) => {
      var type = row.querySelector(".pay-trigger-type").value;
      var val = row.querySelector(".pay-trigger-val").value;

      if (type === 'date' && val) {
        var parts = val.split('-');
        // Create date in local time (avoid timezone shifts)
        var d = new Date(parts[0], parts[1] - 1, parts[2]);

        if (d < today) warnings.push(`Row ${index + 1}: Date is in the past.`);
        if (lastDate && d < lastDate) warnings.push(`Row ${index + 1}: Date is out of order.`);

        lastDate = d;
      }
    });

    var warnDiv = document.getElementById("payDateWarning");
    if (warnDiv) {
      if (warnings.length > 0) {
        warnDiv.innerHTML = "⚠️ " + warnings.join("<br>");
        warnDiv.style.display = "block";
      } else {
        warnDiv.style.display = "none";
      }
    }
  }
  document.getElementById("btnAddPayRow").onclick = function () { renderRow(); };

  document.getElementById("payCurrency").onchange = function () {
    currencyCode = this.value;
    currencySymbol = currencyMap[currencyCode] || "$";

    // Update Header
    document.getElementById("headerTotalDisplay").innerText = currencySymbol + formattedTotal;

    // Re-render rows to update labels
    container.innerHTML = "";
    renderRow();
    validateTotal();
  };

  modeRadios.forEach(r => r.onchange = function () {
    currentMode = this.value;
    container.innerHTML = ""; // Clear rows on mode switch to avoid confusion
    renderRow();
    validateTotal();
  });

  // --- AUTO-FILL PARSER ---
  // Try to detect if the AI put a schedule in the string
  var hasSchedule = false;
  var scheduleTextForParsing = scheduleString || "";
  // Remove common prefixes
  scheduleTextForParsing = scheduleTextForParsing.replace(/^(describe the payment schedule\/terms:|payment schedule:|payment terms:|fees:)\s*/i, "").trim();

  if (scheduleTextForParsing.includes("%") || /on\s|upon\s|due\s/i.test(scheduleTextForParsing)) {
    // Simple heuristic parser
    var parts = scheduleTextForParsing.split(/[;,/]|(?:\s+and\s+)/i);
    parts = parts.filter(p => p && p.trim().length > 5); // Filter out empty/short parts

    var validRows = 0;
    parts.forEach(function (part) {
      var amtMatch = part.match(/(\d+\.?\d*)\s*%/); // Handle decimals in percent
      var amt = amtMatch ? amtMatch[1] : "";

      var dateMatch = part.match(/(?:on|upon|due)\s+(.+)/i);
      var dateVal = dateMatch ? dateMatch[1].trim().replace(/\)$/, "") : "";

      if (amt && dateVal) { // Require both to be found
        if (validRows === 0 && scheduleTextForParsing.includes("%")) {
          currentMode = "percent";
          modeRadios.forEach(r => r.checked = (r.value === "percent"));
        }
        renderRow(amt, "event", dateVal);
        validRows++;
      }
    });
    if (validRows > 0) hasSchedule = true;
  }

  if (!hasSchedule) {
    // FIX: If we just have a flat fee (e.g. "100"), default to 1 row with that amount
    if (totalFee > 0) {
      // If mode is percent, default to 100%
      // If mode is fixed, default to totalFee
      var defaultAmt = (currentMode === 'percent') ? "100" : totalFee;
      renderRow(defaultAmt);
    } else {
      renderRow(); // Default empty row
    }
  } else {
    // If we parsed rows, validate them
    validateTotal();
  }

  document.getElementById("btnPayCancel").onclick = function () {
    overlay.remove();
    // Do not execute callback (stops the fill process)
  };

  document.getElementById("btnPayApply").onclick = function () {
    // Generate String 
    var rowElements = document.querySelectorAll(".pay-row");
    var parts = [];

    rowElements.forEach(function (row) {
      var amt = row.querySelector(".pay-amount").value;
      var calc = row.querySelector(".pay-calc").innerText;
      var type = row.querySelector(".pay-trigger-type").value;
      var val = row.querySelector(".pay-trigger-val").value;

      var mainVal = "";
      var subVal = "";

      if (currentMode === 'percent') {
        mainVal = amt + "%";
        subVal = "(" + calc + ")"; // ($500.00)
      } else {
        // Fixed Amount: Force .00 formatting
        var fAmt = parseFloat(amt) || 0;
        mainVal = currencySymbol + fAmt.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        subVal = ""; // Hide percent in document for fixed mode
      }

      var triggerText = val;

      // Keep "on" only for the Date Picker mode
      if (type === 'date' && val) {
        var dateParts = val.split('-');
        if (dateParts.length === 3) {
          var d = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
          triggerText = "on " + d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
        }
      }

      parts.push(`${mainVal} ${subVal} ${triggerText}`.replace(/\s+/g, " ").trim());
    });
    var finalStr = "due as follows: " + parts.join("; ") + ".";

    // Fallback if empty (shouldn't happen with validation)
    if (parts.length === 0) finalStr = "Payment schedule to be determined.";

    overlay.remove();
    callback(finalStr);
  };
}

// NEW FEATURE: NATIVE SAVE AS DIALOG (V7 - File System API)

var GLOBAL_FILE_BLOB = null;
var GLOBAL_FILENAME = "";

// ==========================================
// REPLACE: showSaveAsManager (Asks Scanner for Data)
// ==========================================
function showSaveAsManager(autoSnapshotLabel) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- 1. RENDER UI (Initial State) ---
  // We render the fields immediately so the user sees the interface,
  // but we put "Scanning..." placeholders in them.

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.borderLeft = "5px solid #2e7d32"; // Green border
  card.style.padding = "0";

  // Header
  var headerHtml = `
      <div style="background:linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding:15px; border-bottom:3px solid #c9a227; margin-bottom:15px; display:flex; justify-content:space-between; align-items:center;">
        <div style="font-family:'Georgia', serif; font-size:18px; color:white;">💾 Save Document</div>
        <div onclick="showFileNameSettings()" style="cursor:pointer; font-size:16px; color:#fff; opacity:0.8;" title="File Name Patterns">⚙️</div>
      </div>
  `;

  var html = headerHtml + `
      <div style="padding:15px;">
        <div id="scanLoader" style="background:#e3f2fd; padding:10px; border-radius:4px; margin-bottom:15px; font-size:11px; color:#1565c0; display:flex; align-items:center; gap:10px;">
           <div class="spinner" style="width:12px;height:12px;border-width:2px;"></div>
           <strong>Scanning Document for Metadata...</strong>
        </div>

        <div style="margin-bottom:15px;">
          <label class="tool-label">CP Contracting Entity</label>
          <select id="saveEntity" class="audience-select" style="font-weight:600; margin-bottom:8px;">
            <option value="Colgate">Colgate-Palmolive Company</option>
            <option value="Hills">Hill's Pet Nutrition</option>
          </select>
        </div>

        <div style="margin-bottom:15px;">
          <label class="tool-label">Agreement Type</label>
          <select id="saveType" class="audience-select" style="font-weight:600; margin-bottom:8px;">
            <option value="Non-Disclosure Agreement">Non-Disclosure Agreement</option>
            <option value="Service Agreement">Service Agreement</option>
            <option value="Master Service Agreement">Master Service Agreement</option>
            <option value="Statement of Work">Statement of Work</option>
            <option value="Amendment">Amendment</option>
            <option value="Material Transfer Agreement">Material Transfer Agreement</option>
            <option value="Speaker Agreement">Speaker Agreement</option>
            <option value="Other">Other (Fill in)...</option>
          </select>
          <input type="text" id="saveTypeManual" class="tool-input" placeholder="Type Agreement Name..." style="display:none;">
        </div>

        <div style="margin-bottom:15px;">
          <label class="tool-label">Third Party Name</label>
          <input type="text" id="saveParty" class="tool-input" placeholder="⏳ Waiting for scanner...">
        </div>

        <div style="margin-bottom:20px;">
          <label class="tool-label">CP Ref #</label>
          <input type="text" id="saveRef" class="tool-input" placeholder="e.g. 12345">
        </div>

        <div style="background:#f5f5f5; padding:10px; border-radius:4px; margin-bottom:15px; font-size:10px; color:#666;">
          <strong>Preview Filename:</strong><br>
          <span id="fileNamePreview" style="color:#333; font-family:monospace;">...</span>
        </div>

        <button id="btnOneClickSave" class="btn btn-blue" style="width:100%; margin-bottom:10px;">💾 Save Document</button>
        <button id="btnSkipSave" class="btn-action" style="width:100%;">Skip (Go to Email)</button>
        
        <div id="saveStatus" style="margin-top:10px; font-size:11px; color:#666; text-align:center; display:none;">
          <div class="spinner" style="width:12px;height:12px;border-width:2px;display:inline-block;"></div> processing...
        </div>
      </div>
  `;

  card.innerHTML = html;
  out.appendChild(card);
  // [NEW] 1. Lock the Card & Show Overlay
  card.style.position = "relative"; // Ensures overlay stays inside this box

  var overlayId = "step4_scan_overlay";
  var scanOverlay = document.createElement("div");
  scanOverlay.id = overlayId;
  scanOverlay.style.cssText = "position:absolute; top:0; left:0; width:100%; height:100%; background:rgba(255,255,255,0.96); z-index:10; display:flex; flex-direction:column; align-items:center; justify-content:center; backdrop-filter:blur(2px); border-radius:0px;";

  scanOverlay.innerHTML = `
      <div class="spinner" style="width:40px; height:40px; border-width:4px; border-color:#1565c0 transparent #1565c0 transparent; margin-bottom:15px;"></div>
      <div style="font-family:'Georgia', serif; font-size:18px; color:#1565c0; font-weight:bold; margin-bottom:5px;">Scanning Document...</div>
      <div style="font-family:'Inter', sans-serif; font-size:11px; color:#546e7a;">Extracting Parties & Ref #</div>
  `;
  card.appendChild(scanOverlay);

  // Call the Master Scanner
  runAiEmailScanner(null, function (aiData) {

    // [NEW] 2. Remove Overlay immediately when data returns
    var existingOv = document.getElementById(overlayId);
    if (existingOv) existingOv.remove();

    // Store globally for Step 5
    window.saveManagerData = aiData;

    // Hide old Loader (just in case)
    var loader = document.getElementById("scanLoader");
    if (loader) loader.style.display = "none";

    // Populate Fields
    if (aiData.third_party_entity) document.getElementById("saveParty").value = aiData.third_party_entity;
    document.getElementById("saveParty").placeholder = "e.g. Acme Corp"; // Reset placeholder

    if (aiData.cp_ref) document.getElementById("saveRef").value = aiData.cp_ref;

    // Entity Logic
    var rawEnt = aiData.colgate_entity || aiData.entity || "Colgate";
    var ent = rawEnt.toLowerCase();

    var selEnt = document.getElementById("saveEntity");
    selEnt.value = (ent.includes("hill")) ? "Hills" : "Colgate";

    // Type Logic
    var rawType = (aiData.type || "").toLowerCase();
    var selType = document.getElementById("saveType");
    var manualType = document.getElementById("saveTypeManual");
    var found = false;
    for (var i = 0; i < selType.options.length; i++) {
      if (rawType.includes(selType.options[i].value.toLowerCase().replace("agreement", "").trim())) {
        selType.selectedIndex = i; found = true; break;
      }
    }
    if (!found && rawType.length > 2) {
      selType.value = "Other";
      manualType.style.display = "block";
      manualType.value = aiData.type;
    }

    // Refresh Preview
    updatePreview();
    showToast("✅ Autofill Complete");
  });

  // --- 3. UI LOGIC (Preview & Buttons) ---
  var entitySel = document.getElementById("saveEntity");
  var typeSel = document.getElementById("saveType");
  var typeManual = document.getElementById("saveTypeManual");
  var partyInp = document.getElementById("saveParty");
  var refInp = document.getElementById("saveRef");
  var previewEl = document.getElementById("fileNamePreview");

  typeSel.onchange = function () {
    typeManual.style.display = (this.value === "Other") ? "block" : "none";
    updatePreview();
  };

  function generateFileName() {
    var entity = entitySel.value;
    var type = (typeSel.value === "Other") ? (typeManual.value || "Agreement") : typeSel.value;
    var party = partyInp.value.trim() || "Unknown";
    var ref = refInp.value.trim() || "Draft";
    var today = new Date();
    var dateStr = (today.getMonth() + 1).toString().padStart(2, '0') + "." + today.getDate().toString().padStart(2, '0') + "." + today.getFullYear();
    var patterns = userSettings.fileNamePatterns || { colgate: "[Ref] - CP-[Party] -- [Type] (DRAFT [Date])", hills: "[Ref] - HPN-[Party] -- [Type] (DRAFT [Date])" };
    var pattern = (entity === "Hills") ? patterns.hills : patterns.colgate;
    return pattern.replace("[Ref]", ref).replace("[Party]", party).replace("[Type]", type).replace("[Date]", dateStr) + ".docx";
  }

  function updatePreview() { previewEl.innerText = generateFileName(); }
  [entitySel, typeSel, typeManual, partyInp, refInp].forEach(el => { el.addEventListener('input', updatePreview); el.addEventListener('change', updatePreview); });
  updatePreview(); // Initial preview (Empty)

  // --- HELPER: HANDOFF TO STEP 5 ---
  function proceedToEmail() {
    // 1. Capture current edits from Save Screen
    var uiData = {
      type: typeSel.value,
      entity: entitySel.value,
      third_party_entity: partyInp.value,
      cp_ref: refInp.value
    };

    // 2. Merge with the hidden email data we fetched earlier
    // (This ensures we don't lose the contact info!)
    var fullData = { ...window.saveManagerData, ...uiData };

    // 3. Call Scanner again (It will see 'fullData' and skip re-scanning)
    runAiEmailScanner(fullData);
  }

  // --- SAVE BUTTON ---
  document.getElementById("btnOneClickSave").onclick = function () {
    var btn = this;
    var status = document.getElementById("saveStatus");
    btn.disabled = true;
    btn.innerHTML = "⏳ Generating File...";
    status.style.display = "block";
    var finalName = generateFileName();

    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var myFile = result.value;
        var sliceCount = myFile.sliceCount;
        var slicesReceived = 0;
        var docdata = [];
        var getSlice = function (i) {
          myFile.getSliceAsync(i, function (sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              if (sliceResult.value.data) {
                var data = sliceResult.value.data;
                docdata = docdata.concat(Array.isArray(data) ? data : Array.from(new Uint8Array(data)));
              }
              slicesReceived++;
              if (slicesReceived == sliceCount) {
                myFile.closeAsync();
                var blob = new Blob([new Uint8Array(docdata)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
                saveFileWithDialog(blob, finalName, function () {

                  // AUTO SNAPSHOT
                  btn.innerHTML = "📷 Capturing Version...";
                  createSnapshot("CP Version: " + finalName + " [Original Version]", null, function () {
                    showToast("✅ Saved & Captured!");
                    proceedToEmail(); // GO TO EMAIL
                    status.style.display = "none";
                    btn.innerHTML = "✅ Done";
                  });

                });
              } else { getSlice(slicesReceived); }
            }
          });
        };
        getSlice(0);
      } else {
        showOutput("Error: " + result.error.message, true);
        btn.disabled = false;
        btn.innerHTML = "💾 Save Document";
      }
    }
    );
  };
  // --- SKIP BUTTON ---
  document.getElementById("btnSkipSave").onclick = function () {
    // CHANGE: Create a snapshot before moving to email
    createSnapshot("[SENT VERSION -- NOT SAVED THROUGH ELI]", null, function () {
      showToast("⏩ Snapshot Saved & Opening Emailer...");
      proceedToEmail();
    }, true); // forceNew = true
  };

  // Auto-Snapshot logic (runs silently)

  // Auto-Snapshot logic (runs silently)
  if (autoSnapshotLabel) {
    // CHANGE: Added 'true' as 4th argument to force a NEW entry (no overwrite)
    createSnapshot(autoSnapshotLabel, null, function () {
      console.log("Snapshot Created: " + autoSnapshotLabel);
      showToast("📷 Snapshot Saved: " + autoSnapshotLabel); // Optional visual confirmation
    }, true);
  }
}
// ==========================================
// ==========================================
async function saveFileWithDialog(blob, fileName, onSuccess) {
  // Method A: Use Modern 'showSaveFilePicker' (Chrome/Edge/Desktop)
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: fileName,
        types: [
          {
            description: "Word Document",
            accept: {
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                [".docx"],
            },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();

      if (onSuccess) onSuccess();
      else { showToast("✅ Saved Successfully"); init(); }
    } catch (err) {
      // User hit "Cancel" in the dialog
      console.log("User cancelled save");
    }
  }
  // Method B: Fallback for Older Browsers (Standard Download)
  else {
    var link = document.createElement("a");
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    if (onSuccess) onSuccess();
    else { showToast("✅ File Downloaded"); init(); }
  }
}
// ==========================================
// START: PROMPTS.saveActiveDocCache (New Helper)
// ==========================================
PROMPTS.saveActiveDocCache = function (id, name, data, aiData) {
  return new Promise((resolve, reject) => {
    // Ensure DB is ready
    if (!PROMPTS._dbInstance) {
      PROMPTS.initDocDB().then(() => PROMPTS.saveActiveDocCache(id, name, data, aiData).then(resolve).catch(resolve));
      return;
    }

    const transaction = PROMPTS._dbInstance.transaction([PROMPTS._DOC_STORE_NAME], "readwrite");
    const store = transaction.objectStore(PROMPTS._DOC_STORE_NAME);

    // We overwrite 'currentDoc' to represent the Active Document
    const putRequest = store.put({
      id: id,
      data: null, // No file blob for active doc
      name: name || "Active Document",
      type: "ActiveDocument",
      aiCache: aiData,
      lastModified: Date.now()
    });
    putRequest.onsuccess = resolve;
    putRequest.onerror = (e) => { console.warn("Cache Save Fail", e); resolve(); }; // Soft fail
  });
};
async function runAiEmailScanner(manualData, callbackMode, forceFresh) {

  if (manualData && !callbackMode && (manualData.third_party_entity || manualData.contact_email || manualData.type)) {
    console.log("⚡ Skipping Scan: Data provided by Save Manager");
    showThirdPartySendModal(manualData, null, 'req');
    return;
  }

  //showToast("🔍 1. Starting Email Scan...");

  var filePayload = null;

  // 🔥 ADDED: Helper to identify garbage
  function isClean(val) {
    if (!val) return false;
    var s = String(val).toLowerCase();
    if (s.includes("application/") || s.includes("image/") || s.includes("vnd.") || s.includes("openxml")) {
      return false;
    }
    return true;
  }

  try {
    // --- STEP 1: CHECK RAM (Fastest) ---
    // SURGICAL FIX: Check both window scope and global scope. Always prioritize this if data exists.
    var sourceData = (typeof window.AGREEMENT_DATA !== 'undefined' ? window.AGREEMENT_DATA : null) ||
      (typeof AGREEMENT_DATA !== 'undefined' ? AGREEMENT_DATA : null);

    if (sourceData && (sourceData.text || sourceData.data)) {
      showToast("✅ 2. Re-Loading OARS Data");
      // Normalize payload from RAM
      if (sourceData.isBase64) filePayload = sourceData;
      else filePayload = sourceData.text || sourceData;
    }
    else if (!forceFresh) {
      // --- STEP 2: FORCE DB LOAD (Self-Reliant) ---
      if (!PROMPTS._dbInstance) await PROMPTS.initDocDB();

      var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';
      var dbPromise = PROMPTS.loadDocFromDB(uniqueKey);
      var timerPromise = new Promise(r => setTimeout(() => r('TIMEOUT'), 3000));
      var record = await Promise.race([dbPromise, timerPromise]);

      if (record === 'TIMEOUT') {
        showToast("❌ DB Timeout");
      }
      else if (!record) {
        showToast("⚠️ No OARS Data Found");
      }
      else {
        if (record.aiCache && Array.isArray(record.aiCache)) {
          console.log("⚡ FAST LOAD: Using Saved AI Cache");

          showToast("⚡ Restoring from Cache...");

          var cachedData = {};
          record.aiCache.forEach(function (item) {
            cachedData[item.field] = item.data;
          });

          Object.keys(cachedData).forEach(function (key) {
            if (!isClean(cachedData[key])) {
              console.log("🗑️ Cleaning Garbage from Cache:", key);
              delete cachedData[key];
            }
          });
          // -----------------------------------------------------

          // ROBUST FIND: Map 'type' -> 'agreement_type'
          var foundType = null;
          Object.keys(cachedData).forEach(function (k) {
            if (k.toLowerCase().includes('type') || k.toLowerCase().includes('agreement')) {
              foundType = cachedData[k];
            }
          });

          if (foundType) {
            // CLEAN: Remove quotes/spaces
            var cleanType = foundType.replace(/^["']|["']$/g, "").trim();
            cachedData.agreement_type = cleanType;
            cachedData.type = cleanType;
          }

          // Launch Modal & EXIT immediately (Bypasses file check)
          if (callbackMode) {
            console.log("⚡ Returning Cached Data to Callback");
            callbackMode(cachedData);
          } else {
            // Launch Modal & EXIT immediately
            showThirdPartySendModal(cachedData, null, 'tp');
          }
          return;
        }
        // ---------------------------------------------------------

        // If no cache, THEN check for file data
        if (!record.data) {
          showToast("❌ Record found but DATA is missing");
        }
        else {
          showToast("✅ OARS Data Found: " + record.name);

          // --- STEP 3: EXTRACT CONTENT LOCALLY ---
          if (record.name.toLowerCase().endsWith(".docx")) {
            // Word Doc -> Text
            if (typeof mammoth === 'undefined') {
              showToast("❌ Mammoth Library Missing");
            } else {
              var result = await mammoth.convertToHtml({ arrayBuffer: record.data });
              var cleanText = result.value.replace(/<[^>]*>/g, ' ');
              filePayload = cleanText;
              window.AGREEMENT_DATA = { name: record.name, text: cleanText };
            }
          }
          else if (record.name.toLowerCase().endsWith(".pdf")) {
            // PDF -> Base64
            var reader = new FileReader();
            var b64 = await new Promise(resolve => {
              reader.onloadend = () => resolve(reader.result.split(',')[1]);
              reader.readAsDataURL(new Blob([record.data]));
            });
            filePayload = { name: record.name, isBase64: true, mime: "application/pdf", data: b64 };
            window.AGREEMENT_DATA = filePayload;
          }
          else {
            // Text File -> String
            var dec = new TextDecoder("utf-8");
            filePayload = dec.decode(record.data);
            window.AGREEMENT_DATA = { name: record.name, text: filePayload };
          }
        }
      }
    }
  } catch (e) {
    console.error(e);
    showToast("❌ Error: " + e.message);
  }

  // --- STEP 4: EXECUTE AI ---
  if (filePayload) {
    if (!callbackMode) setLoading(true, "summary");

    // Determine Prompt (Targeted vs Full)
    var promptText = "";
    if (manualData) {
      promptText = `
          You are a Senior Legal Assistant.
          TASK: Scan the ATTACHED document file to populate an email cover letter.
          EXTRACT THESE SPECIFIC FIELDS:
          1. Agreement Type (e.g. NDA, MSA, SOW).
          2. Entities: Colgate/Hill's vs The Third Party.
          3. Contact Info: Look for the primary Third Party Contact (Name & Email) in the "Notices" section or signature block.
          4. Requester: Look for the "Colgate" or "Hill's" employee requesting the agreement (Name & Email).
          5. CP Ref #: Look specifically in headers/footers for "CP Ref", "Ref #", or code like "A-12345".
          6. Additional Participants: Look for an "Additional Participants" field. This is never the VP approver.

          RETURN RAW JSON:
          {
            "type": "Agreement Type",
            "colgate_entity": "Colgate | Hill's",
            "third_party_entity": "Company Name",
            "contact_full_name": "Full Name",
            "contact_email": "email@address.com",
            "requester_name": "Colgate Employee Name",
            "requester_email": "colgate_employee@colpal.com",
            "cp_ref": "A-2025-12345-US",
            "additional_participants": [{"email": "..."}]
          }`;
    } else {
      promptText = `
          You are a Senior Legal Assistant.
          TASK: Scan the ATTACHED document file to populate an email cover letter.
          EXTRACT THESE SPECIFIC FIELDS:
          1. Agreement Type (e.g. NDA, MSA, SOW).
          2. Entities: Colgate/Hill's vs The Third Party.
          3. Contact Info: Look for the primary Third Party Contact (Name & Email) in the "Notices" section or signature block.
          4. Requester: Look for the "Colgate" or "Hill's" employee requesting the agreement (Name & Email).
          5. CP Ref #: Look specifically in headers/footers for "CP Ref", "Ref #", or code like "A-12345".
          6. Additional Participants: Look for an "Additional Participants" field. This is never the VP approver.

          RETURN RAW JSON:
          {
            "type": "Agreement Type",
            "colgate_entity": "Colgate | Hill's",
            "third_party_entity": "Company Name",
            "contact_full_name": "Full Name",
            "contact_email": "email@address.com",
            "requester_name": "Colgate Employee Name",
            "requester_email": "colgate_employee@colpal.com",
            "cp_ref": "A-2025-12345-US",
            "additional_participants": [{"email": "..."}]
          }`;
    }

    try {
      let parts = [];
      // Handle Object (PDF) vs String (Text)
      if (filePayload.isBase64) {
        parts = [{ text: promptText }, { inlineData: { mimeType: filePayload.mime, data: filePayload.data } }];
      } else {
        parts = [{ text: promptText + "\n\nDOCUMENT TEXT:\n" + filePayload.substring(0, 500000) }];
      }

      var result = await callGemini(API_KEY, parts);
      var data = tryParseGeminiJSON(result);

      // --- FIX 2: SAFE CACHE SAVE ---
      // This saves the result so it's fast next time, but checks if the record exists first.

      // ------------------------------

      if (data.additional_participants) {
        data.additional_participants = normalizeAIArray(data.additional_participants);
      }

      // --- MERGE LOGIC ---
      // If we came from Build Draft (manualData exists), overlay those values on top of the AI scan
      if (manualData) {
        for (var key in manualData) {
          if (manualData[key] && manualData[key].trim() !== "") {
            data[key] = manualData[key];
          }
        }
      }

      if (!callbackMode) setLoading(false);

      if (callbackMode) {
        // SILENT MODE: Return data to Save Manager
        callbackMode(data);
      } else {
        // STANDARD MODE: Open Email UI
        var tab = manualData ? 'req' : 'tp';
        showThirdPartySendModal(data, null, tab);
      }

    } catch (e) {
      console.error(e);
      setLoading(false);
      showToast("⚠️ AI Error: " + e.message);
      showThirdPartySendModal(manualData || {}, null, 'tp');
    }
  } else {
    // --- STEP 3: FALLBACK - SCAN ACTIVE WORD DOCUMENT ---
    console.log("Scanning Active Word Document (Fallback)...");
    if (!callbackMode) setLoading(true, "summary");

    var fullTextContent = "";

    try {
      await Word.run(async function (context) {
        var sections = context.document.sections;
        sections.load("items");
        var body = context.document.body;
        body.load("text");
        await context.sync();

        var rangesToLoad = [body];
        sections.items.forEach(function (section) {
          ['Primary', 'FirstPage', 'EvenPages'].forEach(type => {
            try { rangesToLoad.push(section.getFooter(type)); } catch (e) { }
            try { rangesToLoad.push(section.getHeader(type)); } catch (e) { }
          });
        });

        rangesToLoad.forEach(function (r) { r.load("text"); });
        await context.sync();

        rangesToLoad.forEach(function (r) {
          fullTextContent += r.text + "\n";
        });
      });
    } catch (e) {
      console.warn("Background text scan failed:", e);
      fullTextContent = await validateApiAndGetText(false, null);
    }

    const scanPrompt = `
          You are a Senior Legal Assistant. Extract metadata for an email draft.

          TASK: Scan the document text provided below.

          RETURN RAW JSON:
          {
            "type": "NDA",
            "colgate_entity": "Colgate | Hill's",
            "third_party_entity": "Company Name",
            "contact_first_name": "First Name",
            "contact_email": "email@address.com",
            "cp_ref": "A-2025-12345-US",
            "additional_participants": [{"name": "...", "email": "..."}]
          }

          DOCUMENT TEXT:
          """
          ${fullTextContent.substring(0, 500000)}
          """`;

    try {
      const result = await callGemini(API_KEY, scanPrompt);
      const data = tryParseGeminiJSON(result);
      if (data.additional_participants) {
        data.additional_participants = normalizeAIArray(data.additional_participants);
      }

      if (!callbackMode) setLoading(false);

      if (callbackMode) {
        callbackMode(data);
      } else {
        showThirdPartySendModal(data, null, 'tp');
      }

    } catch (e) {
      setLoading(false);
      console.error(e);
      showThirdPartySendModal({}, null, 'tp');
      showToast("⚠️ Scan failed, please fill manually.");
    }
  }
}
// ==========================================
// REPLACE: showThirdPartySendModal (With Firewall)
// ==========================================
async function showThirdPartySendModal(aiData = {}, autoSelectId = null, defaultDest = 'tp') {

  // --- 🛑 GARBAGE FIREWALL: Sanitize Data on Entry ---
  // This ensures the modal NEVER sees the "application/vnd..." string
  var dirtyKeys = ['agreement_type', 'type'];
  dirtyKeys.forEach(function (key) {
    if (aiData[key]) {
      var val = String(aiData[key]).toLowerCase();
      if (val.includes("application/") || val.includes("vnd.") || val.includes("openxml")) {
        console.log("🔥 Firewall deleted garbage in '" + key + "':", aiData[key]);
        showToast("🧹 Auto-Cleaned: " + key); // Tell you it worked
        aiData[key] = ""; // Wipe it clean
      }
    }
  });
  // -----------------------------------------------------

  var out = document.getElementById("output");
  out.innerHTML = "";

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "border-left: 5px solid #455a64; padding: 0; background: white; margin-bottom: 10px;";

  let savedReqs = await loadFeatureData(STORE_REQUESTERS) || [];

  const hideManual = savedReqs.length > 0 && String(autoSelectId) !== "add_other";
  const inputStyle = "font-size:11px; height:28px; padding:4px 8px; width:100%; box-sizing:border-box;";
  const labelStyle = "font-size:10px; margin-bottom:3px; display:block; color:#546e7a; font-weight:700;";

  card.innerHTML = `
      <div style="background:linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding:15px; border-bottom:3px solid #c9a227; margin-bottom:15px;">
        <div style="font-family:'Georgia', serif; font-size:18px; color:white;">✉️ Contract Emailer</div>
      </div>

      <div style="padding:15px; padding-top:0;">
        <div style="margin-bottom:10px; background:#f5f5f5; padding:8px; border-radius:0px; display:flex; gap:15px;">
          <label style="font-size:11px; cursor:pointer; font-weight:600; color:#333;">
            <input type="radio" name="dest" value="tp" ${defaultDest === 'tp' ? 'checked' : ''}> Third Party
          </label>
          <label style="font-size:11px; cursor:pointer; font-weight:600; color:#333;">
            <input type="radio" name="dest" value="req" ${defaultDest === 'req' ? 'checked' : ''}> Requester
          </label>
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; margin-bottom:10px;">
          <div>
            <label class="tool-label" style="${labelStyle}">Entity</label>
            <select id="emColgateEntity" class="audience-select" style="${inputStyle}">
              <option value="Colgate" ${aiData.colgate_entity && aiData.colgate_entity.includes("Colgate") ? "selected" : ""}>Colgate</option>
              <option value="Hill's" ${aiData.colgate_entity && (aiData.colgate_entity.includes("Hill") || aiData.colgate_entity.includes("Nutrition")) ? "selected" : ""}>Hill's</option>
            </select>
          </div>
          <div>
            <label class="tool-label" style="${labelStyle}">Type</label>
            <select id="emAgType" class="audience-select" style="${inputStyle}">
              <option value="Non-Disclosure Agreement">Non-Disclosure Agreement</option>
              <option value="Material Transfer Agreement">Material Transfer Agreement</option>
              <option value="Speaker Agreement">Speaker Agreement</option>
              <option value="Amendment">Amendment</option>
              <option value="Service Agreement">Service Agreement</option>
              <option value="Master Service Agreement">Master Service Agreement</option>
              <option value="Joint Development Agreement">Joint Development Agreement</option>
              <option value="Clinical Study Agreement">Clinical Study Agreement</option>
              <option value="Research Agreement">Research Agreement</option>
              <option value="Other">Other (Fill in)...</option>
            </select>
          </div>
        </div>

        <div style="margin-bottom:10px;">
          <label class="tool-label" style="${labelStyle}">Third Party Entity</label>
          <input type="text" id="emTpEntity" class="tool-input" value="${aiData.third_party_entity || ""}" style="${inputStyle}">
        </div>

        <div id="tpContactSection" style="display:${defaultDest === 'tp' ? 'grid' : 'none'}; grid-template-columns: 1fr 1fr; gap:10px; margin-bottom:10px;">
          <div>
            <label class="tool-label" style="${labelStyle}">TP Full Name</label>
            <input type="text" id="emTpName" class="tool-input" value="${aiData.contact_full_name || aiData.contact_first_name || ""}" style="${inputStyle}">
          </div>
          <div>
            <label class="tool-label" style="${labelStyle}">TP Email</label>
            <input type="text" id="emTpEmail" class="tool-input" value="${aiData.contact_email || ""}" style="${inputStyle}">
          </div>
        </div>

        <div id="requesterSection" style="background:#e3f2fd; padding:10px; border-radius:0px; margin-bottom:10px; border:1px solid #90caf9;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
            <label class="tool-label" style="margin:0; font-size:10px; color:#1565c0; font-weight:700;">REQUESTER</label>
            <button id="btnDeleteReq" style="background:none; border:none; cursor:pointer; color:#c62828; font-size:10px; font-weight:700; text-transform:uppercase;">🗑️ Remove Name</button>
          </div>
          <select id="selRequester" class="audience-select" style="${inputStyle} border-color:#90caf9;">
            <option value="">-- Select Saved --</option>
            ${savedReqs.map((r) => `<option value="${r.id}" data-email="${r.email}" ${String(autoSelectId) === String(r.id) ? "selected" : ""}>${r.name}</option>`).join("")}
            <option value="add_other" style="font-weight:bold; color:#1565c0;">➕ ADD OTHER / NEW...</option>
          </select>

          <div id="reqManualArea" class="${hideManual ? " hidden" : ""}" style="margin-top:10px; border-top:1px dashed #90caf9; padding-top:10px;">
          <div style="margin-bottom:8px;">
            <input type="text" id="emReqFullName" class="tool-input" placeholder="Full Name" style="${inputStyle}">
          </div>
          <div style="margin-bottom:8px;">
            <input type="text" id="emReqEmail" class="tool-input" placeholder="Email Address" style="${inputStyle}">
          </div>
          <div style="display:flex; gap:10px;">
            <button id="btnSaveReq" class="btn-action" style="flex:1; color:#2e7d32; border-color:#2e7d32; font-size:10px; font-weight:700; background:white; height:28px;">💾 Save & Select</button>
            <button id="btnUseOnce" class="btn-action" style="flex:1; color:#1565c0; border-color:#1565c0; font-size:10px; font-weight:700; background:white; height:28px;">⚡ Use Once</button>
          </div>
        </div>
      </div>

      <div style="margin-bottom:10px;">
        <div id="btnAddParticipants" style="font-size:10px; font-weight:700; color:#1565c0; cursor:pointer; display:inline-block; margin-bottom:5px;">
          ➕ Add Additional Participants
        </div>
        <div id="addParticipantsSection" class="${(aiData.additional_participants && aiData.additional_participants.length > 0) ? "" : "hidden"}" style="background:#f1f8e9; padding:10px; border:1px solid #c5e1a5; margin-bottom:5px;">
        <label class="tool-label" style="${labelStyle}">Additional CCs</label>
        <textarea id="emAddCcs" class="tool-input" style="${inputStyle} height:50px; font-family:monospace;" placeholder="Name <email>, ...">${(aiData.additional_participants || []).map(p => (p.name && p.name !== "Unknown") ? `${p.name} <${p.email}>` : p.email).join(', ')}</textarea>
      </div>
    </div>

    <div style="margin-bottom:15px;">
      <label class="tool-label" style="${labelStyle}">CP Ref #</label>
      <input type="text" id="emCpRef" class="tool-input" value="${aiData.cp_ref || ""}" style="${inputStyle}">
    </div>

    <button id="btnGmailDraft" class="btn btn-blue" style="width:100%; background:linear-gradient(to bottom, #f7f9fa 0%, #dce3e8 100%); padding:12px; font-size:13px; font-weight:800; border:1px solid #b0bec5; border-radius:0px; text-transform:uppercase; box-shadow:inset 0 1px 0 #ffffff, 0 1px 2px rgba(0,0,0,0.08); color:#1565c0;">
      🚀 OPEN GMAIL DRAFT
    </button>

    <button id="btnCancelTp" class="btn-action" style="width:100%; margin-top:8px; font-size:11px; height:28px;">🏠 Home</button>
  </div>
  `;

  out.appendChild(card);

  // --- 🔥 AUTO-FILL LOGIC ---
  setTimeout(function () {
    var agTypeSel = document.getElementById("emAgType");
    var rawType = aiData.agreement_type || aiData.type;

    // 1. Debug Toast to Confirm what we see


    if (rawType && agTypeSel) {
      var cleanTarget = String(rawType).toLowerCase().replace(/[^a-z ]/g, "").trim();

      // 2. SCORING RULES
      var rules = [
        { key: "research", score: 100 },
        { key: "disclosure", score: 100 },
        { key: "confidential", score: 100 },
        { key: "nda", score: 100 },
        { key: "material", score: 100 },
        { key: "transfer", score: 50 },
        { key: "mta", score: 100 },
        { key: "master", score: 100 },
        { key: "service", score: 30 },
        { key: "amendment", score: 100 },
        { key: "clinical", score: 100 },
        { key: "joint", score: 100 },
        { key: "development", score: 50 },
        { key: "agreement", score: 0 } // Ignored
      ];

      var bestIndex = -1;
      var bestScore = 0;

      for (var i = 0; i < agTypeSel.options.length; i++) {
        var optText = agTypeSel.options[i].text.toLowerCase();
        var currentScore = 0;

        if (optText === cleanTarget) currentScore += 500;

        for (var k = 0; k < rules.length; k++) {
          var r = rules[k];
          if (cleanTarget.indexOf(r.key) > -1 && optText.indexOf(r.key) > -1) {
            currentScore += r.score;
          }
        }

        // Anti-Match for "Service" vs "Master Service"
        if (optText.indexOf("master") > -1 && cleanTarget.indexOf("master") === -1) {
          currentScore -= 50;
        }

        if (currentScore > bestScore) {
          bestScore = currentScore;
          bestIndex = i;
        }
      }

      if (bestIndex !== -1 && bestScore > 0) {
        agTypeSel.selectedIndex = bestIndex;

      }
    }

    // Requester Logic
    if (aiData.requester_name) {
      var selReq = document.getElementById("selRequester");
      var found = false;
      if (selReq) {
        for (var i = 0; i < selReq.options.length; i++) {
          if (selReq.options[i].text.toLowerCase() === aiData.requester_name.toLowerCase()) {
            selReq.selectedIndex = i;
            found = true;
            break;
          }
        }
        if (!found) {
          selReq.value = "add_other";
          var manArea = document.getElementById("reqManualArea");
          if (manArea) manArea.classList.remove("hidden");
          document.getElementById("emReqFullName").value = aiData.requester_name;
          if (aiData.requester_email) document.getElementById("emReqEmail").value = aiData.requester_email;
        }
      }
    }

    if (aiData.additional_participants && aiData.additional_participants.length > 0) {
      var btnAdd = document.getElementById("btnAddParticipants");
      if (btnAdd) btnAdd.style.display = 'none';
    }
  }, 300);

  // --- BINDINGS ---
  const selReq = document.getElementById("selRequester");
  const manualArea = document.getElementById("reqManualArea");
  const btnDelete = document.getElementById("btnDeleteReq");
  const tpContactSection = document.getElementById("tpContactSection");

  document.querySelectorAll('input[name="dest"]').forEach(radio => {
    radio.addEventListener('change', function () {
      tpContactSection.style.display = this.value === 'tp' ? 'grid' : 'none';
    });
  });

  selReq.onchange = () => manualArea.classList.toggle("hidden", selReq.value !== "add_other");

  document.getElementById("btnSaveReq").onclick = async function () {
    const name = document.getElementById("emReqFullName").value.trim();
    const email = document.getElementById("emReqEmail").value.trim();
    if (!name || !email) { showToast("⚠️ Enter details first"); return; }
    const newId = String(Date.now());
    savedReqs.push({ id: newId, name: name, email: email });
    await saveFeatureData(STORE_REQUESTERS, savedReqs);
    showToast("✅ Requester Saved");
    showThirdPartySendModal(aiData, newId);
  };

  document.getElementById("btnUseOnce").onclick = () => {
    showToast("⚡ Active for this draft only");
    manualArea.classList.add("hidden");
  };

  btnDelete.onclick = function (e) {
    e.preventDefault();
    const selectedId = selReq.value;
    if (!selectedId || selectedId === "add_other") { showToast("⚠️ Select a saved name first"); return; }
    var delOverlay = document.createElement("div");
    delOverlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:30000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(1px);";
    var delBox = document.createElement("div");
    delBox.className = "tool-box slide-in";
    delBox.style.cssText = "background:white; width:260px; padding:20px; border-left:5px solid #c62828; border-radius:4px; text-align:center;";
    delBox.innerHTML = `
                          <div style="font-size:14px; font-weight:700; margin-bottom:10px; color:#c62828;">Delete Requester?</div>
                          <div style="font-size:11px; color:#555; margin-bottom:15px;">Are you sure?</div>
                          <div style="display:flex; gap:10px;">
                            <button id="btnConfirmDel" class="btn btn-blue" style="flex:1; background:#c62828;">Delete</button>
                            <button id="btnCancelDel" class="btn-action" style="flex:1;">Cancel</button>
                          </div>`;
    delOverlay.appendChild(delBox);
    document.body.appendChild(delOverlay);
    document.getElementById("btnCancelDel").onclick = () => delOverlay.remove();
    document.getElementById("btnConfirmDel").onclick = async function () {
      const updated = savedReqs.filter(r => String(r.id) !== String(selectedId));
      await saveFeatureData(STORE_REQUESTERS, updated);
      delOverlay.remove();
      showToast("🗑️ Removed");
      setTimeout(function () { showThirdPartySendModal(aiData || {}); }, 50);
    };
  };

  var btnAddPart = document.getElementById("btnAddParticipants");
  if (btnAddPart) {
    btnAddPart.onclick = function () {
      var sect = document.getElementById("addParticipantsSection");
      sect.classList.remove("hidden");
      this.style.display = 'none';
    };
  }

  document.getElementById("btnCancelTp").onclick = showHome;

  var gmailBtn = document.getElementById("btnGmailDraft");
  gmailBtn.onclick = function () {
    const entity = document.getElementById("emColgateEntity").value;
    const type = document.getElementById("emAgType").value;
    const tpEntity = document.getElementById("emTpEntity").value || "[Third Party Entity]";
    const tpFullName = document.getElementById("emTpName").value || "[Contact]";
    const tpContact = tpFullName.trim().split(" ")[0];
    const tpEmail = (document.getElementById("emTpEmail").value || "").trim();
    const cpRef = document.getElementById("emCpRef").value || "N/A";

    let reqName = "", reqEmail = "";
    if (selReq.value === "add_other" || selReq.value === "") {
      reqName = document.getElementById("emReqFullName").value;
      reqEmail = document.getElementById("emReqEmail").value;
    } else {
      const selectedOpt = selReq.options[selReq.selectedIndex];
      reqName = selectedOpt.text;
      reqEmail = selectedOpt.dataset.email;
    }
    reqEmail = (reqEmail || "").trim();
    const reqFirstName = (reqName || "[Requester]").split(" ")[0];
    const today = new Date().toLocaleDateString();

    const isToRequester = document.querySelector('input[name="dest"]:checked').value === "req";
    const to = isToRequester ? reqEmail : tpEmail;

    let tmpl = (typeof userSettings !== "undefined" && userSettings.emailTemplates) ? userSettings.emailTemplates : { myName: "Andrew", thirdParty: {}, requester: {} };
    if (!tmpl.thirdParty.subject) tmpl = {
      myName: "Andrew",
      thirdParty: { subject: "[Entity] and [TP Entity] -- [Type] ([CP Ref]) ([Date])", body: "Hello [TP Contact],\n\nAttached please find..." },
      requester: { subject: "FYI: [Entity]...", body: "Hi [Requester]..." }
    };

    const activeTmpl = isToRequester ? tmpl.requester : tmpl.thirdParty;
    let rawSubj = activeTmpl.subject || "";
    let rawBody = activeTmpl.body || "";

    const map = {
      "\\[Entity\\]": entity,
      "\\[TP Entity\\]": tpEntity,
      "\\[Type\\]": type,
      "\\[CP Ref\\]": cpRef,
      "\\[Date\\]": today,
      "\\[TP Contact\\]": tpContact,
      "\\[Requester\\]": reqFirstName,
      "\\[My Name\\]": (function () {
        const rawName = (tmpl.myName || "Andrew").trim();
        if (tmpl.useFullName) return rawName;
        return rawName.split(" ")[0];
      })()
    };

    Object.keys(map).forEach((key) => {
      const regex = new RegExp(key, "g");
      rawSubj = rawSubj.replace(regex, map[key]);
      rawBody = rawBody.replace(regex, map[key]);
    });

    var ccList = [];
    if (reqEmail && reqEmail.length > 3) ccList.push(reqEmail);
    var addCcsInput = document.getElementById("emAddCcs");
    if (addCcsInput && addCcsInput.value) {
      var matches = addCcsInput.value.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
      if (matches && matches.length > 0) {
        for (var m = 0; m < matches.length; m++) ccList.push(matches[m]);
      }
    }
    ccList.push('tech_agreements@colpal.com');

    var uniqueCC = [];
    var seen = {};
    var toAddress = to.toLowerCase().trim();
    for (var i = 0; i < ccList.length; i++) {
      var email = ccList[i].trim();
      if (!email) continue;
      var lower = email.toLowerCase();
      if (lower === toAddress) continue;
      if (seen[lower]) continue;
      uniqueCC.push(email);
      seen[lower] = true;
    }
    var ccString = uniqueCC.join(',');

    var gmailUrl = "https://mail.google.com/mail/?view=cm&fs=1" +
      "&to=" + encodeURIComponent(to) +
      "&cc=" + encodeURIComponent(ccString) +
      "&su=" + encodeURIComponent(rawSubj) +
      "&body=" + encodeURIComponent(rawBody);

    this.innerText = "🚀 Opening...";
    var btn = this;
    setTimeout(function () { btn.innerText = "🚀 OPEN GMAIL DRAFT"; }, 2000);

    if (Office.context.ui && Office.context.ui.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(gmailUrl);
    } else {
      window.open(gmailUrl, "_blank");
    }
  };
}

// ==========================================
function tryCollapseToolbar() {
  // 1. Check if Pinned
  var pinChk = document.getElementById("chkPinToolbar");
  if (pinChk && pinChk.checked) return; // <--- STOPS COLLAPSE IF PINNED

  // 2. Collapse Toolbar
  var toolbar = document.getElementById("mainToolbar");
  if (toolbar) toolbar.classList.add("collapsed");

  // 3. Show Bottom Dock (Force Display)
  var expander = document.getElementById("toolbarExpander");
  if (expander) {
    // ONLY use the class. Do NOT use style.display = "block" here.
    expander.classList.add("visible");
    expander.style.display = ""; // Ensure clean slate
  }
}
// ==========================================
// END: tryCollapseToolbar
// ==========================================
// ==========================================
// START: addWordComment
// ==========================================
function addWordComment(textToFind, commentText, btnElement) {
  var originalLabel = btnElement.innerText;
  btnElement.innerText = "Adding...";

  Word.run(function (context) {
    // 1. Find the text using your robust matcher
    return findSmartMatch(context, textToFind).then(function (range) {
      if (range) {
        // 2. Insert the Comment
        range.insertComment(commentText);

        // 3. UI Feedback
        btnElement.innerText = "✅ Comment Added";
        btnElement.style.color = "#2e7d32";
        btnElement.disabled = true;

        return context.sync();
      } else {
        btnElement.innerText = "⚠️ Not Found";
        setTimeout(function () {
          btnElement.innerText = originalLabel;
        }, 2000);
      }
    });
  }).catch(function (e) {
    console.error(e);
    btnElement.innerText = "Error";
  });
}

// GLOBAL DATA STORE (Holds the analysis for toggling views)
window.redlineSessionData = null;
// ==========================================
// START: runRedlineSummary (V5: Correct Async Queueing)
// ==========================================
function runRedlineSummary(btn) {
  setLoading(true, "default");
  if (btn) btn.disabled = true;

  var hiddenCount = 0;

  Word.run(async function (context) {
    var body = context.document.body;
    var changes = body.getTrackedChanges();
    changes.load("items");
    await context.sync();

    if (changes.items.length === 0) return "NO_CHANGES";

    // --- STEP 1: QUEUE LOADS CORRECTLY ---
    var proxies = [];
    for (var i = 0; i < changes.items.length; i++) {
      var c = changes.items[i];
      c.load("type");

      var r = c.getRange();
      r.load("text"); // Load standard text property

      // CRITICAL FIX: Call getOoxml() HERE, before the sync.
      // We store the *result object* (which is a placeholder until sync).
      var ooxmlResult = r.getOoxml();

      // Context (Paragraph)
      var p = null;
      try {
        p = r.paragraphs.getFirst();
        p.load("text");
      } catch (e) { }

      proxies.push({
        change: c,
        range: r,
        para: p,
        xmlResult: ooxmlResult // Store the pending result
      });
    }

    // --- STEP 2: SYNC TO FILL PLACEHOLDERS ---
    await context.sync();

    // --- STEP 3: PROCESS FILLED DATA ---
    var changesList = [];
    for (var i = 0; i < proxies.length; i++) {
      var item = proxies[i];
      var c = item.change;
      var rawText = item.range.text;

      // Extract the XML value now that sync is done
      var rawXml = item.xmlResult.value;

      // 1. DATA RESCUE: If text is empty (Simple Mode), parse XML
      if (!rawText || rawText.trim().length === 0) {
        if (rawXml) {
          var recovered = extractTextFromOoxml(rawXml);
          if (recovered && recovered.trim().length > 0) {
            rawText = recovered;
          }
        }

        // If still empty after XML rescue, handle accordingly
        if (!rawText || rawText.trim().length === 0) {
          if (c.type === "Deletion") {
            rawText = "[Hidden Deletion]";
            hiddenCount++;
          } else {
            continue; // Skip pure formatting noise
          }
        }
      }

      // Normalize
      rawText = rawText.replace(/[\r\n]+/g, " ").trim();
      if (rawText.length > 300) rawText = rawText.substring(0, 300) + "...";

      var contextText = "";
      try { contextText = item.para ? item.para.text : ""; } catch (e) { }
      if (contextText) {
        contextText = contextText.replace(/[\r\n]+/g, " ").trim();
        if (contextText.length > 500) contextText = contextText.substring(0, 500) + "...";
      }

      changesList.push({
        id: i,
        type: c.type,
        text: rawText,
        context: contextText || "Context unavailable."
      });
    }

    return changesList;
  })
    .then(function (cleanData) {
      if (cleanData === "NO_CHANGES" || cleanData.length === 0) {
        setLoading(false);
        showCompareInterface();
        showToast("⚠️ No text changes found.");
        return;
      }

      // If we had to fallback to "[Hidden Deletion]" placeholders
      if (hiddenCount > 0) {
        showOutput(`<div class="placeholder" style="background:#fff3e0; border:1px solid #ffe0b2; color:#e65100; margin-bottom:15px; font-size:11px;">
                    <strong>⚠️ View Mode Warning</strong><br>
                    ${hiddenCount} items were hidden by Simple Markup.<br>For best accuracy, switch to <b>All Markup</b>.
                </div>`);
      }

      window.redlineSessionData = { raw: cleanData };

      var payload = cleanData.length > 1000 ? cleanData.slice(0, 1000) : cleanData;
      return callGemini(API_KEY, PROMPTS.SUMMARIZE_CHANGES(payload));
    })
    .then(function (aiResult) {
      setLoading(false);
      if (!aiResult) return;
      try {
        var analysis = tryParseGeminiJSON(aiResult);
        window.redlineSessionData.analysis = analysis;
        renderRedlineDashboard();
      } catch (e) {
        showToast("Error parsing AI response");
        showCompareInterface();
      }
      if (btn) btn.disabled = false;
    })
    .catch(function (e) {
      setLoading(false);
      console.error("Redline Error:", e);
      if (e.message.indexOf("GeneralException") === -1) {
        showOutput("Error: " + e.message, true);
      }
      if (btn) btn.disabled = false;
    });
}

// ==========================================
// HELPER: Improved OOXML Parser
// ==========================================
function extractTextFromOoxml(xmlString) {
  if (!xmlString) return "";
  // Matches: <w:t> (Text), <w:delText> (Deleted Text)
  // We also strip the tags to get clean content
  var matches = xmlString.match(/<w:(?:t|delText)[^>]*>(.*?)<\/w:(?:t|delText)>/g);
  if (!matches) return "";

  return matches.map(function (tag) {
    return tag.replace(/<[^>]+>/g, "");
  }).join("");
}
// ==========================================
// REPLACE: renderRedlineDashboard (Squared Locate + Collapsible Header)
// ==========================================

// 1. Define Toggle Helpers Globally
window.toggleRedlineCategory = function (id) {
  var el = document.getElementById(id);
  var arrow = document.getElementById("arrow_" + id);
  if (el) {
    if (el.style.display === "none") {
      el.style.display = "block";
      if (arrow) arrow.innerHTML = "▼";
    } else {
      el.style.display = "none";
      if (arrow) arrow.innerHTML = "▶";
    }
  }
};

window.toggleAllRedlineGroups = function (show) {
  var contents = document.querySelectorAll('.redline-group-content');
  var arrows = document.querySelectorAll('.redline-group-arrow');
  contents.forEach(el => el.style.display = show ? 'block' : 'none');
  arrows.forEach(el => el.innerHTML = show ? '▼' : '▶');
};

// NEW: Toggle the Main Analysis Box
window.toggleRedlineSummary = function () {
  var body = document.getElementById("redlineDashBody");
  var arrow = document.getElementById("redlineDashArrow");

  if (body.style.display === "none") {
    body.style.display = "block";
    arrow.innerHTML = "▼";
  } else {
    body.style.display = "none";
    arrow.innerHTML = "▶";
  }
};

window.renderRedlineDashboard = renderRedlineDashboard;

function renderRedlineDashboard(viewMode) {
  var out = document.getElementById("output");
  out.innerHTML = "";

  var data = window.redlineSessionData;

  if (!data || !data.analysis) {
    out.innerHTML = "<div class='placeholder'>⚠️ Analysis data unavailable. Please run the scan again.</div>";
    return;
  }

  var mode = viewMode || "category";
  var changesList = data.analysis.changes || [];
  var total = changesList.length;

  // --- SHARED PLATINUM BUTTON STYLE (SQUARED OFF) ---
  var platinumBtnStyle = `
                                display: inline-flex; align-items: center; justify-content: center; gap: 4px;
                                padding: 4px 10px;
                                height: 24px;
                                font-family: 'Inter', sans-serif;
                                font-size: 9px;
                                font-weight: 800;
                                text-transform: uppercase;
                                letter-spacing: 0.5px;
                                color: #1565c0;
                                background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
                                border: 1px solid #b0bec5;
                                border-bottom: 2px solid #90a4ae;
                                border-radius: 0px; /* SQUARED OFF */
                                cursor: pointer;
                                transition: all 0.2s;
                                box-shadow: 0 1px 2px rgba(0,0,0,0.05);
                                white-space: nowrap;
                                `;
  var hoverOn = "this.style.background='#e3f2fd'; this.style.borderColor='#1565c0'; this.style.transform='translateY(-1px)';";
  var hoverOff = "this.style.background='linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)'; this.style.borderColor='#b0bec5'; this.style.transform='translateY(0)';";

  // --- CALCULATE STATS ---
  var stats = { Legal: 0, Financial: 0, Formatting: 0, Operational: 0 };
  changesList.forEach(function (c) {
    var cat = c.category || "Other";
    stats[cat] = (stats[cat] || 0) + 1;
  });

  // --- HEADER DASHBOARD ---
  var badgesHtml = "";
  var colors = {
    Legal: "#c62828",
    Financial: "#2e7d32",
    Formatting: "#78909c",
    Operational: "#1565c0",
    Other: "#546e7a",
  };
  var headerOrder = ["Legal", "Financial", "Formatting", "Operational", "Other"];

  headerOrder.forEach(function (key) {
    if (stats[key] > 0) {
      var color = colors[key] || "#546e7a";
      badgesHtml += `<span style="font-size:11px; background:#fff; color:${color}; padding:3px 9px; border:1px solid ${color}; margin-right:5px; font-weight:700; margin-bottom:5px; display:inline-block; border-radius:0px;">${key}: ${stats[key]}</span>`;
    }
  });

  // --- EXPAND/COLLAPSE CONTROLS (Moved to footer) ---
  var showExpandCollapse = (mode === 'category' || mode === 'clause');

  // --- 1. RENDER ANALYSIS DASHBOARD (Collapsible) ---
  var dash = document.createElement("div");
  dash.className = "tool-box slide-in";
  dash.style.cssText = "padding:12px 16px; border-left:5px solid #1565c0; background:#f8f9fa; margin-bottom:15px; box-shadow:0 2px 5px rgba(0,0,0,0.05); border-radius:0px; font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;";

  dash.innerHTML = `
                                <div onclick="window.toggleRedlineSummary()" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;">
                                  <div style="display:flex; align-items:center; gap:8px;">
                                    <div id="redlineDashArrow" style="font-size:10px; color:#1565c0;">▼</div>
                                    <div style="font-size:13px; font-weight:800; color:#1565c0; text-transform:uppercase; letter-spacing:0.5px;">REDLINE ANALYSIS</div>
                                  </div>
                                  <div style="font-size:11px; font-weight:700; background:#1565c0; color:white; padding:3px 10px; border-radius:0px;">${total} Changes</div>
                                </div>

                                <div id="redlineDashBody" style="display:block; margin-top:12px; padding-top:10px; border-top:1px solid #e0e0e0;">
                                  <div style="font-size:13px; color:#37474f; line-height:1.5; margin-bottom:12px;">${data.analysis.overview || "Analysis complete."}</div>
                                  <div style="display:flex; flex-wrap:wrap;">${badgesHtml}</div>

                                </div>
                                `;
  out.appendChild(dash);

  // --- 2. RENDER VIEW TOGGLE (Squared Off & External) ---
  var segContainer = document.createElement("div");
  segContainer.style.cssText = "display:flex; justify-content:center; margin-bottom:20px; position:relative;";

  // Style for Active vs Inactive Buttons
  var btnBase = "padding:5px 16px; font-size:10px; font-weight:700; cursor:pointer; min-width:70px; text-align:center; transition:all 0.2s; border-radius:0px;";
  var btnActive = "background:white; color:#1565c0; box-shadow:0 1px 3px rgba(0,0,0,0.15);";
  var btnInactive = "color:#546e7a; background:transparent;";

  segContainer.innerHTML = `
                                <div style="display:flex; background:#e0e0e0; padding:2px; border:1px solid #cfd8dc; border-radius:0px;">
                                  <div onclick="renderRedlineDashboard('category')" style="${btnBase} ${mode === 'category' ? btnActive : btnInactive}">CATEGORY</div>
                                  <div onclick="renderRedlineDashboard('clause')" style="${btnBase} ${mode === 'clause' ? btnActive : btnInactive} border-left:1px solid #ccc; border-right:1px solid #ccc;">CLAUSE</div>
                                  <div onclick="renderRedlineDashboard('sequential')" style="${btnBase} ${mode === 'sequential' ? btnActive : btnInactive}">LIST</div>
                                </div>
                                <div style="position:absolute; right:0; top:50%; transform:translateY(-50%); display:${showExpandCollapse ? 'flex' : 'none'}; gap:10px; font-size:9px; font-weight:700; color:#1565c0;">
                                  <span onclick="window.toggleAllRedlineGroups(true)" style="cursor:pointer; text-decoration:underline; white-space:nowrap;">EXPAND ALL</span>
                                  <span onclick="window.toggleAllRedlineGroups(false)" style="cursor:pointer; text-decoration:underline; white-space:nowrap;">COLLAPSE ALL</span>
                                </div>
                                `;
  out.appendChild(segContainer);

  // --- HELPER FUNCTIONS ---

  function getRaw(id) {
    if (!data.raw) return { text: "??", context: "", type: "Unknown" };
    return data.raw.find((r) => String(r.id) === String(id)) || {
      text: "(Text reference lost)",
      context: "",
      type: "Unknown",
    };
  }

  function checkIsDeletion(rawItem, aiItem) {
    if (rawItem && rawItem.type) {
      var t = rawItem.type.toLowerCase();
      return (t === "delete" || t === "deletion" || t === "deleted");
    }
    if (aiItem && aiItem.summary) {
      var s = aiItem.summary.toLowerCase();
      return (s.includes("delete") || s.includes("remove") || s.includes("struck"));
    }
    return false;
  }

  function identifyClause(raw) {
    if (!raw || !raw.context) return "General / Unnumbered";
    var txt = raw.context.trim();
    var match = txt.match(/^((?:Section|Article|Clause)?\s*[\dIVX]+(?:\.[\d]+)*\.?|\([a-z0-9]+\))\s+([^\.\:\n]+)/i);
    if (match) {
      var header = (match[1] + " " + match[2]).trim();
      header = header.replace(/[:.]$/, "");
      if (header.length > 60) return header.substring(0, 60) + "...";
      return header;
    }
    var colonMatch = txt.match(/^([^:\n]{3, 50}):/);
    if (colonMatch) return colonMatch[1];
    return "General / Unnumbered";
  }

  function getVisualContext(raw, isDel) {
    var rawContext = raw.context || "";
    var rawText = raw.text || "";

    var renderSimple = (lbl, txt, ctx) => {
      var base = isDel
        ? `<div style="color:#c62828; font-size:13px; line-height:1.5;"><strong>🗑️ ${lbl}:</strong> <span style="text-decoration:line-through; background-color:#ffebee;">${txt}</span></div>`
        : `<div style="color:#2e7d32; font-size:13px; line-height:1.5;"><strong>✏️ ${lbl}:</strong> <span style="background-color:#e8f5e9;">${txt}</span></div>`;
      if (ctx) base += `<div style="margin-top:6px; color:#546e7a; font-size:12px; font-style:italic; line-height:1.4;">"...${ctx.substring(0, 140)}..."</div>`;
      return base;
    };

    if (!rawContext || !rawText) return renderSimple(isDel ? "Deleted" : "Inserted", rawText, "");

    try {
      var escaped = rawText.replace(/[.*+?^${ }()|[\]\\]/g, '\\$&');
      var re = new RegExp(escaped, 'gi');
      if (re.test(rawContext)) {
        var styledContext = rawContext.replace(re, function (match) {
          if (isDel) return `<span style="text-decoration:line-through; color:#c62828; background-color:#ffebee; font-weight:700;">${match}</span>`;
          return `<span style="text-decoration:underline; color:#2e7d32; background-color:#e8f5e9; font-weight:700;">${match}</span>`;
        });
        return `<div style="font-size:13px; color:#37474f; line-height:1.5;">"...${styledContext}..."</div>`;
      }
    } catch (e) { console.warn("Context highlight failed", e); }
    return renderSimple(isDel ? "Deleted" : "Inserted", rawText, rawContext);
  }

  // --- CARD RENDERER ---
  function createCardHTML(item, raw, catColor, indexLabel) {
    var isDel = checkIsDeletion(raw, item);
    var icon = isDel ? "🗑️" : "✏️";
    var visualHtml = getVisualContext(raw, isDel);
    var showBadges = (typeof userSettings !== "undefined" && userSettings.showCompareRisk === true);
    var sev = (item.severity || "Low").toLowerCase();
    var sevColor = "#2e7d32"; var sevBg = "#e8f5e9";
    if (sev.includes("critical") || sev.includes("high")) { sevColor = "#c62828"; sevBg = "#ffebee"; }
    else if (sev.includes("medium")) { sevColor = "#e65100"; sevBg = "#fff3e0"; }
    var badgeHtml = showBadges ? `<div style="font-size:10px; font-weight:800; color:${sevColor}; background:${sevBg}; padding:2px 8px; text-transform:uppercase; border:1px solid ${sevColor}; margin-left:auto;">${item.severity || "Info"}</div>` : '';

    return `
                                <div style="padding:8px 12px; background:#fcfcfc; border-bottom:1px solid #f0f0f0; display:flex; justify-content:flex-start; align-items:center;">
                                  ${indexLabel ? `<span style="font-size:10px; background:#eee; padding:2px 6px; margin-right:8px; font-weight:bold;">#${indexLabel}</span>` : ''}
                                  <div style="font-size:11px; font-weight:800; color:${catColor}; text-transform:uppercase; font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">${icon} ${item.category}</div>
                                  ${badgeHtml}
                                </div>
                                <div style="padding:12px;">
                                  <div style="font-family:'Georgia', serif; margin-bottom:12px; background:#f9fafb; padding:10px; border:1px solid #e0e0e0; border-radius:0px;">${visualHtml}</div>
                                  <div style="font-size:14px; color:#212529; font-weight:600; margin-bottom:6px; line-height:1.4; font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">${item.summary || item.explanation}</div>
                                  ${item.legal_impact ? `<div style="font-size:12px; color:#546e7a; background:#f5f9ff; padding:8px; border-left:3px solid #1565c0; line-height:1.4; margin-bottom:8px;">${item.legal_impact}</div>` : ''}

                                  <div style="text-align:right; margin-top:8px;">
                                    <button style="${platinumBtnStyle}"
                                      onmouseover="${hoverOn}" onmouseout="${hoverOff}"
                                      onclick="locateTrackedChange(${item.id})">
                                      <span>📍</span> LOCATE CHANGE
                                    </button>
                                  </div>
                                </div>
                                `;
  }
  if (mode === "category") {
    var groups = {};
    changesList.forEach((c) => {
      var cat = c.category || "Uncategorized";
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(c);
    });
    var sortOrder = ["Legal", "Financial", "Formatting", "Operational", "Other"];
    var sortedKeys = Object.keys(groups).sort((a, b) => {
      var idxA = sortOrder.indexOf(a); if (idxA === -1) idxA = 99;
      var idxB = sortOrder.indexOf(b); if (idxB === -1) idxB = 99;
      return idxA - idxB;
    });

    sortedKeys.forEach(function (cat) {
      var catColor = colors[cat] || "#333";
      var groupId = "cat_" + cat.replace(/\s/g, "_");
      var groupContainer = document.createElement("div");
      groupContainer.style.cssText = "margin-bottom:12px; border:1px solid #e0e0e0; background:white;";

      groupContainer.innerHTML = `
                                <div onclick="window.toggleRedlineCategory('${groupId}')" style="cursor:pointer; display:flex; align-items:center; background:#f5f5f5; padding:10px 15px; border-bottom:1px solid #e0e0e0; border-left:4px solid ${catColor};">
                                  <div id="arrow_${groupId}" class="redline-group-arrow" style="font-size:12px; color:#666; margin-right:10px;">▼</div>
                                  <div style="font-size:13px; font-weight:800; color:${catColor}; text-transform:uppercase; flex:1;">${cat}</div>
                                  <div style="font-size:11px; font-weight:700; color:#555; background:#e0e0e0; padding:2px 8px;">${groups[cat].length}</div>
                                </div>
                                `;
      var contentDiv = document.createElement("div");
      contentDiv.id = groupId;
      contentDiv.className = "redline-group-content";
      contentDiv.style.display = "block";
      contentDiv.style.padding = "10px";
      contentDiv.style.background = "#eeeeee";

      groups[cat].forEach((c) => {
        var raw = getRaw(c.id);
        var card = document.createElement("div");
        card.className = "response-card slide-in";
        card.style.cssText = "background:white; border:1px solid #e0e0e0; border-left:4px solid " + catColor + "; margin-bottom:10px;";
        card.innerHTML = createCardHTML(c, raw, catColor);
        contentDiv.appendChild(card);
      });
      groupContainer.appendChild(contentDiv);
      out.appendChild(groupContainer);
    });
  }
  else if (mode === "clause") {
    var clauseGroups = {};
    changesList.forEach((c) => {
      var clauseName = "General / Unnumbered";
      if (c.clause && c.clause !== "General" && c.clause.length > 3) {
        clauseName = c.clause;
      } else {
        var raw = getRaw(c.id);
        clauseName = identifyClause(raw);
      }
      if (!clauseGroups[clauseName]) clauseGroups[clauseName] = [];
      clauseGroups[clauseName].push(c);
    });

    var sortedClauses = Object.keys(clauseGroups).sort((a, b) => {
      return parseInt(clauseGroups[a][0].id) - parseInt(clauseGroups[b][0].id);
    });

    sortedClauses.forEach(function (clauseName) {
      var groupId = "cl_" + clauseName.replace(/[^a-zA-Z0-9]/g, "_");
      var groupContainer = document.createElement("div");
      groupContainer.style.cssText = "margin-bottom:12px; border:1px solid #e0e0e0; background:white;";
      var safeClauseName = clauseName.replace(/'/g, "\\'");

      // --- HEADER with SQUARED LOCATE BUTTON ---
      groupContainer.innerHTML = `
                                <div onclick="window.toggleRedlineCategory('${groupId}')" style="cursor:pointer; display:flex; align-items:center; background:#fff; padding:8px 12px; border-bottom:1px solid #e0e0e0; border-left:4px solid #1565c0;">

                                  <div id="arrow_${groupId}" class="redline-group-arrow" style="font-size:12px; color:#1565c0; margin-right:10px; width:12px;">▼</div>

                                  <div style="font-size:13px; font-weight:800; color:#1565c0; flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-right:10px;">
                                    ${clauseName}
                                  </div>

                                  <div style="display:flex; align-items:center; gap:8px;">

                                    <button style="${platinumBtnStyle}"
                                      onmouseover="${hoverOn}" onmouseout="${hoverOff}"
                                      onclick="event.stopPropagation(); locateText('${safeClauseName}', this, true)">
                                      <span>📍</span> LOCATE
                                    </button>

                                    <div style="font-size:10px; font-weight:700; color:white; background:#1565c0; padding:2px 8px; border-radius:10px; min-width:16px; text-align:center;">
                                      ${clauseGroups[clauseName].length}
                                    </div>
                                  </div>
                                </div>
                                `;

      var contentDiv = document.createElement("div");
      contentDiv.id = groupId;
      contentDiv.className = "redline-group-content";
      contentDiv.style.display = "block";
      contentDiv.style.padding = "10px";
      contentDiv.style.background = "#eeeeee";

      clauseGroups[clauseName].forEach((c) => {
        var raw = getRaw(c.id);
        var catColor = colors[c.category] || "#37474f";
        var card = document.createElement("div");
        card.className = "response-card slide-in";
        card.style.cssText = "background:white; border:1px solid #e0e0e0; border-left:4px solid " + catColor + "; margin-bottom:10px;";
        card.innerHTML = createCardHTML(c, raw, catColor);
        contentDiv.appendChild(card);
      });
      groupContainer.appendChild(contentDiv);
      out.appendChild(groupContainer);
    });
  }
  else {
    var sequentialList = changesList.slice().sort((a, b) => parseInt(a.id) - parseInt(b.id));
    sequentialList.forEach(function (item, idx) {
      var raw = getRaw(item.id);
      var catColor = colors[item.category] || "#37474f";
      var card = document.createElement("div");
      card.className = "response-card slide-in";
      card.style.cssText = "background:white; border:1px solid #e0e0e0; border-left:4px solid " + catColor + "; margin-bottom:15px;";
      card.innerHTML = createCardHTML(item, raw, catColor, idx + 1);
      out.appendChild(card);
    });
  }

  // Back Button
  var btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "⬅ Back to Smart Compare";
  btnBack.style.marginTop = "20px";
  btnBack.onclick = showCompareInterface;
  out.appendChild(btnBack);
}
window.locateTrackedChange = locateTrackedChange;
function locateTrackedChange(index) {
  Word.run(function (context) {
    var changes = context.document.body.getTrackedChanges();
    changes.load("items");
    return context.sync().then(function () {
      if (index < changes.items.length) {
        var target = changes.items[index];
        target.getRange().select();
        target.getRange().scrollIntoView();
        showToast("📍 Found Change #" + (index + 1));
      } else {
        showToast("⚠️ Change no longer found.");
      }
      return context.sync();
    });
  }).catch(function (e) {
    console.error("Locate Error:", e);
    showToast("⚠️ Could not locate.");
  });
}
// ==========================================
// END: locateTrackedChange
// ==========================================
// ==========================================
// START: runDraftFromFix (Handles the AI Call)
// ==========================================
function runDraftFromFix(fixSuggestion) {
  showToast("⚡ Drafting Clause...");

  // 1. Get Document Context to inform the draft
  validateApiAndGetText(true).then(function (contextText) {

    // 2. Construct Prompt using the PROMPTS object
    var prompt = PROMPTS.DRAFT(fixSuggestion, contextText.substring(0, 20000));

    // 3. Call AI
    return callGemini(API_KEY, prompt);
  })
    .then(function (jsonString) {
      try {
        var data = tryParseGeminiJSON(jsonString);
        // 4. Show the Dialog Box with the result
        showDraftingModal(data);
      } catch (e) {
        showToast("❌ Error parsing draft.");
      }
    })
    .catch(function (e) {
      showToast("❌ Error: " + e.message);
    });
}

// ==========================================
// START: showDraftingModal (The Dialog Box)
// ==========================================
function showDraftingModal(data) {
  // 1. Create Overlay
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  // 2. Create Modal Box
  var dialog = document.createElement("div");
  dialog.className = "tool-box slide-in";
  dialog.style.cssText = "background:white; width:320px; padding:0; border-radius:4px; box-shadow:0 10px 25px rgba(0,0,0,0.3); overflow:hidden; border:1px solid #ccc;";

  // 3. Content
  dialog.innerHTML = `
                                <div style="background:#1565c0; color:white; padding:12px 15px; font-weight:700; font-size:13px; display:flex; justify-content:space-between; align-items:center;">
                                  <span>⚖️ ${data.title || "Draft Clause"}</span>
                                  <span id="btnCloseDraft" style="cursor:pointer; font-size:16px; opacity:0.8;">✕</span>
                                </div>
                                <div style="padding:15px;">
                                  <div style="font-size:11px; color:#555; margin-bottom:10px;">
                                    <strong>Reasoning:</strong> ${data.explanation}
                                  </div>
                                  <div style="background:#f8f9fa; border:1px solid #ddd; padding:10px; font-family:'Georgia', serif; font-size:13px; color:#333; margin-bottom:15px; max-height:200px; overflow-y:auto; line-height:1.5;">
                                    ${data.text.replace(/\n/g, "<br>")}
                                  </div>
                                  <div style="display:flex; gap:10px;">
                                    <button id="btnInsertDraft" class="btn-insert-premium" style="flex:1;">📍 Insert at Location</button>
                                    <button id="btnCopyDraft" class="btn-action" style="flex:1;">📄 Copy</button>
                                  </div>
                                </div>
                                `;

  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  // --- BIND ACTIONS ---

  // Close
  document.getElementById("btnCloseDraft").onclick = function () { overlay.remove(); };

  // Copy
  document.getElementById("btnCopyDraft").onclick = function () {
    navigator.clipboard.writeText(data.text);
    this.innerText = "Copied! ✅";
    setTimeout(() => { this.innerText = "📄 Copy"; }, 1500);
  };

  // Insert (Triggers "Scroll to Location" sub-modal)
  document.getElementById("btnInsertDraft").onclick = function () {
    overlay.remove(); // Remove drafting modal

    // Show the Insertion Confirmation Modal
    var insertOverlay = document.createElement("div");
    insertOverlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

    var insertDialog = document.createElement("div");
    insertDialog.className = "tool-box slide-in";
    insertDialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #2e7d32; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

    insertDialog.innerHTML = `
                                <div style="font-size:24px; margin-bottom:10px;">📍</div>
                                <div style="font-weight:700; color:#2e7d32; margin-bottom:10px; font-size:14px;">Ready to Insert</div>
                                <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">
                                  1. Scroll to the exact location in the document.<br>
                                    2. Click where you want this clause to go.<br>
                                      3. Click the button below.
                                    </div>
                                    <button id="btnConfirmInsertNow" class="btn btn-blue" style="width:100%; background:#2e7d32;">I am there - Insert Here</button>
                                    <button id="btnCancelInsertNow" class="btn-action" style="width:100%; margin-top:8px;">Cancel</button>
                                    `;

    insertOverlay.appendChild(insertDialog);
    document.body.appendChild(insertOverlay);

    document.getElementById("btnCancelInsertNow").onclick = function () { insertOverlay.remove(); };

    document.getElementById("btnConfirmInsertNow").onclick = function () {
      var btn = this;
      btn.innerText = "Inserting...";
      Word.run(function (context) {
        var selection = context.document.getSelection();
        var insertedRange = selection.insertText(data.text, "Replace");
        return context.sync();
      }).then(function () {
        insertOverlay.remove();
        showToast("✅ Clause Inserted");
      }).catch(function (e) {
        console.error(e);
        btn.innerText = "Error";
      });
    };
  };
}
// ==========================================
// START: runRewriteInChat
// ==========================================
function runRewriteInChat(instruction) {
  // 1. UI Feedback
  addChatBubble("✏️ Rewriting: " + instruction, "user");
  var loaderId = addSkeletonBubble();

  // 2. Safety Check for Source Text
  // (pendingRewriteText is set when you click the 'Rewrite' tool button)
  if (typeof pendingRewriteText === "undefined" || !pendingRewriteText) {
    removeBubble(loaderId);
    addChatBubble("⚠️ Error: No text selected to rewrite. Please select text and click the 'Rewrite' button again.", "ai");
    return;
  }

  // 3. Call AI
  callGemini(API_KEY, PROMPTS.REWRITE(instruction, pendingRewriteText))
    .then(function (jsonString) {
      removeBubble(loaderId);
      try {
        var data = tryParseGeminiJSON(jsonString);
        // 4. Render Result
        renderRewriteBubble(data, pendingRewriteText);
      } catch (e) {
        console.error(e);
        addChatBubble("❌ Error parsing AI response.", "ai");
      }
    })
    .catch(function (e) {
      removeBubble(loaderId);
      addChatBubble("❌ Error: " + e.message, "ai");
    });
}
// ==========================================
// START: WORKFLOW MANAGER (PLAYLIST LOGIC)
// ==========================================
var currentWorkflowMode = "default";
function setWorkflowMode(mode, silent) {
  // 1. NORMALIZE: Treat 'all' as 'default' so it matches your saved settings
  if (mode === 'all') mode = 'default';

  currentWorkflowMode = mode;

  // 2. Initialize Default Maps if missing
  if (!userSettings.workflowMaps) {
    // Use the Global Constant we defined earlier
    userSettings.workflowMaps = JSON.parse(JSON.stringify(DEFAULT_WORKFLOW_MAPS));
  }

  // 3. Load the Playlist (Try Custom, then Constant, then Hard Fallback)
  var playlist = userSettings.workflowMaps[mode];

  if (!playlist || playlist.length === 0) {
    // Fallback to the Global Constant source of truth
    if (DEFAULT_WORKFLOW_MAPS[mode]) {
      playlist = DEFAULT_WORKFLOW_MAPS[mode].slice();
    } else {
      // Emergency Fallback
      playlist = ["error", "legal", "summary", "chat", "playbook", "compare", "template", "email"];
    }
  }

  // 4. Apply to Active Tool Order
  userSettings.toolOrder = playlist;
  userSettings.currentWorkflow = mode; // Remember this mode for next reload

  // 5. Update Visibility & Pages based on new order
  // This ensures newly added apps appear on the correct page immediately
  var newPages = {};
  playlist.forEach(function (key, index) {
    newPages[key] = (index < 4) ? 1 : 2;
    userSettings.visibleButtons[key] = true;
  });

  // Merge, don't overwrite (keeps custom page settings for other tools safe)
  userSettings.buttonPages = Object.assign({}, userSettings.buttonPages, newPages);

  // 6. Render & Save
  saveCurrentSettings(true); // Silent save
  renderToolbar();
  showHome();
  updateFooterVisuals(mode);

  if (!silent) {
    var label = (mode === 'default') ? "Standard" : mode.charAt(0).toUpperCase() + mode.slice(1);
    showToast(label + " Workflow Loaded");
  }
}


function updateFooterVisuals(mode) {
  // Reset All
  ['footSegDraft', 'footSegAll', 'footSegReview'].forEach(id => {
    var el = document.getElementById(id);
    if (el) {
      el.style.background = "transparent";
      el.style.color = "#78909c";
      el.style.fontWeight = "600";
    }
  });

  // Set Active
  var activeId = 'footSegAll'; // Default
  if (mode === 'draft') activeId = 'footSegDraft';
  if (mode === 'review') activeId = 'footSegReview';

  var activeEl = document.getElementById(activeId);
  if (activeEl) {
    activeEl.style.background = "#fff";
    activeEl.style.color = "#455a64";
    activeEl.style.fontWeight = "800";
    activeEl.style.boxShadow = "0 1px 2px rgba(0,0,0,0.1)";
  }
}

// ==========================================
// 🚀 MASTER MODULE: HELP CENTER, AI CHAT & TOOLTIPS (FIXED)
// ==========================================

window.switchHelpTab = function (tabId) {
  // A. Visuals: Reset Buttons
  document.querySelectorAll(".help-tab").forEach(t => {
    t.classList.remove("active");
    t.style.borderBottom = "none";
    t.style.color = "#546e7a";
    t.style.fontWeight = "400";
    t.style.background = "transparent";
  });

  // B. Activate Selected
  var btn = document.getElementById("btn-tab-" + tabId);
  if (btn) {
    btn.classList.add("active");
    btn.style.borderBottom = "3px solid #1565c0";
    btn.style.color = "#1565c0";
    btn.style.fontWeight = "700";
    btn.style.background = "#fff";
  }

  // C. Content Switching
  document.querySelectorAll(".help-body").forEach(b => {
    b.style.display = "none";
  });

  var body = document.getElementById("tab-" + tabId);
  if (body) {
    // Chat needs flex layout, others need block
    body.style.display = (tabId === "bot") ? "flex" : "block";
  }
};

// 2. HELPER: Compact Card (Quick Start) - RESIZED LARGER
function renderCompactCard(title, text) {
  return `
                                    <div style="
            background: #ffffff; 
            border: 1px solid #cfd8dc; 
            padding: 12px; /* Increased from 8px */
            border-left: 6px solid #1565c0; /* Thicker accent bar */
            padding: 8px; 
            border-left: 4px solid #1565c0; 
            display: flex; 
            flex-direction: column; 
            justify-content: center; 
            min-height: 90px; /* Forces cards to be taller */
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            min-height: 60px; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        ">
                                      <div style="
                font-weight: 800; 
                color: #1565c0; /* Changed to Blue for better visibility */
                font-size: 10px; /* Increased from 10px */
                margin-bottom: 8px; 
                color: #1565c0; 
                font-size: 9px; 
                margin-bottom: 4px; 
                text-transform: uppercase;
                letter-spacing: 0.5px;
            ">${title}</div>

                                      <div style="
                font-size: 12px; /* Increased from 9px */
                font-size: 10px; 
                color: #455a64; 
                line-height: 1.5; 
                line-height: 1.3; 
            ">${text}</div>
                                    </div>
                                    `;
}

// 3. HELPER: FAQ Item (Restored)
function renderFaqItem(q, a) {
  return `
                                    <div class="faq-item" onclick="this.classList.toggle('open')" style="margin-bottom:8px; border:1px solid #eee; border-radius:0px; cursor:pointer; background:#fff;">
                                      <div class="faq-question" style="padding:10px; font-weight:700; font-size:11px; color:#37474f; display:flex; justify-content:space-between;">
                                        <span>${q}</span> <span style="color:#1565c0;">+</span>
                                      </div>
                                      <div class="faq-answer" style="padding:0 10px 10px 10px; font-size:11px; color:#546e7a; line-height:1.4; display:none;">
                                        ${a}
                                      </div>
                                    </div>
                                    `;
}

// 4. MAIN FUNCTION: SHOW HELP CENTER
function showHelpCenter() {
  var old = document.getElementById("eliHelpOverlay");
  if (old) old.remove();

  var overlay = document.createElement("div");
  overlay.id = "eliHelpOverlay";
  overlay.className = "help-overlay";

  var modal = document.createElement("div");
  modal.className = "help-modal";
  modal.style.cssText = "border-radius: 0px; box-shadow: 0 20px 60px rgba(0,0,0,0.4); border: 1px solid #b0bec5; overflow: hidden; height: 85vh; max-height: 600px; display: flex; flex-direction: column; background: white;";

  // --- HEADER ---
  var headerHtml = `
                                    <div class="help-header" style="
            background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); 
            padding: 15px 20px; 
            border-bottom: 3px solid #c9a227; 
            display: flex; justify-content: space-between; align-items: center;
            flex-shrink: 0;
        ">
                                      <div style="display: flex; align-items: center; gap: 12px;">
                                        <div style="
                    background: rgba(255,255,255,0.1); 
                    border: 1px solid rgba(255,255,255,0.3); 
                    color: #fff; 
                    font-family: 'Georgia', serif; 
                    font-weight: 700; 
                    font-size: 18px; 
                    padding: 4px 8px; 
                    letter-spacing: 1px;
                ">ELI</div>
                                        <div>
                                          <div style="font-size: 14px; color: #fff; font-weight: 700; letter-spacing: 0.5px;">KNOWLEDGE BASE</div>

                                        </div>
                                      </div>
                                      <div id="btnCloseHelp" style="cursor: pointer; color: #90a4ae; font-size: 18px; line-height: 1;"
                                        onmouseover="this.style.color='#fff'" onmouseout="this.style.color='#90a4ae'">✕</div>
                                    </div>
                                    `;

  // --- TABS ---
  var tabsHtml = `
                                    <div class="help-tabs" style="background: #f1f3f4; border-bottom: 1px solid #e0e0e0; flex-shrink: 0; display:flex; padding:0 10px;">
                                      <div id="btn-tab-tut" class="help-tab active" onclick="switchHelpTab('tut')" style="padding:12px 15px; cursor:pointer; font-size:11px; font-weight:700; color:#1565c0; border-bottom:3px solid #1565c0; background:#fff;">🚀 Quick Start</div>
                                      <div id="btn-tab-faq" class="help-tab" onclick="switchHelpTab('faq')" style="padding:12px 15px; cursor:pointer; font-size:11px; color:#546e7a;">❓ FAQ</div>
                                      <div id="btn-tab-bot" class="help-tab" onclick="switchHelpTab('bot')" style="padding:12px 15px; cursor:pointer; font-size:11px; color:#546e7a;">🤖 Ask Support</div>
                                      <div id="btn-tab-about" class="help-tab" onclick="switchHelpTab('about')" style="padding:12px 15px; cursor:pointer; font-size:11px; color:#546e7a;">👤 About</div>
                                    </div>
                                    `;

  // --- BODY ---
  var bodyHtml = `
                                    <div id="help-content-area" style="flex: 1; overflow-y: auto; position: relative;">

                                      <div id="tab-tut" class="help-body" style="padding: 20px; display: block;">
                                        <div style="text-align:center; margin-bottom:15px;">
                                          <div style="font-size:14px; font-weight:700; color:#1565c0;">Welcome to ELI</div>
                                          <div style="font-size:11px; color:#546e7a;">AI-powered contract review & drafting companion.</div>
                                        </div>
                                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                                          ${renderCompactCard("📝 Error Check", "Scans for typos, broken refs, and definitions.")}
                                          ${renderCompactCard("⚖️ Legal Assist", "Flags high-risk clauses (Indemnity, Liability).")}
                                          ${renderCompactCard("📋 Summary", "Extracts key deal points (Fees, Dates) to grid.")}
                                          ${renderCompactCard("📖 Playbook", "Checks against Colgate's playbook and provides fallbacks.")}
                                          ${renderCompactCard("💬 Ask ELI", "Chat with doc. Rewrite, Defend, Translate.")}
                                          ${renderCompactCard("↔️ Compare", "Redline analysis against uploaded files.")}
                                          ${renderCompactCard("⚡ Build Draft", "Build a draft from a template and fill in placeholders.")}
                                          ${renderCompactCard("✉️ Email Tool", "Drafts Outlook-ready cover emails.")}
                                        </div>
                                      </div>

                                      <div id="tab-faq" class="help-body" style="padding: 20px; display: none;">
                                        ${renderFaqItem("Is my data safe?", "Yes. ELI uses a stateless API. No data is stored on external servers.")}
                                        ${renderFaqItem("How do I change the language?", "Go to Settings (Gear Icon) > Global > Language.")}
                                        ${renderFaqItem("What is 'Laptop Mode'?", "It pins the toolbar and shrinks icons for smaller screens.")}
                                        ${renderFaqItem("Can I add my own clauses?", "Yes! Go to Playbook > Settings (Gear) > Local Database.")}
                                      </div>

                                      <div id="tab-bot" class="help-body" style="height: 100%; display: none; flex-direction: column;">
                                        <div class="support-chat-area" style="flex:1; overflow-y:auto; padding:15px; background:#f9fafb;">
                                          <div class="support-messages" id="supportChatHistory">
                                            <div class="chat-bubble chat-ai">
                                              <strong>🤖 ELI Support:</strong> Hi! I'm trained on the user manual. Ask me anything!
                                            </div>
                                          </div>
                                        </div>
                                        <div class="support-input-area" style="padding:15px; border-top:1px solid #e0e0e0; background:#fff; display:flex; gap:8px; flex-shrink: 0;">
                                          <input type="text" id="supportInput" class="tool-input" style="flex:1; margin-bottom:0; font-size:12px; height:36px;" placeholder="E.g. How do I run a comparison?">
                                            <button id="btnSupportSend" class="btn-action" style="width:auto; padding:0 20px; background:#1565c0; color:white; border:none; font-weight:700; border-radius:0px;">SEND</button>
                                        </div>
                                      </div>

                                      <div id="tab-about" class="help-body" style="display: none; height:100%; overflow-y:auto; background:#fff;">

                                        <div style="text-align: center; padding: 40px 20px 30px 20px;">
                                          <div style="
                        width: 60px; height: 60px; margin: 0 auto 15px auto;
                        background: linear-gradient(135deg, #1565c0 0%, #0d47a1 100%);
                        border-radius: 0px; display: flex; align-items: center; justify-content: center;
                        color: white; font-family: 'Georgia', serif; font-size: 24px; font-weight: 700;
                        box-shadow: 0 10px 20px rgba(21, 101, 192, 0.25);
                    ">ELI</div>

                                          <div style="font-family: 'Segoe UI', sans-serif; font-size: 20px; font-weight: 700; color: #263238; margin-bottom: 5px;">
                                            Enhanced Legal Intelligence
                                          </div>
                                          <div style="font-size: 12px; color: #78909c; font-weight: 500; text-transform: uppercase; letter-spacing: 1px;">
                                            Legal  Edition
                                          </div>
                                        </div>

                                        <div style="padding: 0 30px; margin-bottom: 30px; text-align: center;">
                                          <div style="font-size: 13px; line-height: 1.6; color: #546e7a; margin-bottom: 20px;">
                                            A secure, AI-powered add-in designed to streamline contract review, negotiation, and drafting workflows.
                                          </div>


                                        </div>

                                        <div style="height: 1px; background: #e0e0e0; margin: 0 20px 20px 20px;"></div>

                                        <div style="padding: 0 30px 30px 30px;">
                                          <div style="font-size: 10px; font-weight: 800; color: #b0bec5; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">
                                            Developer
                                          </div>

                                          <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 15px;">
                                            <div style="
                            width: 36px; height: 36px; background: #e3f2fd; color: #1565c0; 
                            border-radius: 50%; display: flex; align-items: center; justify-content: center; 
                            font-size: 12px; font-weight: 700; border: 1px solid #bbdefb;
                        ">AN</div>
                                            <div>
                                              <div style="font-size: 13px; font-weight: 700; color: #37474f;">Andrew Nemiroff</div>
                                              <div style="font-size: 11px; color: #78909c;">Chief External Innovation Counsel</div>
                                            </div>
                                          </div>

                                          <button id="btnOpenContact" class="btn-action" style="width: 100%; justify-content: center; border-radius: 4px;">
                                            <span>✉️</span> Contact Developer
                                          </button>
                                        </div>
                                      </div>

                                    </div>
                                    `;

  modal.innerHTML = headerHtml + tabsHtml + bodyHtml;
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // --- EVENTS ---

  // Close
  document.getElementById("btnCloseHelp").onclick = function () { overlay.remove(); };

  // Chat Logic
  var btnSend = document.getElementById("btnSupportSend");
  var input = document.getElementById("supportInput");
  if (btnSend && input) {
    var send = () => handleSupportChat();
    btnSend.onclick = send;
    input.onkeydown = (e) => { if (e.key === "Enter") send(); };
  }

  // FAQ Toggle Logic
  setTimeout(() => {
    var faqs = document.querySelectorAll(".faq-item");
    faqs.forEach(item => {
      item.onclick = function () {
        var ans = this.querySelector(".faq-answer");
        if (ans.style.display === "block") {
          ans.style.display = "none";
          this.querySelector("span:last-child").innerText = "+";
        } else {
          ans.style.display = "block";
          this.querySelector("span:last-child").innerText = "−";
        }
      };
    });
  }, 100);

  // --- EMAIL LOGIC (Uses Gmail Link - WORKING) ---
  var contactBtn = document.getElementById("btnOpenContact");
  if (contactBtn) {
    contactBtn.onclick = function () {
      var email = "andrew_nemiroff@colpal.com";
      var subject = "ELI Tool Feedback";
      var body = "Hi Andrew,\n\nI have some feedback regarding the ELI tool:\n\n";

      var gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${email}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

      if (Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(gmailUrl);
      } else {
        window.open(gmailUrl, "_blank");
      }
    };
  }
}
// ==========================================
// START: parseRobustSupportResponse (Dedicated for Knowledge Base)
// ==========================================
function parseRobustSupportResponse(responseString) {
  if (!responseString) return { answer: "" };

  // 1. Try to Parse as JSON (in case AI follows instructions perfectly)
  try {
    // Basic cleanup (Markdown wrappers)
    var clean = responseString.replace(/```json/g, "").replace(/```/g, "").trim();

    // Attempt parse
    var parsed = JSON.parse(clean);

    // If it's an object, try to find the best text field to display
    if (typeof parsed === 'object' && parsed !== null) {
      return {
        answer: parsed.answer || parsed.content || parsed.response || JSON.stringify(parsed)
      };
    }
    return { answer: String(parsed) };

  } catch (e) {
    // 2. FALLBACK: If Parsing Fails, Treat as Plain Text
    // This is the "Ask ELI" behavior: just show what the AI said.
    return {
      answer: responseString,
      isFallback: true
    };
  }
}
// ==========================================
// START: handleSupportChat (Updated with Robust Parser)
// ==========================================
var supportHistory = [];

function handleSupportChat() {
  var input = document.getElementById("supportInput");
  var container = document.getElementById("supportChatHistory");
  var question = input.value.trim();
  if (!question) return;

  // 1. Add User Msg
  container.innerHTML += `<div class="chat-bubble chat-user" style="align-self:flex-end; background:#e3f2fd; color:#1565c0;">${question}</div>`;
  input.value = "";

  // Scroll to bottom
  setTimeout(() => { container.scrollTop = container.scrollHeight; }, 10);

  // 2. Loading UI
  var loaderId = "sup-load-" + Date.now();
  container.innerHTML += `<div id="${loaderId}" class="chat-bubble chat-ai" style="align-self:flex-start;">Thinking...</div>`;

  // 3. Prepare Context
  var knowledgeString = JSON.stringify(ELI_KNOWLEDGE_BASE, null, 2);

  var historyContext = "";
  if (supportHistory.length > 0) {
    var recent = supportHistory.slice(-6);
    historyContext = "PREVIOUS CHAT:\n" + recent.map(function (item) {
      return (item.role === "user" ? "User: " : "AI: ") + item.content;
    }).join("\n") + "\n";
  }

  var systemPrompt = `
                                    You are the Expert Support Bot for "ELI" (Colgate's Legal AI Tool).

                                    KNOWLEDGE BASE (Prioritize this info):
                                    ${knowledgeString}

                                    ${historyContext}

                                    INSTRUCTIONS:
                                    1. FIRST, check the Knowledge Base. If the answer is there, use it.
                                    2. IF NOT in the Knowledge Base, you MAY use general legal/tech knowledge to be helpful.
                                    3. Write a clear, natural paragraph.
                                    4. Use **bold** for buttons. Use bullet points for steps.

                                    USER QUESTION: "${question}"
                                    `;

  // 4. Call API
  callGemini(API_KEY, systemPrompt)
    .then(function (rawResponse) {
      // --- USE NEW DEDICATED PARSER ---
      var data = parseRobustSupportResponse(rawResponse);
      var text = data.answer;

      // --- FORMATTING (Convert Markdown to HTML) ---
      if (typeof text === 'string') {
        // Remove code blocks if any remain
        text = text.replace(/```json/g, "").replace(/```/g, "").trim();
        if (text.startsWith('"') && text.endsWith('"')) text = text.slice(1, -1);

        // Bold: **text** -> <b>text</b>
        text = text.replace(/\*\*(.*?)\*\*/g, "<b>$1</b>");

        // Italic: *text* -> <i>text</i>
        text = text.replace(/\*(.*?)\*/g, "<i>$1</i>");

        // Lists: * Item -> • Item (Clean Bullets)
        text = text.replace(/^\s*[\*\-]\s+(.*)$/gm, "• $1<br>");

        // Newlines: \n -> <br>
        text = text.replace(/\\n/g, "<br>").replace(/\n/g, "<br>");
      }

      // 5. Update UI
      var l = document.getElementById(loaderId);
      if (l) l.remove();

      container.innerHTML += `<div class="chat-bubble chat-ai" style="align-self:flex-start; line-height:1.5;">${text}</div>`;

      setTimeout(() => { container.scrollTop = container.scrollHeight; }, 10);

      // 6. Save History
      supportHistory.push({ role: "user", content: question });
      supportHistory.push({ role: "ai", content: text });

    })
    .catch(function (e) {
      var l = document.getElementById(loaderId);
      if (l) l.remove();
      container.innerHTML += `<div class="chat-bubble chat-ai" style="color:red;">Error: ${e.message}</div>`;
    });
}
// ==========================================
// END: handleSupportChat
// ==========================================
// 6. FUNCTION: CREATE TOOLTIP (Calls the correct Help function)
function createSmartTooltip(e, tool) {
  if (e && e.stopPropagation) e.stopPropagation();

  var old = document.getElementById("activeTooltip");
  if (old) old.remove();

  var tip = document.createElement("div");
  tip.id = "activeTooltip";
  tip.className = "toolbar-tooltip";

  // 1. Premium Styles (Container)
  tip.style.cssText = `
                                          position: fixed;
                                          z-index: 2147483647;
                                          pointer-events: auto;
                                          width: 320px;
                                          padding: 0;
                                          box-sizing: border-box;
                                          overflow: hidden;
                                          border-radius: 0px;
                                          border: 1px solid #b0bec5;
                                          box-shadow: 0 15px 40px rgba(0,0,0,0.3);
                                          background: white;
                                          font-family: 'Inter', sans-serif;
                                          animation: fadeIn 0.2s ease;
                                          display: block;
                                          `;

  // 2. Content
  tip.innerHTML = `
                                          <div style="
            background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); 
            padding: 20px 24px; 
            border-bottom: 3px solid #c9a227; 
            position: relative;
        ">
                                            <div id="btnCloseTip" style="
                position: absolute; top: 12px; right: 12px; 
                cursor: pointer; color: #90a4ae; font-weight: bold; font-size: 16px; 
                line-height: 1; transition: color 0.2s;
            " onmouseover="this.style.color='#fff'" onmouseout="this.style.color='#90a4ae'">✕</div>

                                            <div style="font-weight:800; color:#fff; font-size:13px; margin-bottom:6px; letter-spacing:1px; text-transform:uppercase;">
                                              ${tool.label}
                                            </div>
                                            <div style="font-size:12px; color:#cfd8dc; line-height:1.5; font-weight:400;">
                                              ${tool.desc || "No description available."}
                                            </div>
                                          </div>

                                          <div style="padding: 16px 20px; background:#fff;">
                                            <button id="btnGoToSupport" style="
                width: 100%;
                background: linear-gradient(to bottom, #ffffff 0%, #eceff1 100%);
                border: 1px solid #b0bec5;
                border-bottom: 2px solid #90a4ae;
                border-radius: 0px;
                color: #1565c0; /* Deep Blue */
                font-family: 'Inter', sans-serif;
                font-size: 11px;
                font-weight: 800;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                padding: 12px;
                cursor: pointer;
                display: flex; align-items: center; justify-content: center; gap: 8px;
                transition: all 0.2s ease;
                box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            ">
                                              <span style="font-size:14px;">💬</span> Ask AI about this Tool
                                            </button>
                                          </div>
                                          `;

  document.body.appendChild(tip);

  // 3. Positioning
  var rect = e.target.getBoundingClientRect();
  var tipRect = tip.getBoundingClientRect();
  var top = rect.top - tipRect.height - 15;
  var left = rect.left - (tipRect.width / 2) + (rect.width / 2);

  if (top < 10) top = rect.bottom + 15;
  if (left + tipRect.width > window.innerWidth) left = window.innerWidth - tipRect.width - 20;
  if (left < 10) left = 10;

  tip.style.top = top + "px";
  tip.style.left = left + "px";

  // 4. Interaction Handlers
  tip.onclick = function (ev) { ev.stopPropagation(); };

  var closeBtn = tip.querySelector("#btnCloseTip");
  if (closeBtn) closeBtn.onclick = function (ev) { ev.stopPropagation(); tip.remove(); };

  // Button Logic
  var goBtn = tip.querySelector("#btnGoToSupport");

  // Hover Effects
  goBtn.onmouseover = function () {
    this.style.background = "#e3f2fd";
    this.style.borderColor = "#90caf9";
    this.style.transform = "translateY(-1px)";
    this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.1)";
  };
  goBtn.onmouseout = function () {
    this.style.background = "linear-gradient(to bottom, #ffffff 0%, #eceff1 100%)";
    this.style.borderColor = "#b0bec5";
    this.style.transform = "translateY(0)";
    this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";
  };

  goBtn.onclick = function () {
    tip.remove(); // Close tooltip

    // Open Help Center
    if (typeof showHelpCenter === 'function') {
      showHelpCenter();

      // Wait for modal to render, then switch tab
      setTimeout(function () {
        if (typeof switchHelpTab === 'function') switchHelpTab('bot');

        var supportInp = document.getElementById("supportInput");
        if (supportInp) {
          supportInp.value = ""; // Clean input
          supportInp.placeholder = `Ask about ${tool.label}...`; // Context via placeholder
          supportInp.focus();
        }
      }, 100);
    }
  };

  // Close on Outside Click
  setTimeout(() => {
    document.addEventListener('click', function remover(ev) {
      if (tip && !tip.contains(ev.target)) {
        if (tip.parentNode) tip.remove();
        document.removeEventListener('click', remover);
      }
    });
  }, 100);
}
// ==========================================
// REPLACE: smartFillInput (Preserves Full Schedule for Modal)
// ==========================================
function smartFillInput(inp, rawValue) {
  if (!rawValue) return;

  var valStr = String(rawValue).trim();

  // 1. CHECK: Is this a payment/fee field?
  // We check for the class OR the special $! tag
  var isPayment = inp.classList.contains("payment-input") ||
    (inp.dataset.target && (inp.dataset.target.indexOf("$![") === 0 || inp.dataset.target.indexOf("![") === 0));

  if (isPayment) {

    // A. STORE FULL CONTEXT
    inp.dataset.fullTerms = valStr;
    // Actually, let's just show the raw value, but keep fullTerms for the Modal
    inp.value = valStr;


    inp.title = "Full Schedule: " + valStr;
    inp.style.backgroundColor = "#e3f2fd";  // Light Blue
    inp.style.fontWeight = "bold";

    inp.oninput = function () {
      delete this.dataset.fullTerms;
      this.style.backgroundColor = "";
      this.style.fontWeight = "";
      this.style.height = "auto";
      this.style.height = this.scrollHeight + "px";
    };

  } else {
    // Standard Field
    inp.value = valStr;
    inp.style.backgroundColor = "";
    inp.style.fontWeight = "normal";
  }


  inp.style.borderColor = "#2e7d32";
  if (inp.classList.contains('auto-expand')) {
    inp.style.height = "auto";
    inp.style.height = inp.scrollHeight + "px";
  }
}
// ==========================================
// REPLACE: showFileNameSettings (Premium Design + Persistence)
// ==========================================
function showFileNameSettings() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. Defaults
  var defaults = {
    colgate: "[Ref] - CP-[Party] -- [Type] (DRAFT [Date])",
    hills: "[Ref] - HPN-[Party] -- [Type] (DRAFT [Date])"
  };

  var current = userSettings.fileNamePatterns || defaults;

  // 2. Main Card Container
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.padding = "0";
  card.style.borderRadius = "0px";
  card.style.border = "1px solid #d1d5db";
  card.style.borderLeft = "5px solid #1565c0";
  card.style.background = "#fff";

  // --- PREMIUM HEADER ---
  var header = document.createElement("div");
  header.style.cssText = `
                                          background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
                                          padding: 12px 15px;
                                          border-bottom: 3px solid #c9a227;
                                          box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                                          margin-bottom: 12px;
                                          position: relative;
                                          overflow: hidden;
                                          `;
  header.innerHTML = `
                                          <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
                                          <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
                                            SYSTEM CONFIGURATION
                                          </div>
                                          <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
                                            <span>⚙️</span> Build Draft Settings
                                          </div>
                                          `;
  card.appendChild(header);

  // --- CONTENT BODY ---
  var contentDiv = document.createElement("div");
  contentDiv.style.padding = "20px";

  var isHighlight = userSettings.highlightFill === true;

  var html = `
                                          <div style="font-size:12px; color:#37474f; font-weight:700; margin-bottom:10px; text-transform:uppercase; border-bottom:1px solid #eee; padding-bottom:5px;">
                                            File Naming Convention
                                          </div>
                                          <div style="font-size:11px; color:#666; margin-bottom:15px; line-height:1.4; background:#f9fafb; padding:10px; border:1px solid #eee;">
                                            Define how filenames are generated automatically.<br>
                                              <strong>Available Tokens:</strong><br>
                                                <span style="font-family:monospace; color:#1565c0;">[Ref]</span>,
                                                <span style="font-family:monospace; color:#1565c0;">[Party]</span>,
                                                <span style="font-family:monospace; color:#1565c0;">[Type]</span>,
                                                <span style="font-family:monospace; color:#1565c0;">[Date]</span>
                                              </div>

                                              <label class="tool-label">Colgate Pattern</label>
                                              <input type="text" id="patColgate" class="tool-input" value="${current.colgate}" style="font-family:monospace; font-size:11px; margin-bottom:15px;">

                                                <label class="tool-label">Hill's Pattern</label>
                                                <input type="text" id="patHills" class="tool-input" value="${current.hills}" style="font-family:monospace; font-size:11px;">

                                                  <div style="margin-top:20px; border-top:1px solid #eee; padding-top:15px;">
                                                    <div style="font-size:12px; color:#37474f; font-weight:700; margin-bottom:10px; text-transform:uppercase;">
                                                      Verification
                                                    </div>
                                                    <label class="settings-item-row" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer; margin-bottom:5px;">
                                                      <span style="font-size:13px; color:#333;">🖊️ Highlight Replacements</span>
                                                      <input type="checkbox" id="chkHighlightFill" ${isHighlight ? "checked" : ""}>
                                                    </label>
                                                    <div style="font-size:10px; color:#888; line-height:1.3;">
                                                      If enabled, filled placeholders will be highlighted. You must verify and approve them to remove the highlights.
                                                    </div>
                                                  </div>

                                                  <div style="text-align:right; margin-top:10px;">
                                                    <button id="btnResetPat" class="btn-action" style="width:auto; padding:4px 10px; font-size:10px; color:#c62828;">Reset to Defaults</button>
                                                  </div>
                                                  `;
  contentDiv.innerHTML = html;
  card.appendChild(contentDiv);

  // --- FOOTER BUTTONS ---
  var footerDiv = document.createElement("div");
  footerDiv.style.cssText = "padding:15px 20px; background:#f8f9fa; border-top:1px solid #e0e0e0; display:flex; gap:10px;";

  var btnCancel = document.createElement("button");
  btnCancel.className = "btn-platinum-config";
  btnCancel.style.flex = "1";
  btnCancel.innerText = "CANCEL";
  btnCancel.onclick = showHome;

  var btnSave = document.createElement("button");
  btnSave.className = "btn-platinum-config primary";
  btnSave.style.cssText = "flex: 1; font-weight:800;";
  btnSave.innerText = "SAVE SETTINGS";

  footerDiv.appendChild(btnCancel);
  footerDiv.appendChild(btnSave);
  card.appendChild(footerDiv);
  out.appendChild(card);

  // --- BIND LOGIC ---
  btnSave.onclick = () => {
    userSettings.fileNamePatterns = {
      colgate: document.getElementById("patColgate").value,
      hills: document.getElementById("patHills").value
    };
    userSettings.highlightFill = document.getElementById("chkHighlightFill").checked;

    saveCurrentSettings(); // PERSIST TO DISK
    showToast("✅ Settings Saved");
    showHome();
  };

  document.getElementById("btnResetPat").onclick = () => {
    document.getElementById("patColgate").value = defaults.colgate;
    document.getElementById("patHills").value = defaults.hills;
  };
}
// ==========================================
// START: AUTO-FILL HUB (PREMIUM LANDING)
// ==========================================
// ==========================================
// START: AUTO-FILL HUB (4-STEP + WIZARD MODE)
// ==========================================
function showAutoFillHub() {
  // 1. Reset Wizard Flag (Safety: Ensure manual usage doesn't auto-jump)
  window.eliWizardMode = false;

  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- PREMIUM HEADER ---
  var header = document.createElement("div");
  header.style.cssText = `
      background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
      padding: 12px 15px;
      border-bottom: 3px solid #c9a227;
      position: relative;
      overflow: hidden;
      margin-bottom: 10px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.1);
  `;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
        DOCUMENT AUTOMATION
      </div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
        <span>⚡</span> Build Draft Hub
      </div>
  `;
  out.appendChild(header);

  // --- NEW: RUN ALL BUTTON ---
  var btnRunAll = document.createElement("button");
  btnRunAll.className = "btn-platinum-config";
  btnRunAll.innerHTML = "🚀 RUN ALL 4 STEPS";
  btnRunAll.style.cssText = "width:100%; margin-bottom:15px; background:linear-gradient(to bottom, #e3f2fd 0%, #bbdefb 100%); color:#1565c0; border:1px solid #90caf9; font-weight:800; font-size:12px; padding:12px; box-shadow:0 2px 5px rgba(0,0,0,0.1);";

  btnRunAll.onclick = function () {
    // START WIZARD MODE
    window.eliWizardMode = true;
    showToast("🚀 Starting Wizard: Step 1...");
    runDocumentScannerBuild(); // Go to Step 1
  };
  out.appendChild(btnRunAll);
  // ---------------------------

  // --- HELPER: RENDER COMPACT STEP CARD ---
  function createStepCard(stepNum, icon, title, desc, action, color) {
    var card = document.createElement("div");
    card.className = "tool-box slide-in";
    card.style.cssText = `padding: 0; border: 1px solid #d1d5db; border-left: 5px solid ${color}; background: #fff; border-radius:0px; margin-bottom: 10px; cursor: pointer; transition: transform 0.2s;`;

    card.onclick = action;
    card.onmouseover = function () { this.style.transform = "translateY(-2px)"; this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.1)"; };
    card.onmouseout = function () { this.style.transform = "translateY(0)"; this.style.boxShadow = "none"; };

    card.innerHTML = `
        <div style="padding: 12px 15px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                <div style="font-size:10px; font-weight:800; color:${color}; text-transform:uppercase; letter-spacing:1px;">STEP ${stepNum}</div>
                <div style="font-size:18px;">${icon}</div>
            </div>
            <div style="font-family:'Georgia', serif; font-size:14px; color:#333; margin-bottom:4px; font-weight:600;">${title}</div>
            <div style="font-size:11px; color:#546e7a; line-height:1.4;">${desc}</div>
        </div>
    `;
    return card;
  }

  // STEP 1
  var card1 = createStepCard("1", "🏗️", "Build Logic", "Generate custom draft from {{IF}} logic.", function () {
    runDocumentScannerBuild();
  }, "#1565c0");
  out.appendChild(card1);

  // STEP 2
  var card2 = createStepCard("2", "✍️", "Fill Fields", "Populate [Bracketed] placeholders.", function () {
    runTemplateScan();
  }, "#2e7d32");
  out.appendChild(card2);

  // STEP 3
  var card3 = createStepCard("3", "📝", "Error Check", "Scan for typos & broken references.", function () {
    // FIX: Pass 'true' as the second argument. 
    // This activates "Fast Mode", which FORCES Track Changes to be OFF regardless of user settings.
    runInteractiveErrorCheck(showAutoFillHub, true);
  }, "#f57f17");
  out.appendChild(card3);

  // STEP 4
  var card4 = createStepCard("4", "💾", "Save and/or Send", "Finalize document and draft email.", function () {
    showSaveAsManager();
  }, "#7b1fa2");
  out.appendChild(card4);

  // HOME
  var btnHome = document.createElement("button");
  btnHome.className = "btn-platinum-config";
  btnHome.innerText = "🏠 RETURN HOME";
  btnHome.style.width = "100%";
  btnHome.style.marginTop = "10px";
  btnHome.onclick = showHome;
  out.appendChild(btnHome);
}

var logicFoundDefinitions = {};
var logicFieldDependencies = {};

// ==========================================
function runDocumentScannerBuild() {
  setLoading(true, "default");

  Word.run(async (context) => {
    const body = context.document.body;

    // 1. SCAN TITLE
    let docTitle = null;
    const titleResults = body.search("\\{\\{TITLE:*\\}\\}", { matchWildcards: true });
    titleResults.load("text");

    // 2. SCAN LOGIC TAGS (Strict Pattern)
    // Matches {{TAG}} or {{TAG:FORMAT}}
    // Updated Regex to allow colon inside the tag name for formatters
    const searchResults = body.search("\\{\\{[!\\}]@\\}\\}", { matchWildcards: true });
    searchResults.load("text");

    await context.sync();

    if (titleResults.items.length > 0) {
      docTitle = titleResults.items[0].text.replace(/[{}]/g, "").replace("TITLE:", "").trim();
    }

    logicFoundDefinitions = {};
    logicFieldDependencies = {};

    let logicStack = [];
    let logicCount = 0;

    for (let item of searchResults.items) {
      // Clean brackets
      let text = item.text.replace(/^\{\{|\}\}$/g, "").trim();

      // Handle Logic Blocks (IF/ELSE/ENDIF) - These don't have formatters
      if (text.startsWith("IF:")) {
        logicStack.push(text.substring(3).trim());
      }
      else if (text.startsWith("ELSEIF:")) {
        if (logicStack.length > 0) {
          logicStack.pop();
          logicStack.push(text.substring(7).trim());
        }
      }
      else if (text.startsWith("ELSE")) {
        if (logicStack.length > 0) {
          logicStack.pop();
          logicStack.push("true");
        }
      }
      else if (text.startsWith("ENDIF")) {
        if (logicStack.length > 0) logicStack.pop();
      }
      else if (text.startsWith("DEF:")) {
        // Definitions are special: DEF:Key|Options
        if (text.includes("|")) {
          parseLogicDefinition(text, logicStack);
          logicCount++;
        }
      }
      // Note: We don't need count standard variables here, only logic definitions.
      // Variables (e.g. {{Type:UPPER}}) are placeholders found during generation.
    }

    setLoading(false);
    renderLogicBuilderForm(logicCount, docTitle);

  }).catch(e => {
    setLoading(false);
    showOutput("Scanner Error: " + e.message, true);
  });
}

function parseLogicDefinition(rawText, currentStack) {
  let clean = rawText.replace("DEF:", "").trim();
  let pipeSplit = clean.split("|");
  let defPart = pipeSplit[0].trim();
  let optionsPart = pipeSplit.length > 1 ? pipeSplit[1].trim() : null;
  let warningText = pipeSplit.length > 2 ? pipeSplit[2].trim() : null;

  if (!optionsPart) return;

  let defSplit = defPart.split("=");
  let key = defSplit[0].trim();

  // Existing default logic (from key=Val)
  let explicitDefault = defSplit.length > 1 ? defSplit[1].trim() : "";

  // DETECT TEXT MODE: If option starts with "*"
  let isTextMode = optionsPart.startsWith("*");

  // New: Extract text default from "*DefaultValue"
  let textDefault = isTextMode ? optionsPart.substring(1).trim() : "";

  // Final Default: Prefer explicit (key=Val), then text default (*Val), then empty
  let finalDefault = explicitDefault || textDefault;

  logicFoundDefinitions[key] = {
    default: finalDefault,
    isText: isTextMode,
    options: isTextMode ? [] : optionsPart.split(",").map(o => o.trim()),
    warning: warningText
  };

  if (currentStack.length > 0) {
    logicFieldDependencies[key] = [...currentStack];
  }
}

function renderLogicBuilderForm(count, docTitle) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  var displayTitle = docTitle || "Document Builder";

  // Header
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); position: relative; overflow: hidden;`;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">LOGIC ENGINE</div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;"><span>🏗️</span> ${displayTitle}</div>
  `;
  out.appendChild(header);

  if (count === 0) {
    out.innerHTML += `
            <div class="text-block" style="background:#f5f5f5; border-color:#ddd; color:#555; margin: 0 15px;">
                <b>No logic tags found.</b><br><br>
                1. Open an <b>ELI Template</b> document.<br>
                2. Ensure it contains <code>{{DEF:...}}</code> and <code>{{IF:...}}</code> tags.
            </div>
        `;

    var btnBack = document.createElement("button");
    btnBack.className = "btn-action";
    btnBack.innerText = "⬅ Back to Hub";
    btnBack.style.width = "calc(100% - 30px)";
    btnBack.style.margin = "20px 15px";
    btnBack.onclick = showAutoFillHub;
    out.appendChild(btnBack);
    return;
  }

  var formContainer = document.createElement("div");
  formContainer.id = "logicForm";
  formContainer.style.padding = "0 15px";

  for (let key in logicFoundDefinitions) {
    let def = logicFoundDefinitions[key];

    let group = document.createElement("div");
    group.className = "question-group";
    group.setAttribute("data-key", key);
    group.style.marginBottom = "10px";

    let label = document.createElement("label");
    label.className = "tool-label";

    // ➤ NEW: Render Warning Badge if present
    if (def.warning) {
      label.innerHTML = `${key} <span style="color:#c62828; background:#ffebee; font-size:9px; padding:2px 5px; border-radius:4px; margin-left:6px; font-weight:700; border:1px solid #ffcdd2;">⚠️ ${def.warning}</span>`;
    } else {
      label.innerText = key;
    }

    group.appendChild(label);

    let inputEl;

    // --- LOGIC: CHOOSE INPUT TYPE ---
    if (def.isText) {
      // Render Text Input
      // --- INSERT THIS BLOCK ---
      let helpText = document.createElement("div");
      helpText.innerText = "Free Form Entry - leave as 'Don't know' if you are not sure of the answer.";
      // Green color, pulled up closer to label (-6px margin)
      helpText.style.cssText = "font-size: 9px; color: #2e7d32; font-style: italic; margin-top: -6px; margin-bottom: 4px;";
      group.appendChild(helpText);
      // ------------------------

      // Render Text Input
      inputEl = document.createElement("input");
      inputEl.type = "text";
      inputEl.className = "tool-input logic-input";

      // Use parsed default (e.g. from *DefaultVal)
      inputEl.value = def.default || "";
      inputEl.placeholder = "Enter " + key + "...";

      // Trigger logic update on typing
      inputEl.oninput = () => updateLogicVisibility();
    } else {
      // Render Dropdown
      inputEl = document.createElement("select");
      inputEl.className = "audience-select logic-input";
      inputEl.style.width = "100%";

      def.options.forEach(opt => {
        let o = document.createElement("option");
        o.value = opt;
        o.innerText = opt;
        if (opt === def.default) o.selected = true;
        inputEl.appendChild(o);
      });

      // Trigger logic update on selection change
      inputEl.onchange = () => updateLogicVisibility();
    }
    // --------------------------------

    inputEl.setAttribute("data-key", key);
    group.appendChild(inputEl);
    formContainer.appendChild(group);
  }
  out.appendChild(formContainer);

  var btnGen = document.createElement("button");
  btnGen.className = "btn-platinum-config primary";
  btnGen.style.cssText = "width:calc(100% - 30px); margin:20px 15px; background:#1565c0; color:white; border-bottom:3px solid #0d47a1; padding:12px; font-size:12px;";
  btnGen.innerText = "🚀 GENERATE DOCUMENT";
  btnGen.onclick = generateDocumentFromLogic;
  out.appendChild(btnGen);

  updateLogicVisibility();
}

function updateLogicVisibility() {
  let currentAnswers = {};
  const inputs = document.querySelectorAll(".logic-input");
  inputs.forEach(input => currentAnswers[input.getAttribute("data-key")] = input.value);

  for (let key in logicFieldDependencies) {
    let rules = logicFieldDependencies[key];
    let isVisible = true;

    for (let rule of rules) {
      if (!evaluateLogicCondition(rule, currentAnswers)) {
        isVisible = false;
        break;
      }
    }

    let input = document.querySelector(`.question-group[data-key="${key}"]`);
    if (input) input.style.display = isVisible ? "block" : "none";
  }
}
// ==========================================
// REPLACE: evaluateLogicCondition (Supports || OR Logic)
// ==========================================
function evaluateLogicCondition(conditionStr, answers) {
  // 1. Simple Pass-through
  if (conditionStr === "true") return true;

  // 2. Identify Operator (== or !=)
  var operator = conditionStr.includes("!=") ? "!=" : "==";
  var parts = conditionStr.split(operator);

  // Safety: If split failed to produce 2 parts, logic is invalid
  if (parts.length < 2) return false;

  var key = parts[0].trim();
  var rawRightSide = parts[1];

  // 3. Parse Right Side (Handle || for multiple options)
  // Example: "Corp || LLC || Inc" -> ["Corp", "LLC", "Inc"]
  var targets = rawRightSide.split("||").map(function (t) {
    return t.trim().replace(/['"]/g, "");
  });

  var actualVal = answers[key];

  // 4. Comparison Logic
  if (actualVal === undefined) return false;

  if (operator === "==") {
    // MATCH LOGIC: True if actual value matches ANY of the targets
    // "Type == A || B" -> True if Type is A OR Type is B
    return targets.indexOf(actualVal) !== -1;
  } else {
    // MISMATCH LOGIC: True if actual value matches NONE of the targets
    // "Type != A || B" -> True if Type is NOT A AND NOT B
    return targets.indexOf(actualVal) === -1;
  }
}

async function generateDocumentFromLogic() {
  setLoading(true, "default");

  let userAnswers = {};
  try {
    const inputs = document.querySelectorAll(".logic-input");
    inputs.forEach(i => userAnswers[i.getAttribute("data-key")] = i.value);
  } catch (e) { console.warn("Input capture warning:", e); }

  try {
    await Word.run(async (context) => {

      // ➤ FORCE TRACK CHANGES OFF (Clean Build)
      context.document.changeTrackingMode = "Off";

      const body = context.document.body;

      // 2. Run Logic Blocks (IF/ELSE)
      await runSegmentedLogic(context, userAnswers);

      // 3. Replace Variables
      for (let key in userAnswers) {
        let rawValue = userAnswers[key];
        let searchPattern = `\\{\\{${key}*\\}\\}`;
        let results = body.search(searchPattern, { matchWildcards: true });
        results.load("text");
        await context.sync();

        for (let i = results.items.length - 1; i >= 0; i--) {
          let item = results.items[i];
          let tagContent = item.text.replace(/^\{\{|\}\}$/g, "").trim();
          let parts = tagContent.split(":");
          let tagKey = parts[0].trim();
          let formatter = parts.length > 1 ? parts[1].trim().toUpperCase() : null;

          if (tagKey !== key) continue;

          let finalValue = rawValue;
          if (formatter === "UPPER") finalValue = finalValue.toUpperCase();
          else if (formatter === "LOWER") finalValue = finalValue.toLowerCase();
          else if (formatter === "TITLE") {
            finalValue = finalValue.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
          }
          // --- UPDATED INSERTION LOGIC ---

          // 1. Normalize for comparison (remove punctuation like apostrophes)
          let cleanVal = finalValue.trim().toLowerCase().replace(/['"’]/g, "");

          let isOther = cleanVal === "other";
          // Check if it contains "dont know" (covers "i don't know", "dont know", etc)
          let isUnknown = cleanVal.includes("dont know");

          // 2. Modify value based on condition
          if (isOther) {
            finalValue = "[" + finalValue + "]";
          } else if (isUnknown) {
            // Convert to "[THIRD PARTY TO FILL IN - LABEL]"
            finalValue = `[THIRD PARTY TO FILL IN - ${key.toUpperCase()}]`;
          }

          // 3. Perform the replacement and capture the range
          let insertedRange = item.insertText(finalValue, "Replace");

          // 4. Apply Red Color & Bold to the WHOLE inserted text
          if (isOther || isUnknown) {
            insertedRange.font.color = "#c62828";
            insertedRange.font.bold = true;
          }
        }
      }

      // 4. CLEANUP: Delete Tag Text
      const defs = body.search("\\{\\{DEF:*\\}\\}", { matchWildcards: true });
      const titles = body.search("\\{\\{TITLE:*\\}\\}", { matchWildcards: true });
      defs.load("items");
      titles.load("items");
      await context.sync();

      // We delete the tag TEXT here, leaving an empty paragraph for the next step to evaluate
      for (let i = defs.items.length - 1; i >= 0; i--) defs.items[i].delete();
      for (let i = titles.items.length - 1; i >= 0; i--) titles.items[i].delete();
      await context.sync();
      // 5a. FORCE DELETE: Handle {{DELETE_LINE}}
      // Finds the tag and deletes the entire paragraph (The opposite of {{ENTER}})
      const delTags = body.search("\\{\\{DELETE_LINE\\}\\}", { matchWildcards: true });
      delTags.load("items");
      await context.sync();

      for (let i = delTags.items.length - 1; i >= 0; i--) {
        // Delete the entire paragraph containing the tag
        delTags.items[i].paragraphs.getFirst().delete();
      }
      // 5. SPACER LOGIC: Protect {{ENTER}}
      // We replace it with a "Zero Width Space". This makes the line technically "not empty".
      const enterTags = body.search("\\{\\{ENTER\\}\\}", { matchWildcards: true });
      enterTags.load("items");
      await context.sync();

      for (let i = enterTags.items.length - 1; i >= 0; i--) {
        enterTags.items[i].insertText("\u200B", "Replace");
      }

      // 6. SMART COLLAPSE (The Fix)
      // Goal: If we see 2+ empty lines, delete the extras. Keep 1.
      var allParas = body.paragraphs;
      allParas.load("text");
      await context.sync();

      // Loop backwards, stopping before index 0
      for (let i = allParas.items.length - 1; i > 0; i--) {
        let currentPara = allParas.items[i];
        let abovePara = allParas.items[i - 1];

        let cText = currentPara.text;
        let aText = abovePara.text;

        // A line is "Empty" if it has no text AND no secret {{ENTER}} token
        let isCurrentEmpty = (cText.trim().length === 0 && !cText.includes("\u200B") && !cText.includes('\f'));
        let isAboveEmpty = (aText.trim().length === 0 && !aText.includes("\u200B") && !aText.includes('\f'));

        // If THIS line is empty AND the one above is empty -> Delete THIS one.
        // This collapses the stack from bottom-up until only 1 remains.
        if (isCurrentEmpty && isAboveEmpty) {
          currentPara.delete();
        }
      }

      await context.sync();
    });

    setLoading(false);
    showToast("✅ Draft Generated!");
    if (window.eliWizardMode) {
      showToast("🚀 Proceeding to Step 2: Fill Fields");
      setTimeout(runTemplateScan, 800); // Auto-jump
    } else {
      setTimeout(showAutoFillHub, 50); // Standard exit
    }
  } catch (e) {
    setLoading(false);
    console.error("Generation Error:", e);
    var out = document.getElementById("output");
    out.innerHTML = `<div class="placeholder" style="color:red">⚠️ Error: ${e.message}</div>`;
  }
}

async function runSegmentedLogic(context, answers) {
  const body = context.document.body;
  const results = body.search("\\{\\{[!\\}]@\\}\\}", { matchWildcards: true });
  results.load(["text", "index"]);
  await context.sync();

  let tags = [];
  for (let i = 0; i < results.items.length; i++) {
    let text = results.items[i].text.replace(/[{}]/g, "").trim();
    if (text.startsWith("IF:") || text.startsWith("ELSE") || text.startsWith("ENDIF")) {
      tags.push({ range: results.items[i], text: text, index: i });
    }
  }

  let blockStack = [];
  let executionPlans = [];

  for (let i = 0; i < tags.length; i++) {
    let tag = tags[i];
    let txt = tag.text;

    if (txt.startsWith("IF:")) {
      let block = { segments: [], startTagIndex: i, depth: blockStack.length };
      block.segments.push({ condition: txt.substring(3).trim(), tagIndex: i });
      blockStack.push(block);
    }
    else if (txt.startsWith("ELSEIF:")) {
      let block = blockStack.length > 0 ? blockStack[blockStack.length - 1] : null;
      if (block) {
        block.segments[block.segments.length - 1].endTagIndex = i;
        block.segments.push({ condition: txt.substring(7).trim(), tagIndex: i });
      }
    }
    else if (txt.startsWith("ELSE")) {
      let block = blockStack.length > 0 ? blockStack[blockStack.length - 1] : null;
      if (block) {
        block.segments[block.segments.length - 1].endTagIndex = i;
        block.segments.push({ condition: "true", tagIndex: i });
      }
    }
    else if (txt.startsWith("ENDIF")) {
      let block = blockStack.pop();
      if (block) {
        block.segments[block.segments.length - 1].endTagIndex = i;
        executionPlans.push({ segments: block.segments, finalEndTagIndex: i, depth: block.depth });
      }
    }
  }

  executionPlans.sort((a, b) => b.depth - a.depth);

  for (let plan of executionPlans) {
    let winningSegmentIndex = -1;
    for (let s = 0; s < plan.segments.length; s++) {
      if (evaluateLogicCondition(plan.segments[s].condition, answers)) {
        winningSegmentIndex = s;
        break;
      }
    }

    if (winningSegmentIndex > -1) {
      for (let s = 0; s < plan.segments.length; s++) {
        let seg = plan.segments[s];
        let startTag = tags[seg.tagIndex].range;
        if (seg.endTagIndex === undefined) continue;
        let endTag = tags[seg.endTagIndex].range;

        if (s === winningSegmentIndex) {
          startTag.delete();
        } else {
          let range = startTag.getRange("Start").expandTo(endTag.getRange("Start"));
          range.delete();
        }
      }
      if (tags[plan.finalEndTagIndex]) tags[plan.finalEndTagIndex].range.delete();
    } else {
      if (plan.segments.length > 0 && plan.finalEndTagIndex !== undefined) {
        let firstTag = tags[plan.segments[0].tagIndex].range;
        let finalTag = tags[plan.finalEndTagIndex].range;
        let range = firstTag.getRange("Start").expandTo(finalTag.getRange("End"));
        range.delete();
      }
    }
  }
  await context.sync();
}
// ==========================================
// START: requireUserName (Prompts for Name if missing)
// ==========================================
function requireUserName(callback) {
  // 1. Check if name exists in settings
  var currentName = "";
  if (userSettings.emailTemplates && userSettings.emailTemplates.myName) {
    currentName = userSettings.emailTemplates.myName.trim();
  }

  // 2. CHECK: Only check if it is empty. 
  // (Removed the check !== "Andrew" so it now accepts Andrew as a valid name)
  if (currentName.length > 0) {
    callback();
    return;
  }

  // 3. If missing, show Modal
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:2147483647; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var dialog = document.createElement("div");
  dialog.className = "tool-box slide-in";
  dialog.style.cssText = "background:white; width:300px; padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #1565c0; box-shadow:0 20px 50px rgba(0,0,0,0.3);";

  dialog.innerHTML = `
                                                        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 15px; border-bottom: 3px solid #c9a227;">
                                                          <div style="font-size:10px; font-weight:800; color:#8fa6c7; text-transform:uppercase;">SETUP REQUIRED</div>
                                                          <div style="font-family:'Georgia', serif; font-size:16px; color:white;">👤 User Profile</div>
                                                        </div>
                                                        <div style="padding:20px;">
                                                          <div style="font-size:12px; color:#546e7a; margin-bottom:15px; line-height:1.4;">
                                                            Please enter your <strong>First Name</strong>.<br>
                                                              This will be used for email signatures and greetings.
                                                          </div>
                                                          <label class="tool-label">First Name</label>
                                                          <input type="text" id="setupFirstName" class="tool-input" placeholder="e.g. Andrew" style="width:100%; box-sizing:border-box;">

                                                            <button id="btnSaveNameSetup" class="btn-platinum-config primary" style="width:100%; margin-top:10px; background:#1565c0; color:white;">SAVE & CONTINUE</button>
                                                        </div>
                                                        `;

  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  // Focus input
  setTimeout(function () { document.getElementById("setupFirstName").focus(); }, 100);

  // Bind Save
  document.getElementById("btnSaveNameSetup").onclick = function () {
    var val = document.getElementById("setupFirstName").value.trim();
    if (!val) {
      showToast("⚠️ Name is required");
      return;
    }

    // Save to Settings
    if (!userSettings.emailTemplates) userSettings.emailTemplates = {};
    userSettings.emailTemplates.myName = val;
    saveCurrentSettings(); // Persist to Local Storage

    overlay.remove();
    showToast("✅ Profile Saved");
    callback(); // Execute the original action
  };
}
// ==========================================
// END: requireUserName
// ==========================================
// ==========================================
// START: locatePlaceholderOrValue (NEW HELPER)
// ==========================================
function locatePlaceholderOrValue(placeholder, filledValue, btnElement) {
  // 1. UI Feedback
  var originalIcon = "Locate 📍";
  btnElement.innerHTML = "⏳"; // Spinner/Hourglass
  btnElement.style.opacity = "1";

  Word.run(async function (context) {
    // A. Try finding the Bracketed Placeholder first
    // (This is the default state before filling)
    var foundPh = await findSmartMatch(context, placeholder);

    if (foundPh) {
      foundPh.select();
      await context.sync();
      return "placeholder";
    }

    // B. If not found, try finding the Filled Value
    // (This logic runs after the user has clicked "Fill All")
    if (filledValue && filledValue.trim().length > 0) {
      // Safety: Don't search for extremely short common words like "a" or "1" to avoid noise
      if (filledValue.trim().length < 2) return "short";

      var foundVal = await findSmartMatch(context, filledValue);
      if (foundVal) {
        foundVal.select();
        await context.sync();
        return "value";
      }
    }

    return "none";
  })
    .then(function (result) {
      if (result === "placeholder") {
        btnElement.innerHTML = "✅";
        showToast("📍 Found Placeholder");
      } else if (result === "value") {
        btnElement.innerHTML = "✅";
        showToast("📍 Found Filled Value");
      } else if (result === "short") {
        btnElement.innerHTML = "⚠️";
        showToast("⚠️ Value too short to search safely");
      } else {
        btnElement.innerHTML = "❌";
        showToast("⚠️ Not Found in Document");
      }

      // Reset Icon after 2s
      setTimeout(function () {
        btnElement.innerHTML = originalIcon;
        btnElement.style.opacity = "0.6";
      }, 2000);
    })
    .catch(function (e) {
      console.error("Locate Error:", e);
      btnElement.innerHTML = "❌";
      showToast("Error locating text");
    });
}

// ==========================================
// START: Onboarding Wizard
// ==========================================
function checkOnboarding() {
  if (userSettings.hasOnboarded === true) {
    showHome();
  } else {
    // Check for Skip Flag (set by Reset)
    if (localStorage.getItem("eli_skip_intro") === "true") {
      localStorage.removeItem("eli_skip_intro");
      showOnboardingSetup();
    } else {
      showOnboardingWelcome();
    }
  }
}

function showOnboardingWelcome() {
  runWelcomeAnimation(function () {
    showOnboardingSetup();
  });
}

function runWelcomeAnimation(onCompleteCallback) {
  // 1. Create Overlay
  var overlay = document.createElement("div");
  overlay.className = "tour-overlay";
  // Wrap in a squared-off dialog box (tool-box) for consistency
  overlay.innerHTML = `
                                                        <div class="tool-box slide-in" style="width:95%; max-width:700px; min-height:600px; padding:0; border-radius:0px; display:flex; align-items:center; justify-content:center; position:relative; box-shadow:0 20px 50px rgba(0,0,0,0.3); background: #ffffff;">
                                                          <div id="tourStage" class="tour-stage"></div>
                                                          <div id="btnSkipTour" style="position:absolute; bottom:20px; right:20px; font-size:11px; color:#b0bec5; cursor:pointer; font-weight:600; text-transform:uppercase; letter-spacing:1px; z-index:30005;">Skip Intro →</div>
                                                        </div>
                                                        `;
  document.body.appendChild(overlay);

  var stage = document.getElementById("tourStage");
  var timer = null;

  // Bind Skip
  document.getElementById("btnSkipTour").onclick = function () {
    if (timer) clearTimeout(timer);
    overlay.remove();
    if (onCompleteCallback) onCompleteCallback();
  };

  // 2. Define Scenes
  var scenes = [
    {
      duration: 3500, // Scene 1: Welcome
      html: `<div class="gavel-icon">⚖️<div class="gavel-ripple"></div></div>
                                                        <div class="tour-title">Welcome to ELI</div>
                                                        <div class="tour-subtitle">Where <strong>Legal Analysis</strong> meets <strong>Artificial Intelligence</strong>.</div>`
    },
    {
      duration: 3000, // Scene 2: Drafting
      html: `<div class="block-container"><div class="block"></div><div class="block"></div><div class="block"></div><div class="block"></div></div>
                                                        <div class="tour-title">Drafting Assistant</div>
                                                        <div class="tour-subtitle">Building complex drafts automatically from your templates.</div>`
    },
    {
      duration: 3500, // Scene 3: Playbook
      html: `<div class="book-icon">📖<div class="ai-badge">🤖</div></div>
                                                        <div class="tour-title">Review Assistant</div>
                                                        <div class="tour-subtitle">Checking documents against <strong>Colgate's Playbook</strong> and providing fallbacks.</div>`
    },
    {
      duration: 3000, // Scene 4: Analysis
      html: `<div class="doc-scan"><div class="scan-line"></div><div class="bug-dot"></div></div>
                                                        <div class="tour-title">Legal Analysis</div>
                                                        <div class="tour-subtitle">Identifying grammatical erros, legal risks, missing clauses, and definitions instantly.</div>`
    },
    {
      duration: 3000, // Scene 5: Redlines
      html: `<div class="redline-box"><span class="txt-del">Mutual</span><span class="txt-add">One-Way</span></div>
                                                        <div class="tour-title">Redline Review and Document Comparison</div>
                                                        <div class="tour-subtitle">Use AI to explain the redlines or compare to another version.</div>`
    },
    {
      duration: 2500, // Scene 6: Email
      html: `<div class="plane-icon">📧</div>
                                                        <div class="tour-title">Auto Email</div>
                                                        <div class="tour-subtitle">Preparing emails to counterparties for you.</div>`
    },
    {
      duration: 0, // Scene 7: Finale (Stops here)
      html: `<div class="final-logo">ELI</div>
                                                        <div class="tour-title">...and so much more.</div>
                                                        <div class="tour-subtitle">Your AI Legal Partner is ready.</div>
                                                        <button id="btnTourFinish" class="tour-btn">GET STARTED 🚀</button>

                                                        <div style="background:#fff3cd; color:#856404; padding:12px; border-radius:0px; font-size:12px; margin-top:20px; text-align:center; border:1px solid #ffeeba; line-height:1.4; opacity:0; animation: tourFadeIn 1s ease 1s forwards;">
                                                          ⚠️ <b>HEADS UP:</b><br>Please expand the size of this plugin window to get the most out of ELI!
                                                        </div>`
    }
  ];

  // 3. Play Function
  function playScene(index) {
    try {
      if (!scenes || index >= scenes.length) return;

      if (!stage) stage = document.getElementById("tourStage");
      if (!stage) return;

      var scene = scenes[index];
      stage.innerHTML = scene.html; // Inject Content

      // Reset animations by triggering reflow
      void stage.offsetWidth;

      // If it's the last scene, bind the button
      if (index === scenes.length - 1) {
        var btnFinish = document.getElementById("btnTourFinish");
        if (btnFinish) {
          btnFinish.onclick = function () {
            overlay.style.transition = "opacity 0.5s ease";
            overlay.style.opacity = "0";
            setTimeout(function () {
              overlay.remove();
              if (onCompleteCallback) onCompleteCallback();
            }, 500);
          };
        }
      }

      // Schedule next scene
      if (scene.duration > 0) {
        timer = window.setTimeout(function () { playScene(index + 1); }, scene.duration);
      }
    } catch (e) {
      console.error("Animation Error:", e);
    }
  }

  // Start immediately
  // Delay slightly to ensure DOM render
  window.setTimeout(function () { playScene(0); }, 50);
}

function showOnboardingSetup() {
  // Create or reuse Overlay
  var overlay = document.getElementById("onboardingOverlay");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "onboardingOverlay";
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(255,255,255,0.95); z-index:20000; display:flex; justify-content:center; align-items:center; overflow-y:auto;";
    document.body.appendChild(overlay);
  }
  overlay.innerHTML = "";

  // Welcome Card
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "width:95%; max-width:700px; min-height:600px; padding:0; border:1px solid #d1d5db; background:#fff; border-radius:0px; overflow:hidden; border-top:5px solid #1565c0; box-shadow: 0 10px 30px rgba(0,0,0,0.25);";

  var header = document.createElement("div");
  header.style.cssText = "padding:20px; text-align:center; background:#f0f4f8;";
  header.innerHTML = `
                                                        <div style="font-size:32px; margin-bottom:10px;">⚙️</div>
                                                        <div style="font-size:18px; font-weight:800; color:#1565c0; margin-bottom:5px;">Customize your ELI</div>
                                                        <div style="font-size:12px; color:#546e7a;">Let's get you set up specifically for your needs.</div>
                                                        `;
  card.appendChild(header);

  var body = document.createElement("div");
  body.style.padding = "20px";

  // 1. Full Name
  var curName = (userSettings.emailTemplates && userSettings.emailTemplates.myName) ? userSettings.emailTemplates.myName : "";
  body.innerHTML += `
                                                        <label class="tool-label">Full Name</label>
                                                        <input type="text" id="obName" class="tool-input" value="${curName}" placeholder="e.g. John Doe" style="margin-bottom:15px;">
                                                          `;

  // 1b. Email Address (NEW)
  var curEmail = userSettings.userEmail || "";
  body.innerHTML += `
                                                          <label class="tool-label">Email Address</label>
                                                          <input type="text" id="obEmail" class="tool-input" value="${curEmail}" placeholder="e.g. john_doe@colgate.com" style="margin-bottom:15px;">
                                                            `;

  // 2. Language
  body.innerHTML += `
                                                            <label class="tool-label">Preferred Language</label>
                                                            <select id="obLang" class="audience-select" style="width:100%; margin-bottom:20px;">
                                                              <option value="en">English</option>
                                                              <option value="es">Español</option>
                                                              <option value="fr">Français</option>
                                                              <option value="de">Deutsch</option>
                                                              <option value="it">Italiano</option>
                                                              <option value="pt">Português</option>
                                                              <option value="zh">中文</option>
                                                              <option value="ja">日本語</option>
                                                              <option value="ru">Русский</option>
                                                            </select>
                                                            
                                                            <label class="tool-label">Group</label>
                                                            <select id="obGroup" class="audience-select" style="width:100%; margin-bottom:20px;">
                                                              <option value="External Innovation">External Innovation</option>
                                                              <option value="Commercial">Commercial</option>
                                                            </select>
                                                            `;

  // 3. Monitor Setup
  body.innerHTML += `
                                                            <label class="tool-label">Screen Setup</label>
                                                            <div style="display:flex; gap:10px; margin-bottom:10px;">
                                                              <label class="checkbox-wrapper" style="flex:1; background:#f5f5f5; padding:10px; border:1px solid #ddd; border-radius:4px; cursor:pointer;" onclick="selectObMonitor(this, 'regular')">
                                                                <input type="radio" name="obMonitor" value="regular" checked style="transform:scale(1.2); margin-right:5px;">
                                                                  <span style="font-size:11px; font-weight:600;">🖥️ Monitor</span>
                                                              </label>
                                                              <label class="checkbox-wrapper" style="flex:1; background:#f5f5f5; padding:10px; border:1px solid #ddd; border-radius:4px; cursor:pointer;" onclick="selectObMonitor(this, 'laptop')">
                                                                <input type="radio" name="obMonitor" value="laptop" style="transform:scale(1.2); margin-right:5px;">
                                                                  <span style="font-size:11px; font-weight:600;">💻 Laptop Mode</span>
                                                              </label>
                                                            </div>

                                                            <div style="font-size:10px; color:#888; line-height:1.3; margin-bottom:20px;">
                                                              Laptop Mode enables compact icons and auto-hides the toolbar to save space.
                                                            </div>
                                                            `;

  // Error Message Container
  var errorMsg = document.createElement("div");
  errorMsg.id = "obErrorMsg";
  errorMsg.style.cssText = "color:#d32f2f; font-size:11px; font-weight:700; text-align:center; margin-bottom:10px; display:none;";
  body.appendChild(errorMsg);

  // Next Button
  var btnNext = document.createElement("button");
  btnNext.className = "btn-platinum-config primary";
  btnNext.style.width = "100%";
  btnNext.style.fontWeight = "800";
  btnNext.innerText = "NEXT: WELCOME PAGE →";
  btnNext.onclick = saveOnboardingSetup;

  body.appendChild(btnNext);
  card.appendChild(body);
  overlay.appendChild(card);
}

function saveOnboardingSetup() {
  var nameEl = document.getElementById("obName");
  var emailEl = document.getElementById("obEmail");
  var errEl = document.getElementById("obErrorMsg");

  // Reset Styles
  nameEl.style.border = "1px solid #ccc";
  emailEl.style.border = "1px solid #ccc";
  errEl.style.display = "none";

  var name = nameEl.value.trim();
  var email = emailEl.value.trim();

  if (!name) {
    nameEl.style.border = "2px solid #d32f2f";
    errEl.innerText = "Please enter your Full Name.";
    errEl.style.display = "block";
    return;
  }
  if (!email) {
    emailEl.style.border = "2px solid #d32f2f";
    errEl.innerText = "Please enter your Email Address.";
    errEl.style.display = "block";
    return;
  }

  var lang = document.getElementById("obLang").value;
  var group = document.getElementById("obGroup").value;
  var monitorEls = document.getElementsByName("obMonitor");
  var mode = "regular";
  for (var i = 0; i < monitorEls.length; i++) { if (monitorEls[i].checked) mode = monitorEls[i].value; }

  // Save Settings
  if (!userSettings.emailTemplates) userSettings.emailTemplates = {};
  userSettings.emailTemplates.myName = name;
  userSettings.userEmail = email;
  userSettings.userGroup = group;

  localStorage.setItem("eli_language", lang);
  if (typeof setLanguage === "function") setLanguage(lang);

  if (mode === "laptop") {
    // Laptop Mode: Unpin Toolbar (Save Space) + Compact Mode
    userSettings.pinToolbar = false;
    var chkPin = document.getElementById("chkPinToolbar");
    if (chkPin) { chkPin.checked = false; if (chkPin.onchange) chkPin.onchange(); }

    var chkLap = document.getElementById("chkLaptopMode");
    if (chkLap) { chkLap.checked = true; if (chkLap.onchange) chkLap.onchange(); }
  } else {
    // Monitor Mode: Pin Toolbar (Easy Access) + Normal Mode
    userSettings.pinToolbar = true;
    var chkPin = document.getElementById("chkPinToolbar");
    if (chkPin) { chkPin.checked = true; if (chkPin.onchange) chkPin.onchange(); }

    var chkLap = document.getElementById("chkLaptopMode");
    if (chkLap) { chkLap.checked = false; if (chkLap.onchange) chkLap.onchange(); }
  }

  saveCurrentSettings(true); // Intermediate save
  showOnboardingTour(name);
}

function showOnboardingTour(fullName) {
  var firstName = fullName.split(" ")[0];

  // Reuse Overlay
  var overlay = document.getElementById("onboardingOverlay");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "onboardingOverlay";
    overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(255,255,255,0.95); z-index:20000; display:flex; justify-content:center; align-items:center; overflow-y:auto;";
    document.body.appendChild(overlay);
  }
  overlay.innerHTML = "";

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "width:95%; max-width:700px; min-height:600px; padding:0; border:1px solid #d1d5db; background:#fff; border-radius:0px; overflow:hidden; border-top:5px solid #d32f2f; box-shadow: 0 10px 30px rgba(0,0,0,0.25);";

  var header = document.createElement("div");
  header.style.cssText = "padding:20px; text-align:center; background:#fff5f5;";
  header.innerHTML = `
                                                            <div style="font-size:32px; margin-bottom:10px;">✨</div>
                                                            <div style="font-size:18px; font-weight:800; color:#c62828; margin-bottom:5px;">Welcome, ${firstName}!</div>
                                                            <div style="font-size:12px; color:#546e7a;">Here are a few quick tips to help you get started.<br><span style="font-size:10px; opacity:0.8;">(All settings from the previous screen can be changed in the main settings menu.)</span></div>
                                                            `;
  card.appendChild(header);

  var body = document.createElement("div");
  body.style.padding = "20px";
  body.style.fontSize = "12px";
  body.style.color = "#333";
  body.style.lineHeight = "1.6";

  var tips = [
    { icon: "🟥", text: "If you get stuck, click the <b>Red ELI Bar</b> at the top to return Home anytime." },
    { icon: "🔄", text: "Need help? Click the <b>ELI Spinning Logo</b> in the top left (or the mini '?' icon in a button) for Help." },
    { icon: "⚙️", text: "Want to customize? Click the <b>Gear Icon</b> in the main bar or in a button to access Settings and the ELI Store. There, you can fully customize your home screen, toolbar and button order as well as add new apps." },
    { icon: "👇", text: "Use <b>Draft & Review</b> modes at the bottom for tailored workflows. These can be further customized." },
    { icon: "🦶", text: "The <b>Footer</b> is where you upload your data and save/view versions of your agreement. Click <b>Data</b> to upload Agreement Data (e.g. from OARS), while the 📷 takes a <b>Snapshot</b> to save to your <b>Versions</b> list." }
  ];

  tips.forEach(t => {
    var row = document.createElement("div");
    row.style.cssText = "display:flex; gap:10px; margin-bottom:15px; align-items:flex-start;";
    row.innerHTML = `
                                                            <div style="font-size:16px; min-width:20px;">${t.icon}</div>
                                                            <div>${t.text}</div>
                                                            `;
    body.appendChild(row);
  });

  var btnFinish = document.createElement("button");
  btnFinish.className = "btn-platinum-config primary";
  btnFinish.style.width = "100%";
  btnFinish.style.fontWeight = "800";
  btnFinish.innerText = "GET STARTED 🚀";
  btnFinish.onclick = finishOnboarding;

  body.appendChild(btnFinish);
  card.appendChild(body);
  overlay.appendChild(card);
}

function finishOnboarding() {
  userSettings.hasOnboarded = true;
  saveCurrentSettings(true);

  // Remove Overlay
  var overlay = document.getElementById("onboardingOverlay");
  if (overlay) overlay.remove();

  showToast("✅ You're all set!");
  updateFooterName();
  showHome();
}

function selectObMonitor(el, val) {
  var radios = document.getElementsByName("obMonitor");
  for (var i = 0; i < radios.length; i++) {
    radios[i].parentElement.style.background = '#f5f5f5';
    radios[i].parentElement.style.borderColor = '#ddd';
  }
  el.style.background = '#e3f2fd';
  el.style.borderColor = '#2196f3';
  el.querySelector('input').checked = true;
}

function updateFooterName() {
  var span = document.getElementById("footerMyName");
  var headerSpan = document.getElementById("headerMyName");
  var name = (userSettings.emailTemplates && userSettings.emailTemplates.myName) ? userSettings.emailTemplates.myName : "User";

  if (span) span.innerText = " " + name;
  if (headerSpan) headerSpan.innerText = name;
}

// ==========================================
// START: updatePlaybookButtonState
// ==========================================
function updatePlaybookButtonState() {
  var btn = document.getElementById("btnPlaybookAnalyze");
  if (!btn) return;

  var count = document.querySelectorAll(".playbook-check:checked").length;
  if (count === 0) {
    btn.disabled = true;
    btn.style.opacity = "0.5";
    btn.style.cursor = "not-allowed";
  } else {
    btn.disabled = false;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
  }
}
// ==========================================
// END: updatePlaybookButtonState
// ==========================================
// ==========================================
// START: Document Viewer (Movable Dialog)
// ==========================================
window.openDocumentViewer = async function () {
  // 1. Create Modal (if not exists)
  var id = "docViewerModal";
  var modal = document.getElementById(id);
  if (!modal) {
    modal = document.createElement("div");
    modal.id = id;
    modal.style.cssText = "position:fixed; top:80px; left:80px; width:700px; height:650px; background:white; border:1px solid #455a64; box-shadow:0 25px 50px rgba(0,0,0,0.5); z-index:99999; display:flex; flex-direction:column; resize:both; overflow:hidden; border-radius:6px;";

    // Header
    var header = document.createElement("div");
    header.style.cssText = "padding:10px 15px; background:#263238; color:white; cursor:move; display:flex; justify-content:space-between; align-items:center; user-select:none; border-bottom:1px solid #37474f;";
    header.innerHTML = `<div style="display:flex;align-items:center;gap:6px;">
                                                              <span style="font-size:16px;">📄</span>
                                                              <span style="font-weight:700; font-size:12px; letter-spacing:0.5px; text-transform:uppercase;">Document Viewer</span>
                                                            </div>
                                                            <button onclick="document.getElementById('${id}').style.display='none'" style="border:none; background:none; cursor:pointer; color:#b0bec5; font-weight:bold; font-size:16px;">✕</button>`;

    // Drag Logic
    let isDragging = false;
    let startX, startY, initialLeft, initialTop;
    header.onmousedown = function (e) {
      if (e.target.tagName === 'BUTTON') return;
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      initialLeft = modal.offsetLeft;
      initialTop = modal.offsetTop;
      document.onmousemove = function (e) {
        if (!isDragging) return;
        e.preventDefault();
        modal.style.left = (initialLeft + e.clientX - startX) + "px";
        modal.style.top = (initialTop + e.clientY - startY) + "px";
      };
      document.onmouseup = function () { isDragging = false; document.onmousemove = null; };
    };

    modal.appendChild(header);

    // Toolbar
    var toolRow = document.createElement("div");
    toolRow.style.cssText = "padding:8px 12px; background:#eceff1; border-bottom:1px solid #cfd8dc; display:flex; gap:8px; align-items:center;";
    toolRow.innerHTML = `
                                                            <button id="btnViewFull" onclick="window.viewDocTab('full')" class="btn-tab-active" style="padding:4px 8px; border:1px solid #b0bec5; cursor:pointer; font-size:11px; font-weight:700; border-radius:3px;">📄 Full Document</button>
                                                            <button id="btnViewAI" onclick="window.viewDocTab('ai')" class="btn-tab" style="padding:4px 8px; border:1px solid transparent; cursor:pointer; font-size:11px; font-weight:700; color:#546e7a; background:transparent;">⚡ Formatted Data</button>
                                                            <div style="flex:1;"></div>
                                                            <button onclick="window.refreshDocViewer()" style="border:none; background:none; cursor:pointer; font-size:14px;" title="Reload from DB">🔄</button>
                                                            `;
    modal.appendChild(toolRow);

    // Content
    var content = document.createElement("div");
    content.id = id + "Content";
    content.style.cssText = "flex:1; overflow:auto; padding:0; font-family:'Calibri', 'Inter', sans-serif; font-size:14px; background:#f5f5f5; line-height:1.5; position:relative;";
    modal.appendChild(content);

    document.body.appendChild(modal);
  } else {
    modal.style.display = "flex";
  }

  // --- STATE MANAGEMENT ---
  window.currentViewMode = "ai"; // Default to 'ai' (Smart Changes)

  window.refreshDocViewer = function () {
    // Logic below
    window.viewDocTab(window.currentViewMode || 'ai', true);
  };

  window.viewDocTab = function (mode, forceRefresh) {
    window.currentViewMode = mode;

    // Update UI Tabs
    var btnFull = document.getElementById("btnViewFull");
    var btnAI = document.getElementById("btnViewAI");

    if (mode === 'full') {
      btnFull.style.cssText = "padding:4px 8px; border:1px solid #b0bec5; cursor:pointer; font-size:11px; font-weight:700; border-radius:3px; background:#fff; color:#000;";
      btnAI.style.cssText = "padding:4px 8px; border:1px solid transparent; cursor:pointer; font-size:11px; font-weight:700; color:#546e7a; background:transparent;";
      loadAndRenderFull(forceRefresh);
    } else {
      btnAI.style.cssText = "padding:4px 8px; border:1px solid #1565c0; cursor:pointer; font-size:11px; font-weight:700; border-radius:3px; background:#e3f2fd; color:#1565c0;";
      btnFull.style.cssText = "padding:4px 8px; border:1px solid transparent; cursor:pointer; font-size:11px; font-weight:700; color:#546e7a; background:transparent;";
      loadAndRenderAI(forceRefresh);
    }
  };

  // --- RENDER 1: FULL DOCUMENT (MAMMOTH) ---
  async function loadAndRenderFull(force) {
    var out = document.getElementById(id + "Content");
    // If already has content and not forced, skip
    if (!force && out.dataset.mode === 'full' && out.innerHTML.length > 100) return;

    out.dataset.mode = 'full';
    out.innerHTML = `<div style="text-align:center; padding:40px; color:#546e7a;">
                                                              <div class="spinner" style="width:24px;height:24px;border-width:3px;display:inline-block;margin-bottom:10px;"></div><br>
                                                                Loading Full Document...
                                                            </div>`;

    try {
      var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';
      var record = await PROMPTS.loadDocFromDB(uniqueKey);
      if (!record || !record.data) { showEmptyState(out); return; }

      if (record.name.toLowerCase().endsWith(".docx")) {
        if (typeof mammoth === "undefined") {
          out.innerHTML = `<div style="color:red; text-align:center; padding:20px;">❌ Mammoth.js library missing.</div>`;
          return;
        }
        var result = await mammoth.convertToHtml({ arrayBuffer: record.data });
        out.innerHTML = `<div style="background:white; padding:50px; box-shadow:0 2px 10px rgba(0,0,0,0.1); min-height:800px; max-width:850px; margin:20px auto; color:#000;">
                                                              ${result.value}
                                                            </div>`;
      } else {
        // Fallback text
        var dec = new TextDecoder("utf-8");
        var text = dec.decode(record.data);
        out.innerHTML = `<pre style="background:white; padding:20px; white-space:pre-wrap; margin:0;">${text}</pre>`;
      }
    } catch (e) {
      out.innerHTML = `<div style="color:red; padding:20px;">Error: ${e.message}</div>`;
    }
  }

  // --- RENDER 2: SMART AI CHANGES ---
  async function loadAndRenderAI(force) {
    var out = document.getElementById(id + "Content");
    if (!force && out.dataset.mode === 'ai' && out.innerHTML.length > 100) return;

    out.dataset.mode = 'ai';
    out.innerHTML = `<div style="text-align:center; padding:40px; color:#1565c0;">
                                                              <div class="spinner" style="width:24px;height:24px;border-width:3px;border-color:#1565c0 transparent #1565c0 transparent; display:inline-block;margin-bottom:10px;"></div><br>
                                                                <strong>Analyze & Filter Rows</strong><br>
                                                                  <span style="font-size:11px; opacity:0.8;">Smart-Parsing Document Structure with Flash AI...</span>
                                                                </div>`;

    try {
      var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';
      var record = await PROMPTS.loadDocFromDB(uniqueKey);
      if (!record || !record.data) { showEmptyState(out); return; }

      // --- 0. CACHE HIT CHECK (Instant Load) ---
      if (record.aiCache && Array.isArray(record.aiCache)) {
        console.log("⚡ INSTANT LOAD: Using Background AI Cache");
        renderAITable(record.aiCache);
        out.innerHTML += `<div style="text-align:center; font-size:10px; color:#aaa; margin-top:5px;">⚡ Pre-loaded from Background Analysis</div>`;
        return;
      }

      if (!record.name.toLowerCase().endsWith(".docx") && !record.name.toLowerCase().endsWith(".pdf")) {
        out.innerHTML = `<div style="padding:40px; text-align:center;">⚠️ AI Analysis only available for DOCX/PDF.</div>`;
        return;
      }

      // 1. STRATEGY CHANGE: Convert DOCX to HTML for faster, robust text analysis
      var contextContent = "";
      var mimeToSend = 'text/html';

      if (record.name.toLowerCase().endsWith(".docx")) {
        if (typeof mammoth === "undefined") {
          out.innerHTML = `<div style="color:red; text-align:center;">❌ Mammoth.js library missing.</div>`;
          return;
        }
        try {
          var conversion = await mammoth.convertToHtml({ arrayBuffer: record.data });
          contextContent = conversion.value;
          contextContent = contextContent.replace(/<table /g, '<table border="1"');
        } catch (err) {
          out.innerHTML = `<div style="color:red;">Mammoth Error: ${err.message}</div>`;
          return;
        }
      } else if (record.name.toLowerCase().endsWith(".pdf")) {
        var blob = new Blob([record.data], { type: "application/pdf" });
        var reader = new FileReader();
        reader.onloadend = function () {
          runGeminiAnalysis(reader.result.split(',')[1], 'application/pdf');
        };
        reader.readAsDataURL(blob);
        return;
      } else {
        var dec = new TextDecoder("utf-8");
        contextContent = dec.decode(record.data);
        mimeToSend = 'text/plain';
      }

      runGeminiAnalysis(contextContent, mimeToSend);

      async function runGeminiAnalysis(dataContent, mimeType) {
        const isText = mimeType.startsWith('text');
        const prompt = `
                                                                Analyze this document table/content.
                                                                TASK: Extract pairs of Field (Left) and Data (Right) from the document.

                                                                RETURN JSON ARRAY:
                                                                [ {"field": "Field Name / Original Text", "data": "Data Value / Modified Text" } ]

                                                                - Exclude rows where "Data" is empty.
                                                                - For tabular docs, Left col = Field, Right col = Data.
                                                                `;

        try {
          let parts;
          if (isText) {
            var safeText = dataContent.substring(0, 1000000);
            parts = [{ text: prompt + "\n\n--- DOCUMENT ---\n" + safeText }];
          } else {
            parts = [{ text: prompt }, { inlineData: { mimeType: mimeType, data: dataContent } }];
          }

          // FORCE FLASH MODEL for Speed
          // Reverted to default (2.5) per user request as 1.5 is 404
          var result = await callGemini(API_KEY, parts);
          var json = tryParseGeminiJSON(result);

          if (Array.isArray(json) && json.length > 0) {
            var table = `<div style="max-width:850px; margin:20px auto; background:white; box-shadow:0 2px 5px rgba(0,0,0,0.1);">
                                                                  <table style="width:100%; border-collapse:collapse; font-size:12px;">
                                                                    <tr style="background:#eceff1; border-bottom:2px solid #cfd8dc;">
                                                                      <th style="padding:12px; width:40%; text-align:left; color:#455a64; border-right:1px solid #eee;">FIELD</th>
                                                                      <th style="padding:12px; width:60%; text-align:left; color:#1565c0;">DATA</th>
                                                                    </tr>`;

            json.forEach((row, i) => {
              var bg = i % 2 === 0 ? '#fff' : '#fafafa';
              // Handle potential key variations from AI
              var left = row.field || row.left || row.Field || "";
              var right = row.data || row.right || row.Data || "";

              table += `<tr style="background:${bg}; border-bottom:1px solid #eee;">
                                                                      <td style="padding:12px; vertical-align:top; border-right:1px solid #eee; color:#37474f; font-weight:600;">${left}</td>
                                                                      <td style="padding:12px; vertical-align:top; color:#263238; background:#fff;">${right}</td>
                                                                    </tr>`;
            });
            table += "</table></div>";
            out.innerHTML = table;
          } else {
            out.innerHTML = `<div style="padding:40px; text-align:center;">
                        ✅ <strong>No Data Found</strong><br>
                        AI didn't detect any Field/Data pairs.
                     </div>`;
          }
        } catch (apiErr) {
          out.innerHTML = `<div style="color:red; padding:20px;">AI Analysis Error: ${apiErr.message}</div>`;
        }
      };
    } catch (e) {
      out.innerHTML = `<div style="color:red; padding:20px;">System Error: ${e.message}</div>`;
    }

    function renderAITable(json) {
      var table = `<div style="max-width:850px; margin:20px auto; background:white; box-shadow:0 2px 5px rgba(0,0,0,0.1);">
                                                                  <table style="width:100%; border-collapse:collapse; font-size:12px;">
                                                                    <tr style="background:#eceff1; border-bottom:2px solid #cfd8dc;">
                                                                      <th style="padding:12px; width:40%; text-align:left; color:#455a64; border-right:1px solid #eee;">FIELD</th>
                                                                      <th style="padding:12px; width:60%; text-align:left; color:#1565c0;">DATA</th>
                                                                    </tr>`;

      json.forEach((row, i) => {
        var bg = i % 2 === 0 ? '#fff' : '#fafafa';
        var left = row.field || row.left || row.Field || "";
        var right = row.data || row.right || row.Data || "";

        table += `<tr style="background:${bg}; border-bottom:1px solid #eee;">
                                                                      <td style="padding:12px; vertical-align:top; border-right:1px solid #eee; color:#37474f; font-weight:600;">${left}</td>
                                                                      <td style="padding:12px; vertical-align:top; color:#263238; background:#fff;">${right}</td>
                                                                    </tr>`;
      });
      table += "</table></div>";
      out.innerHTML = table;
    }

    function renderNoData() {
      out.innerHTML = `<div style="padding:40px; text-align:center;">
                    ✅ <strong>No Data Found</strong><br>
                    AI didn't detect any Field/Data pairs.
                 </div>`;
    }
  }

  function showEmptyState(container) {
    container.innerHTML = `<div style="text-align:center; padding:40px; color:#c62828;">
        ⚠️ No document loaded.<br><span style="font-size:11px;">Load a file in the footer first.</span>
      </div>`;
  }

  // Trigger Default Load
  window.viewDocTab('ai'); // Default to AI
  // --- GLOBAL FIX: ENSURE FOOTER CLOSE WIPES STATE ---
  // --- GLOBAL FIX: ENSURE FOOTER CLOSE WIPES STATE (SAFER VERSION) ---
  if (typeof document !== 'undefined') {
    document.body.addEventListener('click', function (e) {
      var target = e.target;
      if (target.id === 'btnDeleteFile' || target.closest('#btnDeleteFile')) {

        console.log("🗑️ Footer Delete Clicked. Wiping DB & State...");

        // 1. Wipe IndexedDB
        var uniqueKey = window.CURRENT_DOC_ID || 'currentDoc';

        if (typeof PROMPTS !== 'undefined' && PROMPTS.deleteDocFromDB) {
          PROMPTS.deleteDocFromDB(uniqueKey);
          PROMPTS.saveAICacheToDB(uniqueKey, null);
        }

        // 2. Wipe Office Settings (Persistent)
        try {
          if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            Office.context.document.settings.remove("AgreementDataDetails");
            Office.context.document.settings.saveAsync();
          }
        } catch (err) { console.error(err); }

        // 3. Update UI Visuals immediately
        if (window.clearLoadedFile) window.clearLoadedFile();
      }
    });
  }
};
// ==========================================
// REPLACE: renderVersionTimeline (HTML Only - View & Compare)
// ==========================================
function renderVersionTimeline(container) {
  container.innerHTML = `<div class="spinner" style="width:16px;height:16px;border-width:2px;border-color:#1565c0 transparent #1565c0 transparent;"></div> Loading...`;

  getHistoryFromDoc().then((versions) => {
    container.innerHTML = "";

    if (!versions || versions.length === 0) {
      container.innerHTML = `<div style="text-align:center; padding:20px; color:#546e7a; font-size:11px;">No history saved.</div>`;
      return;
    }

    // Sort & Number (Newest First)
    versions.sort((a, b) => a.timestamp - b.timestamp);
    versions.forEach((v, index) => { v.displayNumber = index + 1; });
    versions.reverse();

    versions.forEach((v) => {
      let node = document.createElement("div");
      node.className = "timeline-node slide-in";
      node.style.cssText = `background: #fff; border: 1px solid #e0e0e0; border-left: 4px solid #1565c0; padding: 10px; margin-bottom: 8px; position: relative;`;

      let dateStr = new Date(v.timestamp).toLocaleString();
      let activeBadge = (v.sessionId === CURRENT_SESSION_ID) ? `<span style="background:#e3f2fd; color:#1565c0; font-weight:bold; font-size:9px; padding:2px 4px; border-radius:2px; margin-left:5px;">CURRENT</span>` : "";

      // Header
      let headerRow = document.createElement("div");
      headerRow.style.cssText = "display:flex; justify-content:space-between; align-items:start; margin-bottom:6px;";

      let textDiv = document.createElement("div");
      textDiv.innerHTML = `
          <div style="font-size:11px; font-weight:800; color:#1565c0;">
             <span style="color:#546e7a; margin-right:4px;">#${v.displayNumber}</span> ${v.label} ${activeBadge}
          </div>
          <div style="font-size:10px; color:#78909c;">${dateStr}</div>
      `;
      headerRow.appendChild(textDiv);

      // Delete Button
      let btnDiv = document.createElement("div");
      let delBtn = document.createElement("button");
      delBtn.innerHTML = "🗑️";
      delBtn.style.cssText = "background:transparent; border:none; cursor:pointer; font-size:12px; padding:4px; position:relative; z-index:999;";

      delBtn.onclick = function (e) {
        e.preventDefault(); e.stopPropagation();
        let overlay = document.createElement("div");
        overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;";
        let dialog = document.createElement("div");
        dialog.style.cssText = "background:white; padding:20px; border-radius:4px; width:250px; text-align:center; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-left: 5px solid #c62828;";
        dialog.innerHTML = `<div style="font-weight:bold; color:#c62828; margin-bottom:10px;">Delete Version #${v.displayNumber}?</div><div style="font-size:12px; color:#555; margin-bottom:15px;">Permanent action.</div><div style="display:flex; gap:10px;"><button id="btnConfDel" style="flex:1; background:#c62828; color:white; border:none; padding:8px; font-weight:bold; cursor:pointer;">Delete</button><button id="btnCancelDel" style="flex:1; background:#eee; border:none; padding:8px; cursor:pointer;">Cancel</button></div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        document.getElementById("btnCancelDel").onclick = function () { overlay.remove(); };
        document.getElementById("btnConfDel").onclick = async function () {
          overlay.remove();
          let newHist = versions.filter(x => x.timestamp !== v.timestamp);
          newHist.forEach(h => delete h.displayNumber);
          await saveHistoryToDoc(newHist);
          renderVersionTimeline(container);
          showToast("🗑️ Version Deleted");
        };
      };
      btnDiv.appendChild(delBtn);
      headerRow.appendChild(btnDiv);
      node.appendChild(headerRow);

      // --- ACTION ROW (HTML ONLY) ---
      let actionRow = document.createElement("div");
      actionRow.style.cssText = "display:flex; gap:6px;";
      // 1. VIEW BUTTON (The Fix)
      let btnView = document.createElement("button");
      btnView.className = "btn-platinum-config";
      btnView.style.cssText = "flex:1; font-size:9px; padding:6px; font-weight:700;";
      btnView.innerText = "👁 VIEW";

      btnView.onclick = function () {
        if (v.html) {
          // Use new Safe Modal
          showPreviewModal(`Snapshot: ${v.label}`, v.html);
        } else {
          // Fallback
          let textPreview = cleanOoxmlForAI(v.data);
          let safeContent = `<pre style="white-space:pre-wrap; font-family:monospace;">${textPreview}</pre>`;
          showPreviewModal(`Snapshot: ${v.label}`, safeContent);
        }
      };

      // 2. COMPARE BUTTON
      let btnComp = document.createElement("button");
      btnComp.className = "btn-platinum-config primary";
      btnComp.style.cssText = "flex:1; font-size:9px; font-weight:800; padding:6px;";
      btnComp.innerText = "↔ COMPARE";
      btnComp.onclick = function () { runVersionCompare(v); };

      actionRow.appendChild(btnView);
      actionRow.appendChild(btnComp);
      node.appendChild(actionRow);

      container.appendChild(node);
    });
  });
}
// ==========================================
// REPLACE: runVersionCompare (Fix: Consistent Extraction & Numbering Guardrail)
// ==========================================
function runVersionCompare(historicalVersion) {
  setLoading(true, "redline");

  // 1. Get Current Text using HTML (Same method as historical snapshots)
  // This ensures consistent formatting so the AI doesn't see false "changes".
  Word.run(async (context) => {
    let body = context.document.body;
    let html = body.getHtml(); // <--- CHANGE: Get HTML to preserve structure parity
    await context.sync();
    return html.value;
  })
    .then(currentHtml => {

      // 2. Process both through the SAME cleaner
      // This normalizes newlines and structure for both sides
      let currentText = getCleanTextAndContext(currentHtml).text;
      let oldText = getCleanTextAndContext(historicalVersion.html).text;

      // --- FIX: Dynamic Labels ---
      let histLabel = historicalVersion.displayNumber
        ? "Version #" + historicalVersion.displayNumber
        : "Historical Snapshot";

      let currentLabel = "Current Active Document";

      // 3. Ask AI (With reinforced instructions about Numbering)
      let prompt = `
      You are a Senior Legal Editor.
      
      TASK: Compare two document versions.
      1. NAME A: "${histLabel}"
      2. NAME B: "${currentLabel}"
      
      INPUT TEXT FOR "${histLabel}":
      """${oldText.substring(0, 500000)}"""

      INPUT TEXT FOR "${currentLabel}":
      """${currentText.substring(0, 500000)}"""
      
      CRITICAL ANALYSIS RULES:
      1. IGNORE MISSING NUMBERING: The text extraction process often strips automatic list numbers (e.g. "1.", "2.1"). **DO NOT** report "removed numbering" or "continuous block of text" as a change. Assume the numbering exists in the source.
      2. FOCUS ON MEANING: Only report changes to words, definitions, obligations, or financial terms.
      3. IGNORE WHITESPACE: Ignore extra newlines or spacing differences.
      
      OUTPUT FORMAT (STRICT JSON ONLY):
      {
        "summary": "Start with: Comparing ${histLabel} vs ${currentLabel}...",
        "changes": [
           { "clause": "Clause Name (e.g. Indemnity)", "change": "Description of change...", "impact": "High Risk" | "Neutral" | "Better" }
        ]
      }
    `;

      return callGemini(API_KEY, prompt).then(result => {
        let analysis = tryParseGeminiJSON(result);
        renderVersionReport(analysis, historicalVersion.label, oldText, currentText);
      });
    })
    .catch(handleError)
    .finally(() => setLoading(false));
}
// ==========================================
// REPLACE: renderVersionReport (With Deep Analysis Button)
// ==========================================
function renderVersionReport(data, verLabel, oldText, currentText) {
  // --- SAFETY PATCH START ---
  if (!data) data = {};
  if (!data.summary) data.summary = "Analysis unavailable.";
  if (!Array.isArray(data.changes)) data.changes = [];
  let out = document.getElementById("output");

  // 1. CLEAR THE SCREEN (Hides the default Compare UI)
  out.innerHTML = "";

  // 2. Header
  let header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 12px;`;
  header.innerHTML = `<div style="font-family:'Georgia', serif; font-size:16px; color:#fff;">🕒 Version Report: ${verLabel}</div>`;
  out.appendChild(header);

  // 3. The Summary Table Card
  let card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "border-left:5px solid #2e7d32; background:#fff; padding:15px; margin-bottom: 15px;";

  card.innerHTML = `
       <div style="font-weight:800; color:#2e7d32; font-size:11px; margin-bottom:8px;">EXECUTIVE SUMMARY</div>
       <div style="font-size:13px; color:#333; margin-bottom:15px; line-height:1.5;">${data.summary}</div>
       
       <table style="width:100%; border-collapse:collapse; font-size:11px;">
         <tr style="background:#f5f5f5; text-align:left; border-bottom:1px solid #e0e0e0;">
            <th style="padding:8px; color:#546e7a;">Clause</th>
            <th style="padding:8px; color:#546e7a;">Change</th>
            <th style="padding:8px; color:#546e7a;">Impact</th>
         </tr>
         ${data.changes.map(c => `
            <tr style="border-bottom:1px solid #eee;">
               <td style="padding:8px; font-weight:bold; color:#37474f;">${c.clause}</td>
               <td style="padding:8px; color:#455a64;">${c.change}</td>
               <td style="padding:8px; font-weight:700; color:${(c.impact || '').toLowerCase().includes('risk') ? '#c62828' : '#2e7d32'}">${c.impact}</td>
            </tr>
         `).join('')}
       </table>
  `;
  out.appendChild(card);

  // 4. "Deep Analysis" Button
  let btnDeep = document.createElement("button");
  btnDeep.className = "btn-platinum-config primary";
  btnDeep.style.cssText = "width:100%; margin-bottom:10px; padding:12px; font-size:12px;";
  btnDeep.innerHTML = "⚡ RUN DEEP ANALYSIS AGAINST SAVED VERSION";

  btnDeep.onclick = function () {
    // Calls your existing "Option 2B" logic directly with the text we already have
    // passing 'oldText' as the uploaded content, and "old" mode means "Compare Doc (New) vs This (Old)"

    // We manually trigger the loading state here because runCompareAnalysis expects to do it
    setLoading(true, "default");

    // Call the main compare engine
    // Note: We skip 'validateApiAndGetText' inside runCompareAnalysis by modifying it, 
    // OR we just call the API directly here to be safer.

    // Let's call the API directly to ensure we use the exact text we have here
    callGemini(API_KEY, PROMPTS.COMPARE(oldText, currentText, "old"))
      .then(result => {
        try {
          let deepData = tryParseGeminiJSON(result);
          renderCompareResults(deepData, "old");
        } catch (e) {
          showOutput("Error parsing deep analysis: " + e.message, true);
        }
      })
      .catch(handleError)
      .finally(() => setLoading(false));
  };

  out.appendChild(btnDeep);
  // --- NEW: ANALYZE TRACK CHANGES BUTTON ---
  let btnTrack = document.createElement("button");
  btnTrack.className = "btn-platinum-config";
  // Styled slightly differently (Red/Pink accent) to distinguish from Deep Analysis
  btnTrack.style.cssText = "width:100%; margin-bottom:10px; padding:12px; font-size:12px; color:#b71c1c; border-color:#ffcdd2;";
  btnTrack.innerHTML = "📝 ANALYZE TRACK CHANGES";

  btnTrack.onclick = function () {
    // Trigger existing function
    runRedlineSummary(this);
  };
  out.appendChild(btnTrack);

  // 5. Back Button
  let btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "⬅ Back to Timeline";
  btnBack.style.width = "100%";
  btnBack.onclick = function () { showCompareInterface(); };
  out.appendChild(btnBack);
}
// Helper to strip tags for AI processing
function cleanOoxmlForAI(ooxml) {
  // Simple regex to remove XML tags. 
  // For production, a DOMParser is better, but this works for basic text extraction.
  return ooxml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}
// ==========================================
// HELPER: Get/Create Unique Document ID
// ==========================================
function getUniqueDocumentId() {
  return new Promise((resolve) => {
    // Check if ID exists in Document Settings (saved inside the DOCX)
    var docId = Office.context.document.settings.get("eli_doc_id");
    if (docId) {
      resolve(docId);
    } else {
      // Generate New ID
      var newId = "doc_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      Office.context.document.settings.set("eli_doc_id", newId);
      Office.context.document.settings.saveAsync(function (result) {
        resolve(newId);
      });
    }
  });
}
// ==========================================
// HELPER: Create Snapshot (Capped at Max 5 Sessions)
function createSnapshot(label, container, callback, forceNew) {
  if (container) {
    container.innerHTML = `<div class="spinner" style="width:16px;height:16px;border-width:2px;border-color:#1565c0 transparent #1565c0 transparent; display:inline-block;"></div> Saving...`;
  }

  Word.run(async (ctx) => {
    let body = ctx.document.body;
    let html = body.getHtml();
    await ctx.sync();
    return html.value;
  }).then(async (htmlContent) => {

    // 1. Get existing history
    let history = await getHistoryFromDoc();

    // 2. Check if we already have a save for THIS session
    let existingSessionIndex = history.findIndex(v => v.sessionId === CURRENT_SESSION_ID);

    let newEntry = {
      sessionId: CURRENT_SESSION_ID,
      timestamp: Date.now(),
      label: label || "Current Session",
      author: "Me",
      html: htmlContent
    };

    // LOGIC CHANGE: Only overwrite if it exists AND we are not forcing a new entry
    if (existingSessionIndex !== -1 && !forceNew) {
      // SCENARIO A: Standard Save -> Update the existing "Session" entry
      history[existingSessionIndex] = newEntry;
    } else {
      // SCENARIO B: New Session OR Forced Snapshot (e.g. Email/Save As)
      // This preserves the old one and appends this as a distinct milestone
      history.push(newEntry);
    }

    // 3. CAP THE HISTORY
    history.sort((a, b) => a.timestamp - b.timestamp);
    if (history.length > 5) {
      history = history.slice(history.length - 5);
    }

    // 4. Save back to document
    await saveHistoryToDoc(history);

    if (callback) callback();

  }).catch(e => {
    console.error("Snapshot failed:", e);
    if (container) container.innerHTML = "❌ Save Failed";
  });
}
// ==========================================
// REPLACE: runAutoBaseline (Respects Setting)
// ==========================================
function runAutoBaseline() {
  // 1. CHECK SETTING: If off, exit immediately
  if (!userSettings.autoSnapshotOnLoad) {
    console.log("ℹ️ Auto-Snapshot disabled in settings. Skipping.");
    return;
  }

  console.log("🚀 Checking document history...");

  getUniqueDocumentId().then(function (docId) {
    PROMPTS.getVersionHistory(docId).then(function (versions) {

      // 1. Generate Label with CURRENT Time
      let timeLabel = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      let finalLabel = "Auto-Save (" + timeLabel + ")";

      // 2. If Brand New (No history), create Baseline
      if (!versions || versions.length === 0) {
        createSnapshot(docId, "Original Baseline", null, function () {
          showToast("✅ Baseline Captured");
        });
        return;
      }

      // 3. If Returning (Check for changes)
      Word.run(async (ctx) => {
        let body = ctx.document.body;
        let html = body.getHtml();
        await ctx.sync();
        return html.value;
      }).then(currentContent => {

        // Find if we already have a save for TODAY (Current Session)
        let todayVersion = versions.find(v => v.sessionId === CURRENT_SESSION_ID);

        // If found, update it (Overwrite)
        if (todayVersion) {
          if (currentContent.length !== todayVersion.html.length) {
            console.log("ℹ️ Content changed. Updating snapshot...");
            createSnapshot(docId, finalLabel, null, function () {
              showToast("✅ Snapshot Updated");
            });
          }
        } else {
          // New Session -> New Entry
          createSnapshot(docId, finalLabel, null, function () {
            showToast("✅ Session Snapshot Saved");
          });
        }
      }).catch(e => console.warn("Auto-save check failed", e));
    });
  });
}
// ==========================================
// CUSTOM STORAGE ENGINE (SETTINGS API)
function getHistoryFromDoc() {
  return new Promise((resolve) => {
    try {
      // Access the robust Settings key-value store
      const raw = Office.context.document.settings.get("eli_version_history");
      if (raw) {
        // Determine if it's stringified JSON or an object
        const data = (typeof raw === 'string') ? JSON.parse(raw) : raw;
        resolve(Array.isArray(data) ? data : []);
      } else {
        resolve([]);
      }
    } catch (e) {
      console.warn("History Load Warning:", e);
      resolve([]);
    }
  });
}
// 2. SAVE HISTORY
function saveHistoryToDoc(historyArray) {
  return new Promise((resolve, reject) => {
    try {
      // 1. LIMIT SIZE: Settings has a 2MB limit. 
      // We keep only the last 3 snapshots to be safe.
      if (historyArray.length > 3) {
        historyArray = historyArray.slice(historyArray.length - 3);
      }

      // 2. OPTIMIZE: Strip huge Base64 images to save space
      // (We preserve text formatting, but drop heavy images)
      const optimizedHistory = historyArray.map(h => {
        if (h.html && h.html.length > 50000) {
          // Simple stripper for img tags with base64 data
          h.html = h.html.replace(/<img[^>]*src=["']data:image\/[^;]+;base64,[^"']+["'][^>]*>/gi, "[Image Removed]");
        }
        return h;
      });

      // 3. SAVE
      // Note: We don't need JSON.stringify if saving to Settings directly in some Office versions, 
      // but stringifying ensures consistency across platforms.
      Office.context.document.settings.set("eli_version_history", JSON.stringify(optimizedHistory));

      Office.context.document.settings.saveAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Settings Save Failed:", result.error.message);
          // Soft fail so the UI doesn't freeze
          resolve();
        } else {
          console.log("✅ History Saved to Settings");
          // UPDATE FOOTER COUNT INSTANTLY
          var el = document.getElementById("footerVerCount");
          if (el) el.innerText = "(" + optimizedHistory.length + ")";
          resolve();
        }
      });

    } catch (e) {
      //console.error("Save Logic Error:", e);
      showToast("⚠️ Error Saving History");
      resolve(); // Resolve anyway to unblock UI
    }
  });
}
// ==========================================
// HELPER: Floating Preview Window (Movable & Resizable)
// ==========================================
function showPreviewModal(title, contentHtml) {
  // 1. Close existing window if open (Singleton pattern)
  let existing = document.getElementById("floatPreviewWindow");
  if (existing) existing.remove();

  // 2. Create Floating Container
  let modal = document.createElement("div");
  modal.id = "floatPreviewWindow";
  modal.className = "slide-in";

  // CSS: Fixed position, Resize enabled, Shadow depth
  modal.style.cssText = `
        position: fixed; 
        top: 60px; 
        left: 30px; 
        width: 500px; 
        height: 600px; 
        background: white; 
        display: flex; 
        flex-direction: column; 
        box-shadow: 0 15px 50px rgba(0,0,0,0.3); 
        border: 1px solid #455a64; 
        z-index: 99999;
        resize: both; 
        overflow: hidden; /* Required for resize handle */
        min-width: 300px;
        min-height: 200px;
        border-radius: 4px;
    `;

  // 3. Header (The Drag Handle)
  let header = document.createElement("div");
  header.style.cssText = `
        padding: 10px 15px; 
        background: #263238; 
        color: white; 
        display: flex; 
        justify-content: space-between; 
        align-items: center; 
        border-bottom: 3px solid #c9a227; 
        cursor: move; 
        user-select: none;
        flex-shrink: 0;
    `;
  header.innerHTML = `
        <div style="font-weight:bold; font-size:12px; text-transform:uppercase; pointer-events:none;">👁 ${title}</div>
        <button id="btnCloseFloat" style="background:none; border:none; color:#b0bec5; font-size:16px; cursor:pointer; font-weight:bold;">✕</button>
    `;

  // 4. Content Area (Iframe)
  let contentArea = document.createElement("iframe");
  contentArea.style.cssText = "flex: 1; border: none; width: 100%; height: 100%; background: #f5f5f5;";

  modal.appendChild(header);
  modal.appendChild(contentArea);
  document.body.appendChild(modal);

  // 5. Inject Content
  let doc = contentArea.contentWindow.document;
  doc.open();
  doc.write(`
        <html>
        <head>
            <style>
                body { font-family: 'Segoe UI', sans-serif; padding: 25px; background: white; color: #333; line-height: 1.5; margin:0; }
                table { border-collapse: collapse; width: 100%; }
                td, th { border: 1px solid #ccc; padding: 5px; }
                h1, h2, h3 { color: #1565c0; margin-top: 0; }
                pre { background: #f5f5f5; padding: 10px; border: 1px solid #ddd; overflow-x: auto; }
            </style>
        </head>
        <body>${contentHtml}</body>
        </html>
    `);
  doc.close();

  // 6. Close Handler
  header.querySelector("#btnCloseFloat").onclick = function () {
    modal.remove();
  };

  // --- 7. DRAG LOGIC ---
  let isDragging = false;
  let startX, startY, initialLeft, initialTop;

  header.onmousedown = function (e) {
    // Ignore clicks on close button
    if (e.target.id === 'btnCloseFloat') return;

    isDragging = true;
    startX = e.clientX;
    startY = e.clientY;
    initialLeft = modal.offsetLeft;
    initialTop = modal.offsetTop;

    // TRICK: Create a transparent overlay to cover the iframe.
    // Otherwise, if mouse moves fast over the iframe, the iframe captures the event and drag stops.
    let blocker = document.createElement('div');
    blocker.id = 'dragBlocker';
    blocker.style.cssText = 'position:absolute; top:0; left:0; width:100%; height:100%; z-index:10000; cursor:move;';
    modal.appendChild(blocker);

    document.onmousemove = function (e) {
      if (!isDragging) return;
      e.preventDefault();
      modal.style.left = (initialLeft + e.clientX - startX) + "px";
      modal.style.top = (initialTop + e.clientY - startY) + "px";
    };

    document.onmouseup = function () {
      isDragging = false;
      document.onmousemove = null;
      document.onmouseup = null;
      let b = document.getElementById('dragBlocker');
      if (b) b.remove();
    };
  };
}

// ==========================================
// NEW: injectHistoryButton (Raised Position)
// ==========================================
function injectHistoryButton() {
  // 1. Find Existing HTML Element
  var btnSnap = document.getElementById("btnHeaderSnapshot");
  if (!btnSnap) return;

  // 2. Click Action
  btnSnap.onclick = function (e) {
    e.stopPropagation();
    e.preventDefault();
    if (typeof setActiveButton === 'function') setActiveButton("");
    if (typeof tryCollapseToolbar === 'function') tryCollapseToolbar();
    showVersionHistoryInterface();
  };
}
// ==========================================
// NEW: showVersionHistoryInterface (With "Don't Snapshot" Option)
// ==========================================
function showVersionHistoryInterface(viewOnly) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. HEADER
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); position: relative; overflow: hidden;`;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">
          DOCUMENT LIFECYCLE
      </div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; letter-spacing: 0.5px; display:flex; align-items:center; gap:8px;">
          <span>📷</span> Version History
      </div>
  `;
  out.appendChild(header);

  // 2. EXPLANATION BAR
  var infoBar = document.createElement("div");
  infoBar.style.cssText = "background:#e3f2fd; border:1px solid #90caf9; border-left:4px solid #1565c0; padding:10px 12px; margin-bottom:15px;";
  infoBar.innerHTML = `<div style="font-size:10px; color:#1565c0; line-height:1.4;">
      <strong>Version History:</strong> You can enable auto-save in the settings that will save each version when ELI loads.
  </div>`;
  out.appendChild(infoBar);

  // 3. TIMELINE CONTAINER
  var timelineContainer = document.createElement("div");
  timelineContainer.id = "historyContainer";
  timelineContainer.style.cssText = "padding:5px; padding-bottom:40px;";
  // Placeholder while we wait for user input
  timelineContainer.innerHTML = `<div style="text-align:center; color:#999; font-size:11px; padding:20px;">Waiting for selection...</div>`;
  out.appendChild(timelineContainer);

  // 4. HIDDEN FILE INPUT
  var fileInp = document.createElement("input");
  fileInp.type = "file";
  fileInp.accept = ".docx";
  fileInp.style.display = "none";
  out.appendChild(fileInp);

  // --- LOGIC ---
  getUniqueDocumentId().then(function (docId) {

    // --- HELPER: SHOW LABEL MODAL ---
    function askForLabel() {
      var overlay = document.createElement("div");
      overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

      var dialog = document.createElement("div");
      dialog.className = "tool-box slide-in";
      dialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #1565c0; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

      dialog.innerHTML = `
            <div style="font-size:24px; margin-bottom:10px;">🏷️</div>
            <div style="font-weight:800; color:#1565c0; margin-bottom:10px; font-size:12px; text-transform:uppercase;">Identify Version Source</div>
            <div style="font-size:12px; color:#555; line-height:1.4; margin-bottom:15px;">
              Who authored this current version?
            </div>
            
            <button id="btnSourceCP" class="btn-action" style="width:100%; margin-bottom:8px; text-align:left;">🏢 Colgate (Internal)</button>
            <button id="btnSourceTP" class="btn-action" style="width:100%; margin-bottom:8px; text-align:left;">👤 Third Party (External)</button>
            <button id="btnSourceSkip" class="btn-action" style="width:100%; color:#546e7a; border-color:#cfd8dc; text-align:center; margin-bottom:15px;">Use Generic Label</button>
            
            <div style="border-top:1px dashed #ddd; margin-bottom:10px;"></div>
            
            <button id="btnNoSnapshot" style="width:100%; background:transparent; border:none; color:#c62828; font-size:11px; font-weight:700; cursor:pointer; text-decoration:underline;">🚫 Don't take a snapshot</button>
        `;

      overlay.appendChild(dialog);
      document.body.appendChild(overlay);

      // Handler: Take Snapshot
      var handleSelection = function (suffix) {
        overlay.remove();
        var label = "Auto-Save";
        if (suffix) label += " (" + suffix + ")";

        timelineContainer.innerHTML = `<div class="spinner" style="width:16px;height:16px;border-width:2px;border-color:#1565c0 transparent #1565c0 transparent; display:inline-block;"></div> Capturing <strong>${label}</strong>...`;

        createSnapshot(label, null, function () {
          renderVersionTimeline(timelineContainer, docId);
        });
      };

      // Handler: Skip Snapshot
      var handleSkip = function () {
        overlay.remove();
        showToast("⚠️ Snapshot Skipped");
        // Just load the timeline directly
        renderVersionTimeline(timelineContainer, docId);
      };

      document.getElementById("btnSourceCP").onclick = function () { handleSelection("CP Version"); };
      document.getElementById("btnSourceTP").onclick = function () { handleSelection("Third Party"); };
      document.getElementById("btnSourceSkip").onclick = function () { handleSelection(""); };

      // NEW BUTTON BINDING
      document.getElementById("btnNoSnapshot").onclick = handleSkip;
    }

    // TRIGGER THE MODAL IMMEDIATELY
    if (viewOnly === true) {
      renderVersionTimeline(timelineContainer, docId);
    } else {
      askForLabel();
    }

    // Import Handler
    // Import Handler (Button at Bottom)
    var mpImport = document.createElement("button");
    mpImport.className = "btn-action";
    mpImport.innerHTML = "📂 Import External File";
    mpImport.style.cssText = "width:100%; margin-top:15px; margin-bottom:5px; background: linear-gradient(180deg, #ffffff 0%, #cfd8dc 100%); border: 1px solid #b0bec5; border-bottom: 3px solid #78909c; color: #37474f; font-weight: 700; text-shadow: 0 1px 0 #fff; border-radius:0px; padding:8px;";
    mpImport.onclick = function () { fileInp.click(); };

    // Place: Append to the end, then move Home Button to be after it
    out.appendChild(mpImport);
    if (typeof btnHome !== 'undefined') out.appendChild(btnHome);

    fileInp.onchange = function (e) {
      let file = e.target.files[0];
      if (!file) return;
      handleFileSelect(file, async (text, html) => {
        let newVer = {
          id: Date.now(),
          timestamp: Date.now(),
          label: "Import: " + file.name,
          author: "External",
          text: text,
          html: html
        };
        PROMPTS.saveVersionSnapshot(docId, newVer).then(() => {
          renderVersionTimeline(timelineContainer, docId);
          showToast("✅ Imported Successfully");
        });
      });
    };
  });

  // 5. HOME BUTTON
  var btnHome = document.createElement("button");
  btnHome.className = "btn-action";
  btnHome.innerHTML = "🏠 RETURN HOME";
  btnHome.style.cssText = `width: 100%; margin-top: 20px; border-radius:0px; padding:10px;`;
  btnHome.onclick = showHome;
  out.appendChild(btnHome);
}
// ==========================================
// START: ELI APP STORE UI (WITH INSTRUCTION TOAST)
// ==========================================
function showAppStore() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // --- SAFETY FIX: Ensure the list exists ---
  if (!userSettings.installedStoreApps) {
    userSettings.installedStoreApps = [];
  }

  // 1. Header
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #1565c0; background:#fff;";

  card.innerHTML = `
        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227;">
            <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">MARKETPLACE</div>
            <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;">
                <span>🧩</span> ELI App Store
            </div>
        </div>
        <div id="storeContent" style="padding:15px; background:#f8f9fa;"></div>
        <div style="padding:10px 15px; background:#fff; border-top:1px solid #e0e0e0;">
            <button class="btn-action" onclick="showHome()">🏠 Close Store</button>
        </div>
    `;
  out.appendChild(card);

  var container = card.querySelector("#storeContent");

  // 2. Filter Data
  var featuredApps = ELI_APP_STORE_CATALOG.filter(function (a) { return a.category !== "Core System Tools"; });
  var coreApps = ELI_APP_STORE_CATALOG.filter(function (a) { return a.category === "Core System Tools"; });

  // 3. Render Helper
  function renderAppRow(app, targetDiv) {
    var item = document.createElement("div");
    item.style.cssText = "display:flex; justify-content:space-between; align-items:center; padding:12px; border:1px solid #eee; margin-bottom:8px; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,0.05); border-radius:4px;";

    // Check Installation Status safely
    var isInstalled = userSettings.installedStoreApps.indexOf(app.key) !== -1;
    var isCore = app.category === "Core System Tools";

    var btnHtml = "";

    if (isCore) {
      btnHtml = `<span style="font-size:10px; font-weight:700; color:#90a4ae; background:#f5f5f5; padding:4px 8px; border-radius:4px;">BUILT-IN</span>`;
    }
    else if (isInstalled) {
      btnHtml = `<button class="btn-uninstall" style="background:#fff; color:#c62828; border:1px solid #ef9a9a; padding:6px 12px; font-size:10px; font-weight:bold; cursor:pointer; border-radius:0px;">UNINSTALL</button>`;
    }
    else {
      btnHtml = `<button class="btn-install" style="background:#1565c0; color:#fff; border:none; padding:6px 12px; font-size:10px; font-weight:bold; cursor:pointer; border-radius:0px; box-shadow:0 2px 4px rgba(21,101,192,0.3);">GET +</button>`;
    }

    item.innerHTML = `
            <div style="display:flex; align-items:center; gap:12px;">
                <div style="font-size:20px; background:#f8f9fa; width:36px; height:36px; display:flex; align-items:center; justify-content:center; border-radius:4px; border:1px solid #eee;">${app.icon}</div>
                <div>
                    <div style="font-size:12px; font-weight:700; color:#37474f;">${app.label}</div>
                    <div style="font-size:10px; color:#78909c; line-height:1.2;">${app.desc}</div>
                </div>
            </div>
            <div>${btnHtml}</div>
        `;

    // Bind Actions
    if (!isCore) {
      var btnIn = item.querySelector(".btn-install");
      var btnUn = item.querySelector(".btn-uninstall");

      if (btnIn) {
        btnIn.onclick = function () {
          // 1. Install to Library (This makes it visible in Customize Toolbar)
          if (userSettings.installedStoreApps.indexOf(app.key) === -1) {
            userSettings.installedStoreApps.push(app.key);
          }
          registerAppRuntime(app);
          saveCurrentSettings();
          showAppStore();  // Refresh store to show "Uninstall" button
          showToast("✅ Installed! Go to 'Customize Toolbar' to add it.");
        };
      }

      if (btnUn) {
        btnUn.onclick = function () {
          var idx = userSettings.installedStoreApps.indexOf(app.key);
          if (idx > -1) userSettings.installedStoreApps.splice(idx, 1);
          unregisterAppRuntime(app.key);
          saveCurrentSettings();
          showAppStore();
          showToast("🗑️ App Removed");
        };
      }
    }

    targetDiv.appendChild(item);
  }

  // 4. Render Sections
  var featDiv = document.createElement("div");
  featDiv.innerHTML = `<div style="font-size:10px; font-weight:800; color:#1565c0; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;">✨ Featured Extensions</div>`;
  container.appendChild(featDiv);
  featuredApps.forEach(function (app) { renderAppRow(app, featDiv); });

  // Collapsible Core Tools
  var coreHeader = document.createElement("div");
  coreHeader.style.cssText = "margin-top:20px; margin-bottom:10px; padding:10px; background:#e3f2fd; border:1px solid #bbdefb; color:#1565c0; font-size:11px; font-weight:700; cursor:pointer; display:flex; align-items:center; justify-content:space-between; border-radius:4px;";
  coreHeader.innerHTML = `<span>🛠️ Core System Tools (${coreApps.length})</span> <span id="coreArrow">▼</span>`;

  var coreList = document.createElement("div");
  coreList.id = "coreAppsList";
  coreList.style.display = "none";

  coreHeader.onclick = function () {
    var isHidden = coreList.style.display === "none";
    coreList.style.display = isHidden ? "block" : "none";
    this.querySelector("#coreArrow").innerText = isHidden ? "▲" : "▼";
  };

  container.appendChild(coreHeader);
  container.appendChild(coreList);
  coreApps.forEach(function (app) { renderAppRow(app, coreList); });
}
// ==========================================
// END: ELI APP STORE UI
// ==========================================
// START: DOCU-SCRUB ENGINE (FEATURE RICH)
// ==========================================
function showDocuScrubUI() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. HEADER
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #00acc1; background:#fff;";

  card.innerHTML = `
        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227;">
            <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">HYGIENE</div>
            <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;">
                <span>🧼</span> Docu-Scrub
            </div>
        </div>
        
        <div style="padding:15px;">
            <div style="font-size:11px; color:#546e7a; margin-bottom:15px;">Configure your cleaning preferences below.</div>

            <div style="background:#f0faff; border:1px solid #b3e5fc; padding:10px; margin-bottom:10px;">
                <div style="font-size:10px; font-weight:800; color:#0277bd; text-transform:uppercase; margin-bottom:8px;">🦷 Whitening (Unify Font)</div>
                
                <div style="display:flex; gap:10px; margin-bottom:8px;">
                    <div style="flex:2;">
                        <label class="tool-label" style="margin-bottom:2px;">Font Name</label>
                        <select id="scrubFontName" class="audience-select" style="margin-bottom:0; height:28px; padding:2px;">
                            <option value="Times New Roman" selected>Times New Roman</option>
                            <option value="Arial">Arial</option>
                            <option value="Calibri">Calibri</option>
                            <option value="Segoe UI">Segoe UI</option>
                            <option value="Helvetica">Helvetica</option>
                            <option value="NO_CHANGE">-- Keep Original --</option>
                        </select>
                    </div>
                    <div style="flex:1;">
                        <label class="tool-label" style="margin-bottom:2px;">Size</label>
                        <input type="number" id="scrubFontSize" class="tool-input" value="11" style="margin-bottom:0; height:28px;">
                    </div>
                </div>

                <div>
                    <label class="tool-label" style="margin-bottom:2px;">Font Color</label>
                    <div style="display:flex; align-items:center; gap:8px;">
                        <input type="color" id="scrubFontColor" value="#000000" style="width:30px; height:30px; border:none; padding:0; background:none; cursor:pointer;">
                        <span style="font-size:11px; color:#555;">(Standard Legal Black)</span>
                    </div>
                </div>
            </div>

            <div style="background:#f1f8e9; border:1px solid #c5e1a5; padding:10px; margin-bottom:10px;">
                <div style="font-size:10px; font-weight:800; color:#33691e; text-transform:uppercase; margin-bottom:8px;">🧵 Flossing (Cleanup)</div>
                
                <label class="checkbox-wrapper" style="margin-bottom:6px;">
                    <input type="checkbox" id="chkRemoveEmpty" checked style="margin-right:6px;">
                    <span>Remove Empty Paragraphs/Lines</span>
                </label>
                <label class="checkbox-wrapper">
                    <input type="checkbox" id="chkFixDoubleSpace" checked style="margin-right:6px;">
                    <span>Fix Double Spaces ("  " -> " ")</span>
                </label>
            </div>

            <div style="background:#fff3e0; border:1px solid #ffe0b2; padding:10px; margin-bottom:15px;">
                <div style="font-size:10px; font-weight:800; color:#ef6c00; text-transform:uppercase; margin-bottom:8px;">🦠 Plaque (Metadata)</div>
                
                <label class="checkbox-wrapper">
                    <input type="checkbox" id="chkStripMeta" checked style="margin-right:6px;">
                    <span>Strip Author & Edit Time Metadata</span>
                </label>
            </div>

            <button id="btnRunScrub" class="btn-action" style="width:100%; background:#00acc1; color:white; border:none; font-weight:700;">✨ APPLY CLEANING</button>
        </div>
        
        <div style="padding:10px 15px; background:#fff; border-top:1px solid #e0e0e0;">
            <button class="btn-action" onclick="showHome()">🏠 Cancel</button>
        </div>
    `;
  out.appendChild(card);

  // Bind Button
  document.getElementById("btnRunScrub").onclick = executeDocuScrub;
}

function executeDocuScrub() {
  // 1. Capture Settings
  var settings = {
    fontName: document.getElementById("scrubFontName").value,
    fontSize: parseFloat(document.getElementById("scrubFontSize").value),
    fontColor: document.getElementById("scrubFontColor").value,
    removeEmpty: document.getElementById("chkRemoveEmpty").checked,
    fixSpaces: document.getElementById("chkFixDoubleSpace").checked,
    stripMeta: document.getElementById("chkStripMeta").checked
  };

  var btn = document.getElementById("btnRunScrub");
  btn.disabled = true;
  btn.innerHTML = "⏳ Scrubbing...";

  showToast("🧼 Scrubbing document...");

  Word.run(async function (context) {
    var body = context.document.body;

    // 1. WHITENING: Apply Font Settings
    // We select the whole body range to apply formatting
    var docRange = body.getRange();

    // Only apply if user didn't select "Keep Original"
    if (settings.fontName !== "NO_CHANGE") {
      docRange.font.name = settings.fontName;
    }

    if (settings.fontSize > 0) {
      docRange.font.size = settings.fontSize;
    }

    docRange.font.color = settings.fontColor;

    // 2. PLAQUE REMOVAL: Strip Metadata
    if (settings.stripMeta) {
      var props = context.document.properties;
      props.load("author, lastAuthor, manager, company");
      // Wipe standard properties
      props.author = "Colgate Legal";
      props.lastAuthor = "";
      props.manager = "";
      props.company = "Colgate-Palmolive";
    }

    // 3. FLOSSING: Remove Empty Paragraphs
    var deletedCount = 0;
    if (settings.removeEmpty) {
      var paragraphs = body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      // Reverse loop is safer for deletions
      for (var i = paragraphs.items.length - 1; i >= 0; i--) {
        var p = paragraphs.items[i];
        // Strict check: Is it truly empty or just whitespace?
        if (p.text.trim() === "") {
          // Try/Catch prevents crash on last paragraph deletion
          try {
            p.delete();
            deletedCount++;
          } catch (e) { }
        }
      }
    }

    // 4. FIX SPACES: Remove Double Spaces
    var spaceCount = 0;
    if (settings.fixSpaces) {
      // Using Wildcard search for 2 or more spaces
      var searchResults = body.search(" {2,}", { matchWildcards: true });
      searchResults.load("items");
      await context.sync();

      spaceCount = searchResults.items.length;
      // Replace all matches with single space
      for (var j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(" ", "Replace");
      }
    }

    await context.sync();

    btn.innerHTML = "✅ Done";
    btn.style.background = "#2e7d32";

    var msg = `✨ Cleaned! Applied ${settings.fontName}.`;
    if (settings.removeEmpty) msg += ` Removed ${deletedCount} lines.`;
    if (settings.fixSpaces) msg += ` Fixed ${spaceCount} spaces.`;

    showToast(msg);

    // Reset button after delay
    setTimeout(function () {
      btn.innerHTML = "✨ APPLY CLEANING";
      btn.style.background = "#00acc1";
      btn.disabled = false;
    }, 3000);

  }).catch(function (error) {
    console.error("Scrub Error: " + error);
    btn.innerHTML = "⚠️ Error";
    showToast("⚠️ Scrub error: " + error.message);
    btn.disabled = false;
  });
}
// ==========================================
// START: DEFINED TERM MATRIX ENGINE
// ==========================================
function runTermMatrix() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. HEADER
  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #673ab7; background:#fff;";

  card.innerHTML = `
        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227;">
            <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">NAVIGATION</div>
            <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;">
                <span>🔠</span> Term Matrix
            </div>
        </div>
        <div id="termList" style="padding:0;">
            <div style="padding:20px; text-align:center; color:#666;">
                <div class="spinner" style="width:16px;height:16px;border-width:2px;display:inline-block;"></div> Scanning definitions...
            </div>
        </div>
        <div style="padding:10px 15px; background:#fff; border-top:1px solid #e0e0e0;">
            <button class="btn-action" onclick="showHome()">🏠 Close Matrix</button>
        </div>
    `;
  out.appendChild(card);

  var listContainer = card.querySelector("#termList");

  Word.run(async function (context) {
    var body = context.document.body;

    // 1. FIND DEFINITIONS: Terms in quotes (e.g. "Products")
    // We look for capitalized words inside quotes.
    // Regex: "([A-Z][a-zA-Z\s]+)"
    var defResults = body.search('"([A-Z][a-z]+ ?)+"', { matchWildcards: true });
    defResults.load("text, font");

    // 2. FIND GHOST CANDIDATES: Capitalized words NOT in quotes
    // This is heuristic: any Capitalized Word that appears > 3 times might be a term.
    // We grab the text to parse in JS for speed.
    body.load("text");

    await context.sync();

    // --- A. PROCESS DEFINITIONS ---
    var definitions = [];
    var definedSet = new Set(); // For ghost check

    defResults.items.forEach(function (item) {
      var raw = item.text.replace(/"/g, "").trim();
      // Filter noise (e.g. "The", "Section")
      if (raw.length > 2 && !["The", "Section", "Article", "Exhibit"].includes(raw)) {
        if (!definedSet.has(raw)) {
          definitions.push({
            term: raw,
            range: item // Keep reference for navigation
          });
          definedSet.add(raw);
        }
      }
    });

    // --- B. PROCESS GHOSTS (Orphans) ---
    var fullText = body.text;
    var ghosts = [];

    // Simple regex for Capitalized phrases not at start of sentence
    // (This is a basic heuristic for demo speed)
    var capMatches = fullText.match(/(?<!^|\. )(?<!")\b[A-Z][a-z]+(?: [A-Z][a-z]+)*\b(?!")/g) || [];

    var counts = {};
    capMatches.forEach(m => { counts[m] = (counts[m] || 0) + 1; });

    for (var key in counts) {
      // If it appears often, looks like a term, but isn't in our defined set
      if (counts[key] > 2 && !definedSet.has(key) && !["Colgate", "Company", "Party", "Parties", "Agreement"].includes(key)) {
        ghosts.push({ term: key, count: counts[key] });
      }
    }

    // --- C. RENDER UI ---
    listContainer.innerHTML = "";

    // RENDER: GHOST ALERT
    if (ghosts.length > 0) {
      var ghostDiv = document.createElement("div");
      ghostDiv.style.cssText = "background:#fff3e0; border-bottom:1px solid #ffe0b2; padding:10px 15px;";
      ghostDiv.innerHTML = `<div style="font-size:10px; font-weight:800; color:#e65100; margin-bottom:5px;">👻 GHOST TERMS DETECTED</div>`;

      var ghostList = document.createElement("div");
      ghostList.style.cssText = "display:flex; flex-wrap:wrap; gap:5px;";

      ghosts.slice(0, 5).forEach(g => {
        var chip = document.createElement("span");
        chip.style.cssText = "font-size:10px; background:#fff; border:1px solid #ffcc80; padding:2px 6px; border-radius:4px; color:#e65100;";
        chip.innerText = `${g.term} (${g.count})`;
        ghostList.appendChild(chip);
      });
      ghostDiv.appendChild(ghostList);
      listContainer.appendChild(ghostDiv);
    }

    // RENDER: DEFINED TERMS
    if (definitions.length === 0) {
      listContainer.innerHTML += `<div style="padding:20px; text-align:center; color:#999;">No defined terms found (e.g. "Term").</div>`;
    } else {
      // Sort Alphabetically
      definitions.sort((a, b) => a.term.localeCompare(b.term));

      var ul = document.createElement("div");
      definitions.forEach((def, i) => {
        var row = document.createElement("div");
        row.className = "term-row"; // Hover effect handled by CSS or inline
        row.style.cssText = "padding:10px 15px; border-bottom:1px solid #f0f0f0; cursor:pointer; display:flex; justify-content:space-between; align-items:center; transition:background 0.2s;";
        row.onmouseover = function () { this.style.background = "#f5f5f5"; };
        row.onmouseout = function () { this.style.background = "white"; };

        row.innerHTML = `<span style="font-weight:600; font-size:12px; color:#333;">${def.term}</span> <span style="font-size:10px; color:#673ab7;">Go ➜</span>`;

        // CLICK TO NAVIGATE
        row.onclick = function () {
          // We must re-find the range in a new context for stability, 
          // or use the original if session is active.
          // For simplicity in this structure, we trigger a fresh search for the term.
          locateText('"' + def.term + '"', this, true);
        };

        ul.appendChild(row);
      });
      listContainer.appendChild(ul);
    }

  }).catch(function (error) {
    console.error("Matrix Error: " + error);
    listContainer.innerHTML = `<div style="color:red; padding:15px;">Error: ${error.message}</div>`;
  });
}
// HELPER: Add app to definitions so it appears in "Customize Toolbar"
function registerAppRuntime(app) {
  if (TOOL_DEFINITIONS.find(t => t.key === app.key)) return; // Prevent duplicates

  TOOL_DEFINITIONS.push({
    id: app.id,
    key: app.key,
    icon: app.icon,
    label: app.label,
    desc: app.desc,
    action: function () {
      setActiveButton(app.id);
      tryCollapseToolbar();
      app.action();
    }
  });
}

// HELPER: Remove app from definitions and toolbars
function unregisterAppRuntime(appKey) {
  // 1. Remove from Definitions
  TOOL_DEFINITIONS = TOOL_DEFINITIONS.filter(t => t.key !== appKey);

  // 2. Remove from Active Workflows
  if (userSettings.workflowMaps) {
    Object.keys(userSettings.workflowMaps).forEach(mode => {
      var playlist = userSettings.workflowMaps[mode];
      var idx = playlist.indexOf(appKey);
      if (idx !== -1) playlist.splice(idx, 1);
    });
  }
}
// ==========================================
// START: DEAL TIMELINE APP
// ==========================================
function runTimelineAnalysis() {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. Header
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 15px;`;
  header.innerHTML = `<div style="font-family:'Georgia', serif; font-size:16px; color:#fff;">📅 Deal Timeline</div>`;
  out.appendChild(header);

  // 2. Loading State
  var loader = document.createElement("div");
  loader.innerHTML = `<div class="spinner" style="width:20px;height:20px;border-width:2px;display:inline-block;"></div> Extracting dates & deadlines...`;
  loader.style.cssText = "text-align:center; color:#546e7a; padding:20px;";
  out.appendChild(loader);

  // 3. Scan & Analyze
  validateApiAndGetText(false, null).then(function (text) {

    var prompt = `
      Analyze this contract and extract all critical dates and durations.
      RETURN RAW JSON ARRAY: 
      [
        { "label": "Effective Date", "date": "YYYY-MM-DD", "desc": "Contract Start" },
        { "label": "Initial Term End", "date": "YYYY-MM-DD", "desc": "Expires after 3 years" },
        { "label": "Renewal Notice Deadline", "date": "YYYY-MM-DD", "desc": "Must notify 60 days prior to expiration", "critical": true }
      ]
      
      RULES:
      1. Calculate exact dates if possible (assume today is ${new Date().toISOString().split('T')[0]}).
      2. If a date is relative (e.g. "3 years from Effective Date"), calculate it.
      3. Mark "critical" as true for non-renewal notices or termination deadlines.
      
      TEXT:
      ${text.substring(0, 50000)}
    `;

    return callGemini(API_KEY, prompt);
  }).then(function (result) {
    loader.remove();
    try {
      var data = tryParseGeminiJSON(result);
      if (!Array.isArray(data)) throw new Error("Invalid Data");

      // Sort by date
      data.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });

      // Render Timeline
      var timelineHTML = `<div style="position:relative; padding-left:20px; border-left:2px solid #e0e0e0; margin-left:10px;">`;

      data.forEach(function (item) {
        var isCrit = item.critical;
        var color = isCrit ? "#c62828" : "#1565c0";
        var bg = isCrit ? "#ffebee" : "#e3f2fd";

        timelineHTML += `
          <div class="slide-in" style="margin-bottom:20px; position:relative;">
            <div style="position:absolute; left:-26px; top:0; width:10px; height:10px; background:${color}; border-radius:50%; border:2px solid #fff; box-shadow:0 0 0 1px ${color};"></div>
            
            <div style="background:${bg}; border-left:4px solid ${color}; padding:10px; border-radius:0px; box-shadow:0 1px 3px rgba(0,0,0,0.1);">
                <div style="font-size:10px; font-weight:700; color:${color}; text-transform:uppercase;">${item.date || "TBD"}</div>
                <div style="font-size:13px; font-weight:800; color:#333; margin-bottom:4px;">${item.label}</div>
                <div style="font-size:11px; color:#546e7a;">${item.desc}</div>
            </div>
          </div>
        `;
      });
      timelineHTML += `</div>`;

      // Export Button
      timelineHTML += `
        <button class="btn-action" style="width:100%; margin-top:10px;" onclick="showToast('📅 Added to Outlook (Simulated)')">
            🗓️ Add to Outlook
        </button>
      `;

      var content = document.createElement("div");
      content.style.padding = "0 15px 15px 15px";
      content.innerHTML = timelineHTML;
      out.appendChild(content);

    } catch (e) {
      out.innerHTML += `<div style="color:red; padding:15px;">Could not extract dates.</div>`;
    }
  }).catch(handleError);
}
// HELPER: Get Unique ID for the current Word Document
function getDocumentGuid() {
  return new Promise((resolve) => {
    try {
      // Try to get existing ID saved inside the .docx file
      var docId = Office.context.document.settings.get("eli_doc_guid");

      if (docId) {
        resolve(docId);
      } else {
        // Create new ID, save it to the file, and persist
        var newId = "doc_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
        Office.context.document.settings.set("eli_doc_guid", newId);
        Office.context.document.settings.saveAsync(function () {
          resolve(newId);
        });
      }
    } catch (e) {
      // Fallback for non-Word environments
      console.warn("Guid Error", e);
      resolve("session_fallback");
    }
  });
}
// ==========================================
// START: PRE-SAVE BRACKET SCANNER
// ==========================================
function runPreSaveScan() {
  setLoading(true, "template");

  Word.run(async (context) => {
    // 1. Search for brackets using "Not closing bracket" pattern (Non-greedy)
    // Pattern matches "[" followed by anything that isn't "]", ending in "]"
    var results = context.document.body.search("\\[[!]]@\\]", { matchWildcards: true });
    results.load("items");
    await context.sync();

    var bracketsFound = [];

    // 2. Load properties
    for (var i = 0; i < results.items.length; i++) {
      results.items[i].font.load("color");
      results.items[i].load("text");
    }
    await context.sync();

    // 3. Filter & Fix Logic
    results.items.forEach((item) => {
      var color = (item.font.color || "").toLowerCase();
      var isRed = color === "#c62828" || color === "red" || color === "#ff0000";

      if (!isRed) {
        var text = item.text;

        // --- FIX: DETECT CUT-OFF DOUBLE BRACKETS ---
        // If it starts with "[[" but the search stopped at the first "]"
        // e.g. "[[Schedule]", we manually add the missing "]"
        if (text.startsWith("[[") && !text.endsWith("]]")) {
          text = text + "]";
        }
        // -------------------------------------------

        bracketsFound.push({ text: text });
      }
    });

    return bracketsFound;
  })
    .then((items) => {
      setLoading(false);

      if (items.length > 0) {
        // Issues found? Show the cleaner UI
        showBracketCleaner(items);
      } else {
        // --- SUCCESS: DOCUMENT IS CLEAN ---
        if (window.eliWizardMode) {
          // AUTO MODE: Proceed to Step 3 (Error Check)
          showToast("🚀 Clean! Proceeding to Step 3...");
          setTimeout(function () {
            // Call Error Check (Fast Mode = true)
            // Pass Step 4 (Save Manager) as the callback for when Error Check is done
            runInteractiveErrorCheck(function () {
              showSaveAsManager();
              window.eliWizardMode = false;
            }, true);
          }, 1000);
        } else {
          // MANUAL MODE: Return to Hub
          showToast("✅ Clean & Verified");
          showAutoFillHub();
        }
        // ----------------------------------
      }
    })
    .catch((e) => {
      setLoading(false);
      console.error("Scan error", e);
      showAutoFillHub(); // Fallback to hub on error
    });
}
// ==========================================
// START: DOCUMENT CLEANER UI (DYNAMIC)
// ==========================================
function showBracketCleaner(items) {
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 1. HEADER
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #f57f17; margin-bottom: 12px; position:relative; overflow:hidden;`;
  header.innerHTML = `
      <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(circle at top right, rgba(255,255,255,0.05), transparent 60%); pointer-events: none;"></div>
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #ffb74d; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">ATTENTION NEEDED</div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;">
          <span>🧹</span> Document Cleaner & SOW Helper
      </div>
  `;
  out.appendChild(header);

  // 2. INTRO
  var intro = document.createElement("div");
  intro.style.cssText = "padding:15px; background:#fff3e0; border-left:4px solid #f57f17; margin-bottom:15px; font-size:12px; color:#5d4037; line-height:1.5;";
  intro.innerHTML = `<strong>${items.length} placeholder(s) detected.</strong><br>These brackets appear to be unfilled. Use the locate button and suggestions below to resolve them before saving.`;
  out.appendChild(intro);

  // 3. RENDER LIST
  var list = document.createElement("div");
  list.style.cssText = "padding:0 15px; margin-bottom:20px;";

  items.forEach((item, idx) => {
    var row = document.createElement("div");
    row.className = "slide-in";
    row.style.cssText = "display:flex; align-items:center; justify-content:space-between; background:white; border:1px solid #e0e0e0; border-left:3px solid #b0bec5; padding:10px; margin-bottom:8px; box-shadow:0 2px 4px rgba(0,0,0,0.05);";

    // Clean Label: "[[Expenses]]" -> "Expenses"
    var cleanLabel = item.text.replace(/[\[\]]/g, "").trim();

    // DYNAMIC MATCHING: Check if cleanLabel includes any key from CLEANER_RULES
    var matchedKey = Object.keys(CLEANER_RULES).find(key => cleanLabel.includes(key));
    var rule = matchedKey ? CLEANER_RULES[matchedKey] : null;

    // Left Side
    var left = document.createElement("div");
    left.innerHTML = `<div style="font-weight:700; color:#37474f; font-size:12px;">${item.text}</div>`;

    // Right Side
    var right = document.createElement("div");
    right.style.display = "flex";
    right.style.gap = "6px";

    // A. LOCATE BUTTON
    var btnLoc = document.createElement("button");
    btnLoc.className = "btn-platinum-config";
    btnLoc.innerText = "📍 Locate";
    btnLoc.style.fontSize = "10px";
    btnLoc.onclick = function () { locateText(item.text, btnLoc, true); };
    right.appendChild(btnLoc);

    // B. HELP BUTTON (Dynamic)
    if (rule) {
      var btnHelp = document.createElement("button");
      btnHelp.className = "btn-platinum-config primary";
      btnHelp.innerHTML = "💡 Suggestions";
      btnHelp.style.fontSize = "10px";
      btnHelp.style.borderColor = "#1565c0";

      btnHelp.onclick = function () {
        // Pass the generic rule object
        showGenericCleanerHelper(rule, item.text);
      };
      right.appendChild(btnHelp);
    }

    row.appendChild(left);
    row.appendChild(right);
    list.appendChild(row);
  });

  out.appendChild(list);

  // 4. FOOTER
  var footer = document.createElement("div");
  footer.style.cssText = "padding:15px; border-top:1px solid #e0e0e0; background:#f8f9fa; display:flex; gap:10px;";
  var btnIgnore = document.createElement("button");
  btnIgnore.className = "btn-action";
  btnIgnore.innerHTML = "I'm Done ➜";
  btnIgnore.style.width = "100%";
  // ... inside showBracketCleaner ...

  btnIgnore.onclick = function () {
    // --- UPDATED NAVIGATION LOGIC ---
    if (window.eliWizardMode) {
      // AUTO MODE: Go to Step 3
      runInteractiveErrorCheck(function () {
        showSaveAsManager();
        window.eliWizardMode = false;
      }, true);
    } else {
      // MANUAL MODE: Return to Hub
      showAutoFillHub();
    }
    // --------------------------------
  };
  footer.appendChild(btnIgnore);
  out.appendChild(footer);
}
// ==========================================
// START: GENERIC CLEANER HELPER (DYNAMIC)
// ==========================================
function showGenericCleanerHelper(rule, targetTextString) {
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "background:white; width:320px; padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #2e7d32; box-shadow:0 20px 50px rgba(0,0,0,0.3);";

  card.innerHTML = `
    <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 15px; border-bottom: 3px solid #2e7d32;">
        <div style="font-size:10px; font-weight:800; color:#a5d6a7; text-transform:uppercase;">SOW HELPER</div>
        <div style="font-family:'Georgia', serif; font-size:16px; color:white;">${rule.title}</div>
    </div>
    <div style="padding:15px;">
        <div style="font-size:12px; color:#546e7a; margin-bottom:15px;">Select a clause to replace <strong>${targetTextString}</strong>:</div>
        <div id="cleanerOptionsContainer"></div>
        <button id="btnCancelHelper" class="btn-action" style="width:100%; margin-top:10px;">Cancel</button>
    </div>
  `;

  overlay.appendChild(card);
  document.body.appendChild(overlay);

  var container = card.querySelector("#cleanerOptionsContainer");

  // 1. Render Dynamic Options
  rule.options.forEach(opt => {
    var div = document.createElement("div");
    div.className = "helper-option";
    div.style.cssText = "cursor:pointer; padding:10px; border:1px solid #e0e0e0; margin-bottom:8px; border-radius:4px; transition:background 0.2s;";

    div.innerHTML = `
        <div style="font-weight:700; font-size:11px; color:#1565c0; margin-bottom:4px;">${opt.label}</div>
        <div style="font-size:10px; color:#333;">${opt.desc}</div>
      `;

    div.onmouseover = function () { this.style.backgroundColor = "#f5f9ff"; this.style.borderColor = "#1565c0"; };
    div.onmouseout = function () { this.style.backgroundColor = "transparent"; this.style.borderColor = "#e0e0e0"; };

    div.onclick = function () {
      // --- PASS 'TRUE' HERE ---
      performInsertion(opt.text, targetTextString, null, null, true);
      // ------------------------

      overlay.remove();
      showToast("✅ Clause Inserted");
      setTimeout(runPreSaveScan, 800);
    };

    container.appendChild(div);
  });

  // 2. Add "Fix Manually"
  var manualDiv = document.createElement("div");
  manualDiv.style.cssText = "cursor:pointer; padding:10px; border:1px dashed #b0bec5; margin-bottom:8px; border-radius:4px; transition:background 0.2s; background:#fcfcfc;";
  manualDiv.innerHTML = `
    <div style="font-weight:700; font-size:11px; color:#546e7a; margin-bottom:4px;">🛠️ Fix Manually</div>
    <div style="font-size:10px; color:#666;">I will edit the text myself.</div>
  `;
  manualDiv.onmouseover = function () { this.style.backgroundColor = "#eceff1"; this.style.borderColor = "#78909c"; };
  manualDiv.onmouseout = function () { this.style.backgroundColor = "#fcfcfc"; this.style.borderColor = "#b0bec5"; };

  manualDiv.onclick = function () { overlay.remove(); };
  container.appendChild(manualDiv);

  document.getElementById("btnCancelHelper").onclick = function () { overlay.remove(); };
}
// ==========================================
// REPLACE: loadClauseLibrary (Auto-Fixes Missing IDs)
// ==========================================
async function loadClauseLibrary() {
  try {
    // 1. Try loading from NEW IndexedDB first
    var dbClauses = await loadLibraryFromDB();

    if (dbClauses && dbClauses.length > 0) {
      console.log("✅ Loaded " + dbClauses.length + " clauses from IndexedDB");
      eliClauseLibrary = dbClauses;
    } else {
      // 2. Fallback: Check LocalStorage (Migration)
      console.log("⚠️ DB empty. Checking LocalStorage for migration...");
      var raw = localStorage.getItem("colgate_library");
      if (raw) {
        eliClauseLibrary = JSON.parse(raw);
        // MIGRATE NOW: Save to DB
        await saveLibraryToDB(eliClauseLibrary);
        console.log("🚀 MIGRATION COMPLETE: Moved to IndexedDB.");
        // Optional: Clear LocalStorage to free up space
        // localStorage.removeItem("colgate_library"); 
      } else {
        eliClauseLibrary = [];
      }
    }
  } catch (e) {
    console.error("Library Load Error", e);
    eliClauseLibrary = [];
  }
}
function saveClauseLibrary() {
  // Save to IndexedDB (The new safe way)
  saveLibraryToDB(eliClauseLibrary).then(() => {
    console.log("✅ Library Saved to DB");
  });

  // SAFETY: Keep saving to localStorage for one more day, then delete this line

}
function createLibraryHeader() {
  var header = document.createElement("div");
  // Reduced padding and removed the top border-radius
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 8px 12px; border-bottom: 3px solid #c9a227; margin-bottom: 10px; border-radius: 0px;`;

  // Removed the "KNOWLEDGE BASE" div entirely

  header.innerHTML = `
        <div style="font-size: 14px; font-weight: 400; color: #ffffff; display:flex; justify-content:space-between; align-items:center;">
            
            <div style="display:flex; align-items:center; gap:4px;">
                <span style="font-weight:700; font-size:13px; letter-spacing:0.5px;">📚 Clause Library</span>
                <button id="btnLibReset" title="Reset View" style="background:#0277bd; border:1px solid #4fc3f7; color:white; font-size:9px; padding:2px 6px; cursor:pointer; font-weight:bold; border-radius:2px;">🏠 HOME</button>
            </div>

            <div style="display:flex; gap:5px;">
                <button id="btnAddSelection" title="Add Highlighted Text" style="background:#2e7d32; border:1px solid #66bb6a; color:white; font-size:10px; font-weight:bold; padding:3px 8px; cursor:pointer; display:flex; align-items:center; gap:4px;"><span>➕</span> DOC</button>
                <button id="btnLibClear" title="Clear All" style="background:#b71c1c; border:1px solid #d32f2f; color:white; font-size:12px; padding:3px 8px; cursor:pointer;">🗑️</button>
                <button id="btnLibExport" title="Export JSON" style="background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.3); color:white; font-size:12px; padding:3px 8px; cursor:pointer;">⬇️</button>
                <button id="btnLibImport" title="Import JSON" style="background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.3); color:white; font-size:12px; padding:3px 8px; cursor:pointer;">⬆️</button>
                <button id="btnAiScan" style="background:#fbc02d; color:#37474f; font-size:10px; font-weight:bold; border:none; padding:3px 8px; cursor:pointer;">⚡ EXTRACT</button>
                <button id="btnLibAdd" style="background:#1565c0; border:1px solid #42a5f5; color:white; font-size:10px; font-weight:bold; padding:3px 8px; cursor:pointer;">+ NEW</button>
                
                <div style="width:1px; background:rgba(255,255,255,0.2); margin:0 4px;"></div>

                <button id="btnLibSettings" title="Match Settings" style="background:rgba(255,255,255,0.1); border:1px solid rgba(255,255,255,0.3); color:white; font-size:12px; padding:3px 8px; cursor:pointer;">⚙️</button>
            </div>
        </div>
    `;

  // CRITICAL: Bindings
  setTimeout(() => {
    document.getElementById("btnLibReset").onclick = showClauseLibraryInterface;
    document.getElementById("btnLibClear").onclick = () => promptDelete('all');
    document.getElementById("btnLibExport").onclick = handleExportJSON;
    document.getElementById("btnLibImport").onclick = () => document.getElementById("libBackupInput").click();
    document.getElementById("btnAiScan").onclick = showAIImportOptions;
    document.getElementById("btnLibAdd").onclick = () => showClauseEditor();

    // Bind Settings
    document.getElementById("btnLibSettings").onclick = showMatchSettings;

    document.getElementById("btnAddSelection").onclick = function () {
      validateApiAndGetText(true).then(text => {
        if (!text || text.length < 2) { showToast("⚠️ Select text first."); return; }
        showClauseEditor(null, text);
      });
    };
  }, 0);

  return header;
}
function createHiddenFileInputs() {
  var container = document.createElement("div");
  container.innerHTML = `
        <input type="file" id="libDocInput" accept=".docx,.pdf,.txt" style="display:none;">
        <input type="file" id="libBackupInput" accept=".json" style="display:none;">
    `;
  return container;
}
// ==========================================
// REPLACE: createFilterControls (Fixed Divider Logic)
// ==========================================
function createFilterControls() {
  var controls = document.createElement("div");
  controls.style.cssText = "padding:0 15px 10px 15px; display:flex; flex-direction:column; gap:8px;";

  var style = document.createElement("style");
  style.innerHTML = `
        .search-row { display: flex; alignItems: center; width: 100%; margin-bottom: 5px; }
        .joined-control { height: 28px !important; box-sizing: border-box !important; margin: 0 !important; border: 1px solid #ccc !important; font-size: 11px !important; outline: none !important; }
        .scope-select { width: 95px !important; flex-shrink: 0 !important; border-radius: 0px !important; border-right: 1px solid #eee !important; background-color: #f5f5f5 !important; color: #333 !important; font-weight: 700 !important; cursor: pointer; padding-left: 4px !important; }
        .search-text { flex: 1 !important; border-radius: 0 !important; border-left: none !important; border-right: none !important; padding: 0 8px !important; background: white !important; color: #333 !important; }
        .btn-go { width: 32px !important; flex-shrink: 0 !important; border-radius: 0px !important; border-left: 1px solid #eee !important; background-color: #fff !important; cursor: pointer !important; display: flex !important; align-items: center !important; justify-content: center !important; }
        .match-btn { height: 28px !important; margin-left: 6px !important; border-radius: 0px !important; padding: 0 10px !important; background: #1565c0 !important; border: 1px solid #1565c0 !important; color: white !important; font-weight: bold !important; cursor: pointer !important; font-size: 10px !important; box-sizing: border-box !important; }
        
        /* Updated Toggle Container */
        .toggle-container { display:flex; background:#f5f5f5; border:1px solid #e0e0e0; border-radius:0px; padding:2px; width: 100%; margin-bottom: 5px; box-sizing:border-box; }
        .toggle-btn { flex:1; text-align:center; padding: 4px 2px; font-size: 9px; font-weight: 800; cursor: pointer; border-radius: 0px; border: 1px solid transparent; color: #666; white-space:nowrap; text-transform:uppercase; }
        .toggle-btn.active { background: #1565c0; color: white; border-color: #1565c0; box-shadow: 0 1px 2px rgba(0,0,0,0.1); } 
        
        /* THE DIVIDER CLASS */
        .toggle-divider { border-left: 2px solid #ccc; margin-left: 4px; padding-left: 4px; }

        .filter-drop { flex: 1; margin: 0; height: 24px !important; padding: 2px !important; font-size: 10px !important; border: 1px solid #ccc; background: white; color: black; border-radius: 0px; }
    `;
  controls.appendChild(style);

  controls.innerHTML += `
        <div class="toggle-container" id="scopeToggle">
            <div class="toggle-btn active" data-mode="all" onclick="setScope('all')">📚 INTERNAL</div>
            <div class="toggle-btn" data-mode="standard" onclick="setScope('standard')">🟢 COLGATE</div>
            <div class="toggle-btn" data-mode="thirdparty" onclick="setScope('thirdparty')">🟠 3RD PARTY</div>
            
            <div class="toggle-btn toggle-divider" data-mode="public" onclick="setScope('public')" style="color:#00838f;">🌐 EXTERNAL</div>
        </div>
        <input type="hidden" id="libraryScopeMode" value="all">

        <div class="search-row">
            <select id="libSearchField" class="joined-control scope-select" title="Filter by...">
                <option value="all">All Fields</option>
                <option value="entity">Party</option>
                <option value="title">Title</option>
                <option value="text">Text</option>
                <option value="category">Category</option>
            </select>
            <input type="text" id="libSearchInput" class="joined-control search-text" placeholder="Search...">
            <button id="btnTriggerSearch" class="joined-control btn-go" title="Run Search">🔍</button>
            <button id="btnMatchSelection" title="Smart Compare" class="match-btn">🎯 Match</button>
        </div>
        
        <div style="background:#fff8e1; border:1px solid #ffe0b2; padding:6px; display:flex; gap:4px; align-items:center; border-radius:0px;">
            <span style="font-size:9px; font-weight:800; color:#ef6c00; text-transform:uppercase; white-space:nowrap;">GROUP BY:</span>
            <select id="groupBy1" class="filter-drop"><option value="category" selected>1. Category</option><option value="agreement">1. Agreement</option><option value="entity">1. Party</option><option value="jurisdiction">1. Jurisdiction</option><option value="none">None</option></select>
            <span style="font-size:9px; color:#ef6c00;">→</span>
            <select id="groupBy2" class="filter-drop"><option value="none">2. None</option><option value="entity" selected>2. Party</option><option value="agreement">2. Agreement</option><option value="category">2. Category</option><option value="jurisdiction">2. Jurisdiction</option></select>
            <span style="font-size:9px; color:#ef6c00;">→</span>
            <select id="groupBy3" class="filter-drop"><option value="none" selected>3. None</option><option value="agreement">3. Agreement</option><option value="date">3. Date</option><option value="jurisdiction">3. Jurisdiction</option></select>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:6px;">
             <div style="font-size:9px; font-weight:bold; user-select:none;">
                <span id="btnExpandAll" style="color:#1565c0; cursor:pointer; margin-right:8px;">EXPAND ALL</span>
                <span id="btnCollapseAll" style="color:#546e7a; cursor:pointer;">COLLAPSE ALL</span>
            </div>
        </div>
    `;

  window.setScope = function (mode) {
    document.getElementById("libraryScopeMode").value = mode;

    var btns = document.querySelectorAll(".toggle-btn");
    btns.forEach(b => {
      if (b.dataset.mode === mode) {
        // --- ACTIVE STATE ---
        b.classList.add("active");
        if (mode === 'standard') { b.style.background = '#2e7d32'; b.style.borderColor = '#2e7d32'; b.style.color = 'white'; }
        else if (mode === 'thirdparty') { b.style.background = '#ef6c00'; b.style.borderColor = '#ef6c00'; b.style.color = 'white'; }
        else if (mode === 'public') { b.style.background = '#00acc1'; b.style.borderColor = '#00acc1'; b.style.color = 'white'; }
        else { b.style.background = '#1565c0'; b.style.borderColor = '#1565c0'; b.style.color = 'white'; }

        // Ensure active border overrides the divider border for consistency
        b.style.borderLeftColor = "";

      } else {
        // --- INACTIVE STATE ---
        b.classList.remove("active");
        b.style.background = "transparent";

        // FIX: Only clear Top, Right, Bottom. Restore Left if it's the divider.
        b.style.borderTopColor = "transparent";
        b.style.borderRightColor = "transparent";
        b.style.borderBottomColor = "transparent";

        if (b.classList.contains("toggle-divider")) {
          b.style.borderLeft = "2px solid #ccc"; // FORCE RESTORE
        } else {
          b.style.borderLeftColor = "transparent";
        }

        if (b.dataset.mode === 'public') b.style.color = '#00838f';
        else b.style.color = '#666';
      }
    });

    if (window.currentMatchData) {
      runSmartLibrarySearch(window.currentMatchData.selection);
    } else {
      refreshLibraryList();
    }

    var msg = "📚 Showing Internal Clauses";
    if (mode === "standard") msg = "🟢 Showing Colgate Standard Only";
    if (mode === "thirdparty") msg = "🟠 Showing 3rd Party Only";
    if (mode === "public") msg = "🌐 Showing External Resources";
    showToast(msg);
  };

  setTimeout(() => {
    document.getElementById("btnMatchSelection").onclick = () => runSmartLibrarySearch();
    document.getElementById("btnTriggerSearch").onclick = refreshLibraryList;
    document.getElementById("libSearchInput").onkeyup = function (e) { if (e.key === "Enter") refreshLibraryList(); else refreshLibraryList(); };
    document.getElementById("libSearchField").onchange = refreshLibraryList;
    document.getElementById("btnExpandAll").onclick = function () { window.toggleAllClauses(true); };
    document.getElementById("btnCollapseAll").onclick = function () { window.toggleAllClauses(false); };

    var g1 = document.getElementById("groupBy1");
    var g2 = document.getElementById("groupBy2");
    var g3 = document.getElementById("groupBy3");

    function handleGroupChange() {
      updateDropdownExclusions();
      if (window.currentMatchData) {
        renderComparisonResults(window.currentMatchData.results, window.currentMatchData.selection);
      } else {
        refreshLibraryList();
      }
    }

    g1.onchange = handleGroupChange; g2.onchange = handleGroupChange; g3.onchange = handleGroupChange;
    updateDropdownExclusions();
  }, 0);

  return controls;
}
// ==========================================
// REPLACE: refreshLibraryList (With Dynamic Tab Counts)
// ==========================================
function refreshLibraryList() {
  window.currentMatchData = null;

  var container = document.getElementById("libListContainer");
  if (!container) return;
  container.innerHTML = "";

  var txt = document.getElementById("libSearchInput").value.toLowerCase().trim();
  var field = document.getElementById("libSearchField").value;
  var scopeEl = document.getElementById("libraryScopeMode");
  var scopeMode = scopeEl ? scopeEl.value : "all";

  var g1 = document.getElementById("groupBy1").value;
  var g2 = document.getElementById("groupBy2").value;
  var g3 = document.getElementById("groupBy3").value;
  var groupLevels = [g1, g2, g3].filter(g => g && g !== "none");

  // 1. FILTER FOR DISPLAY
  var filtered = eliClauseLibrary.filter(function (item) {

    // A. Scope Check
    if (scopeMode === "public") {
      if (item.type !== "Public") return false;
    } else if (scopeMode === "standard") {
      if (item.type !== "Standard") return false;
    } else if (scopeMode === "thirdparty") {
      if (item.type === "Standard" || item.type === "Public") return false;
    } else {
      // Mode "All" -> Hide Public
      if (item.type === "Public") return false;
    }

    // B. Text Search Check
    if (!txt) return true;

    if (field === "all") {
      var searchContent = (
        (item.title || "") + " " + (item.text || "") + " " +
        (item.category || "") + " " + (item.agreement || "") + " " +
        (item.entity || "") + " " + (item.jurisdiction || "")
      ).toLowerCase();
      return searchContent.includes(txt);
    } else {
      var targetVal = (item[field] || "").toLowerCase();
      return targetVal.includes(txt);
    }
  });

  // 2. RENDER THE LIST
  if (filtered.length === 0) {
    container.innerHTML = "<div style='text-align:center; padding:20px; color:#999;'>No clauses found.</div>";
  } else {
    if (groupLevels.length === 0) {
      filtered.forEach(item => container.appendChild(createClauseCard(item)));
    } else {
      renderNestedGroups(container, filtered, groupLevels, 0);
    }
  }

  // 3. UPDATE TAB COUNTS (The Magic Step)
  updateTabCounts(txt, field);
}
function renderNestedGroups(parentContainer, items, levels, currentDepth) {
  // BASE CASE: If we ran out of levels, render the cards
  if (currentDepth >= levels.length) {
    items.forEach(item => {
      parentContainer.appendChild(createClauseCard(item));
    });
    return;
  }

  // 1. Group items by the current level's key
  var currentKey = levels[currentDepth];
  var groups = {};

  items.forEach(item => {
    var val = item[currentKey] || "Uncategorized";
    if (!groups[val]) groups[val] = [];
    groups[val].push(item);
  });

  // 2. Sort groups alphabetically
  var sortedKeys = Object.keys(groups).sort();

  // 3. Render Each Group
  sortedKeys.forEach(groupName => {
    var groupItems = groups[groupName];
    var uniqueId = "grp_" + currentDepth + "_" + groupName.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);

    // --- STYLING BASED ON DEPTH ---
    // Depth 0 (Level 1): Dark Grey Box
    // Depth 1 (Level 2): Blue Indented
    // Depth 2 (Level 3): Light Grey Indented

    var headerStyle = "";
    var contentStyle = "display:none;"; // COLLAPSED BY DEFAULT

    if (currentDepth === 0) {
      // 1. Removed margin-bottom on header (was 5px)
      // 2. Reduced content padding (10px -> 5px)
      headerStyle = "background:#f5f5f5; border:1px solid #d1d5db; border-left:4px solid #546e7a; padding:8px 10px; font-weight:700; color:#546e7a; font-size:11px; margin-bottom:0px;";
      contentStyle += "padding:5px; border:1px solid #e0e0e0; border-top:none; background:#fafafa; margin-bottom:5px;";
    } else if (currentDepth === 1) {
      // 3. Reduced margin-top on sub-header (5px -> 1px) to pull it up
      headerStyle = "background:#e3f2fd; border-left:3px solid #1565c0; padding:4px 8px; font-weight:700; color:#1565c0; font-size:10px; margin-top:1px;";
      contentStyle += "padding:2px 0 2px 8px; border-left:1px solid #e3f2fd; margin-bottom:2px;";
    } else {
      headerStyle = "background:#fff; border-bottom:1px solid #eee; padding:5px 10px; font-weight:600; color:#666; font-size:10px; margin-top:5px; font-style:italic;";
      contentStyle += "padding:5px 0 5px 15px;";
    }

    // Create Header
    var header = document.createElement("div");
    header.className = "group-header depth-" + currentDepth; // Class for "Expand All"
    header.style.cssText = headerStyle + " cursor:pointer; display:flex; justify-content:space-between; align-items:center; user-select:none;";

    var label = currentKey === "entity" ? "👤 " : (currentKey === "agreement" ? "📄 " : "");

    header.innerHTML = `
            <div style="display:flex; align-items:center; gap:6px;">
                <span id="arrow_${uniqueId}" class="arrow-icon">▶</span>
                <span>${label}${groupName.toUpperCase()}</span>
            </div>
            <span style="font-size:9px; opacity:0.7;">${groupItems.length}</span>
        `;

    // Create Content Body
    var content = document.createElement("div");
    content.id = uniqueId;
    content.className = "group-content depth-" + currentDepth; // Class for "Expand All"
    content.style.cssText = contentStyle;

    // Toggle Event
    header.onclick = function (e) {
      e.stopPropagation();
      var arrow = document.getElementById("arrow_" + uniqueId);
      if (content.style.display === "none") {
        content.style.display = "block";
        if (arrow) arrow.innerText = "▼";
        // Highlight active header
        if (currentDepth === 0) header.style.background = "#f0f7ff";
      } else {
        content.style.display = "none";
        if (arrow) arrow.innerText = "▶";
        // Reset header
        if (currentDepth === 0) header.style.background = "#f5f5f5";
      }
    };

    // Append to Parent
    parentContainer.appendChild(header);
    parentContainer.appendChild(content);

    // RECURSE: Go deeper into this container
    renderNestedGroups(content, groupItems, levels, currentDepth + 1);
  });
}
// ==========================================
// REPLACE: createClauseCard (Added AI Rewrite Button)
// ==========================================
function createClauseCard(item) {
  var card = document.createElement("div");
  card.className = "slide-in clause-card-item";

  // 1. CONTAINER
  card.style.cssText = "background:white; border:1px solid #ccc; border-left:4px solid #546e7a; margin-bottom:6px; border-radius:0px; overflow:hidden; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;";

  var bodyId = "clause_body_" + Math.floor(Math.random() * 100000);
  var aiBoxId = "ai_box_" + Math.floor(Math.random() * 100000);

  // 2. BADGE STYLE
  var badgeStyle = "font-size:9px; padding:1px 5px; border-radius:0px; font-weight:700; text-transform:uppercase; margin-right:6px; border:1px solid #ccc;";
  var typeColor = "#000";
  var typeBg = "#f5f5f5";
  var typeBorder = "#ccc";

  if (item.type === "Standard") { typeBg = "#e8f5e9"; typeColor = "#006400"; }
  else if (item.type === "Preferred") { typeBg = "#e3f2fd"; typeColor = "#0d47a1"; }
  else if (item.type === "Negotiated") { typeBg = "#f3e5f5"; typeColor = "#4a148c"; }
  else if (item.type === "Public") { typeBg = "#e0f7fa"; typeColor = "#006064"; }

  // 3. META TAGS
  var metaHtml = "";
  if (item.entity && item.entity !== "Standard") metaHtml += `<span style="color:#1565c0; margin-right:6px;">👤 ${item.entity}</span>`;
  if (item.agreement && item.agreement !== "All") metaHtml += `<span style="color:#546e7a; margin-right:6px;">📄 ${item.agreement}</span>`;
  if (item.date) metaHtml += `<span style="color:#1565c0;">🗓️ ${item.date}</span>`;

  card.innerHTML = `
        <div class="clause-header" onclick="toggleClauseBody('${bodyId}', this)" style="padding:10px; cursor:pointer; display:flex; justify-content:space-between; align-items:center; background:white;">
            <div style="flex:1; overflow:hidden;">
                <div style="font-weight:700; font-size:11px; color:#000; margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-transform: none !important;">${item.title}</div>
                
                <div style="font-size:9px; color:#000; display:flex; align-items:center;">
                    <span style="${badgeStyle} background:${typeBg}; color:${typeColor}; border-color:${typeBorder};">${item.type}</span>
                    ${metaHtml}
                </div>
            </div>
            <div class="clause-arrow" style="font-size:10px; color:#000; margin-left:8px;">▶</div>
        </div>

        <div id="${bodyId}" class="clause-body" style="display:none; padding:0 10px 10px 10px; border-top:1px solid #eee;">
            <div style="font-size:11px; color:#000; line-height:1.5; background:#f9f9f9; padding:8px; border:1px solid #ccc; margin:8px 0; white-space: pre-wrap; border-radius:0px;">"${item.text}"</div>
            
            <div id="${aiBoxId}" style="display:none; background:#fff8e1; border:1px solid #ffe0b2; padding:8px; font-size:11px; margin-bottom:8px; color:#000; border-radius:0px;"></div>

            <div style="display:flex; gap:4px;">
                <button class="btn-insert" style="flex:1; background:#1565c0; color:white; border:none; padding:6px; font-size:10px; cursor:pointer; font-weight:bold; border-radius:0px;">INSERT</button>
                
                <button class="btn-rewrite" title="AI Rewrite" style="width:28px; background:#f3e5f5; border:1px solid #e1bee7; color:#7b1fa2; cursor:pointer; padding:0; border-radius:0px;">✨</button>

                <button class="btn-copy" style="flex:1; background:#f0f0f0; border:1px solid #ccc; color:#000; padding:6px; font-size:10px; cursor:pointer; border-radius:0px;">COPY</button>
                
                <button class="btn-edit" title="Edit" style="width:28px; background:white; border:1px solid #ccc; cursor:pointer; padding:0; border-radius:0px;">✏️</button>
                <button class="btn-delete" title="Delete" style="width:28px; background:white; border:1px solid #ccc; cursor:pointer; padding:0; color:#c62828; border-radius:0px;">🗑️</button>
                
                <button class="btn-info" title="Info" style="width:28px; background:white; border:1px solid #ccc; cursor:pointer; padding:0; border-radius:0px;">ℹ️</button>
            </div>
        </div>
    `;

  // --- BIND ACTIONS ---
  card.querySelector(".btn-insert").onclick = (e) => { e.stopPropagation(); insertText(item.text); };

  // NEW BINDING
  card.querySelector(".btn-rewrite").onclick = (e) => { e.stopPropagation(); showRewriteEditor(item.text); };

  card.querySelector(".btn-copy").onclick = (e) => { e.stopPropagation(); navigator.clipboard.writeText(item.text); showToast("📋 Copied"); };
  card.querySelector(".btn-edit").onclick = (e) => { e.stopPropagation(); showClauseEditor(item); };
  card.querySelector(".btn-delete").onclick = (e) => { e.stopPropagation(); promptDelete(item.id); };
  card.querySelector(".btn-info").onclick = (e) => { e.stopPropagation(); showInfoModal(item); };

  return card;
}
// ==========================================
// 2. TOGGLE HELPER (Required for Animation)
// ==========================================
// This matches the Simulator logic exactly to ensure the 'open' state looks right.
window.toggleClauseBody = function (id, header) {
  var body = document.getElementById(id);
  var arrow = header.querySelector(".clause-arrow");

  if (body.style.display === "none") {
    body.style.display = "block";
    arrow.innerText = "▼";
    header.style.background = "#f5f9ff"; // Highlight active state
  } else {
    body.style.display = "none";
    arrow.innerText = "▶";
    header.style.background = "transparent"; // Reset to clear
  }
};
function bindLibraryHeaderEvents() {
  // Standard Add (Empty)
  document.getElementById("btnLibAdd").onclick = () => showClauseEditor();

  // NEW: Add from Selection
  document.getElementById("btnAddSelection").onclick = function () {
    // Use the existing Word API helper
    validateApiAndGetText(true).then(text => {
      if (!text || text.length < 2) {
        showToast("⚠️ Please select text in the document first.");
        return;
      }
      // Open editor with text pre-filled
      showClauseEditor(null, text);
    }).catch(e => showToast("Error: " + e.message));
  };

  // The "EXTRACT" button
  document.getElementById("btnAiScan").onclick = () => showAIImportOptions();

  // File Input Handlers...
  // (Keep the rest of your existing code here for file inputs)
  var docInput = document.getElementById("libDocInput");
  docInput.onchange = function (e) {
    var f = e.target.files[0]; if (!f) return;
    closeFloatingDialogs();
    showToast("⏳ Reading " + f.name + "...");
    handleFileSelect(f, function (payload) {
      showExtractionStrategy(payload, f.name);
    });
    docInput.value = "";
  };

  document.getElementById("btnLibExport").onclick = handleExportJSON;
  document.getElementById("btnLibImport").onclick = () => document.getElementById("libBackupInput").click();
  document.getElementById("libBackupInput").onchange = handleImportJSON;
}
function bindLibraryFilterEvents() {
  // 1. Search Input (Exists)
  var searchInput = document.getElementById("libSearchInput");
  if (searchInput) {
    searchInput.onkeyup = refreshLibraryList;
  }

  // 2. OLD FILTERS REMOVED
  // The lines for filterCat, filterAg, filterEnt, and sortBy have been deleted
  // because those elements were replaced by the "Group By" dropdowns.
  // (The Group By dropdowns are already bound inside createFilterControls)
}
// REPLACE: runSmartLibrarySearch (Respects Active Filter)
// ==========================================
function runSmartLibrarySearch(savedText) {
  var textPromise = savedText ? Promise.resolve(savedText) : validateApiAndGetText(true);

  textPromise.then(selectionText => {
    if (!selectionText || selectionText.trim().length < 2) {
      showToast("⚠️ Select text first.");
      return;
    }

    var cleanSel = selectionText.toLowerCase().trim();

    // 1. GET ACTIVE FILTER
    var scopeEl = document.getElementById("libraryScopeMode");
    var scopeMode = scopeEl ? scopeEl.value : "all";

    // 2. DEFINE THE DATA POOL BASED ON FILTER
    var subset = [];
    var scopeMsg = "";

    if (scopeMode === "public") {
      // SEARCH EXTERNAL ONLY
      subset = eliClauseLibrary.filter(i => i.type === "Public");
      scopeMsg = "External Resources";
    } else if (scopeMode === "standard") {
      // SEARCH COLGATE STANDARD ONLY
      subset = eliClauseLibrary.filter(i => i.type === "Standard");
      scopeMsg = "Colgate Standard";
    } else if (scopeMode === "thirdparty") {
      // SEARCH 3RD PARTY / NEGOTIATED
      subset = eliClauseLibrary.filter(i => i.type !== "Standard" && i.type !== "Public");
      scopeMsg = "3rd Party / Negotiated";
    } else {
      // SEARCH INTERNAL ALL (No Public)
      subset = eliClauseLibrary.filter(i => i.type !== "Public");
      scopeMsg = "Internal Library";
    }

    var results = [];

    // 3. RUN MATCHING LOGIC
    subset.forEach(item => {
      // A. Exact Title Match
      if (item.title && cleanSel === item.title.toLowerCase()) {
        results.push({ item: item, score: 1.0, matchType: "Exact Title" });
        return;
      }

      // B. Similarity Match
      var score = calculateTextSimilarity(selectionText, item.text);

      if (score >= userSettings.matchThreshold) {
        results.push({ item: item, score: score, matchType: "Body Match" });
      }
    });

    // 4. UPDATE STATE
    window.currentMatchData = {
      results: results,
      selection: selectionText
    };

    // 5. RENDER
    if (results.length === 0) {
      var thresholdPct = Math.round(userSettings.matchThreshold * 100);
      showToast(`⚠️ No matches found in ${scopeMsg}.`);
      renderComparisonResults([], selectionText);
      return;
    }

    results.sort((a, b) => b.score - a.score);
    renderComparisonResults(results, selectionText);
    showToast(`🔍 Matched against: ${scopeMsg}`);

  }).catch(e => showToast("Error: " + e.message));
}
function calculateTextSimilarity(s1, s2) {
  if (!s1 || !s2) return 0;

  // A. Clean Punctuation
  var clean1 = s1.toLowerCase().replace(/[^a-z0-9\s]/g, "");
  var clean2 = s2.toLowerCase().replace(/[^a-z0-9\s]/g, "");

  // B. Stop Words (Noise Filter)
  var stopWords = new Set([
    "the", "and", "or", "of", "to", "a", "in", "for", "is", "on", "that", "by", "this", "with", "it", "as", "be", "are", "at", "from", "an", "was", "have", "not",
    "agreement", "party", "parties", "shall", "will", "section", "clause", "hereby", "herein", "thereto", "subject", "terms", "conditions"
  ]);

  function getTokens(str) {
    // Split by space, filter out short words & stop words
    return str.split(/\s+/).filter(w => w.length > 2 && !stopWords.has(w));
  }

  var tokens1 = new Set(getTokens(clean1));
  var tokens2 = new Set(getTokens(clean2));

  // C. Prevent Divide-by-Zero (Fixes Infinity)
  if (tokens1.size === 0 || tokens2.size === 0) return 0;

  // D. Calculate Intersection
  var intersection = 0;
  tokens1.forEach(t => {
    if (tokens2.has(t)) intersection++;
  });

  // E. Calculate Union
  var union = new Set([...tokens1, ...tokens2]).size;

  // Safety check again
  if (union === 0) return 0;

  return intersection / union;
}
// REPLACE: renderComparisonResults (Stable Grouping)
// ==========================================
function renderComparisonResults(results, userSelection) {
  var container = document.getElementById("libListContainer");
  container.innerHTML = "";

  // Summary Header
  var pct = Math.round((userSettings.matchThreshold || 0.5) * 100);
  container.innerHTML = `
        <div style="padding:12px; background:linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); color:#37474f; font-size:11px; margin-bottom:10px; border:1px solid #b0bec5; border-radius:0px; box-shadow:0 2px 5px rgba(0,0,0,0.05);">
            <strong>🎯 Matches Found:</strong> ${results.length}<br>
            <span style="opacity:0.8;">Showing best matches above ${pct}% threshold.</span>
        </div>`;

  // 1. Determine Grouping
  var g1 = document.getElementById("groupBy1").value;
  var g2 = document.getElementById("groupBy2").value;
  var g3 = document.getElementById("groupBy3").value;
  var groupLevels = [g1, g2, g3].filter(g => g && g !== "none");

  // 2. Render Logic
  // FIX: Removed the "results.length > 2" check. 
  // Now it ALWAYS groups if the user selected a group, ensuring UI consistency when switching tabs.
  if (groupLevels.length > 0) {
    renderNestedMatchGroups(container, results, groupLevels, 0, userSelection);
  } else {
    // Flat List (Only if "None" is selected for grouping)
    results.forEach((res, index) => {
      container.appendChild(createMatchCard(res, index, userSelection));
    });
  }
}

// Helper: Recursive Group Renderer for Matches
function renderNestedMatchGroups(parentContainer, items, levels, currentDepth, userSelection) {
  // BASE CASE: Render Cards
  if (currentDepth >= levels.length) {
    items.forEach((res, i) => {
      var idx = Math.floor(Math.random() * 100000);
      parentContainer.appendChild(createMatchCard(res, idx, userSelection));
    });
    return;
  }

  // 1. Group items
  var currentKey = levels[currentDepth];
  var groups = {};

  items.forEach(res => {
    // Access the underlying clause item via res.item
    var val = res.item[currentKey] || "Uncategorized";
    if (!groups[val]) groups[val] = [];
    groups[val].push(res);
  });

  // 2. Sort groups
  var sortedKeys = Object.keys(groups).sort();

  // 3. Render Groups
  sortedKeys.forEach(groupName => {
    var groupItems = groups[groupName];
    var uniqueId = "match_grp_" + currentDepth + "_" + groupName.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);

    // Styling (Same as Library)
    var headerStyle = "";
    var contentStyle = "display:none;"; // Collapsed by default

    if (currentDepth === 0) {
      headerStyle = "background:#f5f5f5; border:1px solid #d1d5db; border-left:4px solid #546e7a; padding:8px 10px; font-weight:700; color:#546e7a; font-size:11px; margin-bottom:5px;";
      contentStyle += "padding:10px; border:1px solid #e0e0e0; border-top:none; background:#fafafa; margin-bottom:5px;";
    } else if (currentDepth === 1) {
      headerStyle = "background:#e3f2fd; border-left:3px solid #1565c0; padding:6px 10px; font-weight:700; color:#1565c0; font-size:10px; margin-top:5px;";
      contentStyle += "padding:5px 0 5px 10px; border-left:1px solid #e3f2fd; margin-bottom:5px;";
    } else {
      headerStyle = "background:#fff; border-bottom:1px solid #eee; padding:5px 10px; font-weight:600; color:#666; font-size:10px; margin-top:5px; font-style:italic;";
      contentStyle += "padding:5px 0 5px 15px;";
    }

    var header = document.createElement("div");
    header.style.cssText = headerStyle + " cursor:pointer; display:flex; justify-content:space-between; align-items:center; user-select:none;";

    var label = currentKey === "entity" ? "👤 " : (currentKey === "agreement" ? "📄 " : "");

    header.innerHTML = `
            <div style="display:flex; align-items:center; gap:6px;">
                <span id="arrow_${uniqueId}" class="arrow-icon">▶</span>
                <span>${label}${groupName.toUpperCase()}</span>
            </div>
            <span style="font-size:9px; opacity:0.7;">${groupItems.length}</span>
        `;

    var content = document.createElement("div");
    content.id = uniqueId;
    content.style.cssText = contentStyle;

    header.onclick = function (e) {
      e.stopPropagation();
      var arrow = document.getElementById("arrow_" + uniqueId);
      if (content.style.display === "none") {
        content.style.display = "block";
        if (arrow) arrow.innerText = "▼";
      } else {
        content.style.display = "none";
        if (arrow) arrow.innerText = "▶";
      }
    };

    parentContainer.appendChild(header);
    parentContainer.appendChild(content);

    // Recurse
    renderNestedMatchGroups(content, groupItems, levels, currentDepth + 1, userSelection);
  });
}
// REPLACE: createMatchCard (Added AI Rewrite Button)
// ==========================================
function createMatchCard(res, index, userSelection) {
  var item = res.item;
  var percentage = Math.floor(res.score * 100);
  var diffHtml = (typeof generateDiffHtml === "function") ? generateDiffHtml(userSelection, item.text) : item.text;

  var badgeColor = res.matchType === "Exact Title" ? "#2e7d32" : "#1565c0";
  var uniqueId = "match_" + index + "_" + Math.floor(Math.random() * 1000);

  var metaParts = [];
  metaParts.push(item.entity || "Standard");
  if (item.agreement && item.agreement !== "All") metaParts.push(item.agreement);
  if (item.date) metaParts.push(item.date);
  var metaString = metaParts.join(" • ");

  var card = document.createElement("div");
  card.className = "tool-box slide-in";
  card.style.cssText = "padding:12px; margin-bottom:12px; background:white; border:1px solid #cfd8dc; border-left:4px solid " + badgeColor + "; box-shadow:0 2px 8px rgba(0,0,0,0.06); border-radius:0px;";

  var btnStyle = `
            background: linear-gradient(to bottom, #ffffff 0%, #f1f1f1 100%);
            border: 1px solid #ccc;
            padding: 6px 2px;
            font-size: 9px;
            font-weight: 800;
            cursor: pointer;
            text-align: center;
            border-radius: 0px; 
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            transition: all 0.2s ease;
            overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
        `;

  card.innerHTML = `
            <div style="display:flex; justify-content:space-between; margin-bottom:4px; align-items:center;">
                <div style="overflow:hidden;">
                    <div style="font-weight:700; font-size:12px; color:#263238;">${item.title}</div>
                    <div style="font-size:10px; color:#546e7a; margin-top:2px; font-weight:500;">
                        ${metaString}
                    </div>
                </div>
                <div style="font-size:9px; background:#eceff1; color:${badgeColor}; padding:2px 8px; border-radius:2px; font-weight:bold; border:1px solid #cfd8dc; white-space:nowrap; margin-left:8px;">
                    ${percentage}%
                </div>
            </div>
            
            <div style="font-size:11px; font-family:'Georgia', serif; line-height:1.6; background:#fafafa; padding:10px; border:1px solid #eee; margin-bottom:10px; margin-top:8px; max-height:150px; overflow-y:auto; border-radius:0px;">
                ${diffHtml}
            </div>

            <div id="ai_box_${uniqueId}" style="display:none; background:#fff8e1; border:1px solid #ffe0b2; padding:10px; font-size:11px; margin-bottom:10px; color:#5d4037; border-radius:0px;"></div>
            
            <div style="display:grid; grid-template-columns: 1fr 1fr 1fr 1fr 1fr; gap:4px;">
                <button class="btn-insert" title="Replace Selection" style="${btnStyle} color:${badgeColor}; border-bottom:2px solid ${badgeColor};">
                    REPLACE
                </button>
                
                <button class="btn-rewrite" title="Rewrite Favorably" style="${btnStyle} color:#7b1fa2; border-bottom:2px solid #9c27b0;">
                    ✨ REWRITE
                </button>

                <button class="btn-comment" title="Cite Precedent" style="${btnStyle} color:#455a64; border-bottom:2px solid #607d8b;">
                    💬 CITE
                </button>

                <button class="btn-analyze" title="AI Analysis" style="${btnStyle} color:#e65100; border-bottom:2px solid #ff9800;">
                    ⚡ DIFF
                </button>

                <button class="btn-copy" title="Copy to Clipboard" style="${btnStyle} color:#333; border-bottom:2px solid #757575;">
                    COPY
                </button>
            </div>
        `;

  // Bindings
  var btns = card.querySelectorAll("button");
  btns.forEach(b => {
    b.onmouseover = function () { this.style.background = "#fff"; };
    b.onmouseout = function () { this.style.background = "linear-gradient(to bottom, #ffffff 0%, #f1f1f1 100%)"; };
  });

  card.querySelector(".btn-insert").onclick = () => { insertText(item.text); };
  card.querySelector(".btn-copy").onclick = () => {
    navigator.clipboard.writeText(item.text);
    showToast("📋 Copied");
  };
  card.querySelector(".btn-comment").onclick = () => { insertClauseAsComment(item); };

  // NEW BINDING
  card.querySelector(".btn-rewrite").onclick = () => { showRewriteEditor(item.text); };

  card.querySelector(".btn-analyze").onclick = function () {
    var btn = this;
    var output = document.getElementById("ai_box_" + uniqueId);
    if (output.style.display !== "none") { output.style.display = "none"; return; }

    btn.innerText = "⏳...";
    btn.disabled = true;
    output.style.display = "block";
    output.innerHTML = "<i>Analyzing legal differences...</i>";
    runAIClauseComparison(userSelection, item.text, output, btn);
  };

  return card;
}
// ==========================================
// 4. THE AI HELPER (Differences + Effect)
// ==========================================
function runAIClauseComparison(userText, libText, outputContainer, btnElement) {
  var prompt = `
        ACT AS: Senior Legal Counsel.
        TASK: Compare the "DRAFT" clause against the "STANDARD" library clause.
        
        1. STANDARD: "${libText}"
        2. DRAFT: "${userText}"
        
        OUTPUT INSTRUCTIONS:
        Return your analysis in two clear sections using this exact format:
        
        <b>1. KEY DIFFERENCES</b>
        • [Bullet point 1: Factual change (e.g. "Cap changed from 5x to 1x")]
        • [Bullet point 2: Factual change (e.g. "Mutual indemnity removed")]
        
        <b>2. LEGAL EFFECT</b>
        • [Bullet point 1: What this means (e.g. "Significantly increases financial exposure")]
        • [Bullet point 2: What this means (e.g. "Removes protection for IP claims")]
        
        Keep it concise and punchy. No intro or outro text.
    `;

  // Call Gemini
  callGemini(null, prompt).then(response => {
    // Format response (Markdown to HTML basics)
    var formatted = response
      .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>") // Bold Markdown -> HTML
      .replace(/\n/g, "<br>"); // Newlines -> <br>

    outputContainer.innerHTML = `<div style="line-height:1.5;">${formatted}</div>`;

    // Update Button State
    btnElement.innerText = "✅ DONE";
    btnElement.style.background = "#2e7d32"; // Green
    btnElement.style.color = "white";
    btnElement.style.border = "1px solid #1b5e20";
  }).catch(e => {
    outputContainer.innerHTML = "Error: " + e.message;
    btnElement.innerText = "RETRY";
    btnElement.disabled = false;
    btnElement.style.background = "#c62828";
  });
}
function insertClauseAsComment(item) {
  Word.run(function (context) {
    // 1. Get current selection
    var selection = context.document.getSelection();

    // 2. Build the "Negotiation Note"
    var note = `PRECEDENT CITE:\n` +
      `We previously agreed to this clause in the "${item.agreement || 'Previous'}" agreement.\n` +
      `Date: ${item.date || 'Unknown'}\n\n` +
      `STANDARD TEXT:\n"${item.text}"`;

    // 3. Add Comment
    selection.insertComment(note);

    return context.sync();
  })
    .then(() => showToast("💬 Comment Attached"))
    .catch(e => showToast("Error: " + e.message));
}
window.toggleClauseBody = function (id, header) {
  var body = document.getElementById(id);
  var arrow = header.querySelector(".clause-arrow");

  if (body.style.display === "none") {
    body.style.display = "block";
    arrow.innerText = "▼";
    header.style.background = "#f5f9ff"; // Highlight active
  } else {
    body.style.display = "none";
    arrow.innerText = "▶";
    header.style.background = "white";
  }
};
window.toggleAllClauses = function (show) {
  var display = show ? 'block' : 'none';
  var arrow = show ? '▼' : '▶';
  var bg = show ? '#f0f7ff' : '#ffffff'; // Highlights header when expanded

  // 1. Toggle Groups (The Folders)
  document.querySelectorAll('.group-content').forEach(el => el.style.display = display);
  document.querySelectorAll('.arrow-icon').forEach(el => el.innerText = arrow);

  // 2. Toggle Individual Clauses (The Cards) <-- THIS WAS MISSING
  document.querySelectorAll('.clause-body').forEach(el => el.style.display = display);
  document.querySelectorAll('.clause-arrow').forEach(el => el.innerText = arrow);

  // 3. Optional: Highlight Headers so you know they are open
  document.querySelectorAll('.clause-header').forEach(el => el.style.background = bg);
};
function updateDropdownExclusions() {
  var ids = ["groupBy1", "groupBy2", "groupBy3"];
  var selects = ids.map(id => document.getElementById(id));
  var values = selects.map(s => s.value);

  selects.forEach((sel, index) => {
    // Loop through every option in this dropdown
    Array.from(sel.options).forEach(opt => {
      // Check if this option's value is selected in another dropdown
      // (We ignore "none" because you can have "none" multiple times)
      var isSelectedElsewhere = values.some((val, vIndex) =>
        val === opt.value && vIndex !== index && val !== "none"
      );

      if (isSelectedElsewhere) {
        opt.style.display = "none"; // Hide it visually
        opt.disabled = true;       // Disable selection
      } else {
        opt.style.display = "block";
        opt.disabled = false;
      }
    });
  });
}
function handleExportJSON() {
  // 1. Get Data
  if (!eliClauseLibrary || eliClauseLibrary.length === 0) {
    showToast("⚠️ Library is empty");
    return;
  }

  // Generate the text
  var jsonStr = JSON.stringify(eliClauseLibrary, null, 2);

  // 2. Try Auto-Download (Best Effort)
  try {
    var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(jsonStr);
    var a = document.createElement('a');
    a.href = dataStr;
    a.download = "ELI_Clause_Library.json";
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (e) {
    console.log("Auto-download blocked");
  }

  // 3. SHOW MANUAL BACKUP MODAL (The Fix)
  // This ensures you can ALWAYS get your data, even if the browser blocks the file.
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:99999; display:flex; align-items:center; justify-content:center;";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  // Using your squared-off styling
  box.style.cssText = "background:white; width:450px; padding:0; border-radius:0px; border:1px solid #1565c0; box-shadow:0 20px 50px rgba(0,0,0,0.3); display:flex; flex-direction:column;";

  box.innerHTML = `
        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding:12px 15px; border-bottom:3px solid #c9a227; display:flex; justify-content:space-between; align-items:center;">
            <div style="font-weight:700; color:#fff; font-size:12px; text-transform:uppercase;">EXPORT DATA</div>
            <button id="btnCloseExport" style="border:none; background:none; cursor:pointer; color:white; font-weight:bold;">✕</button>
        </div>
        
        <div style="padding:15px; background:#f0f7ff; border-bottom:1px solid #bbdefb;">
            <div style="font-size:11px; color:#1565c0; font-weight:bold; margin-bottom:4px;">Download didn't start?</div>
            <div style="font-size:11px; color:#546e7a;">
                Copy the code below and save it as a text file named <strong>ELI_Library.json</strong>.
            </div>
        </div>

        <textarea id="exportArea" style="width:100%; height:250px; border:none; background:#fafafa; padding:10px; font-family:monospace; font-size:10px; resize:none; box-sizing:border-box; outline:none; color:#333;">${jsonStr}</textarea>
        
        <div style="padding:15px; display:flex; gap:10px; background:#fff; border-top:1px solid #e0e0e0;">
            <button id="btnCopyExport" class="btn-action" style="flex:2; background:#1565c0; color:white; border:none; padding:10px; font-weight:bold;">📄 COPY TO CLIPBOARD</button>
            <button id="btnRetryDownload" class="btn-action" style="flex:1; background:#f5f5f5; color:#333; border:1px solid #ccc;">⬇️ Retry File</button>
        </div>
    `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  // Bindings
  document.getElementById("btnCloseExport").onclick = function () { overlay.remove(); };

  document.getElementById("btnCopyExport").onclick = function () {
    var txt = document.getElementById("exportArea");
    txt.select();
    document.execCommand("copy"); // Fallback for older environments
    navigator.clipboard.writeText(txt.value); // Modern approach

    this.innerHTML = "✅ COPIED!";
    this.style.background = "#2e7d32";
    setTimeout(() => {
      this.innerHTML = "📄 COPY TO CLIPBOARD";
      this.style.background = "#1565c0";
    }, 2000);
  };

  document.getElementById("btnRetryDownload").onclick = function () {
    var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(jsonStr);
    var a = document.createElement('a');
    a.href = dataStr;
    a.download = "ELI_Clause_Library.json";
    document.body.appendChild(a);
    a.click();
    a.remove();
  };

  showToast("ℹ️ Export Window Open");
}
function handleImportJSON(e) {
  var f = e.target.files[0];
  if (!f) return;

  var r = new FileReader();
  r.onload = function (evt) {
    try {
      var json = JSON.parse(evt.target.result);
      if (Array.isArray(json)) {
        // NEW WAY: Show Custom Modal
        showImportStrategyModal(json);
      } else {
        showToast("❌ Invalid JSON Structure");
      }
    } catch (err) {
      console.error(err);
      showToast("❌ Invalid JSON File");
    }
  };
  r.readAsText(f);
  e.target.value = ""; // Reset input so you can reload same file if needed
}
// ==========================================
// REPLACE: showImportStrategyModal (Smart Deduplication)
// ==========================================
function showImportStrategyModal(newData) {
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:99999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:380px; padding:0; border-radius:0px; border:1px solid #1565c0; border-left:5px solid #1565c0; box-shadow:0 20px 50px rgba(0,0,0,0.4); display:flex; flex-direction:column;";

  var viewMain = `
        <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding:15px; border-bottom:3px solid #c9a227;">
            <div style="font-weight:800; color:#fff; font-size:12px; text-transform:uppercase; letter-spacing:1px;">IMPORT STRATEGY</div>
            <div style="color:#90a4ae; font-size:11px; margin-top:4px;">Incoming: <strong>${newData.length} clauses</strong>.</div>
        </div>
        
        <div style="padding:25px;">
            <div style="font-size:13px; color:#333; margin-bottom:20px; line-height:1.5;">
                How would you like to handle duplicates?
            </div>

            <button id="btnSmartMerge" class="btn-action" style="width:100%; background:#1565c0; color:white; border:1px solid #0d47a1; padding:12px; font-weight:800; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;">
                <span>🧠 SMART MERGE (Recommended)</span>
                <span style="font-size:9px; opacity:0.9; font-weight:normal; text-transform:none;">Skip Duplicates & Add New</span>
            </button>

            <button id="btnForceAdd" class="btn-action" style="width:100%; background:#fff; color:#1565c0; border:1px solid #bbdefb; padding:10px; font-weight:700; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;">
                <span>➕ FORCE ADD ALL</span>
                <span style="font-size:10px; opacity:0.8; font-weight:normal; text-transform:none; color:#546e7a;">Keep Both Versions</span>
            </button>

            <button id="btnOverwriteRequest" class="btn-action" style="width:100%; background:#fff; color:#c62828; border:1px solid #ffcdd2; padding:10px; font-weight:700; margin-bottom:20px; display:flex; justify-content:space-between; align-items:center;">
                <span>⚠️ OVERWRITE LIBRARY</span>
                <span style="font-size:10px; opacity:0.8; font-weight:normal; text-transform:none; color:#d32f2f;">Delete Current & Replace</span>
            </button>
            
            <button id="btnCancelImport" class="btn-action" style="width:100%; background:#f5f5f5; color:#555; border:1px solid #ddd;">Cancel</button>
        </div>
    `;

  var viewDanger = `
        <div style="background:#ffebee; padding:15px; border-bottom:3px solid #c62828;">
            <div style="font-weight:800; color:#c62828; font-size:12px; text-transform:uppercase; letter-spacing:1px;">⚠️ CONFIRM DELETION</div>
        </div>
        <div style="padding:25px; text-align:center;">
            <div style="font-size:32px; margin-bottom:10px;">🗑️</div>
            <div style="font-size:13px; color:#333; margin-bottom:20px; line-height:1.5;">
                This will <strong>permanently delete</strong> your current library and replace it with the imported file.<br><br>
                <strong>Are you sure?</strong>
            </div>
            <button id="btnConfirmOverwrite" class="btn-action" style="width:100%; background:#c62828; color:white; border:1px solid #b71c1c; padding:12px; font-weight:800; margin-bottom:10px;">YES, OVERWRITE EVERYTHING</button>
            <button id="btnBackToMain" class="btn-action" style="width:100%; background:#fff; color:#555; border:1px solid #ccc;">Cancel / Go Back</button>
        </div>
    `;

  box.innerHTML = viewMain;
  overlay.appendChild(box);
  document.body.appendChild(overlay);

  function bindMainEvents() {
    // --- SMART MERGE LOGIC (De-Duping) ---
    document.getElementById("btnSmartMerge").onclick = function () {
      var addedCount = 0;
      var skippedCount = 0;

      // 1. Create a Set of existing "Signatures" (Title + Text snippet)
      // This lets us detect content duplicates even if IDs don't match.
      var existingSignatures = new Set();
      var existingIds = new Set();

      eliClauseLibrary.forEach(item => {
        if (item.id) existingIds.add(String(item.id));
        // Signature: Title + First 20 chars of text (normalized)
        var sig = (item.title + "|" + (item.text || "").substring(0, 20)).toLowerCase();
        existingSignatures.add(sig);
      });

      // 2. Process Incoming Data
      newData.forEach(item => {
        // A. Ensure ID exists
        if (!item.id) {
          item.id = "cl_imp_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
        }

        // B. Generate Signature
        var itemSig = (item.title + "|" + (item.text || "").substring(0, 20)).toLowerCase();

        // C. Check for Duplicates (ID Match OR Content Match)
        if (existingIds.has(String(item.id)) || existingSignatures.has(itemSig)) {
          skippedCount++;
        } else {
          eliClauseLibrary.push(item);
          addedCount++;
        }
      });

      finalizeImport(overlay, `✅ Added ${addedCount} clauses. (Skipped ${skippedCount} duplicates)`);
    };

    // --- FORCE ADD (Keep All) ---
    document.getElementById("btnForceAdd").onclick = function () {
      // Just ensure unique IDs so we don't break the UI
      newData.forEach(item => {
        // Always regenerate ID to avoid collision
        item.id = "cl_imp_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      });
      eliClauseLibrary = eliClauseLibrary.concat(newData);
      finalizeImport(overlay, "✅ Forced Add: " + newData.length + " clauses.");
    };

    document.getElementById("btnOverwriteRequest").onclick = function () {
      box.innerHTML = viewDanger;
      box.style.borderLeft = "5px solid #c62828";
      bindDangerEvents();
    };
    document.getElementById("btnCancelImport").onclick = function () {
      overlay.remove();
      showToast("Import Cancelled");
    };
  }

  function bindDangerEvents() {
    document.getElementById("btnConfirmOverwrite").onclick = function () {
      // Ensure IDs on the new set before replacing
      newData.forEach(item => {
        if (!item.id) item.id = "cl_imp_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      });

      eliClauseLibrary = newData;
      finalizeImport(overlay, "♻️ Library Replaced.");
    };
    document.getElementById("btnBackToMain").onclick = function () {
      box.innerHTML = viewMain;
      box.style.borderLeft = "5px solid #1565c0";
      bindMainEvents();
    };
  }
  bindMainEvents();
}
function finalizeImport(overlay, msg) {
  saveClauseLibrary();
  overlay.remove();
  showClauseLibraryInterface();
  showToast(msg);
}
// ==========================================
// REPLACE: showClauseEditor (Added Date Picker)
// ==========================================
function showClauseEditor(existing, prefill) {
  var isEdit = !!existing;
  var d = existing || {};
  var initialText = d.text || prefill || "";

  // 1. GATHER EXISTING DATA FOR DROPDOWNS
  var uniqueCats = [...new Set(eliClauseLibrary.map(i => i.category).filter(Boolean))].sort();
  var uniqueEnts = [...new Set(eliClauseLibrary.map(i => i.entity).filter(Boolean))].sort();
  var uniqueAgs = [...new Set(eliClauseLibrary.map(i => i.agreement).filter(Boolean))].sort();

  var buildOpts = (arr) => arr.map(v => `<option value="${v}">`).join("");

  // 2. Create Overlay
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:400px; padding:20px; border-left:5px solid #1565c0; box-shadow:0 10px 25px rgba(0,0,0,0.5); border-radius:0px;";

  var inputStyle = "width:100%; height:28px; font-size:11px; margin-bottom:8px; padding:4px; box-sizing:border-box; border-radius:0px;";

  // --- PREPARE DATE ---
  // Ensure we have a valid YYYY-MM-DD for the input
  var dateVal = d.date || new Date().toISOString().split('T')[0];
  if (dateVal.includes("/")) dateVal = dateVal.replace(/\//g, "-");

  // 3. Render Form
  box.innerHTML = `
        <div style="font-weight:800; color:#1565c0; margin-bottom:15px; font-size:14px; text-transform:uppercase;">${isEdit ? "EDIT CLAUSE" : "SAVE NEW CLAUSE"}</div>
        
        <label class="tool-label">Title</label>
        <input type="text" id="edTitle" class="tool-input" value="${d.title || ""}" placeholder="Clause Name" style="${inputStyle}">

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
            <div>
                <label class="tool-label">Category</label>
                <input type="text" id="edCat" class="tool-input" value="${d.category || "General"}" style="${inputStyle}" list="dlCategory" placeholder="Select or Type...">
                <datalist id="dlCategory">${buildOpts(uniqueCats)}</datalist>
            </div>
            <div>
                <label class="tool-label">Type</label>
                <select id="edType" class="tool-input" style="${inputStyle}">
                    <option value="Standard" ${d.type === 'Standard' ? 'selected' : ''}>Standard</option>
                    <option value="Negotiated" ${d.type === 'Negotiated' ? 'selected' : ''}>Negotiated</option>
                    <option value="Fallback" ${d.type === 'Fallback' ? 'selected' : ''}>Fallback</option>
                    <option value="Preferred" ${d.type === 'Preferred' ? 'selected' : ''}>Preferred</option>
                    <option value="Public" ${d.type === 'Public' ? 'selected' : ''}>Public Resource</option>
                    <option value="Bad Clause" ${d.type === 'Bad Clause' ? 'selected' : ''}>Bad Clause</option>
                </select>
            </div>
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:10px;">
            <div>
                <label class="tool-label">Party Name</label>
                <input type="text" id="edEntity" class="tool-input" value="${d.entity || "Standard"}" placeholder="e.g. Amazon" style="${inputStyle}" list="dlEntity">
                <datalist id="dlEntity">${buildOpts(uniqueEnts)}</datalist>
            </div>
            <div>
                <label class="tool-label">Agreement</label>
                <input type="text" id="edAg" class="tool-input" value="${d.agreement || "All"}" placeholder="e.g. MSA" style="${inputStyle}" list="dlAg">
                <datalist id="dlAg">${buildOpts(uniqueAgs)}</datalist>
            </div>
            <div>
                <label class="tool-label">Date</label>
                <input type="date" id="edDate" class="tool-input" value="${dateVal}" style="${inputStyle} font-family:sans-serif;">
            </div>
        </div>

        <label class="tool-label">Clause Text</label>
        <textarea id="edText" class="tool-input" style="height:100px; font-family:serif; width:100%; box-sizing:border-box; border-radius:0px;">${initialText}</textarea>

        <div style="display:flex; gap:10px; margin-top:15px;">
            <button id="btnSaveClause" class="btn-action" style="flex:1; background:#1565c0; color:white; border:none; padding:10px; font-weight:bold;">SAVE</button>
            <button id="btnCancelClause" class="btn-action" style="flex:1; padding:10px;">Cancel</button>
        </div>
    `;

  // 4. Append to DOM
  overlay.appendChild(box);
  document.body.appendChild(overlay);

  // 5. Bind Events
  var btnCancel = box.querySelector("#btnCancelClause");
  if (btnCancel) {
    btnCancel.onclick = () => overlay.remove();
  }

  var btnSave = box.querySelector("#btnSaveClause");
  if (btnSave) {
    btnSave.onclick = () => {
      // Scoped value retrieval
      var title = box.querySelector("#edTitle").value;
      var txt = box.querySelector("#edText").value;
      var dateInput = box.querySelector("#edDate").value;

      if (!title || !txt) { showToast("⚠️ Title & Text required"); return; }

      // Use input date, or fallback to today if cleared
      var finalDate = dateInput || new Date().toISOString().split('T')[0];

      var newItem = {
        id: isEdit ? d.id : "cl_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
        title: title,
        text: txt,
        category: box.querySelector("#edCat").value,
        type: box.querySelector("#edType").value,
        agreement: box.querySelector("#edAg").value || "All",
        entity: box.querySelector("#edEntity").value || "Standard",
        date: finalDate, // Saved as YYYY-MM-DD
        jurisdiction: d.jurisdiction || "",
        source: d.source || "Manual Entry"
      };

      if (isEdit) {
        var idx = eliClauseLibrary.findIndex(x => x.id === d.id);
        if (idx !== -1) eliClauseLibrary[idx] = newItem;
      } else {
        eliClauseLibrary.push(newItem);
      }

      saveClauseLibrary();
      overlay.remove();
      refreshLibraryList();
      showToast("✅ Saved");
    };
  }
}
// ==========================================
// REPLACE: showAIImportOptions (Added Auto-Generate Button)
// ==========================================
function showAIImportOptions() {
  closeFloatingDialogs();

  var box = document.createElement("div");
  box.id = "floatingDialog";
  box.className = "slide-in";
  box.style.cssText = "position:absolute; top:60px; left:10%; width:80%; background:white; padding:15px; border-top:4px solid #fbc02d; box-shadow:0 10px 30px rgba(0,0,0,0.3); border-radius:0px; z-index:1000;";

  box.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-weight:800; color:#f9a825; font-size:12px; letter-spacing:1px;">⚡ ADD CLAUSES</div>
            <button id="btnCloseDialog" style="border:none; background:none; cursor:pointer; font-size:14px; font-weight:bold; color:#ccc;">✕</button>
        </div>

        <button id="btnAutoGen" style="
            width:100%; margin-bottom:8px; 
            background: #6a1b9a; color: white; 
            border: 1px solid #4a148c; 
            padding: 10px; font-weight: 700; font-size: 11px; 
            cursor: pointer; border-radius: 0px; 
            display: flex; align-items: center; justify-content: center; gap: 6px;
            transition: all 0.2s;
        ">
            <span>🤖</span> Auto-Generate Library (AI)
        </button>

        <button id="btnScanActive" style="
            width:100%; margin-bottom:8px; 
            background: #1565c0; color: white; 
            border: 1px solid #0d47a1; 
            padding: 10px; font-weight: 700; font-size: 11px; 
            cursor: pointer; border-radius: 0px; 
            display: flex; align-items: center; justify-content: center; gap: 6px;
            transition: all 0.2s;
        ">
            <span>📄</span> Scan Active Document
        </button>

        <button id="btnScanUpload" style="
            width:100%; 
            background: linear-gradient(to bottom, #ffffff 0%, #f5f5f5 100%); 
            border: 1px solid #cfd8dc; color: #546e7a; 
            padding: 10px; font-weight: 700; font-size: 11px; 
            cursor: pointer; border-radius: 0px; 
            display: flex; align-items: center; justify-content: center; gap: 6px;
            transition: all 0.2s;
        ">
            <span>📂</span> Upload File (PDF/Docx)
        </button>
    `;

  document.body.appendChild(box);

  // Bindings
  document.getElementById("btnCloseDialog").onclick = closeFloatingDialogs;

  document.getElementById("btnAutoGen").onclick = function () {
    closeFloatingDialogs();
    showAutoGenerateModal(); // Calls the new engine
  };

  document.getElementById("btnScanActive").onclick = function () {
    closeFloatingDialogs();
    setLoading(true);
    validateApiAndGetText().then(text => {
      setLoading(false);
      if (!text) { showToast("⚠️ No text found."); return; }
      showExtractionStrategy(text, "Active Document");
    }).catch(e => {
      setLoading(false);
      showToast("Error: " + e.message);
    });
  };

  document.getElementById("btnScanUpload").onclick = function () {
    closeFloatingDialogs();
    document.getElementById("libDocInput").click();
  };
}
function showExtractionStrategy(payload, sourceName) {
  closeFloatingDialogs();

  // Default Key Clauses List
  // Use saved settings, or fall back to defaults if empty
  var defaultKeys = userSettings.extractionTerms && userSettings.extractionTerms.length > 0
    ? userSettings.extractionTerms
    : ["Indemnity", "Liability", "Termination", "Confidentiality", "Governing Law", "Warranty", "Force Majeure", "IP Rights"];
  var box = document.createElement("div");
  box.id = "floatingDialog";
  box.className = "slide-in";
  box.style.cssText = "position:absolute; top:60px; left:10%; width:80%; background:white; padding:15px; border-top:4px solid #1565c0; box-shadow:0 10px 30px rgba(0,0,0,0.3); border-radius:4px; z-index:1000;";

  box.innerHTML = `
        <div style="font-weight:800; color:#1565c0; font-size:12px; margin-bottom:10px; text-transform:uppercase;">Step 2: Extraction Strategy</div>
        <div style="font-size:11px; color:#666; margin-bottom:15px;">Source: <strong>${sourceName}</strong></div>

        <label style="display:flex; align-items:center; gap:8px; padding:10px; background:#f5f5f5; border-radius:4px; margin-bottom:10px; cursor:pointer;">
            <input type="radio" name="extMode" value="ALL" checked>
            <div>
                <div style="font-weight:bold; font-size:12px; color:#333;">Extract Everything</div>
                <div style="font-size:10px; color:#666;">Exhaustive scan for all substantive clauses.</div>
            </div>
        </label>

        <label style="display:flex; align-items:center; gap:8px; padding:10px; background:#e3f2fd; border:1px solid #bbdefb; border-radius:4px; margin-bottom:10px; cursor:pointer;">
            <input type="radio" name="extMode" value="KEY">
            <div>
                <div style="font-weight:bold; font-size:12px; color:#1565c0;">Key Clauses Only</div>
                <div style="font-size:10px; color:#1565c0;">Target specific high-value clauses.</div>
            </div>
        </label>

        <div id="keyClauseList" style="display:none; margin-left:25px; margin-bottom:15px; border-left:2px solid #ddd; padding-left:10px;">
            <div style="font-size:10px; color:#999; margin-bottom:5px;">SELECT TARGETS:</div>
            <div id="chkContainer" style="display:grid; grid-template-columns: 1fr 1fr; gap:5px; margin-bottom:8px;">
                ${defaultKeys.map(k => `
                    <label style="font-size:11px; display:flex; gap:5px; align-items:center;">
                        <input type="checkbox" value="${k}" checked> ${k}
                    </label>
                `).join('')}
            </div>
            <div style="display:flex; gap:5px;">
                <input type="text" id="newKeyClause" placeholder="+ Add Custom Type" style="font-size:11px; padding:4px; border:1px solid #ccc; flex:1;">
                <button id="btnAddKey" style="font-size:10px; cursor:pointer;">ADD</button>
            </div>
        </div>

        <div style="display:flex; gap:10px; margin-top:15px;">
            <button id="btnCancelStrat" style="flex:1; padding:8px; background:white; border:1px solid #ccc; color:#555; cursor:pointer;">Cancel</button>
            <button id="btnRunStrat" style="flex:2; padding:8px; background:#2e7d32; color:white; border:none; font-weight:bold; cursor:pointer;">🚀 START EXTRACTION</button>
        </div>
    `;

  document.body.appendChild(box);

  // Toggle Checklist Visibility
  var radios = box.querySelectorAll('input[name="extMode"]');
  var listDiv = document.getElementById("keyClauseList");
  radios.forEach(r => r.onchange = () => {
    listDiv.style.display = (r.value === "KEY") ? "block" : "none";
  });

  // Add Custom Tag Logic
  document.getElementById("btnAddKey").onclick = () => {
    var val = document.getElementById("newKeyClause").value.trim();
    if (val) {
      var span = document.createElement("label");
      span.style.cssText = "font-size:11px; display:flex; gap:5px; align-items:center;";
      span.innerHTML = `<input type="checkbox" value="${val}" checked> ${val}`;
      document.getElementById("chkContainer").appendChild(span);
      document.getElementById("newKeyClause").value = "";
    }
  };

  // Cancel
  document.getElementById("btnCancelStrat").onclick = closeFloatingDialogs;

  // Run Logic
  document.getElementById("btnRunStrat").onclick = () => {
    var mode = document.querySelector('input[name="extMode"]:checked').value;
    var targets = "";

    if (mode === "KEY") {
      // Gather checked items
      var checked = Array.from(document.querySelectorAll('#chkContainer input:checked')).map(c => c.value);
      if (checked.length === 0) {
        showToast("⚠️ Select at least one clause type.");
        return;
      }
      targets = checked.join(", ");
    }

    closeFloatingDialogs();
    runAIClauseExtraction(payload, sourceName, targets); // Pass targets to AI
  };
}
// ==========================================
// UPDATED: runAIClauseExtraction (Smart Categorization)
// ==========================================
function runAIClauseExtraction(inputPayload, docName, targetList) {
  setLoading(true);

  var isTargeted = !!targetList;

  var promptText = `
        You are a Senior Legal Knowledge Manager.
        TASK: Extract substantive legal clauses from the document below into a structured JSON format.

        EXTRACTION RULES:

        1. **HIERARCHY & SPLITTING**:
           - **Main Headers:** Treat headers (e.g. "ARTICLE 5") as containers. Do not extract them as text.
           - **Sub-Sections:** Extract numbered sub-sections (e.g. "5.1 Payment") as SEPARATE objects.

        2. **CATEGORIZATION (CRITICAL)**:
           - **Definitions:** If the clause defines a term, set category to "Definition".
           - **Smart Classification:** Analyze the clause content and assign a specific legal category.
             - Examples: "Confidentiality", "Indemnification", "Limitation of Liability", "Termination", "IP Rights", "Warranties", "Governing Law", "Dispute Resolution", "Payment Terms", "Audit Rights".
           - **Fallback:** Use "General" ONLY if the clause does not fit a specific legal topic.

        3. **DATA PARSING RULES**:
           - "clause_number": The alphanumeric identifier (e.g., "2.1", "(a)", "12").
           - "title": The Header text ONLY (e.g., "Payment Terms").
           - "text": The BODY CONTENT only. **DO NOT** include the clause number or title in this field. Start directly with the sentence.
           - "category": The smart category determined above.

        ${isTargeted ? `🚨 STRICT MODE: ONLY EXTRACT CLAUSES RELATED TO: ${targetList}. IGNORE ALL OTHERS.` : "MODE: EXHAUSTIVE extraction."}

        RETURN RAW JSON ARRAY (Maintain Document Order):
        [{
            "clause_number": "2.1",
            "title": "Payment",
            "category": "Payment Terms",
            "text": "Customer shall pay all invoices within thirty (30) days...",
            "agreement": "Agreement Type (inferred)",
            "entity": "Counterparty Name",
            "date": "YYYY-MM-DD",
            "jurisdiction": "State/Country"
        }]`;

  var parts = [];
  if (inputPayload && inputPayload.isBase64) {
    parts = [{ text: promptText }, { inlineData: { mimeType: inputPayload.mime, data: inputPayload.data } }];
  } else {
    var safeText = String(inputPayload).substring(0, 80000);
    parts = [{ text: promptText + "\n\nDOCUMENT TEXT:\n" + safeText }];
  }

  callGemini(null, parts).then(result => {
    setLoading(false);
    try {
      var extracted = tryParseGeminiJSON(result);

      if (Array.isArray(extracted)) {
        // --- CLIENT-SIDE CLEANUP ---
        extracted.forEach(item => {
          // 1. Tag Definitions if needed (Failsafe)
          if (item.title && !item.category) {
            if (item.title.startsWith('"') || item.text.trim().startsWith('"')) {
              item.category = "Definition";
            }
          }

          // 2. Strip Headers from Body Text
          var prefix = ((item.clause_number || "") + " " + (item.title || "")).trim();
          if (item.clause_number && item.text.startsWith(item.clause_number)) {
            item.text = item.text.substring(item.clause_number.length).replace(/^[\.\:\-\s]+/, "").trim();
          }
          if (item.title && item.text.toLowerCase().startsWith(item.title.toLowerCase())) {
            item.text = item.text.substring(item.title.length).replace(/^[\.\:\-\s]+/, "").trim();
          }
        });

        renderImportStaging(extracted, docName);
      } else {
        showToast("⚠️ No clauses found.");
      }
    } catch (e) {
      showToast("❌ Parsing Error");
      console.error(e);
    }
  }).catch(e => {
    setLoading(false);
    showToast("Error: " + e.message);
  });
}
// REPLACE: renderImportStaging (Full Validation: Category Included)
// ==========================================
function renderImportStaging(items, docName) {
  var out = document.getElementById("output");
  out.innerHTML = "";

  // 1. GATHER DATA FOR DROPDOWNS
  var uniqueCats = [...new Set(eliClauseLibrary.map(i => i.category).filter(Boolean))].sort();
  var uniqueEnts = [...new Set(eliClauseLibrary.map(i => i.entity).filter(Boolean))].sort();
  var uniqueAgs = [...new Set(eliClauseLibrary.map(i => i.agreement).filter(Boolean))].sort();

  if (!uniqueCats.includes("Bad Clause")) uniqueCats.unshift("Bad Clause");

  var buildOpts = (arr) => arr.map(v => `<option value="${v}">`).join("");

  var datalistHtml = `
    <datalist id="dlImpCat">${buildOpts(uniqueCats)}</datalist>
    <datalist id="dlImpEnt">${buildOpts(uniqueEnts)}</datalist>
    <datalist id="dlImpAg">${buildOpts(uniqueAgs)}</datalist>
  `;

  // 2. HEADER
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); padding:12px; border-bottom:3px solid #fb8c00; margin-bottom:15px; border-radius:0px;`;

  header.innerHTML = `
    ${datalistHtml}
    <div style="font-size:10px; font-weight:800; color:#e65100; margin-bottom:4px;">IMPORT STAGING</div>
    <div style="font-size:14px; font-weight:700; color:#37474f; margin-bottom:10px;">${items.length} Clauses Extracted</div>
    
    <div style="background:white; padding:10px; border:1px solid #ffcc80; border-radius:0px;">
        <div style="font-size:9px; font-weight:800; color:#e65100; text-transform:uppercase; margin-bottom:6px;">⚡ BULK ASSIGN</div>
        
        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:5px; margin-bottom:5px;">
            <input type="text" id="bulkEnt" placeholder="Party" list="dlImpEnt" class="tool-input" style="height:24px; font-size:10px; margin:0;">
            <input type="text" id="bulkAg" placeholder="Agreement" list="dlImpAg" class="tool-input" style="height:24px; font-size:10px; margin:0;">
            <input type="date" id="bulkDate" class="tool-input" style="height:24px; font-size:10px; margin:0; font-family:sans-serif;">
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:5px; margin-bottom:8px;">
            <input type="text" id="bulkCat" placeholder="Category" list="dlImpCat" class="tool-input" style="height:24px; font-size:10px; margin:0;">
           <select id="bulkType" class="tool-input" style="height:24px; font-size:10px; margin:0; padding:0;">
    <option value="">-- Type --</option>
    <option value="Standard">Standard</option>
    <option value="Public">Public Resource</option>
    <option value="Negotiated">Negotiated</option>
    <option value="Fallback">Fallback</option>
    <option value="Preferred">Preferred</option>
    <option value="Other">Other</option>  <option value="Bad Clause">Bad Clause</option>
</select>
        </div>

        <button id="btnBulkApply" class="btn-action" style="background:#e65100; color:white; border:none; height:24px; font-size:10px; padding:0; width:100%;">APPLY TO ALL</button>
    </div>
  `;
  out.appendChild(header);

  // VIEW CONTROLS
  var viewControls = document.createElement("div");
  viewControls.style.cssText = "display:flex; justify-content:flex-end; gap:10px; padding:0 10px 10px 10px;";
  viewControls.innerHTML = `
    <span id="btnExpAll" style="font-size:10px; font-weight:700; color:#1565c0; cursor:pointer; text-decoration:underline;">Expand All</span>
    <span id="btnColAll" style="font-size:10px; font-weight:700; color:#546e7a; cursor:pointer; text-decoration:underline;">Collapse All</span>
  `;
  out.appendChild(viewControls);

  // 3. LIST CONTAINER
  var listDiv = document.createElement("div");
  listDiv.style.cssText = "padding:0 10px; overflow-y:auto; max-height:calc(100vh - 380px); margin-bottom:15px;";

  items.forEach((item, idx) => {
    var uniqueId = "edit_area_" + idx;
    var card = document.createElement("div");
    card.className = "slide-in clause-card-item";

    function updateCardVisuals() {
      var isDef = item.category === "Definition";
      var isBad = item.type === "Bad Clause" || item.category === "Bad Clause";

      var accentColor = "#1565c0"; // Blue
      if (isDef) accentColor = "#7b1fa2"; // Purple
      if (isBad) accentColor = "#c62828"; // Red

      var labelText = isDef ? "DEFINITION" : (isBad ? "BAD CLAUSE" : "CLAUSE");
      var titlePlaceholder = isDef ? "Definition Name" : "Clause Title";

      card.style.cssText = `background:white; border:1px solid #ddd; border-left:4px solid ${accentColor}; padding:10px; margin-bottom:0px; border-radius:0px; position:relative; transition: border-left-color 0.3s;`;

      var numHtml = "";
      if (item.clause_number) {
        numHtml = `<div style="background:#f5f5f5; color:#555; border:1px solid #ccc; font-size:10px; font-weight:700; padding:3px 6px; margin-right:8px; border-radius:3px; min-width:24px; text-align:center;">${item.clause_number}</div>`;
      }

      var dateValue = item.date || "";
      if (dateValue.includes("/")) dateValue = dateValue.replace(/\//g, "-");
      if (!dateValue || dateValue.length < 10) dateValue = new Date().toISOString().split('T')[0];

      card.innerHTML = `
            <div style="position:absolute; top:2px; right:2px; font-size:8px; font-weight:800; color:${accentColor}; text-transform:uppercase; opacity:0.6;">${labelText}</div>
            
            <div style="display:flex; gap:8px; align-items:flex-start;">
                <input type="checkbox" class="import-chk" checked style="margin-top:6px;">
                
                <div style="flex:1;">
                    <div style="display:flex; align-items:center; margin-bottom:6px;">
                        ${numHtml}
                        <input type="text" class="imp-title tool-input" value="${item.title || ''}" style="font-weight:700; color:${accentColor}; margin:0; height:24px; flex:1; border:none; background:transparent;" placeholder="${titlePlaceholder}">
                        <button id="btnToggle_${uniqueId}" style="background:none; border:none; color:#999; font-size:10px; cursor:pointer; margin-left:5px;">▼</button>
                    </div>

                    <textarea class="imp-text tool-input" style="height:60px; font-family:serif; font-size:11px; margin-bottom:6px; border-color:${isDef ? '#e1bee7' : (isBad ? '#ffcdd2' : '#ccc')};">${item.text || ''}</textarea>

                    <div id="${uniqueId}" class="edit-area hidden" style="background:#f9f9f9; padding:8px; border:1px solid #eee; margin-top:5px;">
                        
                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:5px; margin-bottom:5px;">
                            <div>
                                <label style="font-size:9px; color:#555;">Category <span style="color:red">*</span></label>
                                <input type="text" class="imp-cat tool-input" value="${item.category || 'General'}" list="dlImpCat" placeholder="Required..." style="height:24px; font-size:10px; margin:0;">
                            </div>
                            <div>
                                <label style="font-size:9px; color:#555;">Type <span style="color:red">*</span></label>
                                <select class="imp-type tool-input" style="height:24px; font-size:10px; margin:0; padding:0;">
    <option value="Standard" ${item.type === 'Standard' ? 'selected' : ''}>Standard</option>
    <option value="Public" ${item.type === 'Public' ? 'selected' : ''}>Public Resource</option>
    <option value="Negotiated" ${!item.type || item.type === 'Negotiated' ? 'selected' : ''}>Negotiated</option>
    <option value="Fallback" ${item.type === 'Fallback' ? 'selected' : ''}>Fallback</option>
    <option value="Preferred" ${item.type === 'Preferred' ? 'selected' : ''}>Preferred</option>
    <option value="Other" ${item.type === 'Other' ? 'selected' : ''}>Other</option> <option value="Bad Clause" ${item.type === 'Bad Clause' ? 'selected' : ''}>Bad Clause</option>
</select>
                            </div>
                        </div>

                        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:5px;">
                            <div>
                                <label style="font-size:9px; color:#555;">Party <span style="color:red">*</span></label>
                                <input type="text" class="imp-ent tool-input" value="${item.entity || 'Standard'}" list="dlImpEnt" placeholder="Required..." style="height:24px; font-size:10px; margin:0;">
                            </div>
                            <div>
                                <label style="font-size:9px; color:#555;">Agreement <span style="color:red">*</span></label>
                                <input type="text" class="imp-ag tool-input" value="${item.agreement || 'All'}" list="dlImpAg" placeholder="Required..." style="height:24px; font-size:10px; margin:0;">
                            </div>
                            <div>
                                <label style="font-size:9px; color:#555;">Date <span style="color:red">*</span></label>
                                <input type="date" class="imp-date tool-input" value="${dateValue}" style="height:24px; font-size:10px; margin:0; padding:0 4px; font-family:sans-serif;">
                            </div>
                        </div>

                    </div>
                    
                    <input type="hidden" class="imp-jur" value="${item.jurisdiction || ''}">
                </div>
            </div>
        `;

      var btn = card.querySelector(`#btnToggle_${uniqueId}`);
      if (btn) btn.onclick = () => document.getElementById(uniqueId).classList.toggle('hidden');

      // --- SYNC EVENTS & ERROR CLEARING ---
      card.querySelector(".imp-title").oninput = function () { item.title = this.value; };
      card.querySelector(".imp-text").oninput = function () { item.text = this.value; };

      card.querySelector(".imp-type").onchange = function () {
        item.type = this.value;
        this.style.borderColor = "";
        updateCardVisuals();
      };

      card.querySelector(".imp-ent").oninput = function () {
        item.entity = this.value;
        this.style.borderColor = "";
      };
      card.querySelector(".imp-ag").oninput = function () {
        item.agreement = this.value;
        this.style.borderColor = "";
      };

      card.querySelector(".imp-date").onchange = function () {
        item.date = this.value;
        this.style.borderColor = "";
      };

      var catInput = card.querySelector(".imp-cat");
      if (catInput) {
        catInput.oninput = function () {
          item.category = this.value;
          this.style.borderColor = ""; // Clear Red Border
          updateCardVisuals();
        };
      }
    }

    updateCardVisuals();
    listDiv.appendChild(card);

    // INSERTION BUTTON (+)
    var insertRow = document.createElement("div");
    insertRow.style.cssText = "text-align:center; height:14px; position:relative; cursor:pointer; margin-bottom:10px; opacity:0.3; transition:opacity 0.2s;";
    insertRow.title = "Insert New Clause Here";

    insertRow.innerHTML = `
        <div style="position:absolute; top:50%; left:0; right:0; height:1px; background:#ccc; z-index:0;"></div>
        <span style="position:relative; z-index:1; background:white; border:1px solid #ccc; color:#555; border-radius:50%; width:16px; height:16px; font-size:12px; line-height:14px; display:inline-block; font-weight:bold;">+</span>
    `;

    insertRow.onmouseover = function () { this.style.opacity = "1"; };
    insertRow.onmouseout = function () { this.style.opacity = "0.3"; };

    insertRow.onclick = function () {
      var today = new Date().toISOString().split('T')[0];
      items.splice(idx + 1, 0, {
        title: "New Clause",
        text: "",
        category: "", // Empty to trigger validation
        type: "Standard",
        entity: "",
        agreement: "",
        clause_number: "",
        date: today
      });
      renderImportStaging(items, docName);
    };

    listDiv.appendChild(insertRow);

  });
  out.appendChild(listDiv);

  // 4. FOOTER ACTIONS
  var footer = document.createElement("div");
  footer.style.cssText = "padding:15px; background:#fafafa; border-top:1px solid #eee; display:flex; gap:10px;";
  var btnSave = document.createElement("button");
  btnSave.className = "btn-action";
  btnSave.style.background = "#2e7d32";
  btnSave.style.color = "white";
  btnSave.style.flex = "2";
  btnSave.innerText = "💾 SAVE SELECTED";

  var btnCancel = document.createElement("button");
  btnCancel.className = "btn-action";
  btnCancel.style.flex = "1";
  btnCancel.innerText = "Cancel";

  footer.appendChild(btnCancel);
  footer.appendChild(btnSave);
  out.appendChild(footer);

  // EVENTS
  document.getElementById("btnExpAll").onclick = function () {
    listDiv.querySelectorAll('.edit-area').forEach(el => el.classList.remove('hidden'));
  };
  document.getElementById("btnColAll").onclick = function () {
    listDiv.querySelectorAll('.edit-area').forEach(el => el.classList.add('hidden'));
  };

  document.getElementById("btnBulkApply").onclick = function () {
    var bEnt = document.getElementById("bulkEnt").value.trim();
    var bAg = document.getElementById("bulkAg").value.trim();
    var bCat = document.getElementById("bulkCat").value.trim();
    var bType = document.getElementById("bulkType").value;
    var bDate = document.getElementById("bulkDate").value;

    items.forEach(item => {
      if (bEnt) item.entity = bEnt;
      if (bAg) item.agreement = bAg;
      if (bType) item.type = bType;
      if (bCat) item.category = bCat;
      if (bDate) item.date = bDate;
    });

    renderImportStaging(items, docName);
    showToast(`⚡ Metadata Applied.`);
  };

  btnCancel.onclick = showClauseLibraryInterface;

  // --- SAVE BUTTON WITH EXTENDED VALIDATION ---
  btnSave.onclick = function () {
    var count = 0;
    var cards = listDiv.querySelectorAll(".clause-card-item");
    var hasError = false;

    // 1. Validation Loop
    cards.forEach((c, i) => {
      if (c.querySelector(".import-chk").checked) {
        var itemData = items[i];
        var missing = [];

        if (!itemData.entity || itemData.entity.trim() === "") missing.push("Party");
        if (!itemData.agreement || itemData.agreement.trim() === "") missing.push("Agreement");
        if (!itemData.date) missing.push("Date");
        if (!itemData.type) missing.push("Type");
        if (!itemData.category || itemData.category.trim() === "") missing.push("Category"); // Added Category

        if (missing.length > 0) {
          hasError = true;
          c.querySelector(".edit-area").classList.remove("hidden");

          if (missing.includes("Party")) c.querySelector(".imp-ent").style.borderColor = "red";
          if (missing.includes("Agreement")) c.querySelector(".imp-ag").style.borderColor = "red";
          if (missing.includes("Date")) c.querySelector(".imp-date").style.borderColor = "red";
          if (missing.includes("Type")) c.querySelector(".imp-type").style.borderColor = "red";
          if (missing.includes("Category")) c.querySelector(".imp-cat").style.borderColor = "red";
        }
      }
    });

    if (hasError) {
      showToast("⚠️ Please fill required fields (Red)");
      return;
    }

    // 2. Save Loop
    cards.forEach((c, i) => {
      if (c.querySelector(".import-chk").checked) {
        var itemData = items[i];
        var finalDate = itemData.date || new Date().toISOString().split('T')[0];

        eliClauseLibrary.push({
          id: "cl_scan_" + Date.now() + "_" + Math.floor(Math.random() * 1000) + "_" + i,
          title: itemData.title,
          text: itemData.text,
          category: itemData.category,
          type: itemData.type,
          agreement: itemData.agreement,
          entity: itemData.entity,
          date: finalDate, // YYYY-MM-DD
          jurisdiction: itemData.jurisdiction || "",
          source: docName
        });
        count++;
      }
    });

    saveClauseLibrary();
    showClauseLibraryInterface();
    showToast(`✅ ${count} Clauses Saved!`);
  };
}
// ==========================================
// 2. SETTINGS MODAL (Fixes Key Terms Tweaking)
// ==========================================
// ==========================================
// 2. SETTINGS MODAL (Robust)
// ==========================================
function showMatchSettings() {
  // Close existing to prevent stack
  var existing = document.getElementById("floatingDialog");
  if (existing) existing.remove();

  var box = document.createElement("div");
  box.id = "floatingDialog";
  box.className = "slide-in";
  box.style.cssText = "position:absolute; top:50px; right:15px; width:280px; background:white; padding:15px; border:1px solid #ccc; box-shadow:0 10px 30px rgba(0,0,0,0.2); border-radius:0px; z-index:2000; border-top: 4px solid #333;";

  // Safety Init: Ensure lists exist before rendering
  if (!userSettings.extractionTerms || userSettings.extractionTerms.length === 0) {
    userSettings.extractionTerms = [
      "Indemnity", "Liability", "Termination", "Confidentiality",
      "Governing Law", "Warranty", "Force Majeure", "IP Rights"
    ];
  }
  if (!userSettings.matchThreshold) userSettings.matchThreshold = 0.5;

  var pct = Math.round(userSettings.matchThreshold * 100);

  // Helper: Render Tags
  function renderTags() {
    return userSettings.extractionTerms.map(t =>
      `<span style="background:#e3f2fd; color:#1565c0; padding:3px 6px; font-size:10px; margin-right:4px; margin-bottom:4px; display:inline-block; border:1px solid #bbdefb; font-weight:bold;">
                ${t} <b class="del-tag" data-tag="${t}" style="cursor:pointer; color:#c62828; margin-left:5px;">&times;</b>
            </span>`
    ).join("");
  }

  // UPDATED HEADER: Includes Flexbox justify-content:space-between and the Close X
  box.innerHTML = `
        <div style="font-weight:800; font-size:12px; margin-bottom:15px; color:#333; text-transform:uppercase; border-bottom:1px solid #eee; padding-bottom:5px; display:flex; justify-content:space-between; align-items:center;">
            <span>⚙️ SYSTEM SETTINGS</span>
            <span id="btnCloseX" style="cursor:pointer; font-size:14px; color:#999; font-weight:bold;" title="Close without saving">✕</span>
        </div>
        
        <div style="margin-bottom:20px;">
            <div style="font-size:10px; font-weight:bold; color:#555; margin-bottom:5px; display:flex; justify-content:space-between;">
                <span>MATCH SENSITIVITY</span>
                <span id="lblPct" style="color:#1565c0;">${pct}%</span>
            </div>
            <input type="range" id="rngThreshold" min="10" max="90" step="5" value="${pct}" style="width:100%; cursor:pointer;">
        </div>

        <div style="margin-bottom:10px;">
            <div style="font-size:10px; font-weight:bold; color:#555; margin-bottom:8px;">KEY EXTRACTION TERMS</div>
            
            <div id="tagContainer" style="margin-bottom:10px; min-height:30px; border:1px solid #f5f5f5; padding:5px; background:#fafafa;">
                ${renderTags()}
            </div>

            <div style="display:flex; gap:0px;">
                <input type="text" id="txtNewTerm" placeholder="Add new term..." style="flex:1; font-size:11px; padding:6px; border:1px solid #ccc; border-right:none; border-radius:0px;">
                <button id="btnAddTerm" style="background:#2e7d32; color:white; border:none; padding:0 12px; font-weight:bold; cursor:pointer; font-size:14px; border-radius:0px;">+</button>
            </div>
        </div>

        <button id="btnCloseSettings" style="width:100%; margin-top:15px; padding:10px; background:#333; color:white; border:none; font-size:11px; font-weight:bold; cursor:pointer; text-transform:uppercase;">SAVE & CLOSE</button>
    `;

  document.body.appendChild(box);

  // --- LOGIC ---

  // 1. Close without Saving (The X Button)
  document.getElementById("btnCloseX").onclick = function () {
    // Reloads saved settings from disk to discard any unsaved UI changes
    // This is safer than just closing because the user might have dragged the slider
    var stored = localStorage.getItem("eli_user_settings");
    if (stored) {
      // Revert in-memory object to what is on disk
      userSettings = JSON.parse(stored);
    }
    box.remove();
  };

  // 2. Slider Logic
  var rng = document.getElementById("rngThreshold");
  var lbl = document.getElementById("lblPct");
  rng.oninput = function () {
    lbl.innerText = this.value + "%";
    userSettings.matchThreshold = this.value / 100;
  };

  // 3. Add Term Logic
  document.getElementById("btnAddTerm").onclick = function () {
    var val = document.getElementById("txtNewTerm").value.trim();
    if (val && !userSettings.extractionTerms.includes(val)) {
      userSettings.extractionTerms.push(val);
      document.getElementById("tagContainer").innerHTML = renderTags();
      document.getElementById("txtNewTerm").value = "";
      bindDelButtons(); // Re-bind after render
    }
  };

  // 4. Delete Term (Scoped Binding)
  function bindDelButtons() {
    var dels = box.querySelectorAll(".del-tag"); // Scope to box
    dels.forEach(btn => {
      btn.onclick = function () {
        var tag = this.dataset.tag;
        userSettings.extractionTerms = userSettings.extractionTerms.filter(t => t !== tag);
        document.getElementById("tagContainer").innerHTML = renderTags();
        bindDelButtons(); // Re-bind after re-render
      };
    });
  }
  bindDelButtons(); // Initial bind

  // 5. Save & Close
  document.getElementById("btnCloseSettings").onclick = function () {
    // Uses the global safe save function (assuming saveCurrentSettings exists globally or falling back to local storage)
    if (typeof saveCurrentSettings === 'function') {
      saveCurrentSettings();
    } else {
      localStorage.setItem("eli_user_settings", JSON.stringify(userSettings));
    }
    box.remove();
    if (typeof showToast === 'function') showToast("✅ Settings Saved");
  };
}
var targetToDelete = null; // Global tracker

function promptDelete(target) {
  if (target === 'all') {
    if (typeof showBulkDeleteModal === 'function') {
      showBulkDeleteModal(); // Use advanced wizard if available
    } else {
      // Fallback for simple delete all
      targetToDelete = 'all';
      showDeleteConfirmation("Delete ALL clauses from library?");
    }
  } else {
    targetToDelete = target;
    showDeleteConfirmation("Permanently delete this clause?");
  }
}

function showDeleteConfirmation(msg) {
  // Check if modal already exists
  if (document.getElementById("custom-modal")) {
    document.getElementById("modal-text").innerText = msg;
    document.getElementById("custom-modal").classList.remove("hidden");
    return;
  }

  // Create Modal on the fly if missing
  var modalHTML = `
        <div id="custom-modal" class="modal-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:99999;display:flex;justify-content:center;align-items:center;">
            <div class="modal-content" style="background:#1e1e1e;border:1px solid #333;padding:20px;width:320px;text-align:center;color:white;box-shadow:0 20px 50px rgba(0,0,0,0.5);border-left:5px solid #d32f2f;">
                <h3 style="margin-top:0;color:#ff4d4d;font-size:16px;text-transform:uppercase;">⚠️ Confirmation</h3>
                <p id="modal-text" style="font-size:13px;color:#ccc;line-height:1.5;margin-top:10px;">${msg}</p>
                <div class="modal-actions" style="display:flex;gap:10px;margin-top:20px;">
                    <button id="modalBtnCancel" style="flex:1;padding:10px;background:#444;color:white;border:none;cursor:pointer;">Cancel</button>
                    <button id="modalBtnConfirm" style="flex:1;padding:10px;background:#b71c1c;color:white;border:none;cursor:pointer;">Delete</button>
                </div>
            </div>
        </div>
    `;

  var div = document.createElement('div');
  div.innerHTML = modalHTML;
  document.body.appendChild(div.firstElementChild);

  // Bind
  document.getElementById("modalBtnCancel").onclick = function () {
    document.getElementById("custom-modal").classList.add("hidden");
  };

  document.getElementById("modalBtnConfirm").onclick = function () {
    if (targetToDelete === 'all') {
      eliClauseLibrary = [];
    } else {
      eliClauseLibrary = eliClauseLibrary.filter(x => x.id !== targetToDelete);
    }
    saveClauseLibrary();
    refreshLibraryList();
    showToast("🗑️ Deleted");
    document.getElementById("custom-modal").classList.add("hidden");
  };
}
function showBulkDeleteModal() {
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:450px; padding:0; border-radius:0px; border:1px solid #c62828; border-left:5px solid #c62828; box-shadow:0 20px 50px rgba(0,0,0,0.4); display:flex; flex-direction:column; max-height: 85vh;";

  var hierarchy = {};
  eliClauseLibrary.forEach(item => {
    var p = item.entity || "Unspecified Party";
    var a = item.agreement || "General";
    var d = item.date || "No Date";
    var comboKey = a + "|||" + d;
    if (!hierarchy[p]) hierarchy[p] = {};
    if (!hierarchy[p][comboKey]) hierarchy[p][comboKey] = { count: 0, name: a, date: d };
    hierarchy[p][comboKey].count++;
  });

  var treeHtml = "";
  Object.keys(hierarchy).sort().forEach((party, pIdx) => {
    var groups = hierarchy[party];
    var totalForParty = Object.values(groups).reduce((acc, obj) => acc + obj.count, 0);
    var partySafe = party.replace(/"/g, "&quot;");
    var partyId = `p_${pIdx}`;

    treeHtml += `
            <div class="party-group" style="border:1px solid #eee; margin-bottom:5px; background:#f9f9f9;">
                <div style="padding:8px; display:flex; align-items:center; border-bottom:1px solid #eee;">
                    <input type="checkbox" class="chk-party" id="${partyId}" data-party="${partySafe}" style="margin-right:8px;">
                    <label for="${partyId}" style="flex:1; font-size:12px; font-weight:700; color:#37474f; cursor:pointer;">${party}</label>
                    <span style="font-size:10px; background:#eceff1; padding:2px 6px; border-radius:4px; margin-right:8px;">${totalForParty}</span>
                    <span class="tree-toggle" style="cursor:pointer; font-size:10px; color:#1565c0; padding:4px;">▶</span>
                </div>
                <div class="ag-list" style="padding:5px 0 5px 25px; background:white; display:none;">
        `;

    Object.keys(groups).sort().forEach((key, aIdx) => {
      var item = groups[key];
      var keySafe = key.replace(/"/g, "&quot;");
      var agId = `a_${pIdx}_${aIdx}`;
      treeHtml += `
                <div style="display:flex; align-items:center; padding:4px 0;">
                    <input type="checkbox" class="chk-ag" id="${agId}" data-party="${partySafe}" data-key="${keySafe}" style="margin-right:8px;">
                    <label for="${agId}" style="font-size:11px; color:#546e7a; cursor:pointer;">
                        <span style="font-weight:700;">${item.name}</span> 
                        <span style="color:#90a4ae; margin-left:4px;">(${item.date})</span>
                        <span style="font-size:9px; background:#f5f5f5; border:1px solid #eee; padding:0 4px; border-radius:3px; margin-left:6px;">${item.count}</span>
                    </label>
                </div>`;
    });
    treeHtml += `</div></div>`;
  });

  box.innerHTML = `
        <div style="background: linear-gradient(135deg, #b71c1c 0%, #d32f2f 100%); padding:15px; border-bottom:3px solid #ffcdd2;">
            <div style="font-weight:800; color:#fff; font-size:12px; text-transform:uppercase; letter-spacing:1px;">BULK DELETE</div>
            <div style="color:#ffcdd2; font-size:11px; margin-top:4px;">Library contains <strong>${eliClauseLibrary.length} clauses</strong>.</div>
        </div>
        <div style="padding:15px; overflow-y:auto; flex:1;">
            <div style="display:flex; gap:10px; margin-bottom:15px; border-bottom:1px solid #eee; padding-bottom:10px;">
                <label style="font-size:11px; cursor:pointer; font-weight:700; color:#37474f;"><input type="radio" name="delView" value="tree" checked> By Party & Agreement</label>
                <label style="font-size:11px; cursor:pointer; font-weight:700; color:#37474f;"><input type="radio" name="delView" value="date"> By Date</label>
                <label style="font-size:11px; cursor:pointer; font-weight:700; color:#c62828;"><input type="radio" name="delView" value="all"> Nuke All</label>
            </div>
            <div id="viewTree">
                <div style="font-size:10px; color:#999; margin-bottom:5px;">Expand a party to see agreements:</div>
                ${treeHtml}
            </div>
            <div id="viewDate" style="display:none; padding:20px; background:#f9f9f9; text-align:center;">
                <div style="margin-bottom:10px; font-size:12px; font-weight:bold;">Delete items older than:</div>
                <input type="date" id="delDateInput" class="tool-input" style="width:100%;">
            </div>
            <div id="viewAll" style="display:none; padding:20px; background:#ffebee; text-align:center; border:1px dashed #c62828;">
                <div style="font-size:24px;">☢️</div>
                <div style="font-weight:800; color:#c62828; margin-top:10px;">DANGER ZONE</div>
            </div>
        </div>
        <div style="padding:15px; border-top:1px solid #eee; background:#f5f5f5; display:flex; gap:10px;">
            <button id="btnExecuteDelete" class="btn-action" style="flex:1; background:#c62828; color:white; border:none; font-weight:bold;">DELETE SELECTED</button>
            <button id="btnCancelDelete" class="btn-action" style="flex:1; background:white; color:#555; border:1px solid #ddd;">Cancel</button>
        </div>
    `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  var radios = box.querySelectorAll('input[name="delView"]');
  var vTree = document.getElementById("viewTree"), vDate = document.getElementById("viewDate"), vAll = document.getElementById("viewAll"), btn = document.getElementById("btnExecuteDelete");

  radios.forEach(r => {
    r.onchange = function () {
      vTree.style.display = "none"; vDate.style.display = "none"; vAll.style.display = "none";
      if (this.value === "tree") { vTree.style.display = "block"; btn.innerText = "DELETE SELECTED"; }
      else if (this.value === "date") { vDate.style.display = "block"; btn.innerText = "DELETE OLDER"; }
      else { vAll.style.display = "block"; btn.innerText = "NUKE EVERYTHING"; }
    };
  });

  box.querySelectorAll(".tree-toggle").forEach(arrow => {
    arrow.onclick = function () {
      var list = this.parentElement.nextElementSibling;
      if (list.style.display === "none") { list.style.display = "block"; this.innerText = "▼"; }
      else { list.style.display = "none"; this.innerText = "▶"; }
    };
  });

  box.querySelectorAll(".chk-party").forEach(chk => {
    chk.onchange = function () {
      var children = box.querySelectorAll(`.chk-ag[data-party="${this.dataset.party}"]`);
      children.forEach(c => c.checked = this.checked);
    };
  });

  document.getElementById("btnCancelDelete").onclick = () => overlay.remove();
  document.getElementById("btnExecuteDelete").onclick = function () {
    var mode = document.querySelector('input[name="delView"]:checked').value;
    var startCount = eliClauseLibrary.length;

    if (mode === "all") eliClauseLibrary = [];
    else if (mode === "date") {
      var dateStr = document.getElementById("delDateInput").value;
      if (!dateStr) { showToast("⚠️ Pick a date."); return; }
      eliClauseLibrary = eliClauseLibrary.filter(x => !x.date || x.date >= dateStr);
    } else if (mode === "tree") {
      var checkedAgs = box.querySelectorAll(".chk-ag:checked");
      if (checkedAgs.length === 0) { showToast("⚠️ Nothing selected."); return; }
      var toDelete = new Set();
      checkedAgs.forEach(chk => toDelete.add(chk.dataset.party + "|" + chk.dataset.key));
      eliClauseLibrary = eliClauseLibrary.filter(item => !toDelete.has((item.entity || "Unspecified Party") + "|" + (item.agreement || "General") + "|||" + (item.date || "No Date")));
    }

    saveClauseLibrary();
    refreshLibraryList();
    overlay.remove();
    showToast(`✅ ${startCount - eliClauseLibrary.length} removed`);
  };
}
function generateDiffHtml(text1, text2) {
  // 1. Tokenize (Split into words)
  var arr1 = text1.split(/\s+/);
  var arr2 = text2.split(/\s+/);

  // 2. Create LCS Matrix
  var matrix = Array(arr1.length + 1).fill(null).map(() => Array(arr2.length + 1).fill(0));

  for (var i = 1; i <= arr1.length; i++) {
    for (var j = 1; j <= arr2.length; j++) {
      if (arr1[i - 1] === arr2[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1] + 1;
      } else {
        matrix[i][j] = Math.max(matrix[i - 1][j], matrix[i][j - 1]);
      }
    }
  }

  // 3. Backtrack to generate Diff HTML
  var i = arr1.length;
  var j = arr2.length;
  var result = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && arr1[i - 1] === arr2[j - 1]) {
      result.unshift(`<span style="color:#333;">${arr1[i - 1]}</span>`);
      i--; j--;
    } else if (j > 0 && (i === 0 || matrix[i][j - 1] >= matrix[i - 1][j])) {
      result.unshift(`<span style="color:#2e7d32; font-weight:bold; text-decoration:underline; background:#e8f5e9;">${arr2[j - 1]}</span>`);
      j--;
    } else {
      result.unshift(`<span style="color:#c62828; text-decoration:line-through; background:#ffebee;">${arr1[i - 1]}</span>`);
      i--;
    }
  }
  return result.join(" ");
}
function showClauseLibraryInterface() {
  var out = document.getElementById("output");
  out.innerHTML = "";

  // 1. Build Static Layout
  out.appendChild(createLibraryHeader());
  out.appendChild(createHiddenFileInputs());
  out.appendChild(createFilterControls());

  // 2. Create Scrollable List Container
  var listContainer = document.createElement("div");
  listContainer.id = "libListContainer";
  listContainer.style.cssText = "padding:0 15px 20px 15px; overflow-y:auto; max-height:calc(100vh - 150px);";
  out.appendChild(listContainer);

  // 3. Initial Render
  refreshLibraryList();

  // 4. Bind Global Events
  bindLibraryHeaderEvents();
  bindLibraryFilterEvents();
}
// ==========================================
// ADD: Missing Info Modal Function
// ==========================================
function showInfoModal(item) {
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:99999; display:flex; align-items:center; justify-content:center;";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #1565c0; border-radius:4px; box-shadow:0 10px 25px rgba(0,0,0,0.3);";

  box.innerHTML = `
        <div style="font-weight:800; color:#1565c0; margin-bottom:10px; font-size:12px; text-transform:uppercase;">CLAUSE METADATA</div>
        <div style="font-size:12px; color:#333; line-height:1.6;">
            <strong>📄 Source:</strong> ${item.source || "Manual Entry"}<br>
            <strong>🗓️ Date:</strong> ${item.date || "N/A"}<br>
            <strong>🏢 Counterparty:</strong> ${item.entity || "Standard"}<br>
            <strong>⚖️ Jurisdiction:</strong> ${item.jurisdiction || "N/A"}
        </div>
        <button id="btnCloseInfo" class="btn-action" style="width:100%; margin-top:15px; border-radius:0px;">Close</button>
    `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);
  document.getElementById("btnCloseInfo").onclick = () => overlay.remove();
}
// 2. THE MODAL: Asks user for Redline vs Standard
function showRedlinePrompt(textToInsert) {
  // Prevent duplicates
  if (document.getElementById("redline-modal")) return;

  var overlay = document.createElement("div");
  overlay.id = "redline-modal";
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:99999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:linear-gradient(135deg, #ffffff 0%, #f1f1f1 100%); width:300px; padding:20px; border:1px solid #b0bec5; border-top:5px solid #1565c0; box-shadow:0 20px 50px rgba(0,0,0,0.5); border-radius:0px; text-align:center;";

  box.innerHTML = `
        <div style="font-weight:800; color:#37474f; font-size:13px; margin-bottom:10px; text-transform:uppercase; letter-spacing:1px;">INSERTION MODE</div>
        <div style="font-size:11px; color:#546e7a; margin-bottom:20px; line-height:1.5;">
            How would you like to insert this clause?
        </div>

        <div style="display:flex; flex-direction:column; gap:8px;">
            <button id="btnRedlineYes" style="
                background: white; border: 1px solid #1565c0; border-left: 5px solid #1565c0;
                color: #1565c0; padding:12px; font-weight:800; font-size:10px; cursor:pointer;
                display:flex; justify-content:space-between; align-items:center;
            ">
                <span>📝 ENABLE REDLINES</span>
                <span style="font-size:9px; opacity:0.7;">Track Changes ON</span>
            </button>

            <button id="btnRedlineNo" style="
                background: white; border: 1px solid #cfd8dc; border-left: 5px solid #546e7a;
                color: #546e7a; padding:12px; font-weight:800; font-size:10px; cursor:pointer;
                display:flex; justify-content:space-between; align-items:center;
            ">
                <span>📄 STANDARD INSERT</span>
                <span style="font-size:9px; opacity:0.7;">No Tracking</span>
            </button>
        </div>
        
        <div style="margin-top:15px; font-size:10px; color:#999; cursor:pointer; text-decoration:underline;" id="btnCancelInsert">Cancel</div>
    `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  // Bind Actions
  document.getElementById("btnRedlineYes").onclick = () => {
    executeInsert(textToInsert, true);
    overlay.remove();
  };

  document.getElementById("btnRedlineNo").onclick = () => {
    executeInsert(textToInsert, false);
    overlay.remove();
  };

  document.getElementById("btnCancelInsert").onclick = () => overlay.remove();
}

// 3. THE EXECUTION: Actually puts text into Word
function executeInsert(text, turnOnRedlines) {
  Word.run(function (context) {
    if (turnOnRedlines) {
      context.document.changeTrackingMode = "TrackAll";
    }
    var selection = context.document.getSelection();
    selection.insertText(text, "Replace");
    return context.sync();
  })
    .then(() => {
      showToast(turnOnRedlines ? "📝 Inserted (Redlines ON)" : "✅ Inserted");
    })
    .catch(e => showToast("Error: " + e.message));
}
function closeFloatingDialogs() {
  var existing = document.getElementById("floatingDialog");
  if (existing) existing.remove();
}
// ==========================================
// NEW ENGINE: AI Rewrite Assistant (Interactive)
// ==========================================
function showRewriteEditor(originalText) {
  // 1. Create Overlay
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:450px; padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #7b1fa2; box-shadow:0 20px 50px rgba(0,0,0,0.4); display:flex; flex-direction:column;";

  // 2. Default Instruction
  var defaultInstruction = "Rewrite this clause to be more favorable to Colgate-Palmolive, prioritizing our rights and limiting our liability.";

  // 3. UI Structure
  box.innerHTML = `
    <div style="background: linear-gradient(135deg, #4a148c 0%, #7b1fa2 100%); padding:12px 15px; border-bottom:3px solid #ab47bc; display:flex; justify-content:space-between; align-items:center;">
        <div style="font-weight:700; color:#fff; font-size:12px; text-transform:uppercase; letter-spacing:1px;">✨ AI REWRITE ASSISTANT</div>
        <button id="btnCloseRewrite" style="border:none; background:none; cursor:pointer; color:white; font-weight:bold; font-size:14px;">✕</button>
    </div>
    
    <div style="padding:15px; background:#f9f9f9; border-bottom:1px solid #eee;">
        <div style="font-size:10px; font-weight:800; color:#7b1fa2; margin-bottom:5px; text-transform:uppercase;">AI INSTRUCTIONS</div>
        <div style="display:flex; gap:5px;">
            <input type="text" id="rwInstructions" class="tool-input" value="${defaultInstruction}" style="flex:1; height:28px; font-size:11px; border:1px solid #e1bee7;">
            <button id="btnRegen" class="btn-action" style="width:auto; padding:0 12px; background:#7b1fa2; color:white; border:none; font-weight:700; font-size:10px;">↻ RUN</button>
        </div>
        <div style="font-size:9px; color:#666; margin-top:4px;">Edit instructions above and click RUN to change the output.</div>
    </div>

    <div style="padding:15px; flex:1;">
        <div style="font-size:10px; font-weight:800; color:#4a148c; margin-bottom:5px; text-transform:uppercase;">RESULT PREVIEW (EDITABLE)</div>
        <textarea id="rwResult" class="tool-input" style="width:100%; height:150px; font-family:'Georgia', serif; font-size:12px; line-height:1.5; padding:10px; border:1px solid #ccc; box-sizing:border-box; resize:none;">Loading AI Suggestion...</textarea>
    </div>
    
    <div style="padding:15px; display:flex; gap:10px; background:#fff; border-top:1px solid #e0e0e0;">
        <button id="btnRwInsert" class="btn-action" style="flex:2; background:#2e7d32; color:white; border:none; padding:10px; font-weight:bold;">📍 INSERT INTO DOC</button>
        <button id="btnRwCopy" class="btn-action" style="flex:1; background:#f5f5f5; color:#333; border:1px solid #ccc;">📄 Copy</button>
    </div>
  `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  // --- LOGIC ---
  var txtInst = document.getElementById("rwInstructions");
  var txtRes = document.getElementById("rwResult");
  var btnRegen = document.getElementById("btnRegen");

  // Helper: Run AI
  function runAI() {
    txtRes.value = "⏳ Generating favorable clause...";
    txtRes.style.color = "#999";
    btnRegen.disabled = true;
    btnRegen.style.opacity = "0.5";

    var prompt = `
        ACT AS: Senior Counsel for Colgate-Palmolive.
        TASK: Rewrite the clause below based on these specific instructions:
        "${txtInst.value}"
        
        ORIGINAL CLAUSE:
        "${originalText}"
        
        OUTPUT RULES:
        1. Return ONLY the rewritten legal text. No explanations.
        2. Ensure the tone is professional and precise.
    `;

    callGemini(API_KEY, prompt).then(text => {
      txtRes.value = text.replace(/^"|"$/g, '').trim(); // Clean quotes
      txtRes.style.color = "#333";
      btnRegen.disabled = false;
      btnRegen.style.opacity = "1";
    }).catch(e => {
      txtRes.value = "Error: " + e.message;
      btnRegen.disabled = false;
      btnRegen.style.opacity = "1";
    });
  }

  // 1. Run immediately on open
  runAI();

  // 2. Bind Regenerate
  btnRegen.onclick = runAI;

  // 3. Bind Insert
  document.getElementById("btnRwInsert").onclick = function () {
    var finalVal = txtRes.value;
    if (!finalVal || finalVal.includes("Loading") || finalVal.includes("Error")) {
      showToast("⚠️ Wait for generation to finish.");
      return;
    }
    insertText(finalVal); // Uses your existing global inserter
    overlay.remove();
    showToast("✅ Rewritten Clause Inserted");
  };

  // 4. Bind Copy
  document.getElementById("btnRwCopy").onclick = function () {
    navigator.clipboard.writeText(txtRes.value);
    this.innerText = "✅ COPIED";
    setTimeout(() => { this.innerText = "📄 Copy"; }, 1500);
  };

  // 5. Bind Close
  document.getElementById("btnCloseRewrite").onclick = function () {
    overlay.remove();
  };
}
// ==========================================
// NEW: showAutoGenerateModal (Topic Selector)
// ==========================================
function showAutoGenerateModal() {
  var overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:20000; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);";

  var box = document.createElement("div");
  box.className = "tool-box slide-in";
  box.style.cssText = "background:white; width:350px; padding:0; border-radius:0px; border:1px solid #d1d5db; border-left:5px solid #6a1b9a; box-shadow:0 20px 50px rgba(0,0,0,0.4);";

  box.innerHTML = `
    <div style="background: linear-gradient(135deg, #4a148c 0%, #7b1fa2 100%); padding: 15px; border-bottom: 3px solid #ab47bc;">
        <div style="font-size:10px; font-weight:800; color:#e1bee7; text-transform:uppercase;">AI LIBRARY BUILDER</div>
        <div style="font-family:'Georgia', serif; font-size:16px; color:white;">🤖 Auto-Generate Clauses</div>
    </div>
    
    <div style="padding:20px;">
        <div style="font-size:12px; color:#546e7a; margin-bottom:15px; line-height:1.4;">
            Select a legal domain below. ELI will generate <strong>25 standard clauses</strong> for that topic and prepare them for import.
        </div>

        <label class="tool-label">Legal Domain / Topic</label>
        <select id="genTopic" class="audience-select" style="width:100%; height:32px; font-weight:600; border-color:#6a1b9a; margin-bottom:15px;">
            <option value="General Commercial & Services">General Commercial (MSA, SOW)</option>
            <option value="Intellectual Property & Licensing">IP & Licensing (JDA, License)</option>
            <option value="Clinical Studies & R&D">Clinical & R&D (CSA, MTA)</option>
            <option value="Marketing & Talent">Marketing (Speaker, Influencer)</option>
            <option value="Data Privacy & Compliance">Data Privacy & Compliance</option>
            <option value="Consumer Studies">Consumer Studies & Liability</option>
        </select>

        <button id="btnRunGen" class="btn-action" style="width:100%; background:#6a1b9a; color:white; border:none; padding:12px; font-weight:800; text-transform:uppercase;">
            🚀 GENERATE 25 CLAUSES
        </button>
        
        <button id="btnCancelGen" class="btn-action" style="width:100%; margin-top:10px;">Cancel</button>
    </div>
  `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  document.getElementById("btnCancelGen").onclick = () => overlay.remove();

  document.getElementById("btnRunGen").onclick = function () {
    var topic = document.getElementById("genTopic").value;
    overlay.remove();
    runAutoGeneration(topic);
  };
}
// ==========================================
// NEW: runAutoGeneration (The AI Engine)
// ==========================================
function runAutoGeneration(topic) {
  setLoading(true);
  showToast("🤖 Generating " + topic + " library...");

  var prompt = `
    You are a Senior Legal Operations Architect for Colgate-Palmolive.
    
    TASK: Generate a synthetic JSON dataset of **25 distinct legal clauses** related to: "${topic}".
    
    STRICT JSON STRUCTURE:
    Return ONLY a raw JSON array. Do not wrap in markdown or code blocks.
    [
      {
       "id": "cl_[RANDOM_ID]",
        "title": "Short Descriptive Title",
        "category": "Specific Category (e.g. Indemnity, IP, Termination)",
        "agreement": "Agreement Type (e.g. Master Services Agreement, Non-Disclosure Agreement, Joint Development Agreement, Clinical Study)",
        
        "entity": "External",  <-- FIXED: Distinguishes from "Standard" (Colgate)
        "type": "Public",      <-- FIXED: Matches the "Public Resource" filter key
        
        "date": "2026-01-02",
        "source": "AI Generated / Industry Standard",
        "text": "The full legal clause text...",
        "jurisdiction": "New York",
        "clause_number": ""
      }
    ]

    CONTENT RULES:
    1. **Variety:** Include definitions, obligations, liabilities, warranties, and termination rights relevant to "${topic}".
    2. **Tone:** Corporate, risk-averse, favorable to a large consumer products company (Colgate).
    3. **Quantity:** You MUST provide exactly 25 items.
    4. **Format:** Ensure valid JSON. Escape all quotes within the text field.
  `;

  callGemini(API_KEY, prompt).then(result => {
    setLoading(false);
    try {
      // Use your robust parser
      var clauses = tryParseGeminiJSON(result);

      if (Array.isArray(clauses) && clauses.length > 0) {
        // Send to your existing staging area for review/edit
        renderImportStaging(clauses, "AI Generation: " + topic);
        showToast(`✅ Generated ${clauses.length} clauses!`);
      } else {
        showToast("⚠️ AI generation failed format check.");
      }
    } catch (e) {
      console.error(e);
      showToast("❌ Error parsing AI response.");
    }
  }).catch(e => {
    setLoading(false);
    showToast("Error: " + e.message);
  });
}
// ==========================================
// NEW: updateTabCounts (Dynamic Button Numbers)
// ==========================================
// ==========================================
// NEW: updateTabCounts (Dynamic Button Numbers)
// ==========================================
function updateTabCounts(searchText, searchField) {
  var counts = { all: 0, standard: 0, thirdparty: 0, public: 0 };

  eliClauseLibrary.forEach(item => {
    var isMatch = true;
    if (searchText) {
      if (searchField === "all") {
        var content = ((item.title || "") + (item.text || "") + (item.category || "") + (item.agreement || "") + (item.entity || "")).toLowerCase();
        isMatch = content.includes(searchText);
      } else {
        var val = (item[searchField] || "").toLowerCase();
        isMatch = val.includes(searchText);
      }
    }

    if (isMatch) {
      if (item.type === "Public") {
        counts.public++;
      } else if (item.type === "Standard") {
        counts.standard++;
        counts.all++;
      } else {
        counts.thirdparty++;
        counts.all++;
      }
    }
  });

  var btns = document.querySelectorAll(".toggle-btn");
  btns.forEach(btn => {
    var mode = btn.dataset.mode;
    var count = counts[mode] || 0;

    // Set labels (Make sure "ALL" is renamed to "INTERNAL")
    var label = "";
    if (mode === "all") label = "📚 INTERNAL";
    if (mode === "standard") label = "🟢 COLGATE";
    if (mode === "thirdparty") label = "🟠 3RD PARTY";
    if (mode === "public") label = "🌐 EXTERNAL";

    btn.innerHTML = `${label} <span style="opacity:0.7; font-size:9px;">(${count})</span>`;
  });
}

function checkStorageUsage() {
  // 1. FORCE CLOSE OVERLAYS
  // This loop finds and removes any "Fixed" position modal blocking the screen
  var allDivs = document.querySelectorAll('div');
  for (var i = 0; i < allDivs.length; i++) {
    var d = allDivs[i];
    // Target divs with high z-index (your modals are usually 20000)
    if (d.style.position === 'fixed' && parseInt(d.style.zIndex || 0) >= 20000) {
      d.remove();
    }
  }

  // 2. PREPARE OUTPUT
  var out = document.getElementById("output");
  out.innerHTML = "";
  out.classList.remove("hidden");

  // 3. HEADER
  var header = document.createElement("div");
  header.style.cssText = `background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%); padding: 12px 15px; border-bottom: 3px solid #c9a227; margin-bottom: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);`;
  header.innerHTML = `
      <div style="font-family: 'Inter', sans-serif; font-size: 9px; font-weight: 700; color: #8fa6c7; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 2px;">SYSTEM HEALTH</div>
      <div style="font-family: 'Georgia', serif; font-size: 16px; font-weight: 400; color: #ffffff; display:flex; align-items:center; gap:8px;">
          <span>💾</span> Storage Analyzer
      </div>
  `;
  out.appendChild(header);

  // 4. ANALYZE DATA
  var items = [];
  var total = 0;

  for (var key in localStorage) {
    if (Object.prototype.hasOwnProperty.call(localStorage, key)) {
      var val = localStorage[key];
      var size = ((val.length + key.length) * 2);
      total += size;
      items.push({ key: key, size: size, preview: val.substring(0, 60) });
    }
  }
  items.sort(function (a, b) { return b.size - a.size; });

  var totalKB = (total / 1024).toFixed(2);
  var limit = 5120; // 5MB limit
  var pct = ((totalKB / limit) * 100).toFixed(1);
  var barColor = pct > 80 ? "#c62828" : (pct > 50 ? "#f57f17" : "#2e7d32");

  // 5. RENDER USAGE BAR
  var statsCard = document.createElement("div");
  statsCard.className = "tool-box slide-in";
  statsCard.style.cssText = "padding:15px; background:white; border-left:5px solid " + barColor + "; margin-bottom:20px;";

  statsCard.innerHTML = `
      <div style="display:flex; justify-content:space-between; margin-bottom:5px; font-size:11px; font-weight:700; color:#555;">
          <span>TOTAL USAGE</span>
          <span>${totalKB} KB / 5000 KB</span>
      </div>
      <div style="height:8px; width:100%; background:#eee; border-radius:4px; overflow:hidden;">
          <div style="height:100%; width:${pct}%; background:${barColor}; transition:width 0.5s;"></div>
      </div>
      <div style="text-align:right; font-size:10px; color:${barColor}; margin-top:4px; font-weight:bold;">${pct}% Full</div>
  `;
  out.appendChild(statsCard);

  // 6. RENDER ITEM LIST
  var listContainer = document.createElement("div");
  listContainer.style.padding = "0 5px";

  if (items.length === 0) {
    listContainer.innerHTML = "<div style='text-align:center; color:#999; padding:20px;'>Storage is empty.</div>";
  }

  items.forEach(function (item) {
    var kb = (item.size / 1024).toFixed(2);
    var row = document.createElement("div");
    row.style.cssText = "background:white; border:1px solid #e0e0e0; margin-bottom:8px; padding:10px; display:flex; justify-content:space-between; align-items:center; box-shadow:0 1px 3px rgba(0,0,0,0.05);";

    row.innerHTML = `
          <div style="overflow:hidden; width:65%;">
              <div style="font-weight:700; font-size:11px; color:#1565c0; margin-bottom:2px;">${item.key}</div>
              <div style="font-size:10px; color:#666; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${item.preview}...</div>
          </div>
          <div style="text-align:right;">
              <div style="font-weight:bold; font-size:11px; color:#333;">${kb} KB</div>
              <button class="btn-del-key" style="margin-top:4px; font-size:9px; background:#fff; border:1px solid #c62828; color:#c62828; cursor:pointer; padding:2px 6px; border-radius:2px;">DELETE</button>
          </div>
      `;

    // Custom Delete Dialog (No 'confirm' block)
    row.querySelector(".btn-del-key").onclick = function () {
      var overlay = document.createElement("div");
      overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:99999; display:flex; align-items:center; justify-content:center;";

      var dialog = document.createElement("div");
      dialog.className = "tool-box slide-in";
      dialog.style.cssText = "background:white; width:280px; padding:20px; border-left:5px solid #c62828; box-shadow:0 10px 25px rgba(0,0,0,0.3); border-radius:4px;";

      dialog.innerHTML = `
              <div style="font-size:24px; margin-bottom:10px;">🗑️</div>
              <div style="font-weight:700; color:#c62828; margin-bottom:10px; font-size:14px;">Delete Item?</div>
              <div style="font-size:11px; color:#555; line-height:1.4; margin-bottom:15px; word-break:break-all;">
                  Permanently remove <strong>${item.key}</strong>?
              </div>
              <div style="display:flex; gap:10px;">
                  <button id="btnYesDel" style="flex:1; background:#c62828; color:white; border:none; padding:8px; font-weight:bold; cursor:pointer;">Delete</button>
                  <button id="btnNoDel" class="btn-action" style="flex:1; background:#f5f5f5; color:#333; border:1px solid #ccc;">Cancel</button>
              </div>
          `;

      overlay.appendChild(dialog);
      document.body.appendChild(overlay);

      document.getElementById("btnNoDel").onclick = function () { overlay.remove(); };
      document.getElementById("btnYesDel").onclick = function () {
        localStorage.removeItem(item.key);
        overlay.remove();
        checkStorageUsage(); // Refresh list
        showToast("🗑️ Item Deleted");
      };
    };

    listContainer.appendChild(row);
  });

  out.appendChild(listContainer);

  // 7. BACK BUTTON (Fixed: Goes Home)
  var btnBack = document.createElement("button");
  btnBack.className = "btn-action";
  btnBack.innerText = "🏠 Home"; // Changed label
  btnBack.style.marginTop = "20px";
  btnBack.onclick = function () { showHome(); }; // Changed action
  out.appendChild(btnBack);
}
// ==========================================
// 🚀 FINAL SYSTEM INITIALIZATION
// ==========================================

// Check if Office.js is loaded
if (typeof Office !== "undefined") {
  Office.onReady(function (info) {
    console.log("⚡ Office.onReady Fired. Host:", info.host);

    // 1. Attempt Resize (Safe Wrap)
    try {
      if (Office.extensionLifeCycle && Office.extensionLifeCycle.taskpane) {
        Office.extensionLifeCycle.taskpane.setWidth(750);
      }
    } catch (e) {
      console.warn("Resize not supported on this platform.");
    }

    // 2. Run Init (With Safety Check)
    if (typeof init === "function") {
      // If DOM is already ready, run immediately. Otherwise wait.
      if (document.readyState === "complete" || document.readyState === "interactive") {
        console.log("🚀 DOM Ready. Running init()...");
        init();
      } else {
        console.log("⏳ DOM not ready. Adding listener...");
        document.addEventListener("DOMContentLoaded", init);
      }
    } else {
      console.error("❌ CRITICAL ERROR: 'init()' function is missing or undefined.");
      document.body.innerHTML = "<h3 style='color:red; padding:20px;'>Error: init() function not found. Check script.</h3>";
    }
  });
} else {
  // Fallback if Office.js script tag is missing in HTML
  console.error("❌ Office.js library not loaded.");
  if (document.readyState === "complete" || document.readyState === "interactive") {
     init(); // Try forcing init anyway for browser testing
  } else {
     document.addEventListener("DOMContentLoaded", init);
  }
}
      
