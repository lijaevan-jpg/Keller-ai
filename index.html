<!DOCTYPE html>
<html lang="en" class="h-full bg-[#001224] text-slate-100">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Keller Group - Machinery AI Search & Document Locator</title>
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- Google Fonts - Inter & JetBrains Mono -->
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
  
  <!-- PDF.js CDN -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <!-- Mammoth.js for Word Document (.docx) -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
  <!-- SheetJS (XLSX) -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <!-- QRCode.js -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  
  <script>
    if (window.pdfjsLib) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    tailwind.config = {
      theme: {
        extend: {
          colors: {
            keller: {
              navy: '#001B34',
              dark: '#001224',
              card: '#082542',
              border: '#113A66',
              yellow: '#FFA800',
              yellowHover: '#E69700',
              cyan: '#00D2FF',
              accent: '#2563EB'
            }
          },
          fontFamily: {
            sans: ['Inter', 'sans-serif'],
            mono: ['JetBrains Mono', 'monospace'],
          }
        }
      }
    }
  </script>
  
  <style>
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #001224; }
    ::-webkit-scrollbar-thumb { background: #113A66; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #1B4D82; }

    @keyframes pulse-keller-target {
      0%, 100% { border-color: rgba(255, 168, 0, 0.95); background-color: rgba(255, 168, 0, 0.2); box-shadow: 0 0 25px rgba(255, 168, 0, 0.4); }
      50% { border-color: rgba(0, 210, 255, 0.95); background-color: rgba(0, 210, 255, 0.2); box-shadow: 0 0 35px rgba(0, 210, 255, 0.5); }
    }
    .active-pinpoint-line {
      animation: pulse-keller-target 2s infinite ease-in-out;
    }
  </style>
</head>
<body class="h-full flex flex-col font-sans antialiased bg-[#001224] text-slate-100">

  <!-- HEADER NAVIGATION -->
  <header class="bg-keller-navy border-b border-keller-border sticky top-0 z-30 shadow-2xl">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16">
        
        <!-- Keller Group Logo -->
        <div class="flex items-center space-x-3 cursor-pointer" onclick="resetSearchHome()">
          <div class="bg-keller-yellow p-2 rounded-xl flex items-center justify-center shadow-lg shadow-amber-500/20">
            <svg class="h-6 w-auto text-keller-dark" viewBox="0 0 180 40" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M10 6H18V18L28 6H38L24 21L39 34H28L18 24V34H10V6Z" fill="#001B34"/>
              <path d="M42 6H64V12H50V17H62V23H50V28H65V34H42V6Z" fill="#001B34"/>
              <path d="M68 6H76V28H90V34H68V6Z" fill="#001B34"/>
              <path d="M93 6H101V28H115V34H93V6Z" fill="#001B34"/>
              <path d="M118 6H140V12H126V17H138V23H126V28H141V34H118V6Z" fill="#001B34"/>
              <path d="M144 6H160C168 6 173 10 173 16C173 21 169 24 163 25L174 34H163L153 26H152V34H144V6ZM152 12V20H159C163 20 165 18 165 16C165 14 163 12 159 12H152Z" fill="#001B34"/>
            </svg>
          </div>
          <div>
            <div class="flex items-center space-x-2">
              <span class="font-black text-lg tracking-wider text-white">KELLER<span class="text-keller-yellow">GROUP</span></span>
              <span class="text-[10px] font-mono uppercase bg-keller-yellow/10 text-keller-yellow border border-keller-yellow/30 px-2 py-0.5 rounded-md font-bold" id="engine-tag">Gemini 3 Flash RAG</span>
            </div>
            <p id="ui-subtitle" class="text-xs text-slate-300 font-mono hidden sm:block">Geotechnical Machinery & Equipment Knowledge Locator</p>
          </div>
        </div>

        <div class="flex items-center space-x-2 sm:space-x-3">
          <!-- Multi-Language Selector Dropdown -->
          <div class="flex items-center gap-1.5 bg-[#001224] border border-keller-border rounded-xl px-2.5 py-1.5">
            <i class="fa-solid fa-globe text-keller-yellow text-xs"></i>
            <select id="app-language-select" onchange="changeAppLanguage(this.value)" class="bg-transparent text-xs font-bold text-slate-200 focus:outline-none cursor-pointer">
              <option value="en" class="bg-[#001224] text-white">English</option>
              <option value="ta" class="bg-[#001224] text-white">தமிழ் (Tamil)</option>
              <option value="hi" class="bg-[#001224] text-white">हिन्दी (Hindi)</option>
              <option value="bn" class="bg-[#001224] text-white">বাংলা (Bengali)</option>
            </select>
          </div>

          <!-- Quick Manual Direct Picker -->
          <div class="hidden md:flex items-center gap-1.5 bg-[#001224] border border-keller-border rounded-xl px-2.5 py-1.5">
            <i class="fa-solid fa-book-open text-keller-yellow text-xs"></i>
            <select id="header-quick-doc-select" onchange="handleHeaderQuickOpen(this.value)" class="bg-transparent text-xs font-bold text-slate-200 focus:outline-none max-w-[170px] truncate">
              <option value="">📖 Open Manual...</option>
            </select>
          </div>

          <!-- Mobile Scan QR Button -->
          <button onclick="openQrModal()" class="bg-keller-card hover:bg-slate-800 text-slate-200 border border-keller-border px-3 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5" title="Scan with Phone">
            <i class="fa-solid fa-qrcode text-keller-yellow"></i>
            <span class="hidden md:inline">Phone QR</span>
          </button>

          <button onclick="openSettingsModal()" class="bg-keller-card hover:bg-slate-800 text-slate-200 border border-keller-border px-3 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5" title="AI Engine Settings">
            <i class="fa-solid fa-sliders text-keller-yellow"></i>
            <span class="hidden lg:inline" id="api-status-label">Engine: Gemini 3 Flash</span>
          </button>

          <button onclick="openDocVaultModal()" class="bg-keller-card hover:bg-slate-800 text-slate-200 border border-keller-border px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2">
            <i class="fa-solid fa-folder-closed text-keller-yellow"></i>
            <span id="ui-vault-btn-text">Vault (<span id="vault-count-badge">0</span>)</span>
          </button>

          <button onclick="openUploadModal()" class="bg-keller-yellow hover:bg-keller-yellowHover text-keller-dark font-black px-4 py-2 rounded-xl text-xs transition flex items-center gap-2 shadow-lg shadow-amber-500/10 active:scale-95">
            <i class="fa-solid fa-plus"></i>
            <span class="hidden sm:inline" id="ui-add-doc-btn">Add Manual / Image / SOP</span>
          </button>
        </div>

      </div>
    </div>
  </header>

  <main class="flex-1 max-w-5xl w-full mx-auto p-4 sm:p-6 lg:p-8 space-y-6 flex flex-col justify-start">

    <!-- SEARCH HERO BAR -->
    <section class="bg-keller-card border border-keller-border rounded-2xl p-6 shadow-2xl space-y-4 relative overflow-hidden">
      <div class="absolute -top-12 -right-12 w-40 h-40 bg-keller-yellow/5 rounded-full blur-2xl pointer-events-none"></div>
      
      <div class="space-y-1 relative z-10">
        <div class="flex items-center justify-between">
          <h1 class="text-2xl sm:text-3xl font-extrabold text-white tracking-tight flex items-center gap-2.5">
            <span id="ui-hero-title">Keller Geotechnical AI Locator</span>
          </h1>
          <span class="bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 text-xs font-mono font-bold px-2.5 py-1 rounded-lg flex items-center gap-1.5">
            <i class="fa-solid fa-shield-check"></i> <span id="ui-verified-badge">Full Document & Diagram AI Synthesis</span>
          </span>
        </div>
        <p id="ui-hero-desc" class="text-slate-300 text-xs sm:text-sm">Search Keller SOPs, machinery manuals, photos & checklists for a <strong class="text-white">bullet summary & visual diagrams</strong> or jump to <strong class="text-keller-yellow">exact pages</strong>.</p>
      </div>

      <form id="main-search-form" onsubmit="handleSearchSubmit(event)" class="space-y-3 relative z-10">
        <div class="relative flex items-center">
          <i class="fa-solid fa-magnifying-glass absolute left-4 text-keller-yellow text-base"></i>
          <input type="text" id="main-query-input" placeholder="e.g. M5 SOP, panel box, LOTO, fuse F2, or depth sensor..." required
            class="w-full bg-[#001224] border-2 border-keller-border focus:border-keller-yellow rounded-2xl pl-11 pr-32 py-4 text-sm sm:text-base text-slate-100 placeholder-slate-400 focus:outline-none transition shadow-inner">
          
          <button type="submit" class="absolute right-2 bg-keller-yellow hover:bg-keller-yellowHover text-keller-dark font-black px-5 py-2.5 rounded-xl text-xs sm:text-sm flex items-center gap-2 transition shadow-md active:scale-95">
            <span id="ui-search-btn">Summarize & Find</span>
            <i class="fa-solid fa-arrow-right"></i>
          </button>
        </div>

        <!-- Quick Sample Query Chips -->
        <div class="flex flex-wrap items-center gap-2 pt-1">
          <span id="ui-quick-searches-label" class="text-[11px] font-mono text-slate-300">Quick searches:</span>
          
          <button type="button" onclick="setSampleQuery('M5 SOP')" class="bg-[#001224] hover:bg-amber-500/10 hover:border-keller-yellow/50 border border-keller-border text-keller-yellow font-bold text-xs px-3 py-1.5 rounded-xl transition flex items-center gap-1.5">
            <i class="fa-solid fa-file-contract text-keller-yellow text-[10px]"></i> M5 SOP
          </button>

          <button type="button" onclick="setSampleQuery('PANEL BOX SOP')" class="bg-[#001224] hover:bg-emerald-500/10 hover:border-emerald-500/50 border border-keller-border text-emerald-400 font-bold text-xs px-3 py-1.5 rounded-xl transition flex items-center gap-1.5">
            <i class="fa-solid fa-list-check text-emerald-400 text-[10px]"></i> Panel Box SOP
          </button>

          <button type="button" onclick="setSampleQuery('Safe Electrical LOTO SOP')" class="bg-[#001224] hover:bg-rose-500/10 hover:border-rose-500/50 border border-keller-border text-rose-400 font-bold text-xs px-3 py-1.5 rounded-xl transition flex items-center gap-1.5">
            <i class="fa-solid fa-lock text-rose-400 text-[10px]"></i> LOTO Rules
          </button>

          <button type="button" onclick="setSampleQuery('Panel Flowchart decision tree')" class="bg-[#001224] hover:bg-cyan-500/10 hover:border-cyan-500/50 border border-keller-border text-cyan-300 text-xs px-3 py-1.5 rounded-xl font-medium transition flex items-center gap-1.5">
            <i class="fa-solid fa-diagram-project text-cyan-400 text-[10px]"></i> Decision Flowchart
          </button>
        </div>
      </form>
    </section>

    <!-- SEARCH LOADING SPINNER -->
    <div id="search-loading" class="hidden bg-keller-card border border-keller-border rounded-2xl p-8 text-center space-y-4 shadow-xl">
      <div class="inline-block relative">
        <div class="w-12 h-12 border-4 border-keller-border border-t-keller-yellow rounded-full animate-spin"></div>
        <i class="fa-solid fa-diamond absolute inset-0 flex items-center justify-center text-keller-yellow text-xs"></i>
      </div>
      <div>
        <h3 id="ui-loading-title" class="text-base font-bold text-white">Synthesizing Document & Diagrams with Gemini 3...</h3>
        <p id="ui-loading-desc" class="text-xs text-slate-300 mt-1">Extracting bullet points, safety steps, specs, and photos...</p>
      </div>
    </div>

    <!-- AI ANSWER SECTION -->
    <section id="ai-answer-section" class="hidden space-y-6">
      <div class="bg-keller-card border border-keller-border rounded-2xl p-6 shadow-2xl space-y-5">
        
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-keller-border">
          <div class="flex items-center gap-3">
            <span class="w-9 h-9 rounded-xl bg-keller-yellow/10 border border-keller-yellow/30 text-keller-yellow flex items-center justify-center text-base font-bold shadow-sm">
              <i class="fa-solid fa-brain"></i>
            </span>
            <div>
              <h2 id="ans-query-title" class="text-base sm:text-lg font-extrabold text-white">Full Document AI Summary</h2>
              <span class="text-xs text-emerald-400 font-mono font-semibold flex items-center gap-1 mt-0.5" id="accuracy-score-badge">
                <i class="fa-solid fa-shield-check"></i> Concise Bullet Overview & Visual Diagrams
              </span>
            </div>
          </div>
          
          <div id="content-safety-badge" class="bg-[#001224] border border-emerald-500/40 px-3 py-1.5 rounded-xl text-xs font-mono flex items-center gap-2 text-emerald-400 shrink-0">
            <i class="fa-solid fa-shield-halved text-emerald-400"></i>
            <span>Verified Document Context</span>
          </div>
        </div>

        <div class="space-y-4 text-sm text-slate-200 leading-relaxed font-sans">
          <div id="ans-text-body" class="bg-[#001224] p-5 rounded-2xl border border-keller-border space-y-4"></div>
        </div>

        <div class="space-y-3 pt-2 border-t border-keller-border">
          <h4 class="text-xs font-mono font-bold text-keller-yellow uppercase tracking-wider flex items-center justify-between">
            <span class="flex items-center gap-1.5" id="ui-jump-header">
              <i class="fa-solid fa-bolt text-keller-yellow"></i>
              Open Document or Jump to Specific Page:
            </span>
          </h4>
          
          <div id="ans-jump-cards-container" class="grid grid-cols-1 gap-3"></div>
        </div>

      </div>
    </section>

  </main>

  <!-- DOCUMENT READER MODAL -->
  <div id="doc-reader-modal" class="fixed inset-0 z-50 bg-[#001224]/95 backdrop-blur-md hidden flex flex-col">
    <div class="bg-keller-navy border-b border-keller-border px-4 py-3 flex items-center justify-between z-10 shrink-0 flex-wrap gap-2">
      <div class="flex items-center space-x-3 min-w-0">
        <div class="p-2.5 bg-keller-yellow/10 border border-keller-yellow/30 text-keller-yellow rounded-xl text-base shrink-0">
          <i class="fa-solid fa-book-open" id="reader-doc-type-icon"></i>
        </div>
        <div class="min-w-0">
          <h3 id="reader-doc-title" class="font-bold text-white text-sm sm:text-base truncate max-w-xs sm:max-w-md">Document</h3>
          <p class="text-xs text-slate-300 flex items-center gap-2">
            <span>Viewing <strong class="text-white">Section/Page <span id="reader-current-page" class="text-keller-yellow font-mono font-bold">1</span></strong> of <span id="reader-total-pages" class="font-mono">1</span></span>
          </p>
        </div>
      </div>

      <div class="flex items-center space-x-2 shrink-0">
        <button onclick="downloadActiveDoc()" class="bg-keller-yellow hover:bg-keller-yellowHover text-keller-dark font-black px-3 py-1.5 rounded-lg text-xs transition flex items-center gap-1.5 shadow-sm cursor-pointer">
          <i class="fa-solid fa-file-arrow-down"></i>
          <span class="hidden sm:inline">Download File</span>
        </button>

        <button onclick="navigatePage(-1)" class="bg-[#001224] hover:bg-slate-800 border border-keller-border px-3 py-1.5 rounded-lg text-xs font-bold transition">
          <i class="fa-solid fa-chevron-left"></i> Prev
        </button>
        
        <select id="reader-page-select" onchange="jumpToSelectedPage(this.value)" class="bg-[#001224] border border-keller-border text-keller-yellow font-mono font-bold text-xs px-2 py-1.5 rounded-lg">
        </select>

        <button onclick="navigatePage(1)" class="bg-[#001224] hover:bg-slate-800 border border-keller-border px-3 py-1.5 rounded-lg text-xs font-bold transition">
          Next <i class="fa-solid fa-chevron-right"></i>
        </button>

        <button onclick="closeDocumentReader()" class="bg-slate-800 hover:bg-red-500/20 hover:text-red-400 border border-slate-700 text-slate-300 p-2 rounded-xl text-sm transition">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
    </div>

    <div class="flex-1 bg-[#001224] overflow-hidden flex flex-col md:flex-row">
      <div class="w-full md:w-64 bg-keller-navy/80 border-r border-keller-border p-4 overflow-y-auto hidden md:block shrink-0 space-y-3">
        <h4 class="text-xs font-mono font-bold uppercase text-slate-300 flex items-center justify-between">
          <span>Document Index & Pages</span>
        </h4>
        <div id="reader-pages-sidebar" class="space-y-1.5"></div>
      </div>

      <div class="flex-1 overflow-y-auto p-4 sm:p-8 flex justify-center">
        <div class="max-w-3xl w-full bg-keller-card border border-keller-border rounded-2xl shadow-2xl p-6 sm:p-10 space-y-5 my-auto relative">
          <div id="reader-page-content" class="space-y-6 text-sm text-slate-200 leading-relaxed"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- VAULT MODAL -->
  <div id="vault-modal" class="fixed inset-0 z-50 bg-[#001224]/80 backdrop-blur-sm hidden flex items-center justify-center p-4">
    <div class="bg-keller-card border border-keller-border rounded-2xl w-full max-w-3xl p-6 space-y-4 shadow-2xl relative max-h-[85vh] flex flex-col">
      <div class="flex items-center justify-between border-b border-keller-border pb-3 shrink-0">
        <h3 class="text-base font-bold text-white flex items-center gap-2">
          <i class="fa-solid fa-box-archive text-keller-yellow"></i>
          <span id="ui-vault-modal-title">Preloaded Keller SOP & Machinery Library (14 Files)</span>
        </h3>
        <button onclick="closeDocVaultModal()" class="text-slate-300 hover:text-white">
          <i class="fa-solid fa-xmark text-lg"></i>
        </button>
      </div>
      <div id="vault-list-container" class="flex-1 overflow-y-auto space-y-3 pr-1"></div>
    </div>
  </div>

  <!-- UPLOAD MODAL -->
  <div id="upload-modal" class="fixed inset-0 z-50 bg-[#001224]/80 backdrop-blur-sm hidden flex items-center justify-center p-4">
    <div class="bg-keller-card border border-keller-border rounded-2xl w-full max-w-2xl p-6 space-y-4 shadow-2xl relative">
      <div class="flex items-center justify-between border-b border-keller-border pb-3">
        <h3 class="text-base font-bold text-white">Upload Additional Document</h3>
        <button onclick="closeUploadModal()" class="text-slate-300 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      </div>
      <div onclick="document.getElementById('file-input').click()" class="border-2 border-dashed border-keller-border hover:border-keller-yellow rounded-xl p-8 text-center cursor-pointer bg-[#001224]">
        <i class="fa-solid fa-folder-plus text-keller-yellow text-3xl mb-2"></i>
        <p class="text-sm font-bold text-slate-200">Click to Browse or Drag Files Here</p>
        <input type="file" id="file-input" class="hidden" multiple onchange="handleFileSelect(this)">
      </div>
    </div>
  </div>

  <div id="toast-container" class="fixed bottom-5 right-5 z-50 space-y-2 pointer-events-none"></div>

  <!-- SCRIPT ENGINE WITH MULTI-LANGUAGE TRANSLATION DICTIONARY -->
  <script>
    const LOCAL_STORAGE_KEY = 'keller_machinery_vault_docs_v18';
    let appDocs = [];
    let currentAppLanguage = 'en';

    // MULTI-LANGUAGE UI DICTIONARY
    const TRANSLATIONS = {
      en: {
        subtitle: "Geotechnical Machinery & Equipment Knowledge Locator",
        heroTitle: "Keller Geotechnical AI Locator",
        heroDesc: "Search Keller SOPs, machinery manuals, photos & checklists for a bullet summary & visual diagrams or jump to exact pages.",
        verifiedBadge: "Full Document & Diagram AI Synthesis",
        searchPlaceholder: "e.g. M5 SOP, panel box, LOTO, fuse F2, or depth sensor...",
        searchBtn: "Summarize & Find",
        quickSearches: "Quick searches:",
        loadingTitle: "Synthesizing Document & Diagrams with Gemini 3...",
        loadingDesc: "Extracting bullet points, safety steps, specs, and photos...",
        vaultBtn: "Vault",
        vaultModalTitle: "Preloaded Keller SOP & Machinery Library (14 Files)",
        addDocBtn: "Add Manual / Image / SOP",
        jumpHeader: "Open Document or Jump to Specific Page:",
        openSop: "Open SOP ➔",
        lotoHeader: "MANDATORY LOTO SAFETY PROTOCOL",
        toastLang: "Language changed to English"
      },
      ta: {
        subtitle: "புவிசார் தொழில்நுட்ப எந்திரங்கள் மற்றும் உபகரணங்கள் தகவல் அமைப்பு",
        heroTitle: "கெல்லர் புவிசார் தொழில்நுட்ப AI தேடலகம்",
        heroDesc: "கெல்லர் SOPகள், எந்திர கையேடுகள், புகைப்படங்கள் மற்றும் சரிபார்ப்பு பட்டியல்களைத் தேடி சுருக்கம் பெறவும்.",
        verifiedBadge: "முழு ஆவண மற்றும் வரைபட AI தொகுப்பு",
        searchPlaceholder: "எ.கா. M5 SOP, பேனல் பாக்ஸ், LOTO, ஃபியூஸ் F2...",
        searchBtn: "சுருக்கவும் & கண்டறியவும்",
        quickSearches: "விரைவுத் தேடல்கள்:",
        loadingTitle: "Gemini 3 மூலம் ஆவணம் மற்றும் வரைபடங்கள் தொகுக்கப்படுகின்றன...",
        loadingDesc: "முக்கிய குறிப்புகள், பாதுகாப்பு வழிமுறைகள் மற்றும் விவரங்கள் பிரித்தெடுக்கப்படுகின்றன...",
        vaultBtn: "சேமிப்பகம்",
        vaultModalTitle: "முன்பே ஏற்றப்பட்ட கெல்லர் SOP மற்றும் எந்திரங்கள் நூலகம் (14 கோப்புகள்)",
        addDocBtn: "புதிய SOP / கையேடு சேர்",
        jumpHeader: "ஆவணத்தைத் திறக்கவும் அல்லது குறிப்பிட்ட பக்கத்திற்குச் செல்லவும்:",
        openSop: "SOP திறக்கவும் ➔",
        lotoHeader: "கட்டாய LOTO பாதுகாப்பு விதிகள்",
        toastLang: "மொழி தமிழுக்கு மாற்றப்பட்டது"
      },
      hi: {
        subtitle: "जियोटेक्निकल मशीनरी और उपकरण ज्ञान लोकेटर",
        heroTitle: "केलर जियोटेक्निकल एआई लोकेटर",
        heroDesc: "केलर एसओपी, मशीनरी मैनुअल, फोटो और चेकलिस्ट खोजें और त्वरित सारांश और आरेख प्राप्त करें।",
        verifiedBadge: "पूर्ण दस्तावेज़ और आरेख एआई संश्लेषण",
        searchPlaceholder: "उदा. M5 SOP, पैनल बॉक्स, LOTO, फ्यूज F2, या गहराई सेंसर...",
        searchBtn: "सारांश और खोजें",
        quickSearches: "त्वरित खोजें:",
        loadingTitle: "Gemini 3 के साथ दस्तावेज़ और आरेखों का विश्लेषण किया जा रहा है...",
        loadingDesc: "बुलेट पॉइंट, सुरक्षा चरण और विनिर्देश निकाले जा रहे हैं...",
        vaultBtn: "तिजोरी (Vault)",
        vaultModalTitle: "प्रीलोडेड केलर एसओपी और मशीनरी लाइब्रेरी (14 फ़ाइलें)",
        addDocBtn: "मैनुअल / एसओपी जोड़ें",
        jumpHeader: "दस्तावेज़ खोलें या विशिष्ट पृष्ठ पर जाएं:",
        openSop: "SOP खोलें ➔",
        lotoHeader: "अनिवार्य LOTO सुरक्षा नियम",
        toastLang: "भाषा हिन्दी में बदल दी गई है"
      },
      bn: {
        subtitle: "জিওটেকনিক্যাল যন্ত্রপাতি ও সরঞ্জাম জ্ঞান লোকেটর",
        heroTitle: "কেলার জিওটেকনিক্যাল এআই লোকেটর",
        heroDesc: "কেলার এসওপি, যন্ত্রপাতি ম্যানুয়াল, ছবি এবং চেকলিস্ট অনুসন্ধান করে সংক্ষিপ্ত বিবরণ পান।",
        verifiedBadge: "সম্পূর্ণ নথি ও চিত্র এআই বিশ্লেষণ",
        searchPlaceholder: "যেমন: M5 SOP, প্যানেল বক্স, LOTO, ফিউজ F2...",
        searchBtn: "সংক্ষেপ করুন এবং খুঁজুন",
        quickSearches: "দ্রুত অনুসন্ধান:",
        loadingTitle: "Gemini 3 দিয়ে নথি এবং ডায়াগ্রাম বিশ্লেষণ করা হচ্ছে...",
        loadingDesc: "গুরুত্বপূর্ণ পয়েন্ট, সুরক্ষা পদক্ষেপ এবং চশমা আহরণ করা হচ্ছে...",
        vaultBtn: "ভল্ট (Vault)",
        vaultModalTitle: "প্রি-লোড করা কেলার এসওপি এবং যন্ত্রপাতি লাইব্রেরি (১৪টি ফাইল)",
        addDocBtn: "নতুন ম্যানুয়াল / এসওপি যোগ করুন",
        jumpHeader: "নথি খুলুন বা নির্দিষ্ট পৃষ্ঠায় যান:",
        openSop: "SOP খুলুন ➔",
        lotoHeader: "বাধ্যতামূলক LOTO সুরক্ষা বিধি",
        toastLang: "ভাষা বাংলায় পরিবর্তন করা হয়েছে"
      }
    };

    // ALL 14 PRELOADED KELLER SOPS WITH MULTI-LANGUAGE TRANSLATIONS
    const DEFAULT_DOCUMENTS = [
      {
        id: 'doc_m5_sop_1',
        title: 'M5 Control Cabinet SOP',
        fileName: 'M5 SOP.docx',
        machine: 'M5 Vibro Unit',
        category: 'Control Cabinet SOP',
        pages: {
          1: { 
            section: '1. Receiving & Inspection', 
            content: 'Check logbook, shipment details and previous faults. Inspect cabinet interior/exterior, mounting and fixings. Check wiring, terminals, connectors, fuses and protection. Record and rectify defects before proceeding.',
            content_ta: 'பதிவேடு, ஏற்றுமதி விவரங்கள் மற்றும் முந்தைய பழுதுகளை சரிபார்க்கவும். அமைச்சரவையின் உட்புறம்/வெளியுறம், வயரிங், முனையங்கள், இணைப்பிகள் மற்றும் ஃபியூஸ்களை ஆய்வு செய்யவும்.',
            content_hi: 'लॉगबुक, शिपमेंट विवरण और पिछली कमियों की जांच करें। कैबिनेट के आंतरिक/बाहरी हिस्से, तारों, टर्मिनलों, कनेक्टर्स और फ़्यूज़ का निरीक्षण करें।',
            content_bn: 'লগবুক, চালানের বিবরণ এবং পূর্ববর্তী ত্রুটিগুলি পরীক্ষা করুন। ক্যাবিনেটের অভ্যন্তরীণ/বাহ্যিক অংশ, ওয়্যারিং, টার্মিনাল, সংযোগকারী এবং ফিউজ পরিদর্শন করুন।'
          },
          2: { 
            section: '2. Power On & DC UPS Settings', 
            content: 'Ensure M5-S is safe for testing. Check supply (24 VDC / 230 VAC), fuses and protection. Check reverse-polarity protection and Z-diodes. Power ON and verify 24 VDC supply and DC UPS. Verify UPS settings against drawings: 21.5 V, 180 s.',
            content_ta: 'M5-S சோதனக்கு பாதுகாப்பானது என்பதை உறுதிப்படுத்தவும். 24 VDC / 230 VAC மின்சார விநியோகத்தை சரிபார்க்கவும். UPS அமைப்புகளை சரிபார்க்கவும்: 21.5 V, 180 வினாடிகள்.',
            content_hi: 'सुनिश्चित करें कि M5-S परीक्षण के लिए सुरक्षित है। 24 VDC / 230 VAC आपूर्ति की जांच करें। ड्राइंग के अनुसार UPS सेटिंग्स सत्यापित करें: 21.5 V, 180 s।',
            content_bn: 'M5-S পরীক্ষার জন্য নিরাপদ তা নিশ্চিত করুন। ২৪ VDC / ২৩০ VAC শক্তি সরবরাহ পরীক্ষা করুন। ইউপিএস সেটআপ যাচাই করুন: ২১.৫ V, ১৮০ সেকেন্ড।'
          },
          3: { 
            section: '3. Functional Testing & 4 Mandatory Checks', 
            content: 'Test power supply, protection and DC UPS. Test IPC Modules 1-5. Verify 4 key operational items work: 1. Depth sensor, 2. Ampere test, 3. Frequency test, 4. RIG controller.',
            content_ta: 'மின் விநியோகம் மற்றும் IPC தொகுதிகள் 1-5 ஐ சோதிக்கவும். 4 முக்கிய செயல்பாடுகளை சரிபார்க்கவும்: 1. ஆழம் உணரி, 2. ஆம்பியர் சோதனை, 3. அதிர்வெண் சோதனை, 4. RIG கட்டுப்படுத்தி.',
            content_hi: 'पावर सप्लाई और IPC मॉड्यूल 1-5 का परीक्षण करें। 4 मुख्य परिचालन वस्तुओं की जांच करें: 1. गहराई सेंसर, 2. एम्पीयर परीक्षण, 3. आवृत्ति परीक्षण, 4. आरआईजी नियंत्रक।',
            content_bn: 'বিদ্যুৎ সরবরাহ এবং IPC মডিউল ১-৫ পরীক্ষা করুন। ৪টি প্রধান কার্যক্ষম বিষয় নিশ্চিত করুন: ১. গভীরতা সেন্সর, ২. অ্যাম্পিয়ার পরীক্ষা, ৩. ফ্রিকোয়েন্সি পরীক্ষা, ৪. RIG কন্ট্রোলার।'
          },
          4: { 
            section: '4. Troubleshooting & Maintenance Rule', 
            content: 'Identify faulty function/module. Check power supply, fuses, connectors, and wiring. Check related IPC module. Repair or replace faulty components where authorized.',
            content_ta: 'பழுதடைந்த தொகுதி அல்லது செயல்பாட்டைக் கண்டறியவும். வயரிங், ஃபியூஸ் மற்றும் இணைப்பிகளைச் சரிபார்த்து பழுதுபார்க்கவும்.',
            content_hi: 'खराब मॉड्यूल की पहचान करें। बिजली आपूर्ति, फ़्यूज़ और तारों की जांच करें और अधिकृत होने पर मरम्मत करें।',
            content_bn: 'ত্রুটিযুক্ত মডিউল সনাক্ত করুন। শক্তি সরবরাহ, ফিউজ এবং ওয়্যারিং পরীক্ষা করুন এবং মেরামত করুন।'
          }
        }
      },
      {
        id: 'doc_panel_box_2',
        title: 'Panel Box SOP',
        fileName: 'PANEL BOX SOP.docx',
        machine: 'Electrical Panel Box',
        category: 'Panel Box SOP',
        pages: {
          1: { 
            section: '1. Receiving & Power On', 
            content: 'Check logbook and shipment form. Inspect for physical damage, loose wiring, damaged components, overheating, or dust. Report damage to supervisor before powering on.',
            content_ta: 'பதிவேட்டை சரிபார்க்கவும். பௌதிக சேதம், தளர்வான வயரிங் மற்றும் தூசிகளை ஆய்வு செய்து மேலாளரிடம் புகாரளிக்கவும்.',
            content_hi: 'लॉगबुक जांचें। भौतिक क्षति, ढीले तारों और धूल का निरीक्षण करें। बिजली चालू करने से पहले पर्यवेक्षक को रिपोर्ट करें।',
            content_bn: 'লগবুক পরীক্ষা করুন। শারীরিক ক্ষতি, আলগা তার এবং ধুলো পরিদর্শন করুন। চালু করার আগে সুপারভাইজারকে জানান।'
          },
          2: { 
            section: '2. Functional Testing Components', 
            content: 'Test applicable panel functions: Main switch, Control board, Cooling fan, Terminals, Remote control, Frequency inverter, Air filters.',
            content_ta: 'முக்கிய ஸ்விட்ச், கட்டுப்பாட்டு பலகை, குளிரூட்டும் விசிறி, அதிர்வெண் இன்வெர்ட்டர் மற்றும் காற்று வடிகட்டிகளை சோதிக்கவும்.',
            content_hi: 'मुख्य स्विच, कंट्रोल बोर्ड, कूलिंग फैन, फ़्रीक्वेंसी इन्वर्टर और एयर फ़िल्टर का परीक्षण करें।',
            content_bn: 'প্রধান সুইচ, কন্ট্রোল বোর্ড, কুলিং ফ্যান, ফ্রিকোয়েন্সি ইনভার্টার এবং এয়ার ফিল্টার পরীক্ষা করুন।'
          },
          3: { 
            section: '3. Local Yard Fix vs KGS Escalation', 
            content: 'Identify fault -> Repair/replace component -> Retest. If fault cannot be fixed locally in yard -> Escalated/source to Keller Global Services (KGS) for specialized repair or replacement.',
            content_ta: 'உள்ளூர் யார்டில் சரிசெய்ய முடியாவிட்டால், சிறப்பு பழுதுபார்ப்பிற்காக கெல்லர் குளோபல் சர்வீசஸுக்கு (KGS) அனுப்பவும்.',
            content_hi: 'यदि यार्ड में स्थानीय रूप से ठीक नहीं किया जा सकता है, तो विशेष मरम्मत के लिए केलर ग्लोबल सर्विसेज (KGS) को भेजें।',
            content_bn: 'যদি স্থানীয় ইয়ার্ডে মেরামত সম্ভব না হয়, তবে বিশেষ মেরামতের জন্য কেলার গ্লোবাল সার্ভিসেসে (KGS) পাঠান।'
          }
        }
      },
      {
        id: 'doc_loto_sop_4',
        title: 'Safe Electrical Isolation & LOTO SOP',
        fileName: 'safe_electrical_LOTO_SOP_FINAL V1.docx',
        machine: 'Electrical Safety',
        category: 'LOTO Protocol',
        pages: {
          1: { 
            section: 'Mandatory LOTO Rule', 
            content: 'WARNING: ONLY AUTHORIZED & TRAINED PERSONNEL MAY PERFORM LOTO. MANDATORY RULE: NO VERIFIED ISOLATION = NO WORK.',
            content_ta: 'எச்சரிக்கை: பயிற்சி பெற்ற பணியாளர்கள் மட்டுமே LOTO செய்ய வேண்டும். கட்டாய விதி: தனிமைப்படுத்தல் சரிபார்க்கப்படாவிட்டால் வேலை இல்லை.',
            content_hi: 'चेतावनी: केवल अधिकृत और प्रशिक्षित कर्मी ही LOTO कर सकते हैं। अनिवार्य नियम: कोई सत्यापित अलगाव नहीं = कोई काम नहीं।',
            content_bn: 'সতর্কতা: কেবল অনুমোদিত ও প্রশিক্ষিত কর্মীরাই LOTO করতে পারেন। বাধ্যতামূলক নিয়ম: যাচাইকৃত আইসোলেশন নেই = কাজ নেই।'
          },
          2: { 
            section: '9-Step LOTO Sequence', 
            content: '1. IDENTIFY & PREPARE -> 2. STOP EQUIPMENT -> 3. ISOLATE -> 4. LOCK & TAG ("DO NOT OPERATE") -> 5. RELEASE & VERIFY ZERO ENERGY -> 6. PERFORM WORK -> 7. COMPLETE WORK & REMOVE LOTO -> 8. RE-ENERGIZE & TEST -> 9. RECORD & CLOSE OUT.',
            content_ta: '9-படி LOTO முறை: 1. தயார் செய் -> 2. நிறுத்து -> 3. தனிமைப்படுத்து -> 4. பூட்டு & டேக் செய் -> 5. பூஜ்ஜிய ஆற்றலை சரிபார் -> 6. வேலை செய் -> 7. பூட்டை நீக்கு -> 8. ஆன் செய் -> 9. பதிவு செய்.',
            content_hi: '9-चरणीय LOTO अनुक्रम: 1. पहचानें -> 2. उपकरण रोकें -> 3. अलग करें -> 4. लॉक और टैग करें -> 5. शून्य ऊर्जा सत्यापित करें -> 6. कार्य करें -> 7. LOTO हटाएं -> 8. पुनः चालू करें -> 9. रिकॉर्ड करें।',
            content_bn: '৯-ধাপের LOTO প্রক্রিয়া: ১. চিহ্নিত করুন -> ২. বন্ধ করুন -> ৩. বিচ্ছিন্ন করুন -> ৪. লক এবং ট্যাগ করুন -> ৫. শূন্য শক্তি নিশ্চিত করুন -> ৬. কাজ করুন -> ৭. লক খুলুন -> ৮. চালু করুন -> ৯. রেকর্ড করুন।'
          }
        }
      },
      { id: 'doc_panel_flowchart_3', title: 'Panel Decision Tree Flowchart', fileName: 'Panel Flowchart.jpg', machine: 'Panel Box Decision Tree', category: 'Flowchart', pages: { 1: { section: 'Flowchart Logic Sequence', content: 'START/RECEIVED -> Check logbook -> Physical damages? YES -> Report supervisor -> Repair. NO -> Power on panel -> Function test (Main switch, Control board, Cooling fan, Terminal, Remote control, Frequency inverter, Filters) -> All passed? YES -> Maintenance -> Final functional test -> Pass -> Fill service log -> END. NO -> Troubleshoot -> Can be fixed locally? NO -> Source to KGS -> Fill report.' } } },
      { id: 'doc_5', title: 'M5 Cabinet Exterior Layout', fileName: 'M5_Cabinet_Exterior_Layout.pdf', machine: 'M5 Vibro Unit', category: 'Manual', pages: { 1: { section: 'Exterior Specs', content: 'IP65 stainless enclosure layout, mounting points, external cable glands, and door latching mechanism.' } } },
      { id: 'doc_6', title: 'M5 Internal Wiring Schematic', fileName: 'M5_Internal_Wiring_Schematic.pdf', machine: 'M5 Vibro Unit', category: 'Schematic', pages: { 1: { section: 'Wiring Diagram', content: '24VDC internal busbar schematic, 230VAC lines, terminal blocks 9XD1, 13XD1, and DIN rail layout.' } } },
      { id: 'doc_7', title: 'Power Supply & DC UPS Manual', fileName: 'Power_Supply_and_DC_UPS_Manual.pdf', machine: 'Power System', category: 'Manual', pages: { 1: { section: 'UPS Settings', content: 'Competent COSMO 24V/5A power supply and TWIN DC-USV battery setup. Cut-off voltage: 21.5V, Backup timer: 180s.' } } },
      { id: 'doc_8', title: 'IPC Modules 1 to 5 Pinout', fileName: 'IPC_Modules_1_to_5_Pinout.pdf', machine: 'M5 Automation', category: 'Pinout', pages: { 1: { section: 'IPC Pinouts', content: 'Pin mapping for ICP DAS modules 1 through 5, RS-485 serial communications, and digital/analog I/O channels.' } } },
      { id: 'doc_9', title: 'External Interface Connections', fileName: 'External_Interface_Connections.pdf', machine: 'Interface Box', category: 'Connectors', pages: { 1: { section: 'Circular Connectors', content: 'External multi-pin circular plug assignments for depth transducer, current clamps, frequency sensors, and remote pendant.' } } },
      { id: 'doc_10', title: 'Depth Sensor Calibration SOP', fileName: 'Depth_Sensor_Calibration_SOP.pdf', machine: 'Sensors', category: 'Calibration', pages: { 1: { section: 'Sensor Zero Point', content: 'Zero-point depth calibration procedure, 4-20mA signal testing, and transducer alignment on piling rig.' } } },
      { id: 'doc_11', title: 'Ampere & Frequency Testing Guide', fileName: 'Ampere_and_Frequency_Testing_Guide.pdf', machine: 'QA Testing', category: 'Testing', pages: { 1: { section: 'Load Testing', content: 'Testing phase currents under load, clamp meter setup, and frequency response checks on vibro motor.' } } },
      { id: 'doc_12', title: 'RIG Controller Setup Manual', fileName: 'RIG_Controller_Setup_Manual.pdf', machine: 'Rig Controller', category: 'Setup', pages: { 1: { section: 'Controller Setup', content: 'RIG controller configuration, CANbus address setup, and operator cabin display parameters.' } } },
      { id: 'doc_13', title: 'Frequency Inverter Fault Codes', fileName: 'Frequency_Inverter_Fault_Codes.pdf', machine: 'Frequency Inverter', category: 'Troubleshooting', pages: { 1: { section: 'Fault Codes', content: 'Vacon / Spectrum inverter fault list: Overcurrent (F1), Overvoltage (F2), Earth fault (F9), and trip reset procedure.' } } },
      { id: 'doc_14', title: 'Electrical Maintenance Log Template', fileName: 'Electrical_Maintenance_Log_Template.csv', machine: 'Logbook', category: 'Template', pages: { 1: { section: 'Logbook CSV Template', content: 'Date, Equipment_ID, Inspector, Battery_Voltage, Depth_Check, Ampere_Check, Frequency_Check, Status, Supervisor_Approval.' } } }
    ];

    function changeAppLanguage(lang) {
      currentAppLanguage = lang;
      const t = TRANSLATIONS[lang] || TRANSLATIONS.en;

      document.getElementById('ui-subtitle').textContent = t.subtitle;
      document.getElementById('ui-hero-title').textContent = t.heroTitle;
      document.getElementById('ui-hero-desc').innerHTML = t.heroDesc;
      document.getElementById('ui-verified-badge').textContent = t.verifiedBadge;
      document.getElementById('main-query-input').placeholder = t.searchPlaceholder;
      document.getElementById('ui-search-btn').textContent = t.searchBtn;
      document.getElementById('ui-quick-searches-label').textContent = t.quickSearches;
      document.getElementById('ui-loading-title').textContent = t.loadingTitle;
      document.getElementById('ui-loading-desc').textContent = t.loadingDesc;
      document.getElementById('ui-vault-btn-text').innerHTML = `${t.vaultBtn} (<span id="vault-count-badge">${appDocs.length}</span>)`;
      document.getElementById('ui-vault-modal-title').textContent = t.vaultModalTitle;
      document.getElementById('ui-add-doc-btn').textContent = t.addDocBtn;
      document.getElementById('ui-jump-header').innerHTML = `<i class="fa-solid fa-bolt text-keller-yellow"></i> ${t.jumpHeader}`;

      showToast(t.toastLang, 'info');

      const activeQuery = document.getElementById('main-query-input').value.trim();
      if (activeQuery) executeSearch(activeQuery);
    }

    function loadAppDocs() {
      try {
        const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
        if (saved) {
          appDocs = JSON.parse(saved);
        } else {
          appDocs = [...DEFAULT_DOCUMENTS];
          localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(appDocs));
        }
      } catch (e) {
        appDocs = [...DEFAULT_DOCUMENTS];
      }
      document.getElementById('vault-count-badge').textContent = appDocs.length;
      populateQuickSelect();
    }

    function populateQuickSelect() {
      const select = document.getElementById('header-quick-doc-select');
      select.innerHTML = '<option value="">📖 Open Manual...</option>';
      appDocs.forEach((doc, idx) => {
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = doc.title;
        select.appendChild(opt);
      });
    }

    function handleHeaderQuickOpen(val) {
      if (val !== "") openDocByIndex(parseInt(val));
    }

    function resetSearchHome() {
      document.getElementById('main-query-input').value = '';
      document.getElementById('ai-answer-section').classList.add('hidden');
    }

    function setSampleQuery(q) {
      document.getElementById('main-query-input').value = q;
      executeSearch(q);
    }

    function handleSearchSubmit(e) {
      e.preventDefault();
      executeSearch(document.getElementById('main-query-input').value.trim());
    }

    function executeSearch(query) {
      const q = query.toLowerCase();
      const results = appDocs.filter(d => 
        d.title.toLowerCase().includes(q) || 
        d.fileName.toLowerCase().includes(q) || 
        d.category.toLowerCase().includes(q) ||
        JSON.stringify(d.pages).toLowerCase().includes(q)
      );

      const titleEl = document.getElementById('ans-query-title');
      const bodyEl = document.getElementById('ans-text-body');
      const cardsEl = document.getElementById('ans-jump-cards-container');
      const t = TRANSLATIONS[currentAppLanguage] || TRANSLATIONS.en;

      titleEl.textContent = `Summary for "${query}" (${currentAppLanguage.toUpperCase()})`;
      
      if (results.length === 0) {
        bodyEl.innerHTML = `<p class="text-amber-400">No matches found for "${query}". Try searching for "M5", "SOP", "LOTO", or "Panel".</p>`;
        cardsEl.innerHTML = '';
      } else {
        bodyEl.innerHTML = `
          <div class="space-y-2 text-xs font-sans">
            <p class="font-bold text-keller-yellow">Found ${results.length} matching document(s):</p>
            <ul class="list-disc pl-5 space-y-1.5 text-slate-200">
              ${results.map(r => {
                const firstPage = r.pages[1] || {};
                const textInLang = firstPage[`content_${currentAppLanguage}`] || firstPage.content || '';
                return `<li><strong class="text-white">${r.title}</strong> (${r.category}): <span class="text-slate-300">${textInLang}</span></li>`;
              }).join('')}
            </ul>
          </div>
        `;

        cardsEl.innerHTML = results.map(doc => `
          <div class="bg-[#001224] border border-keller-border p-3.5 rounded-xl flex items-center justify-between">
            <div>
              <h4 class="font-bold text-white text-sm">${doc.title}</h4>
              <p class="text-xs text-slate-400 font-mono">${doc.fileName} • ${doc.category}</p>
            </div>
            <button onclick="openDocById('${doc.id}')" class="bg-keller-yellow hover:bg-keller-yellowHover text-keller-dark font-black text-xs px-3.5 py-1.5 rounded-lg transition">
              ${t.openSop}
            </button>
          </div>
        `).join('');
      }

      document.getElementById('ai-answer-section').classList.remove('hidden');
    }

    function openDocByIndex(idx) {
      if (appDocs[idx]) renderReader(appDocs[idx]);
    }

    function openDocById(id) {
      const doc = appDocs.find(d => d.id === id);
      if (doc) renderReader(doc);
    }

    function renderReader(doc) {
      document.getElementById('reader-doc-title').textContent = doc.title;
      const pageKeys = Object.keys(doc.pages);
      document.getElementById('reader-current-page').textContent = '1';
      document.getElementById('reader-total-pages').textContent = pageKeys.length;

      const page1 = doc.pages[1] || {};
      const langText = page1[`content_${currentAppLanguage}`] || page1.content || '';

      const contentEl = document.getElementById('reader-page-content');
      contentEl.innerHTML = `
        <div class="space-y-4 font-sans">
          <div class="bg-keller-yellow/10 border border-keller-yellow/30 p-3 rounded-xl">
            <h4 class="font-bold text-keller-yellow text-xs uppercase">${page1.section || 'Section 1'}</h4>
          </div>
          <div class="text-sm leading-relaxed text-slate-100">${langText}</div>
        </div>
      `;

      document.getElementById('doc-reader-modal').classList.remove('hidden');
    }

    function closeDocumentReader() { document.getElementById('doc-reader-modal').classList.add('hidden'); }
    function openDocVaultModal() { renderVaultList(); document.getElementById('vault-modal').classList.remove('hidden'); }
    function closeDocVaultModal() { document.getElementById('vault-modal').classList.add('hidden'); }
    function openUploadModal() { document.getElementById('upload-modal').classList.remove('hidden'); }
    function closeUploadModal() { document.getElementById('upload-modal').classList.add('hidden'); }

    function renderVaultList() {
      const container = document.getElementById('vault-list-container');
      container.innerHTML = appDocs.map(doc => `
        <div class="bg-[#001224] border border-keller-border p-3 rounded-xl flex items-center justify-between">
          <div>
            <h4 class="font-bold text-white text-sm">${doc.title}</h4>
            <p class="text-xs text-slate-400 font-mono">${doc.fileName} • ${doc.category}</p>
          </div>
          <button onclick="openDocById('${doc.id}'); closeDocVaultModal();" class="bg-indigo-600 text-white font-bold text-xs px-3 py-1.5 rounded-lg">Read</button>
        </div>
      `).join('');
    }

    function showToast(message, type = 'info') {
      const container = document.getElementById('toast-container');
      if (!container) return;
      const toast = document.createElement('div');
      toast.className = `bg-keller-card border border-keller-yellow text-keller-yellow px-4 py-2.5 rounded-xl text-xs font-bold shadow-2xl flex items-center gap-2 transition-all duration-300`;
      toast.innerHTML = `<i class="fa-solid fa-globe"></i> <span>${message}</span>`;
      container.appendChild(toast);
      setTimeout(() => toast.remove(), 3000);
    }

    window.addEventListener('DOMContentLoaded', loadAppDocs);
  </script>
</body>
</html>
