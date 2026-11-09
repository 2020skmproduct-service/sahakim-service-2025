import { ServiceTicket, Technician, TroubleshootingTemplate } from './types';
import { EMBEDDED_SAHAKIM_LOGO, EMBEDDED_SKM_MASCOT } from './assets/embeddedImages';

export const GOOGLE_APPS_SCRIPT_TEMPLATE = `// ============================================================================
// SAHA KIM MOTOR V12.1 - GOOGLE SHEETS + APPSHEET + LINE OA NOTIFICATION (Code.gs)
// วิธีใช้งาน:
// 1. เปิด Google Sheet -> ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
// 2. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดนี้ลงไป กดบันทึก (ไอคอนแผ่นดิสก์)
// 3. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) > "จัดการการทำให้ใช้งานได้" (Manage deployments)
//    หรือ "การทำให้ใช้งานได้รายการใหม่" (New deployment) -> เลือกประเภท "เว็บแอป" (Web app)
//    - ดำเนินการในฐานะ (Execute as): "ฉัน" (Me)
//    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone)
// 4. คัดลอก Web App URL (ลงท้ายด้วย /exec) มาวางในหน้าตั้งค่าของ Supervisor
// ============================================================================

const SHEET_TICKETS = 'Tickets';

function ensureHeaderRow(sheet) {
  const headers = [
    'TicketID',
    'Company',
    'TaxID',
    'DocNo',
    'Contact',
    'SiteContact',
    'Location',
    'Warranty',
    'ServiceType',
    'Urgency',
    'DesiredDate',
    'ConfirmedDate',
    'Status',
    'AssignedTech',
    'ApprovedBySupervisor',
    'EquipmentSummary',
    'CreatedAt',
    'RawJson'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  } else {
    const firstCell = String(sheet.getRange(1, 1).getValue() || '').trim();
    if (firstCell !== 'TicketID') {
      sheet.getRange(1, 1, 1, 18).setValues([headers]);
      sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    }
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_TICKETS);
  if (!sheet) {
    // หากแผ่นงานแรกสุดยังว่างเปล่า ให้เปลี่ยนชื่อเป็น Tickets ทันทีเพื่อให้ผู้ใช้เห็นข้อมูลในหน้าแรกเลย
    const firstSheet = ss.getSheets()[0];
    if (firstSheet && firstSheet.getLastRow() === 0) {
      firstSheet.setName(SHEET_TICKETS);
      sheet = firstSheet;
    } else {
      sheet = ss.insertSheet(SHEET_TICKETS, 0);
    }
  }
  ensureHeaderRow(sheet);
  try {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  } catch (e) {}
  return sheet;
}

// ฟังก์ชันลดขนาดรูปภาพ Base64 ใน RawJson เพื่อไม่ให้เกินลิมิต 50,000 ตัวอักษรต่อเซลล์ของ Google Sheets
function sanitizeTicketForSheet(t) {
  try {
    const copy = JSON.parse(JSON.stringify(t));
    if (Array.isArray(copy.equipmentList)) {
      copy.equipmentList = copy.equipmentList.map(function(eq) {
        if (Array.isArray(eq.beforePhotos)) {
          eq.beforePhotos = eq.beforePhotos.map(function(p) {
            return (typeof p === 'string' && (p.indexOf('data:') === 0 || p.indexOf('blob:') === 0 || p.length > 800))
              ? 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=600'
              : p;
          });
        }
        if (Array.isArray(eq.afterPhotos)) {
          eq.afterPhotos = eq.afterPhotos.map(function(p) {
            return (typeof p === 'string' && (p.indexOf('data:') === 0 || p.indexOf('blob:') === 0 || p.length > 800))
              ? 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=600'
              : p;
          });
        }
        return eq;
      });
    }
    return copy;
  } catch (e) {
    return t;
  }
}

function doGet(e) {
  const sheet = setupSheets();
  const rows = sheet.getDataRange().getValues();
  const tickets = [];
  for (let i = 1; i < rows.length; i++) {
    const rawJson = rows[i][17];
    if (rawJson) {
      try {
        tickets.push(JSON.parse(rawJson));
      } catch (err) {}
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      tickets: tickets
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    let writtenCount = 0;
    if (Array.isArray(payload.tickets)) {
      const sheet = setupSheets();
      const tickets = payload.tickets;

      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).clearContent();
      }

      const rowsToWrite = tickets.map(function(t) {
        const cleanT = sanitizeTicketForSheet(t);
        const eqSummary = (cleanT.equipmentList || []).map(function(eq) {
          return eq.type + ' (' + eq.brand + ' ' + (eq.model || '') + '): ' + eq.symptom;
        }).join(' | ');

        const techDisplay = cleanT.leadTech
          ? (cleanT.assignedTech || '') + ' (หน.ช่าง: ' + cleanT.leadTech + ')'
          : (cleanT.assignedTech || '');

        let rawJsonStr = JSON.stringify(cleanT);
        if (rawJsonStr.length > 48000) {
          rawJsonStr = rawJsonStr.substring(0, 48000);
        }

        return [
          String(cleanT.id || '').substring(0, 500),
          String(cleanT.company || '').substring(0, 1000),
          String(cleanT.taxId || '').substring(0, 500),
          String(cleanT.docNo || '').substring(0, 500),
          String(cleanT.contact || '').substring(0, 1000),
          String(cleanT.siteContact || '').substring(0, 1000),
          String(cleanT.location || '').substring(0, 2000),
          String(cleanT.warranty || '').substring(0, 500),
          String(cleanT.serviceType || '').substring(0, 500),
          String(cleanT.urgency || '').substring(0, 500),
          String(cleanT.desiredDate || '').substring(0, 500),
          String(cleanT.confirmedDate || '').substring(0, 500),
          String(cleanT.status || '').substring(0, 500),
          String(techDisplay).substring(0, 1000),
          String(cleanT.approvedBySupervisor || '').substring(0, 1000),
          String(eqSummary).substring(0, 5000),
          String(cleanT.createdAt || '').substring(0, 500),
          rawJsonStr
        ];
      });

      if (rowsToWrite.length > 0) {
        sheet.getRange(2, 1, rowsToWrite.length, 18).setValues(rowsToWrite);
        writtenCount = rowsToWrite.length;
      }
      SpreadsheetApp.flush();
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', updatedRows: writtenCount, updatedAt: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

export function generateStandaloneHtmlString(
  currentTickets: ServiceTicket[],
  currentTechnicians: Technician[],
  currentTemplates: TroubleshootingTemplate[],
  customEquipTypes: string[],
  customBrands: string[],
  supPassword: string,
  servicePin = '1234',
  gasWebAppUrl = '',
  lineChannelToken = '',
  lineTargetId = '',
  currentSupervisors: string[] = [
    'หัวหน้าวิศวกรบริการ (Supervisor)',
    'ผู้จัดการฝ่ายบริการ (Service Manager)'
  ]
): string {
  // Sanitize local Vite asset URLs so images render properly on GitHub Pages
  const sanitizedTickets = currentTickets.map((t) => ({
    ...t,
    equipmentList: t.equipmentList.map((eq) => ({
      ...eq,
      beforePhotos: (eq.beforePhotos || []).map((url) =>
        url.startsWith('http') || url.startsWith('data:')
          ? url
          : 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=600'
      ),
      afterPhotos: (eq.afterPhotos || []).map((url) =>
        url.startsWith('http') || url.startsWith('data:')
          ? url
          : 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=600'
      )
    }))
  }));

  const htmlContent = `<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Saha Kim Motor - After-Sales & Service Management System V12.1</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Sarabun', sans-serif; }
        body.dark-theme { background-color: #0b1120 !important; color: #f1f5f9 !important; }
        body.dark-theme .bg-white:not(.print-container):not(.print-container *) { background-color: #111827 !important; color: #f1f5f9 !important; }
        body.dark-theme .bg-slate-50:not(.print-container):not(.print-container *) { background-color: #1e293b !important; color: #e2e8f0 !important; }
        body.dark-theme .bg-slate-100:not(.print-container):not(.print-container *) { background-color: #334155 !important; color: #f1f5f9 !important; }
        body.dark-theme .text-slate-900:not(.print-container):not(.print-container *),
        body.dark-theme .text-slate-800:not(.print-container):not(.print-container *) { color: #f8fafc !important; }
        body.dark-theme .text-slate-700:not(.print-container):not(.print-container *),
        body.dark-theme .text-slate-600:not(.print-container):not(.print-container *) { color: #cbd5e1 !important; }
        body.dark-theme .border-slate-200:not(.print-container):not(.print-container *),
        body.dark-theme .border-slate-300:not(.print-container):not(.print-container *) { border-color: #334155 !important; }
        body.dark-theme input:not([type="checkbox"]):not([type="radio"]):not(.print-container *),
        body.dark-theme select:not(.print-container *),
        body.dark-theme textarea:not(.print-container *) { background-color: #0f172a !important; color: #f8fafc !important; border-color: #475569 !important; }
        @media print {
            * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
            .no-print { display: none !important; }
            body, body.dark-theme { background: white !important; color: #0f172a !important; overflow: visible !important; height: auto !important; }
            body:has(.print-modal-root) main,
            body:has(.print-modal-root) header,
            body:has(.print-modal-root) footer { display: none !important; }
            .print-modal-root { position: static !important; inset: auto !important; background: white !important; padding: 0 !important; margin: 0 !important; overflow: visible !important; height: auto !important; max-height: none !important; display: block !important; }
            .print-modal-content { position: static !important; max-height: none !important; height: auto !important; overflow: visible !important; box-shadow: none !important; border: none !important; padding: 0 !important; margin: 0 !important; max-width: 100% !important; width: 100% !important; }
            .print-container { margin-top: 0 !important; box-shadow: none !important; border: none !important; max-width: 100% !important; width: 100% !important; }
            @page { size: A4; margin: 10mm; }
        }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen flex flex-col">

    <!-- Top Header -->
    <header id="top-app-header" class="bg-slate-900 text-white shadow-md sticky top-0 z-50 no-print transition-colors">
        <div class="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center gap-3">
            <div class="flex items-center space-x-3 cursor-pointer" onclick="app.goHome()">
                <img src="${EMBEDDED_SAHAKIM_LOGO}" alt="บริษัท สหคิมมอเตอร์ จำกัด" class="w-10 h-10 rounded-lg object-contain bg-white p-0.5 border border-slate-700 shadow shrink-0" />
                <div>
                    <h1 class="font-bold text-base sm:text-lg leading-tight tracking-wide">บริษัท สหคิมมอเตอร์ จำกัด</h1>
                    <p class="text-[10px] sm:text-xs text-orange-400 font-medium">SAHA KIM MOTOR CO., LTD. | Service Portal V12.1</p>
                </div>
            </div>
            <div class="flex items-center gap-2">
                <div class="flex items-center bg-slate-800 p-0.5 rounded-lg border border-slate-700 text-xs">
                    <button id="btn-theme-light" type="button" onclick="setStandaloneTheme('light')" class="px-2 py-1 rounded font-bold bg-white text-slate-900 cursor-pointer">☀️ ขาว</button>
                    <button id="btn-theme-dark" type="button" onclick="setStandaloneTheme('dark')" class="px-2 py-1 rounded font-bold text-slate-300 cursor-pointer">🌙 ดำ</button>
                </div>
                <button onclick="openUserManualModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5">
                    <i class="fa-solid fa-book-open"></i>
                    <span class="hidden sm:inline">คู่มือการใช้งาน</span>
                </button>
                <button onclick="openQrShareModal()" class="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5">
                    <i class="fa-solid fa-qrcode"></i>
                    <span class="hidden sm:inline">QR Code แจ้งซ่อม</span>
                </button>
                <button type="button" onclick="resetToFreshProduction()" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 shadow-xs" title="รีเซ็ตล้างใบงานตัวอย่าง เพื่อเริ่มกรอกใบงานจริงทันที (ตารางช่างจะว่างเขียวทั้งหมด)">
                    <i class="fa-solid fa-rotate-left"></i>
                    <span class="hidden sm:inline">รีเซ็ตเริ่มงานจริง</span>
                </button>
                <div id="header-nav" class="flex items-center space-x-2"></div>
            </div>
        </div>
    </header>

    <!-- Main Container: Supervisor Online Sync Bar (เฉพาะสิทธิ์ Supervisor เท่านั้น) -->
    <div id="supervisor-sync-bar-container" class="max-w-7xl w-full mx-auto px-3 sm:px-6 pt-4 no-print hidden">
        <div class="bg-white border-2 border-emerald-500/80 rounded-2xl p-3 sm:p-4 shadow-sm flex flex-col lg:flex-row items-start lg:items-center justify-between gap-3">
            <div class="flex flex-wrap items-center gap-2 text-xs">
                <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-900 font-extrabold border border-emerald-300">
                    <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse inline-block"></span>
                    <span>ONLINE REAL-TIME SYNC</span>
                </span>
                <span id="online-sync-summary" class="font-bold text-slate-900">ระบบซิงค์ข้อมูลออนไลน์อัตโนมัติ</span>
                <span id="online-sync-time" class="text-slate-500 font-mono text-[11px]">(พร้อมซิงค์ข้อมูลออนไลน์)</span>
            </div>
            <div class="flex flex-wrap items-center gap-2 text-xs w-full lg:w-auto">
                <button type="button" onclick="syncOnlineNow(true)" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-sm">
                    <i class="fa-solid fa-rotate"></i>
                    <span>🔄 กดอัปเดตข้อมูลออนไลน์ทันที (Sync ทุกเครื่อง)</span>
                </button>
                <button type="button" onclick="resetToFreshProduction()" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-sm" title="ล้างใบงานตัวอย่าง ตารางช่างจะว่างเขียว 🟢 ทั้งหมด">
                    <i class="fa-solid fa-rotate-left"></i>
                    <span>🔄 ล้างงานตัวอย่าง (ตารางว่างเขียว 🟢)</span>
                </button>
                <button type="button" onclick="openGoogleSheetModal()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer">
                    <i class="fa-solid fa-table"></i>
                    <span>☁️ ซิงค์ Google Sheet</span>
                </button>
                <button type="button" onclick="openMasterDataModal('brands')" class="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer">
                    <i class="fa-solid fa-sliders"></i>
                    <span>⚙️ แก้ไขข้อมูลตั้งต้น</span>
                </button>
                <button type="button" onclick="exportFullSystemDataStandalone()" class="bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 px-2.5 py-1.5 rounded-xl font-bold cursor-pointer">
                    📤 สำรองไฟล์ข้อมูล
                </button>
                <label class="bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 px-2.5 py-1.5 rounded-xl font-bold cursor-pointer inline-flex items-center">
                    <span>📥 นำเข้าไฟล์ข้อมูล</span>
                    <input type="file" accept=".json,application/json" onchange="importFullSystemDataStandalone(event)" class="hidden" />
                </label>
            </div>
        </div>
    </div>
    <main id="main-content" class="flex-grow max-w-7xl w-full mx-auto p-3 sm:p-6"></main>

    <!-- Footer -->
    <footer class="bg-slate-900 text-slate-400 text-[10px] sm:text-xs py-4 text-center border-t border-slate-800 mt-auto no-print">
        <div class="max-w-7xl mx-auto px-4">
            <p>COPYRIGHT © 2026 บริษัท สหคิมมอเตอร์ จำกัด (SAHA KIM MOTOR CO., LTD.) ALL RIGHTS RESERVED. | ENTERPRISE SERVICE SYSTEM V12.1</p>
        </div>
    </footer>

    <script>
        window.addEventListener('error', function(e) {
            console.warn('Sahakim Portal Client Error Guard:', e && e.message);
        });
        window.addEventListener('unhandledrejection', function(e) {
            console.warn('Sahakim Portal Promise Error Guard:', e && e.reason);
        });

        const BASE_EQUIPMENT_MASTERS = [
            "MOTOR", "END SUCTION PUMP", "VERTICAL PUMP", "BOOSTER PUMP", "TRANSFER PUMP",
            "BLOWER", "FIRE PUMP", "GEN SET", "SEWAGE", "EJECTOR", "AIR PUMP", "AERATOR", "CONTROL"
        ];
        const BASE_BRANDS = [
            "BULL POWER", "BROOK CROMPTON", "STAC", "UNOMACH", "VARISCO", "FYBROC",
            "EBARA", "GRUNDFOS", "SIEMENS", "ABB"
        ];

        const GAS_CODE_TEMPLATE = ${JSON.stringify(GOOGLE_APPS_SCRIPT_TEMPLATE)};
        let gasWebAppUrl = localStorage.getItem('sahakim_gas_url') || ${JSON.stringify(gasWebAppUrl)};
        try {
            localStorage.removeItem('sahakim_line_token');
            localStorage.removeItem('sahakim_line_target');
        } catch (e) {}
        let lineChannelToken = '';
        let lineTargetId = '';

        function safeJsonParse(val, fallback) {
            if (!val || typeof val !== 'string') return fallback;
            try { return JSON.parse(val); } catch (e) { return fallback; }
        }

        // ============================================================================
        // INDEXEDDB RESILIENT STORAGE ENGINE (Zero Quota Crash / Full Mobile Stability)
        // ============================================================================
        const IDB_NAME = 'SahakimPortalStorage_v1';
        const IDB_STORE = 'app_state_store';
        let idbPromise = null;

        function getIdb() {
            if (!idbPromise) {
                idbPromise = new Promise((resolve) => {
                    if (!window.indexedDB) { resolve(null); return; }
                    try {
                        const req = indexedDB.open(IDB_NAME, 1);
                        req.onupgradeneeded = (e) => {
                            const db = e.target.result;
                            if (!db.objectStoreNames.contains(IDB_STORE)) {
                                db.createObjectStore(IDB_STORE);
                            }
                        };
                        req.onsuccess = (e) => resolve(e.target.result);
                        req.onerror = () => resolve(null);
                    } catch (err) {
                        resolve(null);
                    }
                });
            }
            return idbPromise;
        }

        async function saveToIndexedDB(key, val) {
            try {
                const db = await getIdb();
                if (!db) return false;
                return new Promise((resolve) => {
                    const tx = db.transaction(IDB_STORE, 'readwrite');
                    const store = tx.objectStore(IDB_STORE);
                    const req = store.put(val, key);
                    req.onsuccess = () => resolve(true);
                    req.onerror = () => resolve(false);
                });
            } catch (e) {
                return false;
            }
        }

        async function loadFromIndexedDB(key) {
            try {
                const db = await getIdb();
                if (!db) return null;
                return new Promise((resolve) => {
                    const tx = db.transaction(IDB_STORE, 'readonly');
                    const store = tx.objectStore(IDB_STORE);
                    const req = store.get(key);
                    req.onsuccess = () => resolve(req.result || null);
                    req.onerror = () => resolve(null);
                });
            } catch (e) {
                return null;
            }
        }

        // ============================================================================
        // MOBILE-SAFE CUSTOM MODAL DIALOGS (Zero Popup Blocker / No Window.Prompt Hang)
        // ============================================================================
        function showAppToast(msg, type = 'info') {
            try {
                let container = document.getElementById('app-toast-container');
                if (!container) {
                    container = document.createElement('div');
                    container.id = 'app-toast-container';
                    container.className = 'fixed bottom-4 right-4 z-[99999] flex flex-col gap-2 max-w-sm pointer-events-none no-print';
                    document.body.appendChild(container);
                }
                const toast = document.createElement('div');
                const bgClass = type === 'success' ? 'bg-emerald-600 text-white' : type === 'error' ? 'bg-rose-600 text-white' : 'bg-slate-900 text-white';
                toast.className = bgClass + ' px-4 py-3 rounded-xl shadow-2xl text-xs font-bold pointer-events-auto flex items-center gap-2 border border-white/20 transition-all duration-300 transform translate-y-3 opacity-0';
                toast.innerHTML = '<span>' + msg + '</span>';
                container.appendChild(toast);
                requestAnimationFrame(() => {
                    toast.classList.remove('translate-y-3', 'opacity-0');
                });
                setTimeout(() => {
                    toast.classList.add('opacity-0', 'translate-y-2');
                    setTimeout(() => toast.remove(), 300);
                }, 3500);
            } catch(e) {}
        }

        function openCustomPrompt(options) {
            return new Promise((resolve) => {
                const title = options.title || 'กรุณาระบุข้อมูล';
                const message = options.message || '';
                const defaultValue = options.defaultValue || '';
                const inputType = options.inputType || 'text';
                const placeholder = options.placeholder || '';
                const confirmText = options.confirmText || 'ตกลง';
                const cancelText = options.cancelText || 'ยกเลิก';

                const modal = document.createElement('div');
                modal.className = 'fixed inset-0 bg-slate-900/70 z-[9999] flex items-center justify-center p-3 animate-fade-in no-print';
                modal.innerHTML = \`
                    <div class="bg-white w-full max-w-md rounded-2xl shadow-2xl p-5 space-y-4 text-xs sm:text-sm border border-slate-200">
                        <div class="space-y-1">
                            <h3 class="font-bold text-base text-slate-900">\${title}</h3>
                            \${message ? \`<p class="text-slate-600 whitespace-pre-line text-xs">\${message}</p>\` : ''}
                        </div>
                        <div>
                            <input type="\${inputType}" id="custom-prompt-input" value="\${defaultValue}" placeholder="\${placeholder}" class="w-full border-2 border-slate-300 focus:border-orange-500 rounded-xl p-2.5 font-medium outline-none text-slate-900 text-sm bg-white">
                        </div>
                        <div class="flex justify-end gap-2 pt-2 border-t">
                            <button type="button" id="custom-prompt-cancel" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer transition">\${cancelText}</button>
                            <button type="button" id="custom-prompt-ok" class="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer transition">\${confirmText}</button>
                        </div>
                    </div>
                \`;
                document.body.appendChild(modal);
                const input = modal.querySelector('#custom-prompt-input');
                if (input) {
                    input.focus();
                    if (defaultValue) input.select();
                }

                const cleanup = (val) => {
                    modal.remove();
                    resolve(val);
                };

                const okBtn = modal.querySelector('#custom-prompt-ok');
                const cancelBtn = modal.querySelector('#custom-prompt-cancel');
                if (okBtn) okBtn.onclick = () => cleanup(input ? input.value.trim() : '');
                if (cancelBtn) cancelBtn.onclick = () => cleanup(null);
                if (input) {
                    input.onkeydown = (e) => {
                        if (e.key === 'Enter') cleanup(input.value.trim());
                        if (e.key === 'Escape') cleanup(null);
                    };
                }
            });
        }

        function openCustomConfirm(options) {
            return new Promise((resolve) => {
                const title = options.title || 'ยืนยันการดำเนินการ';
                const message = options.message || '';
                const confirmText = options.confirmText || 'ยืนยัน';
                const cancelText = options.cancelText || 'ยกเลิก';
                const isDanger = options.isDanger !== false;

                const modal = document.createElement('div');
                modal.className = 'fixed inset-0 bg-slate-900/70 z-[9999] flex items-center justify-center p-3 animate-fade-in no-print';
                modal.innerHTML = \`
                    <div class="bg-white w-full max-w-md rounded-2xl shadow-2xl p-5 space-y-4 text-xs sm:text-sm border border-slate-200">
                        <div class="flex items-start gap-3">
                            <div class="w-10 h-10 rounded-xl \${isDanger ? 'bg-rose-100 text-rose-600' : 'bg-blue-100 text-blue-600'} flex items-center justify-center text-lg shrink-0 font-bold">
                                \${isDanger ? '⚠️' : 'ℹ️'}
                            </div>
                            <div class="space-y-1">
                                <h3 class="font-bold text-base text-slate-900">\${title}</h3>
                                <p class="text-slate-600 whitespace-pre-line text-xs leading-relaxed">\${message}</p>
                            </div>
                        </div>
                        <div class="flex justify-end gap-2 pt-2 border-t">
                            <button type="button" id="custom-confirm-cancel" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer transition">\${cancelText}</button>
                            <button type="button" id="custom-confirm-ok" class="\${isDanger ? 'bg-rose-600 hover:bg-rose-700' : 'bg-emerald-600 hover:bg-emerald-700'} text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer transition">\${confirmText}</button>
                        </div>
                    </div>
                \`;
                document.body.appendChild(modal);

                const cleanup = (val) => {
                    modal.remove();
                    resolve(val);
                };

                const okBtn = modal.querySelector('#custom-confirm-ok');
                const cancelBtn = modal.querySelector('#custom-confirm-cancel');
                if (okBtn) okBtn.onclick = () => cleanup(true);
                if (cancelBtn) cancelBtn.onclick = () => cleanup(false);
            });
        }

        let customEquipList = safeJsonParse(localStorage.getItem('sahakim_custom_equip_types'), ${JSON.stringify(customEquipTypes)});
        let customBrandList = safeJsonParse(localStorage.getItem('sahakim_custom_brands'), ${JSON.stringify(customBrands)});
        let masterEquipmentTypes = safeJsonParse(localStorage.getItem('sahakim_master_equipment_types_v1'), Array.from(new Set([...BASE_EQUIPMENT_MASTERS, ...customEquipList])));
        let masterBrands = safeJsonParse(localStorage.getItem('sahakim_master_brands_v1'), Array.from(new Set([...BASE_BRANDS, ...customBrandList])));
        let companyProfile = safeJsonParse(localStorage.getItem('sahakim_company_profile_v1'), {
            nameTh: 'บริษัท สหคิมมอเตอร์ จำกัด',
            nameEn: 'SAHA KIM MOTOR CO., LTD.',
            phone: '02-XXX-XXXX',
            tagline: 'บริการซ่อม บำรุงรักษา และจำหน่ายปั๊มน้ำ มอเตอร์ Blower, Gen Set, Sewage, Aerator และระบบควบคุมอุตสาหกรรม'
        });
        const CLOUD_SYNC_API_URL = (window.location.origin.includes('run.app') || window.location.origin.includes('localhost'))
            ? (window.location.origin + '/api/online/state')
            : 'https://ais-pre-brppajwjfafqftfzmicbfo-297429063525.asia-southeast1.run.app/api/online/state';

        function getEquipmentMasters() {
            return Array.from(new Set([...masterEquipmentTypes, ...customEquipList, "ETC. (กรอกเพิ่มเองได้)"]));
        }
        function getBrands() {
            return Array.from(new Set([...masterBrands, ...customBrandList, "ETC. (กรอกเพิ่มเองได้)"]));
        }

        function setStandaloneTheme(mode) {
            const isDark = mode === 'dark';
            document.body.classList.toggle('dark-theme', isDark);
            try { localStorage.setItem('sahakim_theme_mode', isDark ? 'dark' : 'light'); } catch (e) {}
            const lightBtn = document.getElementById('btn-theme-light');
            const darkBtn = document.getElementById('btn-theme-dark');
            if (lightBtn && darkBtn) {
                lightBtn.className = isDark ? 'px-2 py-1 rounded font-bold text-slate-300 cursor-pointer' : 'px-2 py-1 rounded font-bold bg-white text-slate-900 cursor-pointer';
                darkBtn.className = isDark ? 'px-2 py-1 rounded font-bold bg-orange-500 text-white cursor-pointer' : 'px-2 py-1 rounded font-bold text-slate-300 cursor-pointer';
            }
        }
        (function initStandaloneTheme() {
            try {
                const saved = localStorage.getItem('sahakim_theme_mode');
                if (saved === 'dark') setStandaloneTheme('dark');
            } catch (e) {}
        })();

        const CHECKLIST_TEMPLATES = {
            "MOTOR": [
                "Stator Winding Insulation Resistance Test (Megger IEEE 43 / EASA AR100)",
                "Phase-to-Phase Winding Resistance Balance Check",
                "Bearing Housing Vibration & Temperature Measurement (ISO 10816)",
                "Terminal Box, Cable Lug Tightness & Grounding Inspection",
                "No-Load / Full-Load Current & Voltage Unbalance Check"
            ],
            "PUMP": [
                "Pump-Motor Laser / Dial Alignment & Coupling Inspection (Pump Handbook)",
                "Insulation Resistance Test (Megger Motor Winding)",
                "Mechanical Seal / Gland Packing Leakage & Flushing Line Check",
                "Suction & Discharge Pressure Gauge (NPSHa vs NPSHr / Cavitation Check)",
                "Impeller Wear Ring Clearance & Bearing Temperature Check"
            ],
            "BLOWER": [
                "V-Belt Tension & Sheave Pulley Alignment Check",
                "Rotary Lobe / Impeller Timing Clearance & Housing Inspection",
                "Inlet Filter Differential Pressure & Safety Relief Valve Test",
                "Gearbox Oil Level & Bearing Greasing Inspection"
            ],
            "GEN SET": [
                "Engine Oil Pressure, Coolant Temperature & Radiator Core Check",
                "Battery Charger Voltage & Cranking Voltage Test",
                "Alternator Winding Insulation & AVR Voltage Regulation Check",
                "ATS Transfer Switch & Auto-Start Simulation Test"
            ],
            "SEWAGE": [
                "Submersible Cable Entry Seal & Moisture Sensor Probe Check",
                "Non-Clog Impeller / Cutter Clearance & Wear Plate Inspection",
                "Mechanical Seal Oil Chamber Inspection (Water Emulsion Check)",
                "Motor Winding Insulation Resistance (Megger @ 500 VDC)"
            ],
            "EJECTOR": [
                "Jet Nozzle & Venturi Diffuser Throat Wear Inspection",
                "Motive Pump Discharge Pressure & Suction Vacuum Check",
                "Check Valve Sealing & Motor Current Check"
            ],
            "AIR PUMP": [
                "Diaphragm / Piston Valve & Air Distribution Valve Inspection",
                "Air Filter Regulator & Lubricator (FRL) Pressure Check",
                "Discharge Manifold & Operating Cycle Leakage Test"
            ],
            "AERATOR": [
                "Impeller / Propeller Balance & Shaft Sleeve Wear Check",
                "Gearbox / Reducer Oil Seal & Lubrication Check",
                "Motor Winding Insulation & Running Current Balance"
            ],
            "CONTROL": [
                "Main Circuit Breaker (MCCB), Magnetic Contactor & Overload Check",
                "VFD / Soft Starter Parameter, Fault Log & Cooling Fan Check",
                "Phase Protection Relay, Timer & Float/Pressure Switch Test",
                "Power & Control Terminal Torque Tightness Check"
            ]
        };

        function getChecklistForType(type) {
            let upper = (type || "").toUpperCase();
            for (let k of Object.keys(CHECKLIST_TEMPLATES)) {
                if (upper.includes(k)) return [...CHECKLIST_TEMPLATES[k]];
            }
            return [...CHECKLIST_TEMPLATES["PUMP"]];
        }

        const TECH_POSITIONS = [
            "ช่างเทคนิค (Technician)",
            "หัวหน้าช่าง (Lead Technician)",
            "วิศวกรบริการ (Service Engineer)"
        ];

        function isDemoTechnician(tc) {
            if (!tc || !tc.name) return false;
            const n = String(tc.name).trim();
            return n.includes('วิชัย') || n.includes('สมชาย') || n.includes('ประเสริฐ');
        }

        function isDemoTicket(t) {
            if (!t) return false;
            const id = String(t.id || '').trim();
            const comp = String(t.company || '').trim();
            if (id === 'SKM-69-001' || id === 'SKM-69-002' || id === 'SKM-69-003') return true;
            if (comp.includes('ตัวอย่าง') || comp.includes('เดโม') || comp.includes('Demo')) return true;
            return false;
        }

        let knowledgeTemplates = safeJsonParse(localStorage.getItem('sahakim_knowledge_templates'), ${JSON.stringify(currentTemplates)});
        let savedSupPassword = localStorage.getItem('sahakim_sup_pass') || ${JSON.stringify(supPassword)};
        let savedServicePin = localStorage.getItem('sahakim_service_pin') || ${JSON.stringify(servicePin)};
        let savedSupervisors = localStorage.getItem('sahakim_supervisors');
        let initialSupervisors = (savedSupervisors ? safeJsonParse(savedSupervisors, ${JSON.stringify(currentSupervisors)}) : ${JSON.stringify(currentSupervisors)}).filter(s => typeof s === 'string' && !s.includes('วิชัย'));
        let savedTechs = localStorage.getItem('sahakim_technicians');
        let parsedTechs = savedTechs ? safeJsonParse(savedTechs, null) : null;
        let rawTechs = Array.isArray(parsedTechs) ? parsedTechs : ${JSON.stringify(currentTechnicians)};
        let initialTechs = (Array.isArray(rawTechs) ? rawTechs : [])
            .filter(tc => !isDemoTechnician(tc))
            .map((tc, i) => ({
                ...tc,
                position: tc.position || (i === 0 ? "หัวหน้าช่าง (Lead Technician)" : "ช่างเทคนิค (Technician)")
            }));
        try { localStorage.setItem('sahakim_technicians', JSON.stringify(initialTechs)); } catch(e) {}
        
        let savedTickets = localStorage.getItem('sahakim_tickets_v121');
        const INITIAL_DEMO_TICKETS = ${JSON.stringify(sanitizedTickets)};
        let parsedSaved = savedTickets ? safeJsonParse(savedTickets, null) : null;
        let initialTickets = (Array.isArray(parsedSaved) ? parsedSaved : INITIAL_DEMO_TICKETS).filter(t => !isDemoTicket(t));

        let state = {
            currentView: 'home',
            userRole: null,
            loggedInTech: null,
            customerCompany: 'บริษัท สยามอุตสาหกรรม จำกัด',
            technicians: initialTechs,
            supervisors: initialSupervisors,
            tickets: initialTickets,
            activeReportId: null
        };
        let techWorkFilter = 'MY_JOBS'; // 'MY_JOBS' | 'AVAILABLE_JOBS' | 'ALL'

        async function restoreFromIndexedDBIfAvailable() {
            try {
                const idbData = await loadFromIndexedDB('full_state');
                if (idbData && typeof idbData === 'object') {
                    let updated = false;
                    if (Array.isArray(idbData.tickets) && idbData.tickets.length > 0) {
                        const cleanedIdbTickets = idbData.tickets.filter(t => !isDemoTicket(t));
                        if (cleanedIdbTickets.length >= state.tickets.length) {
                            state.tickets = cleanedIdbTickets;
                            updated = true;
                        }
                    }
                    if (Array.isArray(idbData.technicians) && idbData.technicians.length > 0) {
                        const cleanedIdbTechs = idbData.technicians.filter(tc => !isDemoTechnician(tc));
                        if (cleanedIdbTechs.length > 0) {
                            state.technicians = cleanedIdbTechs;
                            updated = true;
                        }
                    }
                    if (idbData.masterBrands) masterBrands = idbData.masterBrands;
                    if (idbData.masterEquipmentTypes) masterEquipmentTypes = idbData.masterEquipmentTypes;
                    if (idbData.companyProfile) companyProfile = idbData.companyProfile;
                    if (updated) {
                        render();
                    }
                }
            } catch (e) {
                console.warn('IndexedDB auto-restore:', e);
            }
        }
        setTimeout(restoreFromIndexedDBIfAvailable, 100);

        function compressImageFile(file, maxWidth, maxHeight, quality) {
            maxWidth = maxWidth || 800;
            maxHeight = maxHeight || 800;
            quality = quality || 0.68;
            return new Promise(function(resolve) {
                if (!file || !file.type || !file.type.match(/image.*/)) { resolve(''); return; }
                const reader = new FileReader();
                reader.onerror = function() { resolve(''); };
                reader.onload = function(e) {
                    const img = new Image();
                    img.onerror = function() { resolve(String(e.target.result || '')); };
                    img.onload = function() {
                        let width = img.width;
                        let height = img.height;
                        if (width > maxWidth || height > maxHeight) {
                            if (width > height) {
                                height = Math.round((height * maxWidth) / width);
                                width = maxWidth;
                            } else {
                                width = Math.round((width * maxHeight) / height);
                                height = maxHeight;
                            }
                        }
                        const canvas = document.createElement('canvas');
                        canvas.width = width;
                        canvas.height = height;
                        const ctx = canvas.getContext('2d');
                        if (!ctx) { resolve(String(e.target.result || '')); return; }
                        ctx.fillStyle = '#ffffff';
                        ctx.fillRect(0, 0, width, height);
                        ctx.drawImage(img, 0, 0, width, height);
                        try {
                            resolve(canvas.toDataURL('image/jpeg', quality));
                        } catch(err) {
                            resolve(String(e.target.result || ''));
                        }
                    };
                    img.src = String(e.target.result || '');
                };
                reader.readAsDataURL(file);
            });
        }

        async function handleStandaloneEquipmentPhotoUpload(eqIndex, field, fileList) {
            if (!fileList || fileList.length === 0) return;
            const eq = tempEquipmentList[eqIndex];
            if (!eq) return;
            if (!eq[field]) eq[field] = [];
            for (let i = 0; i < fileList.length; i++) {
                const compressed = await compressImageFile(fileList[i]);
                if (compressed) {
                    eq[field].push(compressed);
                }
            }
            renderEquipmentRows();
        }

        function removeStandaloneEquipmentPhoto(eqIndex, field, photoIndex) {
            const eq = tempEquipmentList[eqIndex];
            if (eq && eq[field]) {
                eq[field].splice(photoIndex, 1);
                renderEquipmentRows();
            }
        }

        async function resetToFreshProduction() {
            const confirmed = await openCustomConfirm({
                title: 'ยืนยันรีเซ็ตระบบเพื่อเริ่มต้นใช้งานจริง?',
                message: '• ใบงานตัวอย่างจะถูกล้างออกทั้งหมด (ใบงาน = 0 รายการ)\\n• ตารางงานช่างจะกลับมาว่างทั้งหมด (สีเขียว 🟢 ทุกวัน) พร้อมรับงานจริงทันที\\n• ข้อมูลรายชื่อช่างและข้อมูลตั้งต้นจะยังคงอยู่ครบถ้วน ไม่สูญหาย',
                confirmText: 'ล้างงานตัวอย่าง เริ่มใช้งานจริง',
                isDanger: true
            });
            if (!confirmed) return;
            state.tickets = [];
            try { localStorage.setItem('sahakim_tickets_v121', '[]'); } catch(e) {}
            persistState();
            showAppToast('ล้างใบงานตัวอย่างเรียบร้อยแล้ว! ตารางงานช่างว่าง (สีเขียว 🟢) พร้อมเริ่มงานจริง', 'success');
            render();
        }

        async function restoreDemoTickets() {
            const confirmed = await openCustomConfirm({
                title: 'กู้คืนข้อมูลใบงานออนไลน์?',
                message: 'ต้องการดึงข้อมูลใบงานออนไลน์ล่าสุดจากเซิร์ฟเวอร์มาแสดงบนเครื่องนี้ ใช่หรือไม่?',
                confirmText: 'ดึงข้อมูลออนไลน์',
                isDanger: false
            });
            if (!confirmed) return;
            const pulled = await pullFromCloudSyncServer(false);
            if (!pulled) {
                state.tickets = JSON.parse(JSON.stringify(INITIAL_DEMO_TICKETS)).filter(t => !isDemoTicket(t));
                try { localStorage.setItem('sahakim_tickets_v121', JSON.stringify(state.tickets)); } catch(e) {}
                persistState();
            }
            showAppToast('ซิงค์ข้อมูลออนไลน์ล่าสุดเรียบร้อยแล้ว (พบใบงาน ' + state.tickets.length + ' รายการ)', 'success');
            render();
        }

        async function handleStandaloneEditingPhotoUpload(ticketId, eqIndex, field, fileList) {
            if (!fileList || fileList.length === 0) return;
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t || !t.equipmentList || !t.equipmentList[eqIndex]) return;
            const eq = t.equipmentList[eqIndex];
            if (!eq[field]) eq[field] = [];
            for (let i = 0; i < fileList.length; i++) {
                const compressed = await compressImageFile(fileList[i]);
                if (compressed) {
                    eq[field].push(compressed);
                }
            }
            openInternalActionModal(ticketId);
        }

        function removeStandaloneEditingPhoto(ticketId, eqIndex, field, photoIndex) {
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t || !t.equipmentList || !t.equipmentList[eqIndex]) return;
            const eq = t.equipmentList[eqIndex];
            if (eq && eq[field]) {
                eq[field].splice(photoIndex, 1);
                openInternalActionModal(ticketId);
            }
        }

        async function deleteTicket(ticketId) {
            const t = state.tickets.find(item => item.id === ticketId);
            if (!t) return;
            const confirmed = await openCustomConfirm({
                title: 'ยืนยันการลบใบงาน ' + t.id,
                message: 'คุณกำลังจะลบใบงานของ ' + t.company + ' ออกจากระบบ\\n\\n• ข้อมูลใบงานและคิวงานช่างจะถูกลบออกทันที\\n• ระบบจะซิงค์การลบข้อมูลไปยังทุกเครื่องออนไลน์',
                confirmText: 'ลบใบงานนี้',
                isDanger: true
            });
            if (!confirmed) return;
            state.tickets = state.tickets.filter(item => item.id !== ticketId);
            try { localStorage.setItem('sahakim_tickets_v121', JSON.stringify(state.tickets)); } catch(e) {}
            persistState();
            syncToGoogleSheet('🗑️ [ลบใบงาน] ' + t.id + ' (' + t.company + ') ถูกลบออกจากระบบแล้ว');
            pushToCloudSyncServer('delete_ticket_' + ticketId);
            showAppToast('ลบใบงาน ' + ticketId + ' ออกจากระบบเรียบร้อยแล้ว', 'success');
            render();
        }

        async function unassignTicket(ticketId) {
            const t = state.tickets.find(item => item.id === ticketId);
            if (!t) return;
            const prevTech = t.assignedTech || 'ช่าง';
            const confirmed = await openCustomConfirm({
                title: 'ต้องการปลดช่างออกจากใบงาน?',
                message: 'ปลดช่าง ' + prevTech + ' ออกจากใบงาน ' + t.id + '\\n\\n• คิวงานของช่าง ' + prevTech + ' จะว่างลงทันที (กลับเป็นสีเขียว 🟢)\\n• สถานะใบงานจะเปลี่ยนกลับเป็น NEW (รอจัดคิวใหม่)',
                confirmText: 'ปลดช่างออก',
                isDanger: false
            });
            if (!confirmed) return;
            t.assignedTech = '';
            t.confirmedDate = '';
            t.status = 'NEW';
            try { localStorage.setItem('sahakim_tickets_v121', JSON.stringify(state.tickets)); } catch(e) {}
            persistState();
            syncToGoogleSheet();
            pushToCloudSyncServer('unassign_ticket_' + ticketId);
            showAppToast('ปลดช่าง ' + prevTech + ' ออกจากใบงาน ' + ticketId + ' เรียบร้อยแล้ว (คิวงานช่างว่างลงแล้ว)', 'success');
            render();
        }

        function updateGsheetBadge() {
            const badge = document.getElementById('gsheet-status-badge');
            if (!badge) return;
            if (gasWebAppUrl) {
                badge.textContent = 'ซิงค์ Google Sheet แล้ว';
            } else {
                badge.textContent = 'ตั้งค่า Google Sheet';
            }
        }

        function formatTicketLineMessage(t, eventHeading) {
            const portalUrl = window.location.href.startsWith('http') && !window.location.href.includes('localhost')
                ? window.location.href
                : 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
            const eqSummary = (t.equipmentList || []).map(eq => eq.type + ' (' + eq.brand + ' ' + (eq.model || '-') + ') อาการ: ' + eq.symptom).join(', ');
            let totalPartsCost = 0;
            let partsList = [];
            (t.equipmentList || []).forEach(eq => {
                (eq.parts || []).forEach(p => {
                    totalPartsCost += (p.total || 0);
                    partsList.push(p.name + ' x' + p.qty + ' (' + p.total.toLocaleString() + '฿)');
                });
            });
            let lines = [
                eventHeading || '🔔 [อัปเดตงานบริการ - สหคิมมอเตอร์]',
                '━━━━━━━━━━━━━━',
                '📋 เลขที่ใบงาน: ' + t.id,
                '🏢 บริษัท: ' + t.company,
                '📌 สถานะ: ' + t.status,
                '🚨 ความเร่งด่วน: ' + (t.urgency || '-'),
                '🔧 อุปกรณ์: ' + (eqSummary || '-'),
                '📍 สถานที่: ' + (t.location || '-'),
                '📞 ผู้ติดต่อ: ' + (t.siteContact || t.contact || '-')
            ];
            if (t.assignedTech) lines.push('👷 ช่างผู้ปฏิบัติงาน: ' + t.assignedTech);
            if (t.leadTech) lines.push('👨‍🔧 หัวหน้าช่างผู้ควบคุมงาน: ' + t.leadTech);
            if (t.approvedBySupervisor) lines.push('✅ ผู้อนุมัติ (Supervisor/หัวหน้าช่าง): ' + t.approvedBySupervisor);
            if (t.confirmedDate || t.desiredDate) lines.push('📅 วันนัดหมาย: ' + (t.confirmedDate || t.desiredDate));
            if (partsList.length > 0) {
                lines.push('⚙️ อะไหล่ที่ขออนุมัติ: ' + partsList.join(', ') + ' [รวม ' + totalPartsCost.toLocaleString() + ' THB]');
            }
            if (t.rejectReason && t.status === 'REJECTED') {
                lines.push('❌ เหตุผลส่งกลับแก้ไข: ' + t.rejectReason);
            }
            lines.push('━━━━━━━━━━━━━━');
            lines.push('🔗 เปิดระบบ/อนุมัติงาน: ' + portalUrl);
            return lines.join('\\n');
        }

        function getTodayDateStr() {
            const now = new Date();
            const yyyy = now.getFullYear();
            const mm = String(now.getMonth() + 1).padStart(2, '0');
            const dd = String(now.getDate()).padStart(2, '0');
            return yyyy + '-' + mm + '-' + dd;
        }

        function extractTicketDueDate(t) {
            const raw = (t.confirmedDate || t.desiredDate || '').trim();
            const match = raw.match(/\\d{4}-\\d{2}-\\d{2}/);
            return match ? match[0] : raw;
        }

        function isTicketOverdue(t) {
            if (!t || t.status === 'CLOSE') return false;
            const dueStr = extractTicketDueDate(t);
            if (!dueStr || !/^\\d{4}-\\d{2}-\\d{2}$/.test(dueStr)) return false;
            return dueStr < getTodayDateStr();
        }

        function getOverdueDays(t) {
            const dueStr = extractTicketDueDate(t);
            const todayStr = getTodayDateStr();
            if (!dueStr || dueStr >= todayStr) return 0;
            const diffMs = new Date(todayStr).getTime() - new Date(dueStr).getTime();
            return Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)));
        }

        function formatOverdueSummaryLineMessage(overdueList) {
            const portalUrl = window.location.href.startsWith('http') && !window.location.href.includes('localhost')
                ? window.location.href
                : 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
            const todayStr = getTodayDateStr();
            let lines = [
                '🚨 [แจ้งเตือนงานค้างเกินกำหนดนัดหมาย (OVERDUE) - สหคิมมอเตอร์]',
                '📅 ข้อมูล ณ วันที่: ' + todayStr,
                '⚠️ พบงานค้างเกินกำหนดนัดหมาย จำนวน ' + overdueList.length + ' ใบงาน:',
                '━━━━━━━━━━━━━━'
            ];
            overdueList.forEach((t, idx) => {
                const dueStr = extractTicketDueDate(t);
                const daysLate = getOverdueDays(t);
                const eqStr = (t.equipmentList || []).map(eq => eq.type + ' (' + eq.brand + ')').join(', ');
                lines.push((idx + 1) + ') 📋 ' + t.id + ' | ' + t.company);
                lines.push('   📌 สถานะ: ' + t.status);
                lines.push('   🗓️ วันนัดหมาย: ' + dueStr + ' (เกินกำหนด ' + daysLate + ' วัน)');
                lines.push('   👷 ช่างผู้รับผิดชอบ: ' + (t.assignedTech || 'ยังไม่ระบุช่าง'));
                lines.push('   🔧 อุปกรณ์: ' + (eqStr || '-'));
                lines.push('   📍 สถานที่: ' + (t.location || '-'));
                lines.push('━━━━━━━━━━━━━━');
            });
            lines.push('🔗 กดเพื่อตรวจสอบและเร่งรัดงาน: ' + portalUrl);
            return lines.join('\\n');
        }

        function formatAllTicketsSummaryLineMessage(allList) {
            const portalUrl = window.location.href.startsWith('http') && !window.location.href.includes('localhost')
                ? window.location.href
                : 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
            const todayStr = getTodayDateStr();
            const overdueCount = allList.filter(t => isTicketOverdue(t)).length;
            const closedCount = allList.filter(t => t.status === 'CLOSE').length;
            const activeCount = allList.length - closedCount;
            let lines = [
                '📊 [อัปเดตสรุปข้อมูลใบงานทั้งหมด - บริษัท สหคิมมอเตอร์ จำกัด]',
                '📅 ข้อมูล ณ วันที่: ' + todayStr,
                '📋 งานทั้งหมด: ' + allList.length + ' ใบงาน (กำลังดำเนินการ ' + activeCount + ' | ปิดงานแล้ว ' + closedCount + ' | เกินกำหนด ' + overdueCount + ')',
                '━━━━━━━━━━━━━━'
            ];
            allList.forEach((t, idx) => {
                const dueStr = extractTicketDueDate(t);
                const overdue = isTicketOverdue(t);
                const eqStr = (t.equipmentList || []).map(eq => eq.type + ' (' + eq.brand + ' ' + (eq.model || '-') + ')').join(', ');
                lines.push((idx + 1) + ') 📋 ' + t.id + ' — ' + t.company);
                lines.push('   📌 สถานะ: ' + t.status + (overdue ? ' [⚠️ เกินกำหนดนัด]' : ''));
                lines.push('   🔧 อุปกรณ์: ' + (eqStr || '-'));
                lines.push('   👷 ช่าง: ' + (t.assignedTech || 'ยังไม่ระบุ') + ' | 🗓️ นัด: ' + (dueStr || '-'));
                lines.push('━━━━━━━━━━━━━━');
            });
            lines.push('🔗 เปิดระบบจัดการใบงาน: ' + portalUrl);
            return lines.join('\\n');
        }

        async function syncAllTicketsUpdateToLine() {
            if (!gasWebAppUrl) {
                openGoogleSheetModal();
                return;
            }
            await syncToGoogleSheet();
            alert('✅ อัปเดตข้อมูลใบงานทั้งหมด (' + state.tickets.length + ' ใบงาน) ลง Google Sheet เรียบร้อยแล้ว!');
        }

        async function notifyOverdueTicketsToLine() {
            const overdueList = state.tickets.filter(t => isTicketOverdue(t));
            if (overdueList.length === 0) {
                alert('✅ ไม่มีงานค้างที่เกินกำหนดนัดหมายในขณะนี้ (ทุกใบงานอยู่ในกำหนดเวลา)');
                return;
            }
            alert('🚨 มีงานค้างเกินกำหนดนัดหมาย ' + overdueList.length + ' ใบงาน กรุณาตรวจสอบในรายการ');
        }

        async function shareTicketToLine(ticketId, customHeading) {
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t) return;
            const msg = formatTicketLineMessage(t, customHeading || '📋 [สรุปรายละเอียดใบงาน - บริษัท สหคิมมอเตอร์ จำกัด]');
            try {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    await navigator.clipboard.writeText(msg);
                }
            } catch (e) {}
            alert('📋 คัดลอกข้อความสรุปใบงาน ' + t.id + ' ลงคลิปบอร์ดเรียบร้อยแล้ว!');
        }

        function sanitizeTicketsForSheet(list) {
            if (!Array.isArray(list)) return [];
            return list.map(t => ({
                ...t,
                equipmentList: (t.equipmentList || []).map(eq => ({
                    ...eq,
                    beforePhotos: (eq.beforePhotos || []).map(p =>
                        (typeof p === 'string' && (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800))
                            ? 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=600'
                            : p
                    ),
                    afterPhotos: (eq.afterPhotos || []).map(p =>
                        (typeof p === 'string' && (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800))
                            ? 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=600'
                            : p
                    )
                }))
            }));
        }

        async function syncToGoogleSheet() {
            if (gasWebAppUrl && !gasWebAppUrl.includes('docs.google.com/spreadsheets')) {
                const payloadStr = JSON.stringify({
                    tickets: sanitizeTicketsForSheet(state.tickets)
                });
                try {
                    await fetch(gasWebAppUrl, {
                        method: 'POST',
                        redirect: 'follow',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: payloadStr
                    }).catch(() =>
                        fetch(gasWebAppUrl, {
                            method: 'POST',
                            mode: 'no-cors',
                            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                            body: payloadStr
                        })
                    );
                } catch (e) {
                    console.warn('Google Sheet Sync warning:', e);
                }
            }
        }

        async function downloadPdfById(elementId, filename) {
            const sourceEl = document.getElementById(elementId);
            if (!sourceEl) {
                window.print();
                return;
            }
            const restoredStyles = [];
            try {
                const cv = document.createElement('canvas');
                cv.width = 1; cv.height = 1;
                const ctx = cv.getContext('2d', { willReadFrequently: true });
                const toRgbColor = (cssColor) => {
                    if (!ctx) return '#1e293b';
                    try {
                        ctx.clearRect(0, 0, 1, 1);
                        ctx.fillStyle = '#0f172a';
                        ctx.fillStyle = cssColor.trim();
                        ctx.fillRect(0, 0, 1, 1);
                        const d = ctx.getImageData(0, 0, 1, 1).data;
                        return d[3] < 255 ? \`rgba(\${d[0]}, \${d[1]}, \${d[2]}, \${(d[3]/255).toFixed(2)})\` : \`rgb(\${d[0]}, \${d[1]}, \${d[2]})\`;
                    } catch (e) { return '#1e293b'; }
                };
                const hasModernColorFn = (val) => /oklch|oklab|color-mix|lch\\(|lab\\(|color\\(/i.test(val || '');
                const sanitizeModernColors = (cssText) => {
                    if (!cssText || !hasModernColorFn(cssText)) return cssText;
                    let out = cssText
                        .replace(/@supports\\s*\\(\\s*color\\s*:\\s*color-mix\\([^)]*\\)\\s*\\)/gi, '@supports (color: never-match-pdf)')
                        .replace(/\\s+in\\s+(?:oklab|oklch|srgb|srgb-linear|lab|lch|hsl|hwb)\\b/gi, '');
                    for (let pass = 0; pass < 3; pass++) {
                        out = out
                            .replace(/oklch\\([^()]+\\)/gi, (m) => toRgbColor(m))
                            .replace(/oklab\\([^()]+\\)/gi, (m) => toRgbColor(m))
                            .replace(/color\\([^()]+\\)/gi, (m) => toRgbColor(m))
                            .replace(/color-mix\\([^()]+\\)/gi, (m) => toRgbColor(m));
                    }
                    return out;
                };

                const COLOR_PROPS = [
                    'color', 'background-color', 'border-top-color', 'border-right-color',
                    'border-bottom-color', 'border-left-color', 'outline-color', 'text-decoration-color',
                    'fill', 'stroke', 'box-shadow', 'text-shadow', 'background-image'
                ];

                document.querySelectorAll('style').forEach((styleEl) => {
                    if (styleEl.textContent && hasModernColorFn(styleEl.textContent)) {
                        styleEl.textContent = sanitizeModernColors(styleEl.textContent);
                    }
                });

                const inlineSafeColors = (doc, nodes, record) => {
                    const win = doc.defaultView || window;
                    nodes.forEach((node) => {
                        if (!node || !node.style) return;
                        const cs = win.getComputedStyle(node);
                        let saved = false;
                        COLOR_PROPS.forEach((prop) => {
                            const val = cs.getPropertyValue(prop);
                            if (val && hasModernColorFn(val)) {
                                if (record && !saved) {
                                    restoredStyles.push({ el: node, cssText: node.style.cssText });
                                    saved = true;
                                }
                                const cleaned = sanitizeModernColors(val);
                                node.style.setProperty(prop, hasModernColorFn(cleaned) ? (prop.includes('shadow') || prop === 'background-image' ? 'none' : '#0f172a') : cleaned, 'important');
                            }
                        });
                    });
                };

                const ancestors = [];
                let curr = sourceEl.parentElement;
                while (curr) { ancestors.push(curr); curr = curr.parentElement; }
                inlineSafeColors(document, [...ancestors, sourceEl, ...Array.from(sourceEl.querySelectorAll('*'))], true);

                const opt = {
                    margin: [8, 8, 8, 8],
                    filename: filename.endsWith('.pdf') ? filename : filename + '.pdf',
                    image: { type: 'jpeg', quality: 0.96 },
                    html2canvas: {
                        scale: 2,
                        useCORS: true,
                        allowTaint: true,
                        backgroundColor: '#ffffff',
                        logging: false,
                        onclone: (clonedDoc) => {
                            clonedDoc.querySelectorAll('style').forEach((styleEl) => {
                                if (styleEl.textContent && hasModernColorFn(styleEl.textContent)) {
                                    styleEl.textContent = sanitizeModernColors(styleEl.textContent);
                                }
                            });
                            clonedDoc.querySelectorAll('.no-print').forEach((el) => el.remove());
                            clonedDoc.querySelectorAll('[id="' + elementId + '"]').forEach((target) => {
                                target.style.maxHeight = 'none';
                                target.style.height = 'auto';
                                target.style.overflow = 'visible';
                                target.style.boxShadow = 'none';
                            });
                            inlineSafeColors(clonedDoc, [clonedDoc.documentElement, clonedDoc.body, ...Array.from(clonedDoc.querySelectorAll('*'))], false);
                        }
                    },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                    pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
                };
                await window.html2pdf().set(opt).from(sourceEl).save();
            } catch (e) {
                console.warn('PDF export fallback to print:', e);
                window.print();
            } finally {
                for (let i = restoredStyles.length - 1; i >= 0; i--) {
                    restoredStyles[i].el.style.cssText = restoredStyles[i].cssText;
                }
            }
        }

        function openUserManualModal() {
            const existing = document.getElementById('user-manual-modal');
            if (existing) existing.remove();
            let modalBg = document.createElement('div');
            modalBg.id = 'user-manual-modal';
            modalBg.className = 'print-modal-root fixed inset-0 bg-slate-900/70 z-50 flex items-center justify-center p-3 overflow-y-auto';
            modalBg.innerHTML = \`
                <div id="user-manual-print-area" class="print-modal-content bg-white w-full max-w-4xl rounded-2xl shadow-2xl p-5 sm:p-8 space-y-5 max-h-[92vh] overflow-y-auto text-xs sm:text-sm">
                    <div class="flex flex-wrap justify-between items-start gap-3 border-b-2 border-orange-500 pb-3">
                        <div>
                            <p class="text-xs font-bold text-orange-600">SAHA KIM MOTOR CO., LTD. · ENTERPRISE SERVICE SYSTEM V12.1</p>
                            <h2 class="text-lg sm:text-xl font-bold text-slate-900">คู่มือการใช้งานระบบบริการหลังการขาย (After-Sales Service Web App Manual)</h2>
                        </div>
                        <div class="flex flex-wrap items-center gap-2 no-print">
                            <button onclick="downloadPdfById('user-manual-print-area', 'Sahakim-After-Sales-Service-Manual.pdf')" class="bg-orange-500 hover:bg-orange-600 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer shadow-sm"><i class="fa-solid fa-file-pdf mr-1"></i> บันทึกคู่มือเป็น PDF (Save PDF)</button>
                            <button onclick="window.print()" class="bg-slate-800 hover:bg-slate-900 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-print mr-1"></i> สั่งพิมพ์</button>
                            <button onclick="document.getElementById('user-manual-modal').remove()" class="bg-slate-200 text-slate-800 px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer">ปิด ✕</button>
                        </div>
                    </div>
                    <div class="bg-slate-900 text-white p-4 rounded-xl space-y-1.5">
                        <p class="font-bold text-orange-400">📌 ลำดับสถานะใบงานมาตรฐาน (Workflow Status):</p>
                        <p class="text-xs leading-relaxed"><b>1. NEW</b> (ลูกค้าแจ้งซ่อมใหม่) ➔ <b>2. ASSIGNED</b> (นัดวัน/มอบหมายช่าง) ➔ <b>3. IN_PROGRESS</b> (ช่างเข้าตรวจเช็ค) ➔ <b>4. PENDING_PARTS_APPROVAL</b> (ขออนุมัติอะไหล่) ➔ <b>5. REPAIRING</b> (กำลังซ่อม) ➔ <b>6. PENDING_SUPERVISOR_APPROVAL</b> (ช่างส่งรายงานให้ตรวจ) ➔ <b>7. CUSTOMER_ACCEPTANCE</b> (Supervisor อนุมัติแล้ว/รอลูกค้าตรวจรับ) ➔ <b>8. CLOSE</b> (ปิดงานสมบูรณ์)</p>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="border border-orange-200 bg-orange-50/50 rounded-xl p-4 space-y-2">
                            <h3 class="font-bold text-orange-800 text-sm">1. สำหรับลูกค้า (Customer Portal)</h3>
                            <ul class="list-disc pl-5 space-y-1 text-slate-700">
                                <li>เลือกเมนู <b>"1. ฝั่งลูกค้า"</b> แล้วระบุชื่อบริษัทของท่าน</li>
                                <li>กดปุ่ม <b>"+ แจ้งซ่อมใหม่"</b> กรอกข้อมูลผู้ติดต่อ สถานที่ เงื่อนไขประกัน และเพิ่มรายการอุปกรณ์พร้อมรูปถ่าย</li>
                                <li>เมื่อกดส่ง ระบบจะสร้างเลขใบงาน (เช่น SKM-69-xxx) และแจ้งเตือนเข้า LINE บอททันที</li>
                                <li>ติดตามสถานะออนไลน์ ดูใบ Service Report และกด <b>"ตรวจรับงาน"</b> เมื่อซ่อมเสร็จ</li>
                            </ul>
                        </div>
                        <div class="border border-blue-200 bg-blue-50/50 rounded-xl p-4 space-y-2">
                            <h3 class="font-bold text-blue-900 text-sm">2. สำหรับเจ้าหน้าที่รับเรื่อง (Service Support - PIN: 1234)</h3>
                            <ul class="list-disc pl-5 space-y-1 text-slate-700">
                                <li>เข้าเมนู <b>"2. เจ้าหน้าที่ภายใน"</b> ➔ เลือก <b>SERVICE TEAM</b> (รหัสเริ่มต้น 1234)</li>
                                <li>กดปุ่ม <b>"นัดวัน / จ่ายงานช่าง"</b> ในใบงานสถานะ NEW เพื่อระบุวันนัดหมาย ช่างผู้ปฏิบัติงาน และหัวหน้าช่าง</li>
                                <li>กดปุ่ม <b>"แจ้งเตือนงานค้างเข้า LINE"</b> หากมีใบงานเลยกำหนดนัดหมาย (Overdue)</li>
                            </ul>
                        </div>
                        <div class="border border-amber-200 bg-amber-50/50 rounded-xl p-4 space-y-2">
                            <h3 class="font-bold text-amber-900 text-sm">3. สำหรับช่างเทคนิค (Technician - PIN: 1111, 2222, 3333)</h3>
                            <ul class="list-disc pl-5 space-y-1 text-slate-700">
                                <li>เข้าเมนู <b>"2. เจ้าหน้าที่ภายใน"</b> ➔ เลือก <b>TECHNICIAN</b> ➔ กดปุ่ม <b>"จัดการ"</b> ในใบงาน</li>
                                <li>เลือกต้นแบบการวิเคราะห์อาการเสียตามมาตรฐาน <b>Pump Handbook / EASA AR100</b> หรือพิมพ์ผลตรวจเช็ค</li>
                                <li>บันทึก Checklist (Before/After), อัปโหลดรูปถ่ายหน้างาน และเพิ่มรายการเบิกอะไหล่</li>
                                <li>กดปุ่มสีม่วง <b>"ส่ง Supervisor อนุมัติรายงาน"</b> เมื่อปฏิบัติงานเสร็จสิ้น</li>
                            </ul>
                        </div>
                        <div class="border border-emerald-200 bg-emerald-50/50 rounded-xl p-4 space-y-2">
                            <h3 class="font-bold text-emerald-900 text-sm">4. สำหรับหัวหน้างาน (Supervisor - รหัสผ่าน: 9999)</h3>
                            <ul class="list-disc pl-5 space-y-1 text-slate-700">
                                <li>กดปุ่มสีเขียว <b>"อนุมัติอะไหล่"</b> หรือ <b>"อนุมัติรายงาน"</b> พร้อมเลือกชื่อผู้อนุมัติ</li>
                                <li>กดปุ่ม <b>"จัดการช่าง"</b> เพื่อเพิ่ม/ลบรายชื่อช่าง หัวหน้าช่าง และเปลี่ยนรหัสผ่าน</li>
                                <li>กดปุ่มสีเขียว <b>"ส่ง LINE"</b> ท้ายใบงานเพื่อให้บอท @389fyhbl ยิงแจ้งเตือนเข้ากลุ่มทันที</li>
                            </ul>
                        </div>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        async function loadFromGoogleSheet(showMessage) {
            if (!gasWebAppUrl) {
                if (showMessage) alert('⚠️ กรุณาวางลิงก์ Google Apps Script Web App URL (ลงท้ายด้วย /exec) ก่อนครับ');
                return;
            }
            if (gasWebAppUrl.includes('docs.google.com/spreadsheets')) {
                if (showMessage) alert('⚠️ ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet (docs.google.com)\\n\\nกรุณาใช้ลิงก์ Web App URL จากเมนู ส่วนขยาย > Apps Script > การทำให้ใช้งานได้ (Deploy) ที่ลงท้ายด้วย /exec ครับ');
                return;
            }
            try {
                const res = await fetch(gasWebAppUrl, { redirect: 'follow' });
                const data = await res.json();
                if (data && Array.isArray(data.tickets) && data.tickets.length > 0) {
                    state.tickets = data.tickets;
                    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(state.tickets));
                    render();
                    if (showMessage) alert('✅ ดึงข้อมูลล่าสุดจาก Google Sheet สำเร็จ (' + data.tickets.length + ' ใบงาน)');
                } else if (showMessage) {
                    alert('✅ เชื่อมต่อสำเร็จ แต่ยังไม่มีข้อมูลใน Google Sheet ระบบกำลังส่งข้อมูลปัจจุบัน (' + state.tickets.length + ' ใบงาน) ขึ้นไปบันทึกให้ทันที');
                    await syncToGoogleSheet();
                }
            } catch (e) {
                if (showMessage) alert('❌ ไม่สามารถดึงข้อมูลจาก URL นี้ได้ กรุณาตรวจสอบว่าตั้งค่า Deploy Web App เป็น "ทุกคน (Anyone)" และใช้ลิงก์ที่ลงท้ายด้วย /exec แล้ว');
            }
        }

        let internalActiveTab = 'tickets'; // 'tickets' | 'new_ticket' | 'schedule' | 'dashboard' | 'all'
        let customerActiveTab = 'list'; // 'list' | 'new' | 'schedule'

        async function pushToCloudSyncServer(reason) {
            try {
                const payload = {
                    tickets: sanitizeTicketsForSheet(state.tickets),
                    technicians: state.technicians,
                    supervisors: state.supervisors,
                    knowledgeTemplates: knowledgeTemplates,
                    masterBrands: masterBrands,
                    masterEquipmentTypes: masterEquipmentTypes,
                    masterChecklists: CHECKLIST_TEMPLATES,
                    companyProfile: companyProfile,
                    customEquipTypes: customEquipList,
                    customBrands: customBrandList,
                    supPassword: savedSupPassword,
                    servicePin: savedServicePin,
                    gasWebAppUrl: gasWebAppUrl,
                    lineChannelToken: lineChannelToken,
                    lineTargetId: lineTargetId,
                    updatedBy: reason || 'WEB_USER'
                };
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 6000);
                await fetch(CLOUD_SYNC_API_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                    signal: controller.signal
                });
                clearTimeout(timeoutId);
                const timeEl = document.getElementById('online-sync-time');
                if (timeEl) timeEl.textContent = '(อัปเดตออนไลน์ล่าสุด: ' + new Date().toLocaleTimeString('th-TH') + ')';
            } catch (e) {
                const timeEl = document.getElementById('online-sync-time');
                if (timeEl) timeEl.textContent = '(ทำงานแบบออฟไลน์ - บันทึกในเครื่องแล้ว)';
            }
        }

        async function pullFromCloudSyncServer(showMsg) {
            try {
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 6000);
                const res = await fetch(CLOUD_SYNC_API_URL, { signal: controller.signal });
                clearTimeout(timeoutId);
                if (!res.ok) return false;
                const json = await res.json();
                const d = json && (json.tickets ? json : (json.data || json));
                if (d && Array.isArray(d.tickets) && d.tickets.length > 0) {
                    const serverTickets = d.tickets.filter(t => !isDemoTicket(t));
                    const serverIds = new Set(serverTickets.map(t => t.id));
                    const localOnly = (state.tickets || []).filter(t => !serverIds.has(t.id) && !isDemoTicket(t));
                    state.tickets = [...serverTickets, ...localOnly];
                    if (Array.isArray(d.technicians) && d.technicians.length > 0) {
                        state.technicians = d.technicians.filter(tc => !isDemoTechnician(tc));
                    }
                    if (Array.isArray(d.supervisors) && d.supervisors.length > 0) state.supervisors = d.supervisors;
                    if (Array.isArray(d.masterBrands) && d.masterBrands.length > 0) masterBrands = d.masterBrands;
                    if (Array.isArray(d.masterEquipmentTypes) && d.masterEquipmentTypes.length > 0) masterEquipmentTypes = d.masterEquipmentTypes;
                    if (d.masterChecklists && typeof d.masterChecklists === 'object') Object.assign(CHECKLIST_TEMPLATES, d.masterChecklists);
                    if (d.companyProfile && typeof d.companyProfile === 'object') companyProfile = { ...companyProfile, ...d.companyProfile };
                    
                    persistState();
                    render();
                    const timeEl = document.getElementById('online-sync-time');
                    if (timeEl) timeEl.textContent = '(ซิงค์ออนไลน์ล่าสุด: ' + new Date().toLocaleTimeString('th-TH') + ')';
                    if (showMsg) showAppToast('ซิงค์และอัปเดตข้อมูลออนไลน์ล่าสุดสำเร็จ! (ใบงาน ' + state.tickets.length + ' รายการ)', 'success');
                    return true;
                }
            } catch (e) {
                console.warn('Pull cloud sync notice:', e);
            }
            return false;
        }

        async function syncOnlineNow(showMsg) {
            await pushToCloudSyncServer('ONE_CLICK_SYNC');
            if (gasWebAppUrl) await syncToGoogleSheet();
            await pullFromCloudSyncServer(false);
            render();
            if (showMsg) {
                showAppToast('อัปเดตและซิงค์ข้อมูลออนไลน์ไปยังทุกเครื่องเรียบร้อยแล้ว! (ใบงาน ' + state.tickets.length + ' รายการ · ช่าง ' + state.technicians.length + ' คน)', 'success');
            }
        }

        async function exportFullSystemDataStandalone() {
            if (state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 สิทธิ์เฉพาะ SUPERVISOR เท่านั้น',
                    message: 'กรุณากรอกรหัสผ่าน Supervisor เพื่อสำรองไฟล์ข้อมูลระบบ:',
                    inputType: 'password',
                    placeholder: 'รหัสเริ่มต้นคือ admin123'
                });
                if (pass === savedSupPassword || pass === 'admin123') {
                    // authenticated
                } else {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    return;
                }
            }
            const payload = {
                version: '12.1-ONLINE',
                exportedAt: new Date().toISOString(),
                tickets: state.tickets,
                technicians: state.technicians,
                supervisors: state.supervisors,
                masterBrands: masterBrands,
                masterEquipmentTypes: masterEquipmentTypes,
                masterChecklists: CHECKLIST_TEMPLATES,
                companyProfile: companyProfile
            };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'SahaKim_Online_Data_' + new Date().toISOString().slice(0, 10) + '.json';
            a.click();
            showAppToast('สำรองไฟล์ข้อมูลระบบเรียบร้อยแล้ว', 'success');
        }

        async function importFullSystemDataStandalone(event) {
            if (state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 สิทธิ์เฉพาะ SUPERVISOR เท่านั้น',
                    message: 'กรุณากรอกรหัสผ่าน Supervisor เพื่อนำเข้าไฟล์ข้อมูล:',
                    inputType: 'password',
                    placeholder: 'รหัสเริ่มต้นคือ admin123'
                });
                if (pass !== savedSupPassword && pass !== 'admin123') {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    if (event && event.target) event.target.value = '';
                    return;
                }
            }
            const file = event.target.files && event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const parsed = JSON.parse(reader.result);
                    if (Array.isArray(parsed.tickets)) state.tickets = parsed.tickets.filter(t => !isDemoTicket(t));
                    if (Array.isArray(parsed.technicians)) state.technicians = parsed.technicians.filter(tc => !isDemoTechnician(tc));
                    if (Array.isArray(parsed.supervisors)) state.supervisors = parsed.supervisors;
                    if (Array.isArray(parsed.masterBrands)) masterBrands = parsed.masterBrands;
                    if (Array.isArray(parsed.masterEquipmentTypes)) masterEquipmentTypes = parsed.masterEquipmentTypes;
                    if (parsed.masterChecklists) Object.assign(CHECKLIST_TEMPLATES, parsed.masterChecklists);
                    if (parsed.companyProfile) companyProfile = { ...companyProfile, ...parsed.companyProfile };
                    persistState();
                    render();
                    showAppToast('นำเข้าไฟล์ข้อมูลและอัปเดตออนไลน์เรียบร้อยแล้ว!', 'success');
                } catch (e) {
                    showAppToast('ไฟล์ข้อมูลไม่ถูกต้อง กรุณาเลือกไฟล์ .json ที่ส่งออกจากระบบ', 'error');
                }
            };
            reader.readAsText(file);
        }

        async function openMasterDataModal(activeTab) {
            if (state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 สิทธิ์เฉพาะ SUPERVISOR เท่านั้น',
                    message: 'กรุณากรอกรหัสผ่าน Supervisor เพื่อจัดการข้อมูลตั้งต้น:',
                    inputType: 'password',
                    placeholder: 'รหัสเริ่มต้นคือ admin123'
                });
                if (pass === savedSupPassword || pass === 'admin123') {
                    state.userRole = 'SUPERVISOR';
                    render();
                } else {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    return;
                }
            }
            const tab = activeTab || 'brands';
            const existing = document.getElementById('master-data-modal');
            if (existing) existing.remove();
            const modalBg = document.createElement('div');
            modalBg.id = 'master-data-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/75 z-50 flex items-center justify-center p-3 overflow-y-auto no-print';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-3xl rounded-2xl shadow-2xl p-5 space-y-4 max-h-[92vh] overflow-y-auto text-xs sm:text-sm">
                    <div class="flex justify-between items-center border-b pb-3">
                        <div>
                            <span class="bg-emerald-100 text-emerald-800 text-[10px] font-bold px-2.5 py-0.5 rounded-full">MASTER DATA & SYSTEM CONFIGURATION</span>
                            <h3 class="font-bold text-base sm:text-lg text-slate-900 mt-0.5">⚙️ แก้ไขข้อมูลตั้งต้น (ยี่ห้อ · ประเภทเครื่องจักร · ข้อมูลบริษัท · ช่าง)</h3>
                        </div>
                        <button onclick="document.getElementById('master-data-modal').remove()" class="bg-slate-900 text-white px-3 py-1 rounded-lg font-bold cursor-pointer">✕ ปิด</button>
                    </div>
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
                        <button onclick="openMasterDataModal('brands')" class="p-2.5 rounded-xl border font-bold cursor-pointer \${tab === 'brands' ? 'bg-orange-500 text-white border-orange-600' : 'bg-slate-50 text-slate-800'}">🏷️ 1. ยี่ห้อ (\${masterBrands.length})</button>
                        <button onclick="openMasterDataModal('equipment')" class="p-2.5 rounded-xl border font-bold cursor-pointer \${tab === 'equipment' ? 'bg-blue-600 text-white border-blue-700' : 'bg-slate-50 text-slate-800'}">⚙️ 2. เครื่องจักร (\${masterEquipmentTypes.length})</button>
                        <button onclick="openMasterDataModal('company')" class="p-2.5 rounded-xl border font-bold cursor-pointer \${tab === 'company' ? 'bg-indigo-600 text-white border-indigo-700' : 'bg-slate-50 text-slate-800'}">🏢 3. ชื่อ/เบอร์บริษัท</button>
                        <button onclick="document.getElementById('master-data-modal').remove(); openTechnicianManager();" class="p-2.5 rounded-xl border font-bold cursor-pointer bg-slate-900 text-white">👷 4. จัดการช่าง (\${state.technicians.length})</button>
                    </div>
                    \${tab === 'brands' ? \`
                        <div class="space-y-3 bg-slate-50 p-4 rounded-xl border">
                            <div class="flex gap-2">
                                <input type="text" id="new-master-brand-input" placeholder="พิมพ์ชื่อยี่ห้อใหม่ (เช่น KSB, TSURUMI, WEG)..." class="flex-1 border rounded-lg px-3 py-2 bg-white font-bold">
                                <button onclick="const v = document.getElementById('new-master-brand-input').value.trim().toUpperCase(); if(v && !masterBrands.includes(v)){ masterBrands.push(v); persistState(); openMasterDataModal('brands'); render(); }" class="bg-orange-500 text-white px-4 py-2 rounded-lg font-bold cursor-pointer">+ เพิ่มยี่ห้อ</button>
                            </div>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 max-h-72 overflow-y-auto">
                                \${masterBrands.map((b, idx) => \`
                                    <div class="flex items-center gap-1.5 bg-white p-2 rounded-lg border">
                                        <span class="font-mono text-slate-400 w-6">\${idx + 1}.</span>
                                        <input type="text" value="\${b}" onchange="masterBrands[\${idx}] = this.value.trim().toUpperCase(); persistState(); render();" class="flex-1 border rounded px-2 py-1 font-bold">
                                        <button onclick="masterBrands.splice(\${idx}, 1); persistState(); openMasterDataModal('brands'); render();" class="text-red-500 font-bold px-2 cursor-pointer">ลบ</button>
                                    </div>
                                \`).join('')}
                            </div>
                        </div>
                    \` : ''}
                    \${tab === 'equipment' ? \`
                        <div class="space-y-3 bg-slate-50 p-4 rounded-xl border">
                            <div class="flex gap-2">
                                <input type="text" id="new-master-eq-input" placeholder="พิมพ์ประเภทเครื่องจักรใหม่ (เช่น SPLIT CASE PUMP)..." class="flex-1 border rounded-lg px-3 py-2 bg-white font-bold">
                                <button onclick="const v = document.getElementById('new-master-eq-input').value.trim().toUpperCase(); if(v && !masterEquipmentTypes.includes(v)){ masterEquipmentTypes.push(v); persistState(); openMasterDataModal('equipment'); render(); }" class="bg-blue-600 text-white px-4 py-2 rounded-lg font-bold cursor-pointer">+ เพิ่มประเภทเครื่องจักร</button>
                            </div>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 max-h-72 overflow-y-auto">
                                \${masterEquipmentTypes.map((eq, idx) => \`
                                    <div class="flex items-center gap-1.5 bg-white p-2 rounded-lg border">
                                        <span class="font-mono text-slate-400 w-6">\${idx + 1}.</span>
                                        <input type="text" value="\${eq}" onchange="masterEquipmentTypes[\${idx}] = this.value.trim().toUpperCase(); persistState(); render();" class="flex-1 border rounded px-2 py-1 font-bold">
                                        <button onclick="masterEquipmentTypes.splice(\${idx}, 1); persistState(); openMasterDataModal('equipment'); render();" class="text-red-500 font-bold px-2 cursor-pointer">ลบ</button>
                                    </div>
                                \`).join('')}
                            </div>
                        </div>
                    \` : ''}
                    \${tab === 'company' ? \`
                        <div class="space-y-3 bg-slate-50 p-4 rounded-xl border">
                            <div><label class="block font-bold mb-1">ชื่อบริษัท (ภาษาไทย)</label><input type="text" id="comp-th-input" value="\${companyProfile.nameTh}" class="w-full border rounded-lg p-2 bg-white font-bold"></div>
                            <div><label class="block font-bold mb-1">ชื่อบริษัท (ภาษาอังกฤษ)</label><input type="text" id="comp-en-input" value="\${companyProfile.nameEn}" class="w-full border rounded-lg p-2 bg-white font-bold"></div>
                            <div><label class="block font-bold mb-1">เบอร์โทรศัพท์ / ช่องทางติดต่อ</label><input type="text" id="comp-phone-input" value="\${companyProfile.phone}" class="w-full border rounded-lg p-2 bg-white font-bold"></div>
                            <button onclick="companyProfile.nameTh = document.getElementById('comp-th-input').value; companyProfile.nameEn = document.getElementById('comp-en-input').value; companyProfile.phone = document.getElementById('comp-phone-input').value; persistState(); render(); showAppToast('บันทึกข้อมูลบริษัทเรียบร้อย!', 'success');" class="bg-indigo-600 text-white px-4 py-2 rounded-xl font-bold cursor-pointer">บันทึกข้อมูลบริษัท</button>
                        </div>
                    \` : ''}
                    <div class="flex justify-end gap-2 pt-2 border-t">
                        <button onclick="syncOnlineNow(true); document.getElementById('master-data-modal').remove();" class="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-xl font-bold cursor-pointer">✓ บันทึกและอัปเดตออนไลน์</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function persistState(lineNotifyText) {
            // 1. Asynchronously persist full-fidelity state to IndexedDB (No 5MB storage limit)
            const fullSnapshot = {
                tickets: state.tickets,
                technicians: state.technicians,
                supervisors: state.supervisors,
                supPassword: savedSupPassword,
                servicePin: savedServicePin,
                knowledgeTemplates: knowledgeTemplates,
                customEquipTypes: customEquipList,
                customBrands: customBrandList,
                masterBrands: masterBrands,
                masterEquipmentTypes: masterEquipmentTypes,
                companyProfile: companyProfile,
                savedAt: new Date().toISOString()
            };
            saveToIndexedDB('full_state', fullSnapshot).catch(e => console.warn('IDB save warning:', e));

            // 2. Safely persist to localStorage with QuotaExceededError protection
            try {
                localStorage.setItem('sahakim_tickets_v121', JSON.stringify(state.tickets));
                localStorage.setItem('sahakim_technicians', JSON.stringify(state.technicians));
                localStorage.setItem('sahakim_supervisors', JSON.stringify(state.supervisors));
                localStorage.setItem('sahakim_sup_pass', savedSupPassword);
                localStorage.setItem('sahakim_service_pin', savedServicePin);
                localStorage.setItem('sahakim_knowledge_templates', JSON.stringify(knowledgeTemplates));
                localStorage.setItem('sahakim_custom_equip_types', JSON.stringify(customEquipList));
                localStorage.setItem('sahakim_custom_brands', JSON.stringify(customBrandList));
                localStorage.setItem('sahakim_master_brands_v1', JSON.stringify(masterBrands));
                localStorage.setItem('sahakim_master_equipment_types_v1', JSON.stringify(masterEquipmentTypes));
                localStorage.setItem('sahakim_company_profile_v1', JSON.stringify(companyProfile));
            } catch (quotaError) {
                console.warn('LocalStorage quota reached, creating lightweight fallback in localStorage (IndexedDB retains 100% full photos):', quotaError);
                try {
                    // Truncate heavy base64 strings so localStorage retains ticket IDs, customer names, dates, and statuses
                    const lightTickets = state.tickets.map(t => ({
                        ...t,
                        equipmentList: (t.equipmentList || []).map(eq => ({
                            ...eq,
                            beforePhotos: (eq.beforePhotos || []).map(p => (typeof p === 'string' && p.length > 500) ? 'data:image-saved-in-indexeddb' : p),
                            afterPhotos: (eq.afterPhotos || []).map(p => (typeof p === 'string' && p.length > 500) ? 'data:image-saved-in-indexeddb' : p)
                        }))
                    }));
                    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(lightTickets));
                } catch(e2) {}
            }

            syncToGoogleSheet(lineNotifyText);
            pushToCloudSyncServer('STATE_UPDATED');
        }

        async function openGoogleSheetModal() {
            if (state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 สิทธิ์เฉพาะ SUPERVISOR เท่านั้น',
                    message: 'กรุณากรอกรหัสผ่าน Supervisor เพื่อตั้งค่าการซิงค์:',
                    inputType: 'password',
                    placeholder: 'รหัสเริ่มต้นคือ admin123'
                });
                if (pass === savedSupPassword || pass === 'admin123') {
                    state.userRole = 'SUPERVISOR';
                    render();
                } else {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    return;
                }
            }
            let modalBg = document.createElement('div');
            modalBg.id = 'gsheet-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-3 overflow-y-auto';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-2xl rounded-2xl shadow-2xl p-5 space-y-4 max-h-[92vh] overflow-y-auto text-xs sm:text-sm">
                    <div class="flex justify-between items-center border-b pb-2">
                        <h3 class="font-bold text-base text-slate-900"><i class="fa-solid fa-table text-emerald-600 mr-1.5 text-lg"></i> ตั้งค่าฐานข้อมูล Google Sheets (Real-Time Cloud Sync)</h3>
                        <button onclick="document.getElementById('gsheet-modal').remove()" class="text-xl font-bold cursor-pointer">&times;</button>
                    </div>

                    <div class="bg-emerald-50 border border-emerald-200 p-3 rounded-xl space-y-1 text-xs text-emerald-900">
                        <p class="font-bold">📌 วิธีเชื่อมต่อให้ข้อมูลอัปเดตลง Google Sheet อัตโนมัติ 100%:</p>
                        <p>1. ใน Google Sheet ไปที่เมนู <b>ส่วนขยาย (Extensions) &gt; Apps Script</b> ลบโค้ดเก่าแล้ววางโค้ดด้านล่างลงไป กดบันทึก 💾</p>
                        <p>2. กดปุ่มสีน้ำเงินมุมขวาบน <b>"การทำให้ใช้งานได้ (Deploy)" &gt; "การทำให้ใช้งานได้รายการใหม่ (New deployment)"</b> เลือกประเภท <b>เว็บแอป (Web app)</b> และตั้งค่าผู้มีสิทธิ์เข้าถึงเป็น <b>"ทุกคน (Anyone)"</b></p>
                        <p>3. คัดลอก <b>Web App URL (ที่ลงท้ายด้วย /exec)</b> มาวางในช่องขั้นตอนที่ 2 ด้านล่าง แล้วกดปุ่มสีส้มบันทึก</p>
                    </div>

                    <div class="space-y-2">
                        <label class="block font-bold text-slate-800">ขั้นตอนที่ 1: คัดลอกโค้ด Apps Script สำหรับ Google Sheets ด้านล่างไปวางใน Apps Script</label>
                        <textarea id="gas-code-box" readonly rows="5" class="w-full border rounded-xl p-2 font-mono text-[11px] bg-slate-900 text-emerald-300"></textarea>
                        <button type="button" onclick="navigator.clipboard.writeText(GAS_CODE_TEMPLATE); alert('คัดลอกโค้ด Google Apps Script เรียบร้อย! นำไปวางใน Google Sheet > ส่วนขยาย > Apps Script แล้วกด Deploy เป็นรายการใหม่ได้เลย');" class="bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-copy mr-1"></i> คัดลอกโค้ด Apps Script</button>
                    </div>

                    <div class="space-y-2 border-t pt-3">
                        <label class="block font-bold text-slate-800">ขั้นตอนที่ 2: วางลิงก์ Google Apps Script Web App URL (ต้องขึ้นต้นด้วย https://script.google.com/... และลงท้ายด้วย /exec)</label>
                        <div class="flex gap-2 items-center">
                            <input type="url" id="gas-url-input" value="${gasWebAppUrl}" placeholder="https://script.google.com/macros/s/......./exec" class="flex-1 border rounded-xl p-2 font-mono text-xs">
                            <button type="button" onclick="clearGasUrl()" title="ลบลิงก์เดิมเพื่อใส่ลิงก์ใหม่" class="bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 px-3 py-2 rounded-xl text-xs font-bold cursor-pointer whitespace-nowrap"><i class="fa-solid fa-trash mr-1"></i> ล้างลิงก์เดิม</button>
                        </div>
                        <div class="flex flex-wrap gap-2 pt-1">
                            <button type="button" onclick="saveGasUrlAndSync(false)" class="bg-orange-500 hover:bg-orange-600 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-cloud-arrow-up mr-1"></i> อัปเดตข้อมูลปัจจุบันขึ้น Google Sheet ทันที</button>
                            <button type="button" onclick="loadFromGoogleSheet(true)" class="bg-slate-800 hover:bg-slate-900 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-cloud-arrow-down mr-1"></i> ดึงข้อมูลจาก Google Sheet ลงเว็บ</button>
                        </div>
                    </div>

                    <div class="flex flex-wrap justify-end gap-2 pt-2 border-t">
                        <button type="button" onclick="document.getElementById('gsheet-modal').remove()" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer">ปิด</button>
                        <button type="button" onclick="saveGasUrlAndSync(true)" class="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">บันทึกการตั้งค่าและซิงก์ลง Sheet</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
            document.getElementById('gas-code-box').value = GAS_CODE_TEMPLATE;
        }

        async function saveGasUrlAndSync(closeAfter) {
            const val = document.getElementById('gas-url-input').value.trim();
            if (val && val.includes('docs.google.com/spreadsheets') && !val.includes('script.google.com')) {
                alert('⚠️ ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet (docs.google.com)\\n\\nกรุณาใช้ลิงก์ Web App URL จากเมนู ส่วนขยาย > Apps Script > การทำให้ใช้งานได้ (Deploy) ที่ลงท้ายด้วย /exec แทนครับ');
                return;
            }
            gasWebAppUrl = val;
            lineChannelToken = '';
            lineTargetId = '';
            localStorage.setItem('sahakim_gas_url', gasWebAppUrl);
            localStorage.removeItem('sahakim_line_token');
            localStorage.removeItem('sahakim_line_target');
            updateGsheetBadge();
            if (gasWebAppUrl) {
                await syncToGoogleSheet();
                alert('✅ บันทึกลิงก์และส่งข้อมูลใบงานทั้งหมด (' + state.tickets.length + ' ใบงาน) ไปยังแท็บ Tickets ใน Google Sheet เรียบร้อยแล้ว!');
            } else {
                alert('⚠️ บันทึกการตั้งค่าแล้ว (ยังไม่มีการเชื่อมต่อ Google Sheets)');
            }
            if (closeAfter) {
                const m = document.getElementById('gsheet-modal');
                if (m) m.remove();
            }
        }

        async function clearGasUrl() {
            const confirmed = await openCustomConfirm({
                title: '🗑️ ล้างลิงก์ Web App เดิม',
                message: 'คุณต้องการลบลิงก์ Web App เดิมออก เพื่อตัดการเชื่อมต่อหรือวางลิงก์ใหม่ใช่หรือไม่?',
                confirmText: 'ล้างลิงก์',
                cancelText: 'ยกเลิก'
            });
            if (confirmed) {
                const inp = document.getElementById('gas-url-input');
                if (inp) inp.value = '';
                gasWebAppUrl = '';
                localStorage.removeItem('sahakim_gas_url');
                updateGsheetBadge();
                showAppToast('ล้างลิงก์เดิมเรียบร้อยแล้ว สามารถวางลิงก์ใหม่ได้เลย', 'info');
            }
        }

        function openQrShareModal() {
            const currentUrl = window.location.href.startsWith('http') && !window.location.href.includes('localhost')
                ? window.location.href
                : 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
            const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=280x280&margin=10&data=' + encodeURIComponent(currentUrl);
            let modalBg = document.createElement('div');
            modalBg.id = 'qr-modal';
            modalBg.className = 'print-modal-root fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-3 overflow-y-auto';
            modalBg.innerHTML = \`
                <div id="qr-card-print-area" class="print-modal-content bg-white w-full max-w-md rounded-2xl shadow-2xl p-6 space-y-4 text-center text-xs sm:text-sm">
                    <div class="flex justify-between items-center border-b pb-2 no-print">
                        <h3 class="font-bold text-base text-slate-900"><i class="fa-solid fa-qrcode text-orange-500 mr-2"></i> QR Code & ลิงก์ระบบแจ้งซ่อมออนไลน์</h3>
                        <button onclick="document.getElementById('qr-modal').remove()" class="text-xl font-bold cursor-pointer no-print">&times;</button>
                    </div>
                    <div class="border-2 border-slate-900 rounded-2xl p-5 bg-slate-50 space-y-2">
                        <p class="text-[11px] font-bold text-orange-600 uppercase tracking-wider">SAHA KIM MOTOR CO., LTD.</p>
                        <h4 class="text-base font-extrabold text-slate-900">สแกน QR Code เพื่อแจ้งซ่อม / ติดตามงาน</h4>
                        <div class="bg-white p-3 rounded-xl border inline-block mx-auto shadow-sm">
                            <img crossorigin="anonymous" src="\${qrApiUrl}" alt="QR Code Saha Kim Motor" class="w-48 h-48 mx-auto" />
                        </div>
                        <p class="text-xs font-bold text-slate-800 pt-1">บริษัท สหคิมมอเตอร์ จำกัด</p>
                        <p class="text-[11px] text-slate-500 break-all">\${currentUrl}</p>
                    </div>
                    <div class="flex flex-wrap justify-center gap-2 pt-2 no-print">
                        <button type="button" onclick="navigator.clipboard.writeText('\${currentUrl}'); alert('คัดลอกลิงก์เรียบร้อย! ส่งเข้า LINE ให้ลูกค้าหรือทีมงานได้เลย');" class="bg-slate-900 text-white px-4 py-2 rounded-xl font-bold cursor-pointer"><i class="fa-solid fa-link mr-1"></i> คัดลอกลิงก์</button>
                        <button type="button" onclick="downloadPdfById('qr-card-print-area', 'Sahakim-QR-Service-Sticker.pdf')" class="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-xl font-bold cursor-pointer"><i class="fa-solid fa-file-pdf mr-1"></i> Save PDF (ป้าย QR)</button>
                        <button type="button" onclick="window.print()" class="bg-slate-200 text-slate-800 px-3.5 py-2 rounded-xl font-bold cursor-pointer"><i class="fa-solid fa-print mr-1"></i> สั่งพิมพ์</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function playAlertSound() {
            try {
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                if (!AudioCtx) return;
                const ctx = new AudioCtx();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.type = 'sine';
                osc.frequency.setValueAtTime(880, ctx.currentTime);
                gain.gain.setValueAtTime(0.06, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.3);
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.3);
            } catch(e) {}
        }

        function exportToExcel() {
            let csv = '\\uFEFFTicketID,Company,Contact,ServiceType,Warranty,Urgency,Status,AssignedTech,CreatedAt\\n';
            state.tickets.forEach(t => {
                csv += \`"\${t.id}","\${t.company}","\${t.contact}","\${t.serviceType}","\${t.warranty}","\${t.urgency}","\${t.status}","\${t.assignedTech || ''}","\${t.createdAt}"\\n\`;
            });
            let blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            let url = URL.createObjectURL(blob);
            let a = document.createElement('a');
            a.href = url;
            a.download = \`SahaKim_Service_Report_\${new Date().toISOString().slice(0,10)}.csv\`;
            a.click();
        }

        let customerLookupQuery = '';
        let selectedCustomerDate = new Date().toISOString().slice(0, 10);
        let supportScheduleOffset = 0;
        let supportSelectedDate = new Date().toISOString().slice(0, 10);

        function getActionableStatusButton(t) {
            let status = t.status;
            const isSupportOrSup = state.userRole === 'SERVICE TEAM' || state.userRole === 'SUPERVISOR';
            if (status === 'NEW') {
                return \`<div class="flex flex-wrap gap-1.5 items-center">
                    <button onclick="openAssignDispatchModal('\${t.id}')" class="bg-blue-600 hover:bg-blue-700 text-white text-xs px-3 py-1.5 rounded-xl font-bold shadow flex items-center gap-1 cursor-pointer"><i class="fa-solid fa-calendar-check"></i> NEW: เลือกวัน/เลือกช่างตามคิวว่าง</button>
                    \${isSupportOrSup ? \`<button onclick="supportAcceptOnBehalf('\${t.id}')" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-xl font-bold shadow cursor-pointer">✅ รับงานแทนช่างเลย</button>\` : ''}
                </div>\`;
            } else if (status === 'ASSIGNED') {
                return \`<div class="flex flex-wrap gap-1.5 items-center">
                    <span class="bg-amber-100 text-amber-900 border border-amber-300 text-xs px-2.5 py-1 rounded-lg font-bold">⏳ รอช่างตอบรับงาน (\${t.assignedTech || 'ยังไม่ระบุช่าง'})</span>
                    \${state.userRole === 'TECHNICIAN' ? \`<button onclick="techAcceptJobStandalone('\${t.id}')" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-bold shadow cursor-pointer">👷 ช่างกดตอบรับงาน</button>\` : ''}
                    \${isSupportOrSup ? \`<button onclick="supportAcceptOnBehalf('\${t.id}')" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-bold shadow cursor-pointer">✅ รับงานแทนช่างเลย</button><button onclick="openAssignDispatchModal('\${t.id}')" class="bg-blue-600 hover:bg-blue-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-bold shadow cursor-pointer">🔄 เปลี่ยนช่างใหม่/เปลี่ยนวัน</button>\` : ''}
                </div>\`;
            } else if (status === 'INSPECTION' || status === 'REPAIRING') {
                return \`<div class="flex flex-wrap gap-1.5 items-center">
                    <span class="bg-emerald-100 text-emerald-900 border border-emerald-300 text-xs px-2.5 py-1 rounded-lg font-bold">🔧 กำลังดำเนินการซ่อม (\${t.assignedTech || 'ช่างเทคนิค'})</span>
                    \${isSupportOrSup && !t.isLocked ? \`<button onclick="openAssignDispatchModal('\${t.id}')" class="bg-slate-800 hover:bg-slate-900 text-white text-xs px-2.5 py-1 rounded-lg font-bold cursor-pointer">📅 เปลี่ยนวัน/เปลี่ยนช่าง</button>\` : ''}
                </div>\`;
            } else if (status === 'PENDING_PARTS_APPROVAL') {
                return \`<div class="flex flex-wrap gap-2 items-center"><span class="bg-rose-100 text-rose-700 text-xs px-3 py-1 rounded-full font-bold">⏳ รออนุมัติอะไหล่</span>
                    \${state.userRole === 'SUPERVISOR' ? \`<button onclick="supervisorApproveParts('\${t.id}')" class="bg-emerald-600 text-white text-xs px-2.5 py-1 rounded-lg font-bold cursor-pointer">✔ อนุมัติอะไหล่</button><button onclick="supervisorReject('\${t.id}')" class="bg-red-600 text-white text-xs px-2.5 py-1 rounded-lg font-bold cursor-pointer">❌ ไม่อนุมัติ</button>\` : ''}</div>\`;
            } else if (status === 'PENDING_SUPERVISOR_APPROVAL') {
                return \`<div class="flex flex-wrap gap-2 items-center"><span class="bg-purple-100 text-purple-700 text-xs px-3 py-1 rounded-full font-bold">✍️ รออนุมัติรายงาน</span>
                    \${state.userRole === 'SUPERVISOR' ? \`<button onclick="supervisorApproveReport('\${t.id}')" class="bg-emerald-600 text-white text-xs px-2.5 py-1 rounded-lg font-bold cursor-pointer">✔ อนุมัติรายงาน</button><button onclick="supervisorReject('\${t.id}')" class="bg-red-600 text-white text-xs px-2.5 py-1 rounded-lg font-bold cursor-pointer">❌ ส่งกลับ</button>\` : ''}</div>\`;
            } else if (status === 'REJECTED') {
                return \`<span class="bg-red-100 text-red-700 text-xs px-3 py-1 rounded-full font-bold">❌ ส่งกลับแก้ไข: \${t.rejectReason || 'โปรดแก้ไขข้อมูล'}</span>\`;
            } else if (status === 'CUSTOMER_ACCEPTANCE') {
                return \`<span class="bg-cyan-100 text-cyan-700 text-xs px-3 py-1 rounded-full font-bold">🤝 รอตรวจรับงานจากลูกค้า</span>\`;
            } else if (status === 'CLOSE') {
                return \`<span class="bg-emerald-100 text-emerald-700 text-xs px-3 py-1 rounded-full font-bold">✅ ปิดงานสำเร็จ (ล็อกข้อมูล)</span>\`;
            }
            return \`<span class="bg-slate-100 text-slate-700 text-xs px-3 py-1 rounded-full font-bold">\${status}</span>\`;
        }

        function buildStandaloneScheduleDays(baseOffset) {
            const thaiDays = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
            const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
            const baseToday = new Date();
            baseToday.setHours(0, 0, 0, 0);
            const list = [];
            for (let idx = 0; idx < 14; idx++) {
                const offset = (baseOffset || 0) + idx;
                const d = new Date(baseToday);
                d.setDate(baseToday.getDate() + offset);
                const yyyy = d.getFullYear();
                const mm = String(d.getMonth() + 1).padStart(2, '0');
                const dd = String(d.getDate()).padStart(2, '0');
                const dateStr = yyyy + '-' + mm + '-' + dd;
                const dayLabel = thaiDays[d.getDay()] + ' ' + d.getDate() + ' ' + thaiMonths[d.getMonth()];
                const ticketsOnDay = state.tickets.filter(t => t.status !== 'CLOSE' && extractTicketDueDate(t) === dateStr);
                const totalJobs = ticketsOnDay.length;
                const level = totalJobs >= 3 ? 'RED' : totalJobs >= 1 ? 'YELLOW' : 'GREEN';
                const techAvailability = state.technicians.map(tc => {
                    const assignedTickets = ticketsOnDay.filter(t => t.assignedTech === tc.name);
                    const jobCount = assignedTickets.length;
                    const techLevel = jobCount >= 2 ? 'RED' : jobCount === 1 ? 'YELLOW' : 'GREEN';
                    return {
                        ...tc,
                        jobCount,
                        assignedTickets,
                        techLevel,
                        availabilityLabel: techLevel === 'RED' ? ('คิวแน่น (' + jobCount + ' งาน)') : techLevel === 'YELLOW' ? ('รับงานเพิ่มได้ (มี ' + jobCount + ' งาน)') : 'ว่าง (พร้อมรับงาน)'
                    };
                });
                const freeTechCount = techAvailability.filter(tc => tc.jobCount === 0).length;
                list.push({
                    dateStr,
                    dayLabel,
                    totalJobs,
                    ticketsOnDay,
                    techAvailability,
                    freeTechCount,
                    level,
                    statusText: level === 'RED' ? 'เต็มมาก' : level === 'YELLOW' ? 'รับงานได้' : 'ว่าง',
                    detailText: level === 'RED' ? ('คิวแน่น (' + totalJobs + ' งาน)') : level === 'YELLOW' ? ('รับงานได้ (' + totalJobs + ' งาน)') : 'คิวช่างว่าง'
                });
            }
            return list;
        }

        function getCustomerScheduleDays() {
            return buildStandaloneScheduleDays(0);
        }

        function getSupportScheduleDays() {
            return buildStandaloneScheduleDays(supportScheduleOffset);
        }

        function getStandaloneDateInfo(dateStr) {
            const list = getSupportScheduleDays();
            const found = list.find(d => d.dateStr === dateStr);
            if (found) return found;
            const ticketsOnDay = state.tickets.filter(t => t.status !== 'CLOSE' && extractTicketDueDate(t) === dateStr);
            const totalJobs = ticketsOnDay.length;
            const level = totalJobs >= 3 ? 'RED' : totalJobs >= 1 ? 'YELLOW' : 'GREEN';
            const techAvailability = state.technicians.map(tc => {
                const assignedTickets = ticketsOnDay.filter(t => t.assignedTech === tc.name);
                const jobCount = assignedTickets.length;
                const techLevel = jobCount >= 2 ? 'RED' : jobCount === 1 ? 'YELLOW' : 'GREEN';
                return {
                    ...tc,
                    jobCount,
                    assignedTickets,
                    techLevel,
                    availabilityLabel: techLevel === 'RED' ? ('คิวแน่น (' + jobCount + ' งาน)') : techLevel === 'YELLOW' ? ('รับงานเพิ่มได้ (มี ' + jobCount + ' งาน)') : 'ว่าง (พร้อมรับงาน)'
                };
            });
            return {
                dateStr,
                dayLabel: dateStr,
                totalJobs,
                ticketsOnDay,
                techAvailability,
                freeTechCount: techAvailability.filter(tc => tc.jobCount === 0).length,
                level,
                statusText: level === 'RED' ? 'เต็มมาก' : level === 'YELLOW' ? 'รับงานได้' : 'ว่าง',
                detailText: level === 'RED' ? ('คิวแน่น (' + totalJobs + ' งาน)') : level === 'YELLOW' ? ('รับงานได้ (' + totalJobs + ' งาน)') : 'คิวช่างว่าง'
            };
        }

        function supportAcceptOnBehalf(id) {
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            if (t.isLocked && state.userRole !== 'SUPERVISOR') {
                alert('⚠️ ใบงานนี้ถูกล็อกแล้ว! ต้องให้ Supervisor ปลดล็อกเท่านั้น');
                return;
            }
            const dateToUse = extractTicketDueDate(t) || t.desiredDate || new Date().toISOString().slice(0, 10);
            const dateInfo = getStandaloneDateInfo(dateToUse);
            const freeTech = (dateInfo.techAvailability.find(tc => tc.jobCount === 0) || {}).name;
            const chosenTech = t.assignedTech || freeTech || (state.technicians[0] && state.technicians[0].name) || 'ช่างเทคนิค';
            const chosenLead = t.leadTech || (state.technicians.find(tc => (tc.position || '').includes('หัวหน้าช่าง')) || {}).name || '';
            t.confirmedDate = dateToUse + ' (Service Support รับงานแทนช่าง ' + chosenTech + ' แล้ว)';
            t.assignedTech = chosenTech;
            t.leadTech = chosenLead;
            t.status = 'REPAIRING';
            const lineMsg = formatTicketLineMessage(t, '✅ [Service Support รับงานแทนช่าง (' + chosenTech + ') เรียบร้อยแล้ว - เริ่มดำเนินการ]');
            persistState(lineMsg);
            render();
        }

        function techAcceptJobStandalone(id) {
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            const dateToUse = extractTicketDueDate(t) || t.desiredDate || new Date().toISOString().slice(0, 10);
            const chosenTech = t.assignedTech || (state.technicians[0] && state.technicians[0].name) || 'ช่างเทคนิค';
            t.confirmedDate = dateToUse + ' (ช่าง ' + chosenTech + ' กดตอบรับงานแล้ว)';
            t.assignedTech = chosenTech;
            t.status = 'REPAIRING';
            const lineMsg = formatTicketLineMessage(t, '👷 [ช่าง ' + chosenTech + ' กดตอบรับงานแล้ว - เริ่มดำเนินการ]');
            persistState(lineMsg);
            render();
        }

        function openAssignDispatchModal(ticketId, preferredDate, preferredTech) {
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t) return;
            if (t.isLocked && state.userRole !== 'SUPERVISOR') {
                alert('⚠️ ใบงานนี้ถูกล็อกแล้ว! ต้องให้ Supervisor ปลดล็อกเท่านั้น');
                return;
            }
            const existing = document.getElementById('dispatch-modal');
            if (existing) existing.remove();

            const targetDate = preferredDate || extractTicketDueDate(t) || t.desiredDate || new Date().toISOString().slice(0, 10);
            const dateInfo = getStandaloneDateInfo(targetDate);
            const firstFree = (dateInfo.techAvailability.find(tc => tc.jobCount === 0) || {}).name;
            const currentTech = preferredTech || t.assignedTech || firstFree || (state.technicians[0] && state.technicians[0].name) || '';
            const defaultLead = t.leadTech || (state.technicians.find(tc => (tc.position || '').includes('หัวหน้าช่าง')) || {}).name || '';

            const modalBg = document.createElement('div');
            modalBg.id = 'dispatch-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 overflow-y-auto';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-lg rounded-2xl shadow-2xl p-5 space-y-4 text-xs sm:text-sm">
                    <div class="flex justify-between items-center border-b pb-2">
                        <div>
                            <p class="text-[10px] font-bold text-orange-600 uppercase">SERVICE SUPPORT · DISPATCH & SCHEDULE</p>
                            <h3 class="font-bold text-base text-slate-900">กำหนด/เปลี่ยนวันรับงานช่าง และเลือกช่างตามคิวว่าง (\${t.id})</h3>
                        </div>
                        <button onclick="document.getElementById('dispatch-modal').remove()" class="text-xl font-bold cursor-pointer">&times;</button>
                    </div>
                    <div class="space-y-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">1. กำหนดหรือเปลี่ยนวันที่รับงานช่าง (YYYY-MM-DD) *</label>
                            <input type="date" id="dispatch-date-input" value="\${targetDate}" onchange="openAssignDispatchModal('\${t.id}', this.value)" class="w-full border rounded-xl p-2 font-mono font-bold">
                            <p class="text-[11px] mt-1 font-bold \${dateInfo.level === 'RED' ? 'text-red-600' : dateInfo.level === 'YELLOW' ? 'text-amber-600' : 'text-emerald-600'}">
                                สถานะคิววันที่เลือก (\${targetDate}): \${dateInfo.statusText} (\${dateInfo.detailText}) · ช่างว่าง \${dateInfo.freeTechCount} คน
                            </p>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">2. เลือกช่างตามความเหมาะสมและคิวว่างในวันที่ \${targetDate} *</label>
                            <select id="dispatch-tech-input" class="w-full border rounded-xl p-2 bg-white font-bold">
                                \${dateInfo.techAvailability.map(tc => \`
                                    <option value="\${tc.name}" \${tc.name === currentTech ? 'selected' : ''}>
                                        \${tc.techLevel === 'GREEN' ? '🟢 [ว่าง] ' : tc.techLevel === 'YELLOW' ? '🟡 [มี ' + tc.jobCount + ' งาน] ' : '🔴 [คิวแน่น ' + tc.jobCount + ' งาน] '}\${tc.name} [\${tc.position || 'ช่างเทคนิค'}] (\${tc.specialization})
                                    </option>
                                \`).join('')}
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">3. เลือกหัวหน้าช่างผู้ควบคุมงาน (Lead Technician)</label>
                            <input type="text" id="dispatch-lead-input" value="\${defaultLead}" placeholder="ระบุชื่อหัวหน้าช่าง (ถ้ามี)" class="w-full border rounded-xl p-2">
                        </div>
                        <div class="bg-emerald-50 border border-emerald-200 p-3 rounded-xl text-xs text-emerald-900">
                            <b>💡 สิทธิ์ของ Service Support:</b> หากช่างยังไม่กดตอบรับงาน สามารถกดปุ่มสีเขียว <b>"มอบหมาย & รับงานแทนช่างเลย"</b> เพื่อรับงานแทนช่างทันที หรือเปลี่ยนช่างคนใหม่ได้ตลอดเวลา
                        </div>
                        <div class="flex flex-wrap justify-end gap-2 pt-2 border-t">
                            <button type="button" onclick="document.getElementById('dispatch-modal').remove()" class="bg-slate-200 text-slate-700 px-3.5 py-2 rounded-xl font-bold cursor-pointer">ยกเลิก</button>
                            <button type="button" onclick="confirmDispatchAssign('\${t.id}', false)" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl font-bold cursor-pointer">บันทึกวันนัด / มอบหมายช่าง</button>
                            <button type="button" onclick="confirmDispatchAssign('\${t.id}', true)" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl font-bold cursor-pointer">✅ มอบหมาย & รับงานแทนช่างเลย</button>
                        </div>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function confirmDispatchAssign(ticketId, acceptImmediately) {
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t) return;
            const dateVal = document.getElementById('dispatch-date-input').value;
            const techVal = document.getElementById('dispatch-tech-input').value;
            const leadVal = document.getElementById('dispatch-lead-input').value.trim();
            if (!dateVal || !techVal) return;

            t.desiredDate = dateVal;
            t.assignedTech = techVal;
            t.leadTech = leadVal;
            if (acceptImmediately) {
                t.confirmedDate = dateVal + ' (Service Support รับงานแทนช่าง ' + techVal + ' แล้ว)';
                t.status = 'REPAIRING';
            } else {
                t.confirmedDate = dateVal + ' (ยืนยันโดย Service Support - รอช่างตอบรับ)';
                t.status = 'ASSIGNED';
            }
            supportSelectedDate = dateVal;
            const m = document.getElementById('dispatch-modal');
            if (m) m.remove();
            const lineMsg = formatTicketLineMessage(
                t,
                acceptImmediately
                    ? '✅ [Service Support กำหนดวันนัดและรับงานแทนช่าง (' + techVal + ') ทันที]'
                    : '👷 [กำหนดวันนัดหมายและมอบหมายช่างแล้ว - รอช่างตอบรับ]'
            );
            persistState(lineMsg);
            render();
        }

        function updateTicketStatus(id, newStatus) {
            if (newStatus === 'ASSIGNED') {
                openAssignDispatchModal(id);
                return;
            }
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            if (t.isLocked && state.userRole !== 'SUPERVISOR') {
                alert('⚠️ ใบงานนี้ถูกล็อกแล้ว! ต้องให้ Supervisor ปลดล็อกเท่านั้น');
                return;
            }
            t.status = newStatus;
            persistState();
            render();
        }

        function applyStandaloneCustomerLookup(rawVal) {
            const trimmed = (rawVal || '').trim();
            if (!trimmed) return;
            customerLookupQuery = trimmed;
            const byId = state.tickets.find(t => t.id.toLowerCase() === trimmed.toLowerCase() || t.id.toLowerCase().includes(trimmed.toLowerCase()));
            if (byId) {
                state.customerCompany = byId.company;
            } else {
                const byComp = state.tickets.find(t => (t.company || '').toLowerCase().includes(trimmed.toLowerCase()));
                state.customerCompany = byComp ? byComp.company : trimmed;
            }
        }

        function pickCustomerScheduleDate(dateStr, level) {
            selectedCustomerDate = dateStr;
            render();
            switchCustTab('new');
            const dateInput = document.getElementById('new-date');
            if (dateInput) dateInput.value = dateStr;
        }

        let currentSigTicketId = null;
        let currentSigType = null;
        let currentSigColor = '#0f172a';
        let sigDrawn = false;
        let sigCtx = null;
        let sigCanvas = null;
        let isDrawingSig = false;
        let lastSigPos = null;

        function setSigColor(c) {
            currentSigColor = c;
            if (sigCtx) sigCtx.strokeStyle = c;
        }

        function clearSigCanvas() {
            if (!sigCanvas || !sigCtx) return;
            sigCtx.clearRect(0, 0, sigCanvas.width, sigCanvas.height);
            sigDrawn = false;
            const wm = document.getElementById('sig-watermark');
            if (wm) wm.style.display = 'flex';
        }

        function closeSignatureModal() {
            const el = document.getElementById('online-signature-modal');
            if (el) el.remove();
            currentSigTicketId = null;
            currentSigType = null;
        }

        function openSignatureModal(ticketId, type) {
            currentSigTicketId = ticketId;
            currentSigType = type;
            const t = state.tickets.find(x => x.id === ticketId);
            if (!t) return;

            const existingModal = document.getElementById('online-signature-modal');
            if (existingModal) existingModal.remove();

            const isTech = type === 'technician';
            const title = isTech ? '✍️ ลายเซ็นช่างผู้ปฏิบัติงาน (Technician Signature)' : '✍️ ลายเซ็นลูกค้าผู้ตรวจรับงาน (Customer Acceptance)';
            const signerName = isTech 
                ? (t.assignedTech || 'ช่างเทคนิค') 
                : (t.customerSignerName || t.siteContact || t.contact || '');
            const currentSig = isTech ? t.technicianSignature : t.customerSignature;

            const modal = document.createElement('div');
            modal.id = 'online-signature-modal';
            modal.className = 'fixed inset-0 bg-slate-900/80 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print';
            modal.innerHTML = \`
                <div class="bg-white w-full max-w-lg rounded-2xl shadow-2xl p-5 sm:p-6 space-y-4 text-xs sm:text-sm">
                    <div class="flex justify-between items-start border-b pb-3">
                        <div>
                            <span class="text-[10px] font-bold px-2 py-0.5 rounded \${isTech ? 'bg-blue-100 text-blue-800' : 'bg-emerald-100 text-emerald-800'} uppercase">
                                \${isTech ? 'TECHNICIAN E-SIGNATURE' : 'CUSTOMER SIGN-OFF'}
                            </span>
                            <h3 class="text-base sm:text-lg font-bold text-slate-900 mt-1">\${title}</h3>
                            <p class="text-xs text-slate-500">ใบงาน: <b class="text-orange-600">\${t.id}</b> | บริษัท: <b>\${t.company}</b></p>
                        </div>
                        <button onclick="closeSignatureModal()" class="text-2xl font-bold text-slate-400 hover:text-slate-600 cursor-pointer">&times;</button>
                    </div>

                    <div class="space-y-3">
                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">
                                \${isTech ? 'ชื่อ-นามสกุล ช่างผู้ปฏิบัติงาน *' : 'ชื่อ-นามสกุล ลูกค้าผู้ตรวจรับงาน *'}
                            </label>
                            <input type="text" id="signer-name-input" value="\${signerName}" placeholder="ระบุชื่อผู้ลงนาม..." class="w-full border rounded-xl p-2.5 text-xs sm:text-sm font-semibold bg-slate-50">
                        </div>

                        <div>
                            <div class="flex justify-between items-center mb-1">
                                <label class="text-xs font-bold text-slate-700">ใช้นิ้วหรือปากกาเขียนลายเซ็นลงในกรอบด้านล่าง:</label>
                                <div class="flex items-center space-x-1.5">
                                    <span class="text-[10px] text-slate-500">สีหมึก:</span>
                                    <button type="button" onclick="setSigColor('#0f172a')" class="w-5 h-5 rounded-full bg-slate-900 border-2 border-white ring-1 ring-slate-400 cursor-pointer" title="สีดำ"></button>
                                    <button type="button" onclick="setSigColor('#1d4ed8')" class="w-5 h-5 rounded-full bg-blue-700 border-2 border-white ring-1 ring-blue-400 cursor-pointer" title="สีน้ำเงิน"></button>
                                </div>
                            </div>
                            <div class="relative border-2 border-dashed border-slate-300 rounded-xl bg-slate-50 overflow-hidden" style="touch-action: none;">
                                <canvas id="sig-pad-canvas" class="w-full h-44 bg-white cursor-crosshair block"></canvas>
                                <div id="sig-watermark" class="pointer-events-none absolute inset-0 flex items-center justify-center text-slate-300 text-xs font-bold tracking-widest uppercase select-none">
                                    ✍️ วาดลายเซ็นที่นี่ (Sign Here)
                                </div>
                            </div>
                        </div>

                        \${currentSig ? \`
                            <div class="p-2.5 bg-slate-50 rounded-xl border flex items-center justify-between text-xs">
                                <div class="flex items-center gap-2">
                                    <span class="text-[10px] font-bold text-emerald-700">มีลายเซ็นเดิมในระบบแล้ว</span>
                                    <img src="\${currentSig}" class="h-8 max-w-[120px] object-contain border bg-white rounded p-0.5">
                                </div>
                                <span class="text-[10px] text-slate-500">เซ็นใหม่ด้านบนเพื่อแทนที่</span>
                            </div>
                        \` : ''}
                    </div>

                    <div class="flex items-center justify-between pt-2 border-t">
                        <button type="button" onclick="clearSigCanvas()" class="px-3.5 py-2 text-xs font-bold text-red-600 bg-red-50 hover:bg-red-100 rounded-xl cursor-pointer">
                            <i class="fa-solid fa-eraser mr-1"></i> ล้างลายเซ็น (Clear)
                        </button>
                        <div class="flex items-center space-x-2">
                            <button type="button" onclick="closeSignatureModal()" class="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-xl cursor-pointer">
                                ยกเลิก
                            </button>
                            <button type="button" onclick="saveSignatureModal()" class="px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-xl shadow cursor-pointer">
                                <i class="fa-solid fa-check mr-1"></i> ยืนยันบันทึกลายเซ็น
                            </button>
                        </div>
                    </div>
                </div>
            \`;
            document.body.appendChild(modal);

            setTimeout(() => {
                initSigCanvasEvents();
            }, 50);
        }

        function initSigCanvasEvents() {
            sigCanvas = document.getElementById('sig-pad-canvas');
            if (!sigCanvas) return;

            const rect = sigCanvas.getBoundingClientRect();
            const dpr = Math.min(window.devicePixelRatio || 1, 2);
            sigCanvas.width = Math.max(300, Math.floor(rect.width * dpr));
            sigCanvas.height = Math.max(140, Math.floor(rect.height * dpr));

            sigCtx = sigCanvas.getContext('2d');
            sigCtx.lineCap = 'round';
            sigCtx.lineJoin = 'round';
            sigCtx.strokeStyle = currentSigColor;
            sigCtx.lineWidth = 2.4 * dpr;

            sigDrawn = false;
            isDrawingSig = false;
            lastSigPos = null;

            function getCoords(e) {
                const r = sigCanvas.getBoundingClientRect();
                let clientX = e.clientX;
                let clientY = e.clientY;
                if (e.touches && e.touches.length > 0) {
                    clientX = e.touches[0].clientX;
                    clientY = e.touches[0].clientY;
                }
                const sx = sigCanvas.width / r.width;
                const sy = sigCanvas.height / r.height;
                return {
                    x: (clientX - r.left) * sx,
                    y: (clientY - r.top) * sy
                };
            }

            function onStart(e) {
                e.preventDefault();
                try {
                    if (e.pointerId && sigCanvas.setPointerCapture) {
                        sigCanvas.setPointerCapture(e.pointerId);
                    }
                } catch (_) {}
                isDrawingSig = true;
                const pt = getCoords(e);
                lastSigPos = pt;
                sigDrawn = true;
                const wm = document.getElementById('sig-watermark');
                if (wm) wm.style.display = 'none';

                sigCtx.beginPath();
                sigCtx.arc(pt.x, pt.y, sigCtx.lineWidth / 2, 0, Math.PI * 2);
                sigCtx.fillStyle = currentSigColor;
                sigCtx.fill();
            }

            function onMove(e) {
                if (!isDrawingSig || !lastSigPos) return;
                e.preventDefault();
                const pt = getCoords(e);
                sigCtx.beginPath();
                sigCtx.moveTo(lastSigPos.x, lastSigPos.y);
                sigCtx.lineTo(pt.x, pt.y);
                sigCtx.stroke();
                lastSigPos = pt;
            }

            function onEnd(e) {
                if (!isDrawingSig) return;
                isDrawingSig = false;
                lastSigPos = null;
            }

            sigCanvas.addEventListener('pointerdown', onStart);
            sigCanvas.addEventListener('pointermove', onMove);
            sigCanvas.addEventListener('pointerup', onEnd);
            sigCanvas.addEventListener('pointercancel', onEnd);

            sigCanvas.addEventListener('touchstart', onStart, { passive: false });
            sigCanvas.addEventListener('touchmove', onMove, { passive: false });
            sigCanvas.addEventListener('touchend', onEnd);
        }

        function saveSignatureModal() {
            if (!currentSigTicketId || !currentSigType) return;
            const t = state.tickets.find(x => x.id === currentSigTicketId);
            if (!t) return;

            const existingSig = currentSigType === 'technician' ? t.technicianSignature : t.customerSignature;
            if (!sigDrawn && !existingSig) {
                alert('⚠️ กรุณาใช้นิ้วหรือปากกาเขียนลายเซ็นลงในกรอบก่อนกดยืนยัน');
                return;
            }

            const nameInput = document.getElementById('signer-name-input');
            const signerName = nameInput ? nameInput.value.trim() : '';
            const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 16);

            let dataUrl = existingSig;
            if (sigDrawn && sigCanvas) {
                const exp = document.createElement('canvas');
                exp.width = 400;
                exp.height = 140;
                const expCtx = exp.getContext('2d');
                if (expCtx) {
                    expCtx.drawImage(sigCanvas, 0, 0, exp.width, exp.height);
                    dataUrl = exp.toDataURL('image/png');
                } else {
                    dataUrl = sigCanvas.toDataURL('image/png');
                }
            }

            let lineNotice = '';
            if (currentSigType === 'technician') {
                t.technicianSignature = dataUrl;
                t.technicianSignedAt = nowStr;
                if (signerName) t.assignedTech = signerName;
                lineNotice = '✍️ [ช่าง ' + (t.assignedTech || 'เทคนิค') + ' ลงลายเซ็นรับรองการปฏิบัติงานแล้ว (' + t.id + ')]';
                alert('✅ บันทึกลายเซ็นช่างผู้ปฏิบัติงานเรียบร้อย!');
            } else {
                t.customerSignature = dataUrl;
                t.customerSignedAt = nowStr;
                if (signerName) t.customerSignerName = signerName;
                if (t.status === 'CUSTOMER_ACCEPTANCE' || state.userRole === 'CUSTOMER') {
                    t.status = 'CLOSE';
                    t.isLocked = true;
                    lineNotice = '🎉 [ลูกค้าเซ็นชื่อตรวจรับงานบนหน้าจอและปิดจบใบงาน (CLOSE) เรียบร้อยแล้ว (' + t.id + ')]';
                    alert('🎉 บันทึกลายเซ็นลูกค้าตรวจรับงาน และปิดจบใบงานเรียบร้อย!');
                } else {
                    lineNotice = '✍️ [ลูกค้าลงลายมือชื่อตรวจรับงานแล้ว (' + t.id + ')]';
                    alert('✅ บันทึกลายเซ็นลูกค้าผู้ตรวจรับงานเรียบร้อย!');
                }
            }

            persistState(lineNotice ? formatTicketLineMessage(t, lineNotice) : '');
            closeSignatureModal();
            render();
        }

        const app = {
            goHome: () => { state.currentView = 'home'; render(); },
            openCustomerPortal: async () => {
                let comp = await openCustomPrompt({
                    title: 'ตรวจสอบสถานะงาน / จองคิวช่าง (Service Request)',
                    message: 'กรุณากรอก ชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม (เช่น SKM-69-001):',
                    defaultValue: customerLookupQuery || state.customerCompany || 'บริษัท สยามอุตสาหกรรม จำกัด',
                    placeholder: 'ชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม'
                });
                if (!comp) return;
                applyStandaloneCustomerLookup(comp);
                state.currentView = 'customer_portal';
                render();
            },
            openInternalLogin: () => { state.currentView = 'internal_login'; render(); },
            loginInternalSubmit: (event) => {
                event.preventDefault();
                const pin = document.getElementById('emp-pin').value.trim();
                const role = document.getElementById('emp-role').value;
                if (role === 'SUPERVISOR') {
                    if (pin !== savedSupPassword && pin !== 'admin123') {
                        showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง', 'error');
                        return;
                    }
                    state.userRole = role;
                    state.loggedInTech = null;
                } else if (role === 'SERVICE TEAM') {
                    if (pin !== savedServicePin && pin !== savedSupPassword && pin !== 'admin123' && !state.technicians.some(tc => tc.pin === pin)) {
                        showAppToast('รหัสผ่านเจ้าหน้าที่ Service ไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง', 'error');
                        return;
                    }
                    state.userRole = role;
                    state.loggedInTech = null;
                } else if (role === 'TECHNICIAN') {
                    const techSelectEl = document.getElementById('tech-user-select');
                    const selectedId = techSelectEl ? techSelectEl.value : '';
                    let matched = state.technicians.find(tc => tc.id === selectedId);
                    if (!matched) {
                        matched = state.technicians.find(tc => tc.pin === pin);
                    }
                    if (pin !== savedSupPassword && pin !== 'admin123' && (!matched || matched.pin !== pin)) {
                        showAppToast('รหัส PIN ช่างเทคนิคไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง', 'error');
                        return;
                    }
                    state.loggedInTech = matched || state.technicians[0] || { id: 'TECH-01', name: 'ช่างเทคนิค', position: 'ช่างเทคนิค (Technician)', pin: pin };
                    state.userRole = role;
                    techWorkFilter = 'MY_JOBS';
                }
                state.currentView = 'internal_dashboard';
                render();
                showAppToast('ยินดีต้อนรับ ' + state.userRole + (state.loggedInTech ? (' (' + state.loggedInTech.name + ')') : ''), 'success');
            },
            logoutInternal: () => { state.userRole = null; state.loggedInTech = null; state.currentView = 'home'; render(); }
        };

        async function switchRoleWithAuth(newRole) {
            if (newRole === 'SUPERVISOR' && state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 กรุณากรอกรหัสผ่าน Supervisor',
                    message: 'เฉพาะผู้ดูแลระบบและหัวหน้างานเท่านั้น (รหัสเริ่มต้น admin123):',
                    inputType: 'password',
                    placeholder: 'admin123'
                });
                if (pass !== savedSupPassword && pass !== 'admin123') {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    return;
                }
                state.loggedInTech = null;
            } else if (newRole === 'TECHNICIAN') {
                if (!state.loggedInTech && state.technicians.length > 0) {
                    state.loggedInTech = state.technicians[0];
                }
                techWorkFilter = 'MY_JOBS';
            }
            state.userRole = newRole;
            render();
        }

        function render() {
            updateGsheetBadge();
            const syncBarEl = document.getElementById('supervisor-sync-bar-container');
            if (syncBarEl) {
                syncBarEl.classList.toggle('hidden', state.userRole !== 'SUPERVISOR');
            }
            const sumEl = document.getElementById('online-sync-summary');
            if (sumEl) {
                sumEl.textContent = 'ระบบซิงค์ข้อมูลออนไลน์อัตโนมัติ (ใบงาน ' + state.tickets.length + ' รายการ · ช่าง ' + state.technicians.length + ' คน)';
            }
            const main = document.getElementById('main-content');
            const headerNav = document.getElementById('header-nav');
            main.innerHTML = '';
            headerNav.innerHTML = '';

            if (state.currentView === 'home') {
                main.innerHTML = renderHome();
            } else if (state.currentView === 'customer_portal') {
                headerNav.innerHTML = \`<button onclick="app.goHome()" class="bg-slate-700 hover:bg-slate-600 text-white px-3 py-1.5 rounded text-xs sm:text-sm transition cursor-pointer"><i class="fa-solid fa-house mr-1"></i> หน้าหลัก</button>\`;
                main.innerHTML = renderCustomerPortal();
            } else if (state.currentView === 'internal_login') {
                headerNav.innerHTML = \`<button onclick="app.goHome()" class="bg-slate-700 hover:bg-slate-600 text-white px-3 py-1.5 rounded text-xs sm:text-sm transition cursor-pointer"><i class="fa-solid fa-house mr-1"></i> หน้าหลัก</button>\`;
                main.innerHTML = renderInternalLogin();
            } else if (state.currentView === 'internal_dashboard') {
                let urgentCount = state.tickets.filter(t => t.status === 'NEW' || t.status === 'PENDING_SUPERVISOR_APPROVAL' || t.status === 'PENDING_PARTS_APPROVAL').length;
                if (urgentCount > 0) playAlertSound();
                headerNav.innerHTML = \`
                    <div class="flex items-center space-x-2 text-xs sm:text-sm">
                        \${urgentCount > 0 ? \`<span class="bg-red-600 text-white text-[10px] sm:text-xs px-2 py-0.5 rounded-full font-bold shadow"><i class="fa-solid fa-bell mr-1"></i> งานรออนุมัติ (\${urgentCount})</span>\` : ''}
                        <span class="bg-orange-500 text-white px-2 py-1 rounded-full font-bold text-[10px] sm:text-xs"><i class="fa-solid fa-user-shield mr-1"></i> \${state.userRole}</span>
                        <button onclick="app.logoutInternal()" class="bg-red-600 hover:bg-red-700 text-white px-2.5 py-1 rounded text-xs transition cursor-pointer"><i class="fa-solid fa-right-from-bracket"></i></button>
                    </div>
                \`;
                main.innerHTML = renderInternalDashboard();
            } else if (state.currentView === 'report_view') {
                headerNav.innerHTML = \`<button onclick="state.currentView = state.userRole ? 'internal_dashboard' : 'customer_portal'; render();" class="bg-slate-700 hover:bg-slate-600 text-white px-3 py-1.5 rounded text-xs sm:text-sm transition cursor-pointer"><i class="fa-solid fa-arrow-left mr-1"></i> ย้อนกลับ</button>\`;
                main.innerHTML = renderReportView(state.activeReportId);
            }
        }

        function renderHome() {
            return \`
                <div class="py-6 px-4 text-center max-w-4xl mx-auto space-y-6">
                    <div>
                        <span class="bg-orange-100 text-orange-600 text-[11px] font-bold px-3 py-1 rounded-full uppercase">Saha Kim Motor Platform V12.1</span>
                        <h2 class="text-2xl sm:text-4xl font-extrabold text-slate-900 mt-3">ระบบบริหารงานบริการหลังการขาย</h2>
                        <p class="text-slate-600 mt-1 font-bold text-base sm:text-lg">\${companyProfile.nameTh} (\${companyProfile.nameEn})</p>
                    </div>

                    <!-- ปุ่มทางลัดเมนูใช้งานหลัก (กดเลือกเมนูเพื่อเข้าใช้งานได้ทันที) -->
                    <div class="bg-white p-4 rounded-2xl border-2 border-slate-200 shadow-sm text-left space-y-2.5">
                        <p class="text-xs font-bold text-slate-700">⚡ เมนูทางลัดเข้าใช้งานด่วน (คลิกเลือกเมนูที่ต้องการ):</p>
                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
                            <button type="button" onclick="state.currentView = 'customer_portal'; render(); setTimeout(() => switchCustTab('new'), 50);" class="p-3 rounded-xl border-2 border-orange-200 bg-orange-50/70 hover:bg-orange-500 hover:text-white text-slate-900 font-bold text-left transition cursor-pointer">
                                <p class="text-sm font-extrabold">➕ แจ้งซ่อมใหม่</p>
                                <p class="text-[11px] opacity-80 mt-0.5">เปิดใบแจ้งซ่อม Service Request</p>
                            </button>
                            <button type="button" onclick="state.currentView = 'customer_portal'; render(); setTimeout(() => switchCustTab('schedule'), 50);" class="p-3 rounded-xl border-2 border-blue-200 bg-blue-50/70 hover:bg-blue-600 hover:text-white text-slate-900 font-bold text-left transition cursor-pointer">
                                <p class="text-sm font-extrabold">📅 ตารางงานช่าง</p>
                                <p class="text-[11px] opacity-80 mt-0.5">เช็คคิวว่าง แดง/เหลือง/เขียว</p>
                            </button>
                            <button type="button" onclick="state.currentView = 'customer_portal'; render(); setTimeout(() => switchCustTab('list'), 50);" class="p-3 rounded-xl border-2 border-emerald-200 bg-emerald-50/70 hover:bg-emerald-600 hover:text-white text-slate-900 font-bold text-left transition cursor-pointer">
                                <p class="text-sm font-extrabold">🔍 ติดตามงานซ่อม</p>
                                <p class="text-[11px] opacity-80 mt-0.5">ค้นหาสถานะ & ใบรายงาน</p>
                            </button>
                            <button type="button" onclick="internalActiveTab = 'dashboard'; app.openInternalLogin();" class="p-3 rounded-xl border-2 border-indigo-200 bg-indigo-50/70 hover:bg-indigo-600 hover:text-white text-slate-900 font-bold text-left transition cursor-pointer">
                                <p class="text-sm font-extrabold">📊 Dashboard & ช่าง</p>
                                <p class="text-[11px] opacity-80 mt-0.5">สำหรับเจ้าหน้าที่ภายใน</p>
                            </button>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mt-4">
                        <div onclick="app.openCustomerPortal()" class="bg-white p-6 sm:p-8 rounded-2xl shadow-lg border-2 border-orange-500 hover:shadow-xl transition cursor-pointer flex flex-col justify-between group">
                            <div>
                                <div class="w-16 h-16 bg-orange-500 text-white rounded-2xl flex items-center justify-center text-2xl mb-3 shadow">
                                    <i class="fa-solid fa-users"></i>
                                </div>
                                <h3 class="text-xl font-bold text-slate-900">1. Service Request (สำหรับลูกค้า)</h3>
                                <p class="text-slate-600 text-xs sm:text-sm mt-2">ติดตามสถานะงานซ่อม แจ้งซ่อมใหม่ และตรวจสอบตารางคิวงานช่าง</p>
                            </div>
                            <span class="mt-6 inline-flex items-center text-orange-600 font-bold text-xs sm:text-sm group-hover:translate-x-1 transition-transform">เข้าสู่หน้า Service Request <i class="fa-solid fa-arrow-right ml-1.5"></i></span>
                        </div>
                        <div onclick="app.openInternalLogin()" class="bg-white p-6 sm:p-8 rounded-2xl shadow-lg border-2 border-slate-900 hover:shadow-xl transition cursor-pointer flex flex-col justify-between group">
                            <div>
                                <div class="flex items-center gap-3.5 mb-3">
                                    <img src="${EMBEDDED_SKM_MASCOT}" alt="เจ้าหน้าที่ภายใน SKM SERVICE" class="w-20 h-20 rounded-2xl object-cover border-2 border-orange-500 bg-slate-900 shadow shrink-0" />
                                    <div>
                                        <span class="inline-block px-2.5 py-0.5 rounded-md text-[10px] font-bold bg-orange-100 text-orange-800 border border-orange-200 mb-1">SKM SERVICE · เชี่ยวชาญงาน ใส่ใจบริการ</span>
                                        <h3 class="text-xl font-bold text-slate-900">2. เจ้าหน้าที่ภายใน (Internal Staff)</h3>
                                    </div>
                                </div>
                                <p class="text-slate-600 text-xs sm:text-sm mt-2">จัดการใบงานซ่อม ปฏิบัติงานช่าง ตารางงานช่าง สรุป Dashboard และระบบตั้งค่าสำหรับเจ้าหน้าที่</p>
                            </div>
                            <span class="mt-6 inline-flex items-center text-slate-900 font-bold text-xs sm:text-sm group-hover:translate-x-1 transition-transform">เข้าสู่ระบบภายใน <i class="fa-solid fa-arrow-right ml-1.5"></i></span>
                        </div>
                    </div>
                </div>
            \`;
        }

        let tempEquipmentList = [];

        function renderCustomerPortal() {
            const q = (customerLookupQuery || state.customerCompany || '').trim().toLowerCase();
            const customerTickets = state.tickets.filter(t =>
                !q ? false : ((t.id || '').toLowerCase().includes(q) || (t.company || '').toLowerCase().includes(q))
            );
            const scheduleDays = getCustomerScheduleDays();

            return \`
                <div class="space-y-4">
                    <div class="bg-slate-900 text-white p-4 sm:p-6 rounded-2xl shadow-lg space-y-3">
                        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                            <div>
                                <span class="bg-orange-500 text-white text-[10px] px-2.5 py-0.5 rounded-full font-bold uppercase">SERVICE REQUEST PORTAL</span>
                                <h2 class="text-xl sm:text-2xl font-bold mt-1">Service Request — \${state.customerCompany}</h2>
                                <p class="text-xs text-slate-300">ตรวจสอบสถานะงานด้วยชื่อบริษัทหรือเลขที่ใบแจ้งซ่อม และเลือกวันขอช่างเบื้องต้นจากตารางงานช่างด้านล่าง</p>
                            </div>
                            <button onclick="app.openCustomerPortal()" class="bg-white text-slate-900 font-bold px-3 py-1.5 rounded-xl text-xs hover:bg-slate-100 shadow cursor-pointer">ค้นหาบริษัท/เลขใบงานอื่น</button>
                        </div>
                        <div class="bg-slate-800 border border-slate-700 p-3 rounded-xl flex flex-col sm:flex-row gap-2">
                            <input type="text" id="cust-inline-search" value="\${customerLookupQuery || state.customerCompany}" placeholder="พิมพ์ชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม เช่น SKM-69-001..." class="flex-1 bg-white text-slate-900 rounded-lg px-3 py-2 text-xs sm:text-sm font-medium">
                            <button type="button" onclick="applyStandaloneCustomerLookup(document.getElementById('cust-inline-search').value); render();" class="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-magnifying-glass mr-1"></i> ค้นหาสถานะงาน</button>
                        </div>
                    </div>

                    <!-- ตารางงานช่าง (สีแดงเต็มมาก / สีเหลืองรับงานได้ / สีเขียวว่าง) -->
                    <div class="bg-white p-4 sm:p-5 rounded-2xl shadow border-2 border-slate-200 space-y-3">
                        <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-2 border-b pb-2.5">
                            <div>
                                <h3 class="font-bold text-slate-900 text-sm sm:text-base"><i class="fa-solid fa-calendar-check text-orange-500 mr-1.5"></i> ตารางงานช่าง (คลิกที่วันเพื่อเลือกระบุวันขอช่างเบื้องต้น)</h3>
                                <p class="text-xs text-slate-500">ตรวจสอบคิวว่างของทีมช่างก่อนนัดหมาย — คลิกเลือกวันที่ต้องการเพื่อกรอกลงในฟอร์มแจ้งซ่อมอัตโนมัติ</p>
                            </div>
                            <div class="flex flex-wrap gap-2 text-xs">
                                <span class="px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-900 font-bold">🟢 สีเขียว = ว่าง</span>
                                <span class="px-2.5 py-1 rounded-lg bg-amber-50 border border-amber-300 text-amber-900 font-bold">🟡 สีเหลือง = รับงานได้</span>
                                <span class="px-2.5 py-1 rounded-lg bg-red-50 border border-red-300 text-red-900 font-bold">🔴 สีแดง = เต็มมาก</span>
                            </div>
                        </div>
                        <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2">
                            \${scheduleDays.map(day => {
                                const isSel = selectedCustomerDate === day.dateStr;
                                const cardCls = day.level === 'RED'
                                    ? 'bg-red-50 border-red-400 text-red-950 hover:bg-red-100'
                                    : day.level === 'YELLOW'
                                    ? 'bg-amber-50 border-amber-400 text-amber-950 hover:bg-amber-100'
                                    : 'bg-emerald-50 border-emerald-400 text-emerald-950 hover:bg-emerald-100';
                                const badgeCls = day.level === 'RED'
                                    ? 'bg-red-600 text-white'
                                    : day.level === 'YELLOW'
                                    ? 'bg-amber-500 text-white'
                                    : 'bg-emerald-600 text-white';
                                return \`
                                    <button type="button" onclick="pickCustomerScheduleDate('\${day.dateStr}', '\${day.level}')" class="p-2.5 rounded-xl border-2 text-left transition cursor-pointer \${cardCls} \${isSel ? 'ring-2 ring-slate-900 shadow-md' : ''}">
                                        <p class="text-xs font-extrabold">\${day.dayLabel}</p>
                                        <p class="text-[10px] font-mono opacity-75 mb-1.5">\${day.dateStr}</p>
                                        <span class="inline-block text-[10px] font-bold px-2 py-0.5 rounded \${badgeCls}">\${day.statusText}</span>
                                        <p class="text-[10px] font-semibold mt-1">\${day.detailText}</p>
                                    </button>
                                \`;
                            }).join('')}
                        </div>
                    </div>

                    <div class="flex border-b border-slate-200 overflow-x-auto text-xs sm:text-sm">
                        <button onclick="switchCustTab('list')" id="btn-cust-list" class="py-2 px-4 font-bold border-b-2 border-orange-500 text-orange-600 whitespace-nowrap cursor-pointer">สถานะงานซ่อม (\${customerTickets.length})</button>
                        <button onclick="switchCustTab('new')" id="btn-cust-new" class="py-2 px-4 font-bold border-b-2 border-transparent text-slate-500 whitespace-nowrap cursor-pointer"><i class="fa-solid fa-plus-circle mr-1"></i> แจ้งซ่อมใหม่ (Service Request)</button>
                    </div>

                    <div id="cust-tab-list" class="space-y-3">
                        \${customerTickets.length === 0 ? \`
                            <div class="bg-white p-8 rounded-2xl text-center shadow border">
                                <p class="text-slate-500 text-xs sm:text-sm">ไม่พบประวัติใบแจ้งซ่อมสำหรับ "\${customerLookupQuery || state.customerCompany}"</p>
                                <button onclick="switchCustTab('new')" class="mt-3 bg-orange-500 text-white px-4 py-2 rounded-xl text-xs font-bold shadow cursor-pointer">+ สร้างใบ Service Request ใหม่</button>
                            </div>
                        \` : customerTickets.map(t => \`
                            <div class="bg-white p-4 rounded-xl shadow border flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
                                <div class="space-y-1 text-xs sm:text-sm">
                                    <div class="flex items-center space-x-2 flex-wrap gap-1">
                                        <span class="font-bold text-orange-600 text-base">\${t.id}</span>
                                        <span class="bg-slate-100 text-slate-800 text-xs px-2 py-0.5 rounded font-bold">\${t.company}</span>
                                        <span class="bg-orange-100 text-orange-800 text-xs px-2.5 py-0.5 rounded-full font-bold">สถานะ: \${t.status}</span>
                                    </div>
                                    <p class="text-slate-700"><b>บริการ:</b> \${t.serviceType} | <b>ประกัน:</b> <span class="text-indigo-600 font-bold">\${t.warranty}</span></p>
                                    <p class="text-slate-500 text-xs">วันขอช่าง/นัดหมาย: <span class="font-bold text-emerald-600">\${t.confirmedDate || t.desiredDate}</span> | สถานที่: \${t.location || '-'} | ช่าง: \${t.assignedTech || 'อยู่ระหว่างจัดสรรช่าง'}</p>
                                </div>
                                <div class="flex items-center space-x-2">
                                    <button onclick="openServiceReport('\${t.id}')" class="bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-print mr-1"></i> รายงาน</button>
                                    <button onclick="openSignatureModal('\${t.id}', 'customer')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-signature mr-1"></i> \${t.customerSignature ? '✓ ลายเซ็นรับงานแล้ว' : (t.status === 'CUSTOMER_ACCEPTANCE' ? '✍️ เซ็นตรวจรับงานและปิดจบ' : '✍️ เซ็นตรวจรับงานบนจอ')}</button>
                                </div>
                            </div>
                        \`).join('')}
                    </div>

                    <div id="cust-tab-new" class="hidden bg-white p-4 sm:p-6 rounded-2xl shadow border">
                        <h3 class="text-base font-bold text-slate-900 mb-3 pb-2 border-b"><i class="fa-solid fa-file-circle-plus text-orange-500 mr-2"></i> แบบฟอร์ม Service Request (แจ้งซ่อมและระบุวันขอช่างเบื้องต้น)</h3>
                        <form onsubmit="submitNewTicket(event)" class="space-y-4 text-xs sm:text-sm">
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                <div><label class="block font-bold text-slate-700 mb-1">ชื่อบริษัท *</label><input type="text" id="new-company" value="\${state.customerCompany}" required class="w-full border rounded-xl p-2 bg-slate-50"></div>
                                <div><label class="block font-bold text-slate-700 mb-1">TAX ID</label><input type="text" id="new-taxid" placeholder="0105552012345" class="w-full border rounded-xl p-2"></div>
                                <div><label class="block font-bold text-slate-700 mb-1">เลขที่เอกสาร / ใบรับประกัน (ถ้ามี)</label><input type="text" id="new-docno" placeholder="เช่น DOC-2026-999" class="w-full border rounded-xl p-2"></div>
                                <div><label class="block font-bold text-slate-700 mb-1">ผู้ติดต่อและเบอร์โทร *</label><input type="text" id="new-contact" placeholder="คุณสมชาย (081-234-5678)" required class="w-full border rounded-xl p-2"></div>
                                <div>
                                    <label class="block font-bold text-slate-700 mb-1">เงื่อนไขการรับประกัน / คิดเงิน *</label>
                                    <select id="new-warranty" class="w-full border rounded-xl p-2 bg-white font-bold text-orange-600">
                                        <option>In-Warranty (อยู่ในประกัน - ฟรี / FOC)</option>
                                        <option>Out-of-Warranty (หมดประกัน - คิดค่าบริการ / Charge)</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block font-bold text-slate-700 mb-1">ประเภทบริการ *</label>
                                    <select id="new-servicetype" class="w-full border rounded-xl p-2 bg-white font-bold text-slate-700">
                                        <option>ซ่อมหน้างาน</option>
                                        <option>รับมาซ่อม</option>
                                        <option>งานตรวจเช็ค</option>
                                        <option>คอมมิชชั่นนิ่ง (Commissioning)</option>
                                        <option>เทสรัน (Test Run)</option>
                                        <option>PM (Preventive Maintenance)</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block font-bold text-slate-700 mb-1">ความเร่งด่วนของงาน *</label>
                                    <select id="new-urgency" class="w-full border rounded-xl p-2 bg-white font-bold text-red-600">
                                        <option>ปกติ (Normal)</option>
                                        <option>ด่วน (High)</option>
                                        <option>ด่วนที่สุด (Urgent)</option>
                                    </select>
                                </div>
                                <div><label class="block font-bold text-slate-700 mb-1">สถานที่ปฏิบัติงาน / โลเคชั่น *</label><input type="text" id="new-location" placeholder="ระบุสถานที่ หรือลิงก์ Google Maps" required class="w-full border rounded-xl p-2"></div>
                                <div><label class="block font-bold text-slate-700 mb-1">วันที่ขอช่างเบื้องต้น (คลิกเลือกจากตารางงานช่างด้านบนได้) *</label><input type="date" id="new-date" value="\${selectedCustomerDate}" required class="w-full border rounded-xl p-2 font-mono"></div>
                                <div class="sm:col-span-2">
                                    <label class="block font-bold text-slate-700 mb-1">อัปโหลดรูปภาพอาการเสีย (ถ่ายรูปผ่านมือถือหรือเลือกไฟล์)</label>
                                    <input type="file" id="new-image-file" accept="image/*" capture="environment" class="w-full border rounded-xl p-2 bg-slate-50 text-xs">
                                </div>
                            </div>
                            <div class="border-t pt-3">
                                <div class="flex justify-between items-center mb-2">
                                    <h4 class="font-bold text-slate-900 text-xs">รายการอุปกรณ์ (เพิ่มอุปกรณ์ได้หลายรายการ เลือก ETC. เพื่อพิมพ์เพิ่มเองได้)</h4>
                                    <button type="button" onclick="addEquipmentRow()" class="bg-slate-900 text-white text-xs px-3 py-1.5 rounded-lg font-bold cursor-pointer"><i class="fa-solid fa-plus mr-1"></i> เพิ่มอุปกรณ์</button>
                                </div>
                                <div id="equipment-container" class="space-y-3"></div>
                            </div>
                            <div class="flex justify-end space-x-2 pt-3 border-t">
                                <button type="button" onclick="switchCustTab('list')" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer">ยกเลิก</button>
                                <button type="submit" class="bg-orange-500 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">ส่งใบแจ้งซ่อม (Service Request)</button>
                            </div>
                        </form>
                    </div>
                </div>
            \`;
        }

        function switchCustTab(tab) {
            ['list', 'new'].forEach(t => {
                let el = document.getElementById(\`cust-tab-\${t}\`);
                let btn = document.getElementById(\`btn-cust-\${t}\`);
                if (el) el.classList.add('hidden');
                if (btn) btn.className = "py-2 px-4 font-bold border-b-2 border-transparent text-slate-500 whitespace-nowrap cursor-pointer";
            });
            let targetEl = document.getElementById(\`cust-tab-\${tab}\`);
            let targetBtn = document.getElementById(\`btn-cust-\${tab}\`);
            if (targetEl) targetEl.classList.remove('hidden');
            if (targetBtn) targetBtn.className = "py-2 px-4 font-bold border-b-2 border-orange-500 text-orange-600 whitespace-nowrap cursor-pointer";
            if (tab === 'new' && tempEquipmentList.length === 0) addEquipmentRow();
        }

        function addEquipmentRow() {
            tempEquipmentList.push({ type: "END SUCTION PUMP", customType: "", brand: "STAC", customBrand: "", model: "", tag: "", symptom: "", beforePhotos: [], afterPhotos: [] });
            renderEquipmentRows();
        }

        function removeEquipmentRow(index) {
            tempEquipmentList.splice(index, 1);
            if (tempEquipmentList.length === 0) addEquipmentRow();
            renderEquipmentRows();
        }

        function renderEquipmentRows() {
            const container = document.getElementById('equipment-container');
            if (!container) return;
            const eqMasters = getEquipmentMasters();
            const brands = getBrands();
            container.innerHTML = tempEquipmentList.map((eq, index) => \`
                <div class="bg-slate-50 p-3 rounded-xl border relative space-y-2 text-xs">
                    <div class="flex justify-between items-center">
                        <span class="font-bold text-orange-600">อุปกรณ์ #\${index + 1}</span>
                        \${tempEquipmentList.length > 1 ? \`<button type="button" onclick="removeEquipmentRow(\${index})" class="text-red-500 font-bold cursor-pointer">ลบ</button>\` : ''}
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-4 gap-2">
                        <div>
                            <select onchange="tempEquipmentList[\${index}].type = this.value; renderEquipmentRows();" class="w-full border rounded p-1.5 bg-white font-bold">
                                \${eqMasters.map(m => \`<option \${eq.type === m ? 'selected' : ''}>\${m}</option>\`).join('')}
                            </select>
                            \${eq.type.startsWith('ETC.') ? \`<input type="text" value="\${eq.customType || ''}" oninput="tempEquipmentList[\${index}].customType = this.value" placeholder="ระบุประเภทอุปกรณ์เอง..." required class="w-full border border-orange-400 rounded p-1.5 bg-orange-50 mt-1 font-bold">\` : ''}
                        </div>
                        <div>
                            <select onchange="tempEquipmentList[\${index}].brand = this.value; renderEquipmentRows();" class="w-full border rounded p-1.5 bg-white font-bold">
                                \${brands.map(b => \`<option \${eq.brand === b ? 'selected' : ''}>\${b}</option>\`).join('')}
                            </select>
                            \${eq.brand.startsWith('ETC.') ? \`<input type="text" value="\${eq.customBrand || ''}" oninput="tempEquipmentList[\${index}].customBrand = this.value" placeholder="ระบุแบรนด์เอง..." required class="w-full border border-orange-400 rounded p-1.5 bg-orange-50 mt-1 font-bold">\` : ''}
                        </div>
                        <input type="text" value="\${eq.model}" oninput="tempEquipmentList[\${index}].model = this.value" placeholder="รุ่น (Model)..." class="border rounded p-1.5 bg-white">
                        <input type="text" value="\${eq.symptom}" oninput="tempEquipmentList[\${index}].symptom = this.value" placeholder="อาการเสีย *" required class="border rounded p-1.5 bg-white">
                    </div>

                    <!-- รูปภาพก่อนและหลังซ่อมต่อ Equipment (บีบอัดอัตโนมัติ) -->
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-2 border-t border-slate-200">
                        <div class="bg-white p-2 rounded-lg border space-y-1.5">
                            <div class="flex justify-between items-center">
                                <span class="font-bold text-orange-700"><i class="fa-solid fa-camera mr-1"></i> รูปก่อนซ่อม (Before):</span>
                                <span class="text-[10px] text-slate-500">\${(eq.beforePhotos || []).length} รูป (บีบอัดอัตโนมัติ)</span>
                            </div>
                            <div class="flex flex-wrap items-center gap-1.5">
                                <label class="cursor-pointer bg-orange-600 hover:bg-orange-700 text-white px-2 py-1 rounded text-[11px] font-bold shadow-xs">
                                    <i class="fa-solid fa-camera mr-1"></i> ถ่ายรูป (กล้อง)
                                    <input type="file" accept="image/*" capture="environment" onchange="handleStandaloneEquipmentPhotoUpload(\${index}, 'beforePhotos', this.files)" class="hidden">
                                </label>
                                <label class="cursor-pointer bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-2 py-1 rounded text-[11px] font-bold shadow-xs">
                                    <i class="fa-solid fa-images mr-1"></i> อัลบั้ม / ไฟล์
                                    <input type="file" accept="image/*" multiple onchange="handleStandaloneEquipmentPhotoUpload(\${index}, 'beforePhotos', this.files)" class="hidden">
                                </label>
                            </div>
                            \${(eq.beforePhotos || []).length > 0 ? \`
                                <div class="flex flex-wrap gap-1.5 pt-1">
                                    \${eq.beforePhotos.map((p, pI) => \`
                                        <div class="relative w-12 h-12 rounded border border-orange-200 overflow-hidden shadow-xs">
                                            <img src="\${p}" class="w-full h-full object-cover">
                                            <button type="button" onclick="removeStandaloneEquipmentPhoto(\${index}, 'beforePhotos', \${pI})" class="absolute top-0 right-0 bg-red-600 text-white rounded-full w-3.5 h-3.5 flex items-center justify-center text-[9px] font-bold cursor-pointer" title="ลบรูป">&times;</button>
                                        </div>
                                    \`).join('')}
                                </div>
                            \` : ''}
                        </div>
                        <div class="bg-white p-2 rounded-lg border space-y-1.5">
                            <div class="flex justify-between items-center">
                                <span class="font-bold text-emerald-700"><i class="fa-solid fa-camera mr-1"></i> รูปหลังซ่อม (After):</span>
                                <span class="text-[10px] text-slate-500">\${(eq.afterPhotos || []).length} รูป (บีบอัดอัตโนมัติ)</span>
                            </div>
                            <div class="flex flex-wrap items-center gap-1.5">
                                <label class="cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white px-2 py-1 rounded text-[11px] font-bold shadow-xs">
                                    <i class="fa-solid fa-camera mr-1"></i> ถ่ายรูป (กล้อง)
                                    <input type="file" accept="image/*" capture="environment" onchange="handleStandaloneEquipmentPhotoUpload(\${index}, 'afterPhotos', this.files)" class="hidden">
                                </label>
                                <label class="cursor-pointer bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-2 py-1 rounded text-[11px] font-bold shadow-xs">
                                    <i class="fa-solid fa-images mr-1"></i> อัลบั้ม / ไฟล์
                                    <input type="file" accept="image/*" multiple onchange="handleStandaloneEquipmentPhotoUpload(\${index}, 'afterPhotos', this.files)" class="hidden">
                                </label>
                            </div>
                            \${(eq.afterPhotos || []).length > 0 ? \`
                                <div class="flex flex-wrap gap-1.5 pt-1">
                                    \${eq.afterPhotos.map((p, pI) => \`
                                        <div class="relative w-12 h-12 rounded border border-emerald-200 overflow-hidden shadow-xs">
                                            <img src="\${p}" class="w-full h-full object-cover">
                                            <button type="button" onclick="removeStandaloneEquipmentPhoto(\${index}, 'afterPhotos', \${pI})" class="absolute top-0 right-0 bg-red-600 text-white rounded-full w-3.5 h-3.5 flex items-center justify-center text-[9px] font-bold cursor-pointer" title="ลบรูป">&times;</button>
                                        </div>
                                    \`).join('')}
                                </div>
                            \` : ''}
                        </div>
                    </div>
                </div>
            \`).join('');
        }

        function submitNewTicket(event) {
            event.preventDefault();
            let contactVal = document.getElementById('new-contact').value;
            let docNoVal = document.getElementById('new-docno').value;
            let serviceTypeVal = document.getElementById('new-servicetype').value;
            let warrantyVal = document.getElementById('new-warranty').value;
            let urgencyVal = document.getElementById('new-urgency').value;
            let locationVal = document.getElementById('new-location').value;
            let fileInput = document.getElementById('new-image-file');
            let imgUrl = "";
            if (fileInput.files && fileInput.files[0]) {
                imgUrl = URL.createObjectURL(fileInput.files[0]);
            }

            const resolvedEquip = tempEquipmentList.map(eq => {
                let finalType = eq.type.startsWith('ETC.') && eq.customType.trim() ? eq.customType.trim().toUpperCase() : eq.type;
                let finalBrand = eq.brand.startsWith('ETC.') && eq.customBrand.trim() ? eq.customBrand.trim().toUpperCase() : eq.brand;
                if (eq.type.startsWith('ETC.') && eq.customType.trim()) customEquipList.push(finalType);
                if (eq.brand.startsWith('ETC.') && eq.customBrand.trim()) customBrandList.push(finalBrand);
                return {
                    type: finalType,
                    brand: finalBrand,
                    model: eq.model || '-',
                    tag: eq.tag || 'N/A',
                    symptom: eq.symptom,
                    detail: '',
                    beforePhotos: (eq.beforePhotos && eq.beforePhotos.length > 0) ? eq.beforePhotos : (imgUrl ? [imgUrl] : []),
                    afterPhotos: eq.afterPhotos || [],
                    analysis: "",
                    cause: "",
                    standardsReference: "",
                    faultOrigin: warrantyVal.includes('In-Warranty') ? 'เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)' : 'เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)',
                    correctiveAction: "",
                    parts: [],
                    partsApproved: false,
                    checklistResults: getChecklistForType(finalType).map(item => ({ item, beforeStatus: 'NG', afterStatus: 'OK' }))
                };
            });

            const newTicket = {
                id: \`SKM-69-\${String(state.tickets.length + 1).padStart(3, '0')}\`,
                company: document.getElementById('new-company').value,
                taxId: document.getElementById('new-taxid').value,
                docNo: docNoVal,
                contact: contactVal,
                siteContact: contactVal,
                location: locationVal,
                warranty: warrantyVal,
                serviceType: serviceTypeVal,
                urgency: urgencyVal,
                desiredDate: document.getElementById('new-date').value,
                confirmedDate: '',
                status: "NEW",
                isLocked: false,
                rejectReason: "",
                createdAt: new Date().toISOString().replace('T', ' ').substring(0, 16),
                equipmentList: resolvedEquip,
                assignedTech: "",
                approvedBySupervisor: ""
            };
            state.tickets.unshift(newTicket);
            const lineMsg = formatTicketLineMessage(newTicket, '🔔 [มีใบแจ้งซ่อมใหม่! (NEW) - สหคิมมอเตอร์]');
            persistState(lineMsg);
            tempEquipmentList = [];
            alert('สร้างใบแจ้งซ่อมสำเร็จ! ระบบบันทึกข้อมูลและส่งแจ้งเตือนไปยังทีมบริการแล้ว');
            switchCustTab('list');
            render();
        }

        function customerAcceptJob(id) {
            const t = state.tickets.find(x => x.id === id);
            if (t) {
                t.status = "CLOSE";
                t.isLocked = true;
                persistState();
                alert('ตรวจรับงานและปิดจบเรียบร้อย!');
                render();
            }
        }

        function renderInternalLogin() {
            const hasTechs = state.technicians && state.technicians.length > 0;
            return \`
                <div class="max-w-md mx-auto bg-white p-6 sm:p-8 rounded-2xl shadow-xl border mt-6">
                    <div class="text-center mb-5">
                        <img src="${EMBEDDED_SKM_MASCOT}" alt="เจ้าหน้าที่ภายใน SKM SERVICE" class="w-20 h-20 rounded-2xl mx-auto object-cover border-2 border-orange-500 bg-slate-900 shadow mb-2" />
                        <h2 class="text-xl font-bold text-slate-900">เข้าสู่ระบบภายใน</h2>
                        <p class="text-xs text-slate-500">บริษัท สหคิมมอเตอร์ จำกัด</p>
                    </div>
                    <form onsubmit="app.loginInternalSubmit(event)" class="space-y-3 text-xs sm:text-sm">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">เลือกบทบาท *</label>
                            <select id="emp-role" onchange="const row = document.getElementById('tech-selector-row'); if(row) row.style.display = this.value === 'TECHNICIAN' ? 'block' : 'none';" class="w-full border rounded-xl p-2 bg-slate-50 font-bold text-slate-800">
                                <option value="SERVICE TEAM">SERVICE TEAM (รับงาน, จัดการคิว/นัดวัน)</option>
                                <option value="TECHNICIAN">TECHNICIAN (ช่าง: รับงาน, วิเคราะห์ AI, รูปถ่าย, อะไหล่)</option>
                                <option value="SUPERVISOR">SUPERVISOR (จัดการพนักงาน, อนุมัติ, Dashboard)</option>
                            </select>
                        </div>
                        <div id="tech-selector-row" style="display:none;" class="space-y-1">
                            <label class="block font-bold text-slate-700 mb-1">เลือกชื่อช่างของคุณ *</label>
                            \${hasTechs ? \`
                                <select id="tech-user-select" class="w-full border-2 border-teal-500 rounded-xl p-2 bg-teal-50 font-bold text-teal-900">
                                    \${state.technicians.map(tc => \`<option value="\${tc.id}">\${tc.name} (\${tc.position || 'ช่างเทคนิค'}) [PIN: \${tc.pin || '1111'}]</option>\`).join('')}
                                </select>
                            \` : \`
                                <div class="bg-amber-50 border border-amber-300 p-2 rounded-xl text-amber-900 text-xs">
                                    ยังไม่มีรายชื่อช่างในระบบ คุณสามารถกรอกรหัส Supervisor (admin123) เพื่อเข้าไปเพิ่มช่างได้
                                </div>
                            \`}
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">รหัสพนักงาน (PIN) หรือรหัสผ่าน Supervisor *</label>
                            <input type="password" id="emp-pin" placeholder="กรอกรหัสผ่าน / PIN เจ้าหน้าที่" required class="w-full border rounded-xl p-2 text-center tracking-widest font-bold bg-white">
                        </div>
                        <button type="submit" class="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold py-2.5 rounded-xl shadow transition cursor-pointer">เข้าสู่ระบบ</button>
                    </form>
                </div>
            \`;
        }

        function switchInternalTab(tab) {
            internalActiveTab = tab;
            render();
            if (tab === 'new_ticket' && tempEquipmentList.length === 0) {
                addEquipmentRow();
            }
        }

        function renderInternalDashboard() {
            let totalTickets = state.tickets.length;
            let closedTickets = state.tickets.filter(t => t.status === 'CLOSE').length;
            let focCount = state.tickets.filter(t => t.warranty && t.warranty.includes('In-Warranty')).length;
            let chargeCount = totalTickets - focCount;
            let totalMoney = 0;
            let equipCountMap = {};
            state.tickets.forEach(t => {
                t.equipmentList.forEach(eq => {
                    equipCountMap[eq.type] = (equipCountMap[eq.type] || 0) + 1;
                    if (eq.parts) eq.parts.forEach(p => totalMoney += (p.total || 0));
                });
            });
            let overdueTickets = state.tickets.filter(t => isTicketOverdue(t));
            let overdueCount = overdueTickets.length;
            let onTimeCount = Math.max(0, totalTickets - overdueCount);
            let topEquip = Object.keys(equipCountMap).sort((a,b) => equipCountMap[b] - equipCountMap[a])[0] || "BOOSTER PUMP";

            const roleBannerBg = state.userRole === 'SERVICE TEAM'
                ? 'bg-gradient-to-r from-blue-950 via-slate-900 to-sky-950 border-l-8 border-blue-400'
                : state.userRole === 'TECHNICIAN'
                ? 'bg-gradient-to-r from-teal-950 via-slate-900 to-emerald-950 border-l-8 border-teal-400'
                : 'bg-gradient-to-r from-indigo-950 via-slate-900 to-purple-950 border-l-8 border-indigo-400';
            const roleBadgeText = state.userRole === 'SERVICE TEAM'
                ? '🔵 โทนสีน้ำเงิน (Service Support / Dispatch Center)'
                : state.userRole === 'TECHNICIAN'
                ? '🟢 โทนสีเขียวมรกต-ทีล (Technician / Field Operations)'
                : '🟣 โทนสีม่วงคราม (Supervisor / Executive Approval)';

            return \`
                <div class="space-y-4">
                    <div class="\${roleBannerBg} text-white p-4 sm:p-6 rounded-2xl shadow-lg flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
                        <div>
                            <span class="bg-orange-500 text-white text-[10px] px-2.5 py-0.5 rounded-full font-bold">\${roleBadgeText}</span>
                            <h2 class="text-xl sm:text-2xl font-bold mt-1">บทบาทปัจจุบัน: \${state.userRole}</h2>
                            \${state.userRole === 'TECHNICIAN' && state.loggedInTech ? \`
                                <div class="mt-2 flex flex-wrap items-center gap-2">
                                    <span class="bg-teal-800/90 border border-teal-300 text-teal-100 text-xs px-3 py-1 rounded-full font-bold inline-flex items-center gap-1.5 shadow-sm">
                                        <span>👷 ช่างประจำตัว:</span>
                                        <strong class="underline text-white font-extrabold">\${state.loggedInTech.name}</strong>
                                        <span class="text-[11px] opacity-90">(\${state.loggedInTech.position || 'ช่างเทคนิค'})</span>
                                    </span>
                                    <button type="button" onclick="state.loggedInTech = null; state.currentView = 'internal_login'; render();" class="text-teal-200 hover:text-white underline text-xs font-semibold cursor-pointer">🔄 เปลี่ยนช่าง</button>
                                </div>
                            \` : ''}
                        </div>
                        <div class="flex flex-wrap items-center gap-2 text-xs">
                            <div class="flex items-center bg-slate-800/90 p-1 rounded-xl border border-slate-600 gap-1">
                                <button onclick="switchRoleWithAuth('SERVICE TEAM')" class="px-2.5 py-1 rounded-lg font-bold cursor-pointer \${state.userRole==='SERVICE TEAM'?'bg-blue-600 text-white':'text-slate-300'}">SERVICE TEAM</button>
                                <button onclick="switchRoleWithAuth('TECHNICIAN')" class="px-2.5 py-1 rounded-lg font-bold cursor-pointer \${state.userRole==='TECHNICIAN'?'bg-teal-600 text-white':'text-slate-300'}">TECHNICIAN</button>
                                <button onclick="switchRoleWithAuth('SUPERVISOR')" class="px-2.5 py-1 rounded-lg font-bold cursor-pointer \${state.userRole==='SUPERVISOR'?'bg-indigo-600 text-white':'text-slate-300'}">SUPERVISOR</button>
                            </div>
                            \${state.userRole === 'SUPERVISOR' ? \`
                                <button onclick="syncAllTicketsUpdateToLine()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg font-bold shadow cursor-pointer"><i class="fa-solid fa-cloud-arrow-up mr-1"></i> อัปเดตลง Sheet</button>
                                \${overdueCount > 0 ? \`<button onclick="notifyOverdueTicketsToLine()" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-lg font-bold shadow cursor-pointer"><i class="fa-solid fa-triangle-exclamation mr-1"></i> งานค้าง (\${overdueCount})</button>\` : ''}
                                <button onclick="openTechnicianManager()" class="bg-orange-500 text-white px-3 py-1.5 rounded-lg font-bold shadow cursor-pointer"><i class="fa-solid fa-users-gear mr-1"></i> จัดการช่าง</button>
                            \` : ''}
                            <button onclick="exportToExcel()" class="bg-emerald-700 text-white px-3 py-1.5 rounded-lg font-bold shadow cursor-pointer"><i class="fa-solid fa-file-excel mr-1"></i> Export Excel</button>
                        </div>
                    </div>

                    <!-- INTERNAL NAVIGATION MENU TABS (กดเลือกเมนูเข้าใช้งานได้ทันที) -->
                    <div class="bg-white p-3.5 sm:p-4 rounded-2xl border-2 border-slate-200 shadow-sm space-y-2.5">
                        <div class="flex flex-wrap items-center justify-between gap-2">
                            <p class="text-xs font-bold text-slate-800">📌 เมนูเข้าใช้งานสำหรับเจ้าหน้าที่ภายใน (คลิกเลือกหมวดที่ต้องการดู):</p>
                            <div class="flex flex-wrap items-center gap-2">
                                \${state.userRole === 'SUPERVISOR' ? \`
                                    <button type="button" onclick="openMasterDataModal('brands')" class="px-3 py-1 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer">⚙️ เพิ่ม/แก้ไขข้อมูลตั้งต้น (ยี่ห้อ · เครื่องจักร · เบอร์โทร · ช่าง)</button>
                                \` : ''}
                                <button type="button" onclick="switchInternalTab(internalActiveTab === 'all' ? 'tickets' : 'all')" class="px-3 py-1 rounded-lg text-xs font-bold border cursor-pointer \${internalActiveTab === 'all' ? 'bg-slate-900 text-white border-slate-900' : 'bg-slate-100 text-slate-700 border-slate-300'}">
                                    \${internalActiveTab === 'all' ? '✓ กำลังแสดงทุกหมวดในหน้าเดียว' : '👁️ แสดงรวมทุกหมวดในหน้าเดียว'}
                                </button>
                            </div>
                        </div>
                        <div class="grid grid-cols-2 lg:grid-cols-4 gap-2.5">
                            <button type="button" onclick="switchInternalTab('tickets')" class="p-3.5 rounded-xl border-2 text-left transition flex items-center justify-between gap-2 cursor-pointer \${internalActiveTab === 'tickets' ? 'bg-slate-900 border-slate-900 text-white shadow' : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-800'}">
                                <div>
                                    <p class="text-xs sm:text-sm font-bold">1. 📋 รายการใบงานซ่อม</p>
                                    <p class="text-[11px] \${internalActiveTab === 'tickets' ? 'text-orange-300' : 'text-slate-500'}">รับงาน, วิเคราะห์ AI & อนุมัติงาน</p>
                                </div>
                                <span class="px-2.5 py-1 rounded-full text-xs font-mono font-bold \${internalActiveTab === 'tickets' ? 'bg-orange-500 text-white' : 'bg-slate-200 text-slate-800'}">\${totalTickets}</span>
                            </button>
                            <button type="button" onclick="switchInternalTab('new_ticket')" class="p-3.5 rounded-xl border-2 text-left transition flex items-center justify-between gap-2 cursor-pointer \${internalActiveTab === 'new_ticket' ? 'bg-orange-500 border-orange-600 text-white shadow' : 'bg-orange-50/70 hover:bg-orange-100 border-orange-200 text-slate-900'}">
                                <div>
                                    <p class="text-xs sm:text-sm font-bold">2. ➕ แจ้งซ่อมใหม่</p>
                                    <p class="text-[11px] \${internalActiveTab === 'new_ticket' ? 'text-orange-100' : 'text-slate-600'}">เปิดใบงานแจ้งซ่อมใหม่เข้าระบบ</p>
                                </div>
                                <i class="fa-solid fa-plus text-base"></i>
                            </button>
                            <button type="button" onclick="switchInternalTab('schedule')" class="p-3.5 rounded-xl border-2 text-left transition flex items-center justify-between gap-2 cursor-pointer \${internalActiveTab === 'schedule' ? 'bg-blue-600 border-blue-700 text-white shadow' : 'bg-blue-50/70 hover:bg-blue-100 border-blue-200 text-slate-900'}">
                                <div>
                                    <p class="text-xs sm:text-sm font-bold">3. 📅 ตารางงานช่าง</p>
                                    <p class="text-[11px] \${internalActiveTab === 'schedule' ? 'text-blue-100' : 'text-slate-600'}">เช็คคิวว่าง แดง/เหลือง/เขียว</p>
                                </div>
                                <span class="text-xs">🟢🟡🔴</span>
                            </button>
                            <button type="button" onclick="switchInternalTab('dashboard')" class="p-3.5 rounded-xl border-2 text-left transition flex items-center justify-between gap-2 cursor-pointer \${internalActiveTab === 'dashboard' ? 'bg-indigo-600 border-indigo-700 text-white shadow' : 'bg-indigo-50/70 hover:bg-indigo-100 border-indigo-200 text-slate-900'}">
                                <div>
                                    <p class="text-xs sm:text-sm font-bold">4. 📊 Dashboard</p>
                                    <p class="text-[11px] \${internalActiveTab === 'dashboard' ? 'text-indigo-100' : 'text-slate-600'}">สรุปยอดงาน, ต้นทุน & SLA</p>
                                </div>
                                <i class="fa-solid fa-chart-column text-base"></i>
                            </button>
                        </div>
                    </div>

                    \${internalActiveTab === 'new_ticket' ? \`
                        <div class="bg-white p-4 sm:p-6 rounded-2xl shadow border-2 border-orange-300 space-y-4">
                            <div class="flex justify-between items-center border-b pb-2">
                                <h3 class="text-base font-bold text-slate-900"><i class="fa-solid fa-file-circle-plus text-orange-500 mr-2"></i> เปิดใบแจ้งซ่อมใหม่เข้าระบบ (Service Request)</h3>
                                <button type="button" onclick="switchInternalTab('tickets')" class="bg-slate-100 text-slate-700 px-3 py-1 rounded-lg text-xs font-bold cursor-pointer">ดูรายการใบงาน (\${totalTickets})</button>
                            </div>
                            <form onsubmit="submitNewTicket(event)" class="space-y-4 text-xs sm:text-sm">
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                    <div><label class="block font-bold text-slate-700 mb-1">ชื่อบริษัทลูกค้า *</label><input type="text" id="new-company" value="\${state.customerCompany}" required class="w-full border rounded-xl p-2 bg-slate-50"></div>
                                    <div><label class="block font-bold text-slate-700 mb-1">TAX ID</label><input type="text" id="new-taxid" placeholder="0105552012345" class="w-full border rounded-xl p-2"></div>
                                    <div><label class="block font-bold text-slate-700 mb-1">เลขที่เอกสาร / ใบรับประกัน</label><input type="text" id="new-docno" placeholder="เช่น DOC-2026-999" class="w-full border rounded-xl p-2"></div>
                                    <div><label class="block font-bold text-slate-700 mb-1">ผู้ติดต่อและเบอร์โทร *</label><input type="text" id="new-contact" placeholder="คุณสมชาย (081-234-5678)" required class="w-full border rounded-xl p-2"></div>
                                    <div>
                                        <label class="block font-bold text-slate-700 mb-1">เงื่อนไขการรับประกัน *</label>
                                        <select id="new-warranty" class="w-full border rounded-xl p-2 bg-white font-bold text-orange-600">
                                            <option>In-Warranty (อยู่ในประกัน - ฟรี / FOC)</option>
                                            <option>Out-of-Warranty (หมดประกัน - คิดค่าบริการ / Charge)</option>
                                        </select>
                                    </div>
                                    <div>
                                        <label class="block font-bold text-slate-700 mb-1">ประเภทบริการ *</label>
                                        <select id="new-servicetype" class="w-full border rounded-xl p-2 bg-white font-bold text-slate-700">
                                            <option>ซ่อมหน้างาน</option>
                                            <option>รับมาซ่อม</option>
                                            <option>คอมมิชชั่นนิ่ง (Commissioning)</option>
                                            <option>เทสรัน (Test Run)</option>
                                            <option>PM (Preventive Maintenance)</option>
                                        </select>
                                    </div>
                                    <div>
                                        <label class="block font-bold text-slate-700 mb-1">ความเร่งด่วนของงาน *</label>
                                        <select id="new-urgency" class="w-full border rounded-xl p-2 bg-white font-bold text-red-600">
                                            <option>ปกติ (Normal)</option>
                                            <option>ด่วน (High)</option>
                                            <option>ด่วนที่สุด (Urgent)</option>
                                        </select>
                                    </div>
                                    <div><label class="block font-bold text-slate-700 mb-1">สถานที่ปฏิบัติงาน *</label><input type="text" id="new-location" placeholder="ระบุสถานที่หน้างาน" required class="w-full border rounded-xl p-2"></div>
                                    <div><label class="block font-bold text-slate-700 mb-1">วันที่นัดหมาย/ขอช่าง *</label><input type="date" id="new-date" value="\${selectedCustomerDate}" required class="w-full border rounded-xl p-2 font-mono"></div>
                                    <div><label class="block font-bold text-slate-700 mb-1">รูปภาพอาการเสีย</label><input type="file" id="new-image-file" accept="image/*" class="w-full border rounded-xl p-2 bg-slate-50 text-xs"></div>
                                </div>
                                <div class="border-t pt-3">
                                    <div class="flex justify-between items-center mb-2">
                                        <h4 class="font-bold text-slate-900 text-xs">รายการอุปกรณ์</h4>
                                        <button type="button" onclick="addEquipmentRow()" class="bg-slate-900 text-white text-xs px-3 py-1.5 rounded-lg font-bold cursor-pointer">+ เพิ่มอุปกรณ์</button>
                                    </div>
                                    <div id="equipment-container" class="space-y-3"></div>
                                </div>
                                <div class="flex justify-end gap-2 pt-3 border-t">
                                    <button type="button" onclick="switchInternalTab('tickets')" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer">ยกเลิก</button>
                                    <button type="submit" onclick="setTimeout(() => { internalActiveTab = 'tickets'; render(); }, 100);" class="bg-orange-500 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">บันทึกและเปิดใบแจ้งซ่อม</button>
                                </div>
                            </form>
                        </div>
                    \` : ''}

                    <!-- SERVICE SUPPORT: TECHNICIAN AVAILABILITY CALENDAR -->
                    \${(internalActiveTab === 'schedule' || internalActiveTab === 'all') ? \`
                    <div class="bg-white p-4 sm:p-6 rounded-2xl shadow border-2 border-slate-200 space-y-4">
                        <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-3 border-b pb-3">
                            <div>
                                <span class="bg-orange-500 text-white text-[10px] font-bold px-2.5 py-0.5 rounded-full">SERVICE SUPPORT · TECHNICIAN AVAILABILITY CALENDAR</span>
                                <h3 class="text-base sm:text-lg font-bold text-slate-900 mt-1">ตารางงานช่างทั้งหมด (เลื่อนดูสัปดาห์ก่อนหน้า/ถัดไปได้) แยกสี แดง / เหลือง / เขียว</h3>
                                <p class="text-xs text-slate-600">มีสิทธิ์เปลี่ยนหรือกำหนดวันรับงานช่าง เลือกวันและเลือกช่างตามความเหมาะสมและว่าง · <b class="text-emerald-700">ถ้าช่างยังไม่กดตอบรับงาน สามารถกด "รับงานแทนช่างเลย" หรือ "เปลี่ยนช่างใหม่" ได้ทันที</b></p>
                            </div>
                            <div class="flex flex-wrap gap-2 text-xs">
                                <span class="px-3 py-1.5 rounded-lg bg-red-50 border border-red-300 text-red-900 font-bold">🔴 สีแดง = เต็มมาก (คิวแน่นมาก)</span>
                                <span class="px-3 py-1.5 rounded-lg bg-amber-50 border border-amber-300 text-amber-900 font-bold">🟡 สีเหลือง = รับงานได้</span>
                                <span class="px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-900 font-bold">🟢 สีเขียว = ว่าง</span>
                            </div>
                        </div>
                        <div class="flex flex-wrap items-center justify-between gap-2 text-xs">
                            <div class="flex items-center gap-1.5">
                                <button type="button" onclick="supportScheduleOffset -= 7; render();" class="px-3 py-1.5 rounded-lg border bg-slate-50 hover:bg-slate-100 font-bold cursor-pointer">← สัปดาห์ก่อนหน้า (7 วัน)</button>
                                <button type="button" onclick="supportScheduleOffset = 0; supportSelectedDate = new Date().toISOString().slice(0,10); render();" class="px-3 py-1.5 rounded-lg border bg-white hover:bg-slate-50 font-bold cursor-pointer">วันนี้</button>
                                <button type="button" onclick="supportScheduleOffset += 7; render();" class="px-3 py-1.5 rounded-lg border bg-slate-50 hover:bg-slate-100 font-bold cursor-pointer">สัปดาห์ถัดไป (7 วัน) →</button>
                            </div>
                            <span class="text-slate-500 font-semibold">คลิกที่วันในตารางเพื่อดูช่างที่ว่าง กำหนดวันนัดหมาย หรือกดรับงานแทนช่าง</span>
                        </div>
                        <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5">
                            \${getSupportScheduleDays().map(day => {
                                const isSel = supportSelectedDate === day.dateStr;
                                const cardCls = day.level === 'RED'
                                    ? 'bg-red-50 border-red-400 text-red-950 hover:bg-red-100'
                                    : day.level === 'YELLOW'
                                    ? 'bg-amber-50 border-amber-400 text-amber-950 hover:bg-amber-100'
                                    : 'bg-emerald-50 border-emerald-400 text-emerald-950 hover:bg-emerald-100';
                                const badgeCls = day.level === 'RED'
                                    ? 'bg-red-600 text-white'
                                    : day.level === 'YELLOW'
                                    ? 'bg-amber-500 text-white'
                                    : 'bg-emerald-600 text-white';
                                return \`
                                    <button type="button" onclick="supportSelectedDate = '\${day.dateStr}'; render();" class="p-3 rounded-xl border-2 text-left transition flex flex-col justify-between gap-2 cursor-pointer \${cardCls} \${isSel ? 'ring-2 ring-slate-900 shadow-md' : ''}">
                                        <div>
                                            <p class="text-xs font-extrabold">\${day.dayLabel}</p>
                                            <p class="text-[10px] font-mono opacity-75">\${day.dateStr}</p>
                                        </div>
                                        <div class="space-y-1">
                                            <div class="flex items-center justify-between gap-1">
                                                <span class="inline-block text-[10px] font-bold px-2 py-0.5 rounded \${badgeCls}">\${day.statusText}</span>
                                                <span class="text-[10px] font-bold">ว่าง \${day.freeTechCount}/\${state.technicians.length}</span>
                                            </div>
                                            <p class="text-[11px] font-semibold">\${day.detailText}</p>
                                        </div>
                                    </button>
                                \`;
                            }).join('')}
                        </div>
                        \${(() => {
                            const selInfo = getStandaloneDateInfo(supportSelectedDate);
                            return \`
                                <div class="bg-slate-900 text-white p-4 rounded-xl space-y-3 text-xs">
                                    <div class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-700 pb-2">
                                        <span class="font-bold text-orange-400">📅 คิวช่างประจำวันที่เลือก: \${selInfo.dateStr} (\${selInfo.statusText} · \${selInfo.detailText})</span>
                                        <span>ช่างว่างพร้อมรับงาน \${selInfo.freeTechCount}/\${state.technicians.length} คน</span>
                                    </div>
                                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                                        \${selInfo.techAvailability.map(tc => \`
                                            <div class="bg-slate-800 border border-slate-700 p-3 rounded-xl flex flex-col gap-2">
                                                <div class="flex items-center justify-between gap-2">
                                                    <div>
                                                        <p class="font-bold text-white text-xs">\${tc.name} <span class="text-[10px] text-slate-400">(\${tc.position || 'ช่างเทคนิค'})</span></p>
                                                        <p class="text-[10px] text-slate-400">\${tc.specialization}</p>
                                                    </div>
                                                    <span class="text-[10px] font-bold px-2 py-0.5 rounded-full \${tc.techLevel === 'RED' ? 'bg-red-600 text-white' : tc.techLevel === 'YELLOW' ? 'bg-amber-500 text-white' : 'bg-emerald-600 text-white'}">\${tc.availabilityLabel}</span>
                                                </div>
                                                \${tc.assignedTickets.length > 0 ? \`
                                                    <div class="bg-slate-900/90 p-2 rounded-lg border border-slate-700 space-y-1">
                                                        <p class="text-[10px] font-bold text-amber-300">งานที่รับในวันนี้ (\${tc.assignedTickets.length} งาน):</p>
                                                        \${tc.assignedTickets.map(t => \`
                                                            <div class="flex items-center justify-between gap-1 text-[11px] bg-slate-950 p-1.5 rounded">
                                                                <span class="font-bold font-mono text-orange-400 mr-1">\${t.id}</span>
                                                                <span class="text-white truncate max-w-[100px]">\${t.company}</span>
                                                                <div class="flex items-center gap-1 shrink-0">
                                                                    <button onclick="unassignTicket('\${t.id}')" class="bg-amber-600 hover:bg-amber-500 text-white px-1.5 py-0.5 rounded text-[9px] font-bold cursor-pointer" title="ปลดช่างออกจากงานนี้ คิวจะว่างลงทันที"><i class="fa-solid fa-user-xmark mr-0.5"></i> ปลดช่าง</button>
                                                                    <button onclick="deleteTicket('\${t.id}')" class="bg-red-600 hover:bg-red-500 text-white px-1.5 py-0.5 rounded text-[9px] font-bold cursor-pointer" title="ลบใบงานนี้ออกจากระบบ"><i class="fa-solid fa-trash mr-0.5"></i> ลบ</button>
                                                                </div>
                                                            </div>
                                                        \`).join('')}
                                                    </div>
                                                \` : \`
                                                    <p class="text-[10px] text-emerald-400">✓ คิวว่าง ไม่มีงานค้างในวันนี้</p>
                                                \`}
                                            </div>
                                        \`).join('')}
                                    </div>
                                </div>
                            \`;
                        })()}
                    </div>
                    \` : ''}

                    \${(internalActiveTab === 'dashboard' || internalActiveTab === 'all') ? \`
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                        <div class="bg-white p-3 rounded-xl shadow border">
                            <p class="text-slate-500 font-bold">งานทั้งหมดต่อเดือน</p>
                            <h3 class="text-xl font-extrabold text-slate-900 mt-1">\${totalTickets} งาน (ปิดแล้ว \${closedTickets})</h3>
                        </div>
                        <div class="bg-white p-3 rounded-xl shadow border">
                            <p class="text-slate-500 font-bold">ประเภทงาน FOC / คิดเงิน</p>
                            <h3 class="text-sm font-bold text-indigo-600 mt-1">FOC: \${focCount} | คิดเงิน: \${chargeCount}</h3>
                            <p class="text-[10px] text-slate-500">ยอดอะไหล่รวม: \${totalMoney.toLocaleString()} THB</p>
                        </div>
                        <div class="bg-white p-3 rounded-xl shadow border">
                            <p class="text-slate-500 font-bold">สินค้าที่เสียบ่อยสุด</p>
                            <h3 class="text-sm font-bold text-orange-600 mt-1">\${topEquip}</h3>
                        </div>
                        <div class="bg-white p-3 rounded-xl shadow border">
                            <p class="text-slate-500 font-bold">เวลาการทำงาน (SLA วันนัดปัจจุบัน)</p>
                            <h3 class="text-sm font-bold text-emerald-600 mt-1">ตรงเวลา: \${onTimeCount} | <span class="text-rose-600">เกินกำหนด: \${overdueCount}</span></h3>
                        </div>
                    </div>
                    \` : ''}

                    \${(() => {
                        const myTechName = state.loggedInTech ? state.loggedInTech.name : '';
                        const myAssignedCount = state.tickets.filter(t => myTechName && (t.assignedTech === myTechName || (t.assignedTech && t.assignedTech.includes(myTechName)))).length;
                        const availableJobsCount = state.tickets.filter(t => (!t.assignedTech || t.assignedTech === 'ยังไม่ระบุ' || t.assignedTech === '') && (t.status === 'NEW' || t.status === 'ASSIGNED')).length;

                        let displayedTickets = state.tickets;
                        if (state.userRole === 'TECHNICIAN') {
                            if (techWorkFilter === 'MY_JOBS') {
                                displayedTickets = state.tickets.filter(t => myTechName && (t.assignedTech === myTechName || (t.assignedTech && t.assignedTech.includes(myTechName))));
                            } else if (techWorkFilter === 'AVAILABLE_JOBS') {
                                displayedTickets = state.tickets.filter(t => (!t.assignedTech || t.assignedTech === 'ยังไม่ระบุ' || t.assignedTech === '') && (t.status === 'NEW' || t.status === 'ASSIGNED'));
                            }
                        }

                        return \`
                        <div class="bg-white p-4 sm:p-6 rounded-2xl shadow border space-y-3">
                            <div class="flex flex-wrap justify-between items-center gap-2 border-b pb-2">
                                <h3 class="text-base font-bold text-slate-900"><i class="fa-solid fa-list-check text-orange-500 mr-2"></i> จัดการรายการคำขอและสถานะ (Workflow)</h3>
                                <div class="flex flex-wrap gap-2">
                                    <button onclick="switchInternalTab('new_ticket')" class="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer">+ เปิดใบแจ้งซ่อมใหม่</button>
                                    \${state.userRole === 'SUPERVISOR' && overdueCount > 0 ? \`<button onclick="notifyOverdueTicketsToLine()" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-triangle-exclamation mr-1"></i> ส่งแจ้งเตือนงานค้าง (\${overdueCount} งาน) เข้ากลุ่ม LINE</button>\` : ''}
                                </div>
                            </div>

                            \${state.userRole === 'TECHNICIAN' ? \`
                                <div class="bg-gradient-to-r from-teal-50 via-emerald-50 to-teal-50 border-2 border-teal-300 rounded-2xl p-3.5 flex flex-col md:flex-row justify-between items-start md:items-center gap-3 shadow-xs">
                                    <div>
                                        <div class="flex items-center gap-2">
                                            <span class="bg-teal-700 text-white text-xs px-2.5 py-0.5 rounded-full font-bold">👷 มุมมองช่าง</span>
                                            <span class="font-bold text-slate-800 text-xs sm:text-sm">ช่าง: \${state.loggedInTech ? state.loggedInTech.name : 'ช่างเทคนิค'}</span>
                                        </div>
                                        <p class="text-[11px] text-slate-600 mt-0.5">เลือกดูงานที่มอบหมายให้ตนเอง หรือกดรับงานใหม่ที่ยังไม่มีช่าง</p>
                                    </div>
                                    <div class="flex flex-wrap items-center gap-1.5 text-xs">
                                        <button type="button" onclick="setTechWorkFilter('MY_JOBS')" class="px-3 py-1.5 rounded-xl font-bold transition cursor-pointer \${techWorkFilter === 'MY_JOBS' ? 'bg-teal-700 text-white shadow' : 'bg-white text-slate-700 border border-slate-300'}">
                                            📌 งานของฉัน (\${myAssignedCount})
                                        </button>
                                        <button type="button" onclick="setTechWorkFilter('AVAILABLE_JOBS')" class="px-3 py-1.5 rounded-xl font-bold transition cursor-pointer \${techWorkFilter === 'AVAILABLE_JOBS' ? 'bg-orange-600 text-white shadow' : 'bg-white text-slate-700 border border-slate-300'}">
                                            📥 งานรอช่างรับงาน (\${availableJobsCount})
                                        </button>
                                        <button type="button" onclick="setTechWorkFilter('ALL')" class="px-3 py-1.5 rounded-xl font-bold transition cursor-pointer \${techWorkFilter === 'ALL' ? 'bg-slate-800 text-white shadow' : 'bg-white text-slate-700 border border-slate-300'}">
                                            📋 งานทั้งหมด (\${state.tickets.length})
                                        </button>
                                    </div>
                                </div>
                            \` : ''}

                            <div class="space-y-3">
                                \${displayedTickets.length === 0 ? \`
                                    <div class="text-center py-8 bg-slate-50 border border-dashed border-slate-300 rounded-xl p-4 text-xs space-y-2">
                                        <p class="font-bold text-slate-700 text-sm">
                                            \${state.userRole === 'TECHNICIAN'
                                                ? (techWorkFilter === 'MY_JOBS'
                                                    ? 'ยังไม่มีใบงานที่มอบหมายให้ช่าง "' + (state.loggedInTech ? state.loggedInTech.name : 'ช่างเทคนิค') + '" ในขณะนี้'
                                                    : 'ไม่มีงานใหม่ที่รอช่างรับงานในขณะนี้')
                                                : 'ไม่พบรายการใบงาน'}
                                        </p>
                                        \${state.userRole === 'TECHNICIAN' && techWorkFilter === 'MY_JOBS' && availableJobsCount > 0 ? \`
                                            <p class="text-amber-800 font-semibold">
                                                💡 มีงานรอช่างรับงาน \${availableJobsCount} งานในระบบ!
                                                <button onclick="setTechWorkFilter('AVAILABLE_JOBS')" class="underline font-bold text-orange-600 cursor-pointer ml-1">คลิกที่นี่เพื่อกดรับงาน</button>
                                            </p>
                                        \` : ''}
                                    </div>
                                \` : displayedTickets.map(t => {
                                    const overdue = isTicketOverdue(t);
                                    const dueStr = extractTicketDueDate(t);
                                    const daysLate = overdue ? getOverdueDays(t) : 0;
                                    const isUnassigned = !t.assignedTech || t.assignedTech === 'ยังไม่ระบุ' || t.assignedTech === '';
                                    return \`
                                    <div class="bg-slate-50 p-3 sm:p-4 rounded-xl border \${overdue ? 'border-rose-300 bg-rose-50/30' : ''} flex flex-col lg:flex-row justify-between items-start lg:items-center gap-3">
                                        <div class="space-y-1 text-xs">
                                            <div class="flex items-center space-x-2 flex-wrap gap-1">
                                                <span class="font-bold text-orange-600">\${t.id}</span>
                                                \${getActionableStatusButton(t)}
                                                <span class="text-slate-700 font-bold">\${t.company}</span>
                                                \${overdue ? \`<span class="bg-rose-600 text-white text-[10px] px-2 py-0.5 rounded font-bold">⚠️ เกินกำหนดนัด \${daysLate} วัน (นัด \${dueStr})</span>\` : ''}
                                                \${t.isLocked ? '<span class="bg-red-100 text-red-600 text-[10px] px-2 py-0.5 rounded font-bold"><i class="fa-solid fa-lock mr-1"></i> ล็อกแล้ว</span>' : '<span class="bg-emerald-100 text-emerald-600 text-[10px] px-2 py-0.5 rounded font-bold"><i class="fa-solid fa-lock-open mr-1"></i> เปิดแก้ไข</span>'}
                                            </div>
                                            <p class="text-slate-500">วันนัดหมาย: <span class="font-bold \${overdue ? 'text-rose-600' : 'text-slate-800'}">\${t.confirmedDate || t.desiredDate || '-'}</span> | สถานที่: <span class="text-slate-800 font-bold">\${t.location || '-'}</span> | ช่าง: <span class="font-bold text-orange-600">\${t.assignedTech || 'ยังไม่ระบุ'}</span> \${t.leadTech ? \`| หัวหน้าช่าง: <span class="font-bold text-indigo-700">\${t.leadTech}</span>\` : ''} \${t.approvedBySupervisor ? \`| ผู้อนุมัติ: <span class="font-bold text-emerald-700">\${t.approvedBySupervisor}</span>\` : ''}</p>
                                        </div>
                                        <div class="flex flex-wrap items-center gap-1.5">
                                            \${state.userRole === 'TECHNICIAN' && isUnassigned ? \`
                                                <button onclick="techAcceptJob('\${t.id}')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer inline-flex items-center gap-1">
                                                    <i class="fa-solid fa-hand-holding-hand"></i> ✋ กดรับงานนี้ทันที
                                                </button>
                                            \` : ''}
                                            \${t.isLocked && state.userRole === 'SUPERVISOR' ? \`<button onclick="unlockTicket('\${t.id}')" class="bg-purple-600 text-white px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-key mr-1"></i> ปลดล็อก</button>\` : ''}
                                            <button onclick="shareTicketToLine('\${t.id}')" class="bg-slate-700 hover:bg-slate-800 text-white px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-copy mr-1"></i> คัดลอกสรุป</button>
                                            \${t.assignedTech ? \`<button onclick="unassignTicket('\${t.id}')" class="bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300 px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer" title="ปลดช่างออกจากใบงานนี้"><i class="fa-solid fa-user-xmark mr-1"></i> ปลดช่าง</button>\` : ''}
                                            <button onclick="openSignatureModal('\${t.id}', 'technician')" class="bg-blue-600 hover:bg-blue-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-signature mr-1"></i> \${t.technicianSignature ? '✓ ลายเซ็นช่าง' : '✍️ ลายเซ็นช่าง'}</button>
                                            <button onclick="openInternalActionModal('\${t.id}')" class="bg-orange-500 text-white px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer"><i class="fa-solid fa-pen-to-square mr-1"></i> ปฏิบัติงาน/วิเคราะห์ AI</button>
                                            <button onclick="openServiceReport('\${t.id}')" class="bg-slate-800 text-white px-2.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer"><i class="fa-solid fa-print mr-1"></i> รายงาน</button>
                                            <button onclick="deleteTicket('\${t.id}')" class="bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 px-2.5 py-1.5 rounded-lg text-xs font-bold shadow cursor-pointer" title="ลบใบงานนี้ออกจากระบบ"><i class="fa-solid fa-trash mr-1"></i> ลบ</button>
                                        </div>
                                    </div>
                                \`;}).join('')}
                            </div>
                        </div>
                        \`;
                    })()}
                    \` : ''}
                </div>
            \`;
        }

        function setTechWorkFilter(filter) {
            techWorkFilter = filter;
            render();
        }

        function techAcceptJob(ticketId) {
            const t = state.tickets.find(item => item.id === ticketId);
            if (!t) return;
            const techName = state.loggedInTech ? state.loggedInTech.name : (state.technicians[0] ? state.technicians[0].name : 'ช่างเทคนิค');
            t.assignedTech = techName;
            t.status = 'SCHEDULED';
            if (!t.confirmedDate) {
                t.confirmedDate = t.desiredDate || new Date().toISOString().split('T')[0];
            }
            const lineMsg = '👷 [ช่างรับงานแล้ว - SKM Portal]\\nใบงาน: ' + t.id + '\\nลูกค้า: ' + t.company + '\\nช่างผู้รับผิดชอบ: ' + techName + '\\nสถานะ: กำหนดวันนัดหมายแล้ว (' + t.confirmedDate + ')';
            persistState(lineMsg);
            techWorkFilter = 'MY_JOBS';
            render();
            showAppToast('✓ ช่าง ' + techName + ' รับงาน ' + t.id + ' สำเร็จแล้ว!', 'success');
        }

        function getApproverOptionsList() {
            let leadNames = state.technicians
                .filter(tc => (tc.position || '').includes('หัวหน้าช่าง'))
                .map(tc => tc.name + ' (หัวหน้าช่าง)');
            return Array.from(new Set([...(state.supervisors || []), ...leadNames]));
        }

        async function openTechnicianManager() {
            if (state.userRole !== 'SUPERVISOR') {
                const pass = await openCustomPrompt({
                    title: '🔒 สิทธิ์เฉพาะ SUPERVISOR เท่านั้น',
                    message: 'กรุณากรอกรหัสผ่าน Supervisor เพื่อจัดการรายชื่อและรหัสผ่าน:',
                    inputType: 'password',
                    placeholder: 'รหัสเริ่มต้นคือ admin123'
                });
                if (pass === savedSupPassword || pass === 'admin123') {
                    if (pass === 'admin123' && savedSupPassword !== 'admin123') {
                        savedSupPassword = 'admin123';
                        persistState();
                    }
                    state.userRole = 'SUPERVISOR';
                    render();
                } else {
                    if (pass !== null) showAppToast('รหัสผ่าน Supervisor ไม่ถูกต้อง!', 'error');
                    return;
                }
            }
            let supChips = (state.supervisors || []).map((supName, sIdx) => \`
                <span class="inline-flex items-center gap-1.5 bg-white border border-emerald-300 text-slate-800 px-2.5 py-1 rounded-lg font-bold text-xs">
                    <span>\${supName}</span>
                    \${state.supervisors.length > 1 ? \`<button onclick="removeSupervisorName(\${sIdx})" class="text-red-500 hover:text-red-700 font-bold cursor-pointer">&times;</button>\` : ''}
                </span>
            \`).join('');

            let techRows = state.technicians.length === 0 ? \`
                <div class="text-center py-6 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-xs space-y-1.5">
                    <div class="w-10 h-10 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto text-lg font-bold">✓</div>
                    <p class="font-bold text-sm text-slate-800">ลบรายชื่อช่าง DEMO ออกทั้งหมดเรียบร้อยแล้ว (ปัจจุบันมีช่าง 0 คน)</p>
                    <p class="text-slate-500">ระบบพร้อมรองรับข้อมูลช่างจริงของบริษัท สามารถเริ่มกรอกชื่อช่างคนแรกได้ที่ <b>ข้อ 4. ด้านล่าง</b> (ระบบจะรันรหัสเริ่มต้นเป็น TECH-01 ให้อัตโนมัติ)</p>
                </div>
            \` : state.technicians.map((tc, idx) => \`
                <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs space-y-2.5 shadow-sm">
                    <div class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200/80 pb-2">
                        <div class="flex items-center gap-2">
                            <span class="font-mono font-bold text-slate-600 bg-slate-200 px-2 py-0.5 rounded text-[11px]">\${tc.id}</span>
                            <span class="px-2 py-0.5 rounded text-[10px] font-bold \${(tc.position || '').includes('หัวหน้าช่าง') ? 'bg-indigo-600 text-white' : 'bg-teal-700 text-white'}">\${tc.position || 'ช่างเทคนิค'}</span>
                        </div>
                        <div class="flex items-center gap-1.5">
                            <span class="text-[11px] font-bold text-slate-600">ตำแหน่ง:</span>
                            <select id="edit-tech-pos-\${idx}" class="border rounded-lg px-2 py-1 text-[11px] font-semibold bg-white text-slate-800">
                                \${TECH_POSITIONS.map(p => \`<option value="\${p}" \${(tc.position || 'ช่างเทคนิค (Technician)') === p ? 'selected' : ''}>\${p}</option>\`).join('')}
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        <div>
                            <label class="block text-[10px] font-bold text-slate-600 mb-0.5">ชื่อ-นามสกุลช่าง *</label>
                            <input type="text" id="edit-tech-name-\${idx}" value="\${tc.name}" placeholder="ชื่อ-นามสกุลช่าง" class="w-full border rounded-lg px-2.5 py-1.5 font-bold text-slate-900 bg-white">
                        </div>
                        <div>
                            <label class="block text-[10px] font-bold text-slate-600 mb-0.5">ความเชี่ยวชาญเฉพาะทาง</label>
                            <input type="text" id="edit-tech-spec-\${idx}" value="\${tc.specialization || ''}" placeholder="เช่น PUMP & MOTOR" class="w-full border rounded-lg px-2.5 py-1.5 text-slate-700 bg-white">
                        </div>
                    </div>

                    <div class="flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-slate-200/60">
                        <div class="flex items-center gap-1.5">
                            <span class="text-[11px] font-bold text-slate-700">รหัส PIN ช่าง:</span>
                            <input type="text" id="edit-tech-pin-\${idx}" value="\${tc.pin || '1111'}" class="w-20 border rounded-lg px-2 py-1 font-mono font-bold text-center text-orange-600 bg-white" title="PIN ประจำตัวช่าง">
                        </div>
                        <div class="flex items-center gap-1.5">
                            <button type="button" onclick="saveTechnicianRow(\${idx})" class="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg font-bold text-xs cursor-pointer shadow-sm">
                                💾 บันทึก
                            </button>
                            <button type="button" onclick="confirmRemoveTechnician(\${idx}, this)" class="bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded-lg font-bold text-xs cursor-pointer shadow-sm">
                                🗑️ ลบ
                            </button>
                        </div>
                    </div>
                </div>
            \`).join('');

            let modalBg = document.createElement('div');
            modalBg.id = 'tech-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-3 overflow-y-auto';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-xl rounded-2xl shadow-2xl p-5 space-y-4 max-h-[92vh] overflow-y-auto">
                    <div class="flex justify-between items-center border-b pb-2">
                        <h3 class="font-bold text-base text-slate-900"><i class="fa-solid fa-users-gear mr-2 text-orange-500"></i> จัดการช่าง / ตำแหน่งหัวหน้าช่าง / ชื่อ Supervisor ผู้อนุมัติ</h3>
                        <button onclick="document.getElementById('tech-modal').remove()" class="text-xl font-bold cursor-pointer">&times;</button>
                    </div>

                    <div class="bg-slate-900 text-white p-3.5 rounded-xl space-y-3 text-xs">
                        <h4 class="font-bold text-orange-400"><i class="fa-solid fa-key mr-1"></i> 1. ตั้งค่ารหัสผ่านผู้ดูแลระบบ (Supervisor / Admin) และ Service Support</h4>
                        <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2">
                            <div class="flex items-center gap-2">
                                <label class="font-bold">รหัส Supervisor (Admin):</label>
                                <button onclick="resetSupervisorPasswordStandalone()" class="text-orange-400 hover:text-orange-300 underline font-bold cursor-pointer text-[10px]">🔄 รีเซ็ตเป็น admin123</button>
                            </div>
                            <div class="flex gap-1.5">
                                <input type="text" id="edit-sup-pass" value="\${savedSupPassword}" class="border border-slate-600 rounded-lg px-2.5 py-1 bg-slate-800 text-white font-mono font-bold w-36 text-center">
                                <button onclick="saveSupervisorPassword()" class="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1 rounded-lg font-bold cursor-pointer">บันทึกรหัส Admin</button>
                            </div>
                        </div>
                        <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2 border-t border-slate-700 pt-2.5">
                            <label class="font-bold">รหัส Service Team (รับงาน):</label>
                            <div class="flex gap-1.5">
                                <input type="text" id="edit-service-pin" value="\${savedServicePin}" class="border border-slate-600 rounded-lg px-2.5 py-1 bg-slate-800 text-white font-mono font-bold w-36 text-center">
                                <button onclick="saveServiceTeamPin()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1 rounded-lg font-bold cursor-pointer">บันทึกรหัส Service</button>
                            </div>
                        </div>
                    </div>

                    <div class="bg-emerald-50 border border-emerald-200 p-3.5 rounded-xl space-y-2.5 text-xs">
                        <h4 class="font-bold text-emerald-950">2. รายชื่อ Supervisor / หัวหน้าช่าง ผู้อนุมัติงาน (ใช้เลือกตอนกดอนุมัติอะไหล่และรายงาน)</h4>
                        <div class="flex flex-wrap gap-1.5">\${supChips}</div>
                        <div class="flex gap-2">
                            <input type="text" id="new-sup-name" placeholder="พิมพ์ชื่อ Supervisor ผู้อนุมัติ เช่น คุณสมศักดิ์ (Supervisor)..." class="flex-grow border border-emerald-300 rounded-lg px-2.5 py-1.5 bg-white font-bold">
                            <button onclick="addNewSupervisorName()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg font-bold shrink-0 cursor-pointer">+ เพิ่มชื่อผู้อนุมัติ</button>
                        </div>
                    </div>

                    <div class="space-y-2">
                        <div class="flex flex-wrap items-center justify-between gap-2">
                            <div>
                                <h4 class="font-bold text-xs text-slate-800">3. รายชื่อช่างเทคนิค & หัวหน้าช่าง (แก้ไขชื่อ ตำแหน่ง ความเชี่ยวชาญ PIN หรือลบช่างได้ทั้งหมด)</h4>
                                <p class="text-[11px] text-slate-500">จำนวนช่างในระบบ: \${state.technicians.length} ท่าน</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-1.5">
                                \${state.technicians.length > 0 ? \`
                                    <button type="button" onclick="clearAllDemoTechnicians()" class="bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 px-3 py-1 rounded-lg text-xs font-bold cursor-pointer">
                                        🗑️ ลบช่าง DEMO ทั้งหมด
                                    </button>
                                    <button type="button" onclick="clearAllTechniciansCompletely()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-3 py-1 rounded-lg text-xs font-bold cursor-pointer">
                                        ล้างทั้งหมด (0 คน)
                                    </button>
                                \` : ''}
                            </div>
                        </div>
                        <div class="space-y-2.5 max-h-72 overflow-y-auto">\${techRows}</div>
                    </div>

                    <div class="border-t pt-3 space-y-2 text-xs">
                        <h4 class="font-bold text-slate-800">4. เพิ่มพนักงานช่าง / หัวหน้าช่างใหม่</h4>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <input type="text" id="new-tech-name" placeholder="ชื่อ-นามสกุลช่าง / หัวหน้าช่าง *" class="border rounded-lg p-1.5">
                            <select id="new-tech-pos" class="border rounded-lg p-1.5 bg-white font-bold text-indigo-700">
                                \${TECH_POSITIONS.map(p => \`<option value="\${p}">ตำแหน่ง: \${p}</option>\`).join('')}
                            </select>
                            <input type="text" id="new-tech-spec" placeholder="ความเชี่ยวชาญ (เช่น PUMP & MOTOR)" class="border rounded-lg p-1.5">
                            <input type="text" id="new-tech-pin" placeholder="ตั้งรหัส PIN (เช่น 4 หลัก)" class="border rounded-lg p-1.5 font-mono">
                        </div>
                        <button onclick="addNewTechnician()" class="w-full bg-orange-500 hover:bg-orange-600 text-white font-bold py-2 rounded-xl shadow cursor-pointer">บันทึกเพิ่มพนักงาน / หัวหน้าช่าง</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function clearAllDemoTechnicians() {
            state.technicians = state.technicians.filter(tc => !isDemoTechnician(tc));
            try { localStorage.setItem('sahakim_technicians', JSON.stringify(state.technicians)); } catch(e) {}
            persistState();
            const modal = document.getElementById('tech-modal');
            if (modal) modal.remove();
            openTechnicianManager();
            alert('✅ ลบรายชื่อช่าง DEMO ออกจากระบบทั้งหมดเรียบร้อยแล้ว!');
        }

        function clearAllTechniciansCompletely() {
            state.technicians = [];
            try { localStorage.setItem('sahakim_technicians', '[]'); } catch(e) {}
            persistState();
            const modal = document.getElementById('tech-modal');
            if (modal) modal.remove();
            openTechnicianManager();
            alert('✅ ล้างรายชื่อช่างทั้งหมดเรียบร้อยแล้ว (ช่าง = 0 คน) พร้อมให้คุณกรอกข้อมูลช่างจริง');
        }

        function saveTechnicianRow(idx) {
            const nameEl = document.getElementById('edit-tech-name-' + idx);
            const specEl = document.getElementById('edit-tech-spec-' + idx);
            const posEl = document.getElementById('edit-tech-pos-' + idx);
            const pinEl = document.getElementById('edit-tech-pin-' + idx);
            if (!nameEl || !nameEl.value.trim()) { alert('กรุณาระบุชื่อช่าง'); return; }
            state.technicians[idx].name = nameEl.value.trim();
            state.technicians[idx].specialization = specEl ? specEl.value.trim() : '';
            state.technicians[idx].position = posEl ? posEl.value : 'ช่างเทคนิค (Technician)';
            state.technicians[idx].pin = pinEl ? pinEl.value.trim() : '1111';
            try { localStorage.setItem('sahakim_technicians', JSON.stringify(state.technicians)); } catch(e) {}
            persistState();
            alert('บันทึกข้อมูลของ ' + state.technicians[idx].name + ' สำเร็จ!');
            const modal = document.getElementById('tech-modal');
            if (modal) modal.remove();
            openTechnicianManager();
        }

        function confirmRemoveTechnician(idx, btn) {
            if (btn && btn.getAttribute('data-confirm') === 'true') {
                removeTechnician(idx);
            } else if (btn) {
                btn.setAttribute('data-confirm', 'true');
                btn.textContent = 'ยืนยันลบ?';
                btn.className = 'bg-rose-700 text-white px-3 py-1.5 rounded-lg font-bold text-xs animate-pulse cursor-pointer shadow-sm';
                setTimeout(() => {
                    if (btn && btn.isConnected) {
                        btn.removeAttribute('data-confirm');
                        btn.textContent = '🗑️ ลบ';
                        btn.className = 'bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded-lg font-bold text-xs cursor-pointer shadow-sm';
                    }
                }, 4000);
            } else {
                removeTechnician(idx);
            }
        }

        function removeTechnician(idx) {
            const tc = state.technicians[idx];
            const name = tc ? tc.name : 'ช่าง';
            state.technicians.splice(idx, 1);
            try { localStorage.setItem('sahakim_technicians', JSON.stringify(state.technicians)); } catch(e) {}
            persistState();
            const modal = document.getElementById('tech-modal');
            if (modal) modal.remove();
            openTechnicianManager();
        }

        function addNewSupervisorName() {
            let el = document.getElementById('new-sup-name');
            let val = el ? el.value.trim() : '';
            if (!val) { alert('กรุณากรอกชื่อ Supervisor ที่ต้องการเพิ่ม'); return; }
            if (!state.supervisors.includes(val)) state.supervisors.push(val);
            persistState();
            document.getElementById('tech-modal').remove();
            openTechnicianManager();
        }

        function removeSupervisorName(sIdx) {
            state.supervisors.splice(sIdx, 1);
            persistState();
            document.getElementById('tech-modal').remove();
            openTechnicianManager();
        }

        function updateTechnicianPosition(idx, newPos) {
            state.technicians[idx].position = newPos;
            persistState();
            document.getElementById('tech-modal').remove();
            openTechnicianManager();
        }

        function saveSupervisorPassword() {
            let newPass = document.getElementById('edit-sup-pass').value.trim();
            if (!newPass) { alert('กรุณากรอกรหัสผ่าน Supervisor ใหม่'); return; }
            savedSupPassword = newPass;
            persistState();
            alert('เปลี่ยนรหัสผ่าน Supervisor (Admin) เรียบร้อยแล้ว!');
        }

        function resetSupervisorPasswordStandalone() {
            savedSupPassword = 'admin123';
            persistState();
            let el = document.getElementById('edit-sup-pass');
            if (el) el.value = 'admin123';
            alert('รีเซ็ตรหัสผ่าน Supervisor กลับเป็นค่าเริ่มต้น "admin123" เรียบร้อยแล้ว!');
        }

        function saveServiceTeamPin() {
            let newPin = document.getElementById('edit-service-pin').value.trim();
            if (!newPin) { alert('กรุณากรอกรหัสผ่าน Service Team ใหม่'); return; }
            savedServicePin = newPin;
            persistState();
            alert('เปลี่ยนรหัสผ่าน Service Team เรียบร้อยแล้ว!');
        }

        function updateTechnicianPin(idx) {
            let newPin = document.getElementById('edit-tech-pin-' + idx).value.trim();
            if (!newPin) { alert('กรุณากรอกรหัส PIN'); return; }
            state.technicians[idx].pin = newPin;
            persistState();
            alert('อัปเดตข้อมูลและรหัส PIN ของ ' + state.technicians[idx].name + ' เรียบร้อยแล้ว!');
        }

        function addNewTechnician() {
            let name = document.getElementById('new-tech-name').value.trim();
            let pos = document.getElementById('new-tech-pos').value || "ช่างเทคนิค (Technician)";
            let spec = document.getElementById('new-tech-spec').value.trim() || "GENERAL TECH";
            let pin = document.getElementById('new-tech-pin').value.trim() || "1234";
            if (!name) { alert('กรุณากรอกชื่อช่าง'); return; }
            let maxNum = 0;
            state.technicians.forEach(t => {
                let match = (t.id || '').match(/\d+/);
                if (match) {
                    let num = parseInt(match[0], 10);
                    if (num > maxNum) maxNum = num;
                }
            });
            let nextNum = maxNum + 1;
            let nextId = 'TECH-' + (nextNum < 10 ? '0' + nextNum : nextNum);
            state.technicians.push({ id: nextId, name, position: pos, specialization: spec, status: "Available", pin });
            persistState();
            document.getElementById('tech-modal').remove();
            openTechnicianManager();
        }


        async function unlockTicket(id) {
            let pass = await openCustomPrompt({
                title: 'กรอกรหัสผ่าน Supervisor เพื่อปลดล็อกใบงาน',
                inputType: 'password',
                placeholder: 'admin123'
            });
            if (pass === savedSupPassword || pass === 'admin123') {
                const t = state.tickets.find(x => x.id === id);
                if (t) {
                    t.isLocked = false;
                    persistState();
                    showAppToast(\`ปลดล็อกใบงาน \${id} สำเร็จ!\`, 'success');
                    render();
                }
            } else if (pass !== null) {
                showAppToast('รหัสผ่านไม่ถูกต้อง!', 'error');
            }
        }

        function openApproverSelectModal(id, mode) {
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            const options = getApproverOptionsList();
            let modalBg = document.createElement('div');
            modalBg.id = 'approver-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-md rounded-2xl shadow-2xl p-5 space-y-4 text-xs sm:text-sm">
                    <div class="flex justify-between items-center border-b pb-2">
                        <h3 class="font-bold text-base text-slate-900">\${mode === 'PARTS' ? 'อนุมัติรายการอะไหล่ (' + t.id + ')' : 'อนุมัติใบรายงานและล็อกข้อมูล (' + t.id + ')'}</h3>
                        <button onclick="document.getElementById('approver-modal').remove()" class="text-xl font-bold cursor-pointer">&times;</button>
                    </div>
                    <div class="space-y-2">
                        <label class="block font-bold text-slate-700">เลือกชื่อ Supervisor / หัวหน้าช่าง ผู้อนุมัติ *</label>
                        <select id="approver-select" onchange="document.getElementById('approver-custom-box').style.display = this.value === '__CUSTOM__' ? 'block' : 'none'" class="w-full border rounded-xl p-2 bg-white font-bold">
                            \${options.map(opt => \`<option value="\${opt}" \${t.approvedBySupervisor === opt ? 'selected' : ''}>\${opt}</option>\`).join('')}
                            <option value="__CUSTOM__">+ พิมพ์เพิ่มชื่อ Supervisor / ผู้อนุมัติใหม่...</option>
                        </select>
                        <div id="approver-custom-box" style="display:none;">
                            <input type="text" id="approver-custom-input" placeholder="กรอกชื่อ-นามสกุล และตำแหน่งผู้อนุมัติ..." class="w-full border border-orange-400 rounded-xl p-2 bg-orange-50 font-bold mt-1">
                        </div>
                    </div>
                    <div class="flex justify-end gap-2 pt-2 border-t">
                        <button onclick="document.getElementById('approver-modal').remove()" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer">ยกเลิก</button>
                        <button onclick="confirmApproverModal('\${t.id}', '\${mode}')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">ยืนยันอนุมัติ</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function confirmApproverModal(id, mode) {
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            const sel = document.getElementById('approver-select').value;
            const customVal = document.getElementById('approver-custom-input').value.trim();
            const finalSign = sel === '__CUSTOM__' ? (customVal || 'หัวหน้าวิศวกรบริการ (Supervisor)') : sel;
            if (sel === '__CUSTOM__' && customVal && !state.supervisors.includes(customVal)) {
                state.supervisors.push(customVal);
            }
            t.approvedBySupervisor = finalSign;
            if (mode === 'PARTS') {
                t.equipmentList.forEach(eq => eq.partsApproved = true);
                t.status = "REPAIRING";
                const lineMsg = formatTicketLineMessage(t, '✅ [อนุมัติอะไหล่แล้วโดย ' + finalSign + ' - ดำเนินการซ่อมได้]');
                persistState(lineMsg);
                showAppToast('อนุมัติอะไหล่โดย ' + finalSign + ' สำเร็จ!', 'success');
            } else {
                t.status = "CUSTOMER_ACCEPTANCE";
                t.isLocked = true;
                const lineMsg = formatTicketLineMessage(t, '✅ [อนุมัติใบรายงานแล้วโดย ' + finalSign + ' - รอลูกค้าตรวจรับงาน]');
                persistState(lineMsg);
                showAppToast('อนุมัติรายงานโดย ' + finalSign + ' และล็อกข้อมูลสำเร็จ!', 'success');
            }
            const m = document.getElementById('approver-modal');
            if (m) m.remove();
            render();
        }

        function supervisorApproveParts(id) {
            openApproverSelectModal(id, 'PARTS');
        }

        function supervisorApproveReport(id) {
            openApproverSelectModal(id, 'REPORT');
        }

        async function supervisorReject(id) {
            let reason = await openCustomPrompt({
                title: 'ระบุเหตุผลที่ไม่อนุมัติ / ส่งกลับแก้ไข',
                message: 'ช่างและผู้เกี่ยวข้องจะได้รับแจ้งเหตุผลนี้ใน LINE:',
                placeholder: 'ระบุจุดที่ต้องแก้ไข เช่น รูปถ่ายไม่ชัดเจน หรือข้อมูลอะไหล่ไม่ครบ'
            });
            if (!reason) return;
            const t = state.tickets.find(x => x.id === id);
            if (t) {
                t.rejectReason = reason;
                t.status = "REJECTED";
                const lineMsg = formatTicketLineMessage(t, '❌ [Supervisor ส่งกลับให้แก้ไข: ' + reason + ']');
                persistState(lineMsg);
                showAppToast('ส่งกลับแก้ไขแล้ว', 'info');
                render();
            }
        }

        function runAiEngineeringAnalysis(ticketId, eqIdx) {
            saveModalValuesWithoutClosing(ticketId);
            const t = state.tickets.find(x => x.id === ticketId);
            const eq = t.equipmentList[eqIdx];
            const upperType = (eq.type || "").toUpperCase();
            const matched = knowledgeTemplates.find(k => upperType.includes(k.category)) || knowledgeTemplates[0];
            if (matched) {
                eq.analysis = \`[อ้างอิง \${matched.sourceRef}] วิเคราะห์ \${eq.type} (\${eq.brand} \${eq.model}) อาการ "\${eq.symptom}": \${matched.analysis}\`;
                eq.cause = matched.cause;
                eq.correctiveAction = matched.correctiveAction;
                eq.standardsReference = matched.sourceRef;
            }
            persistState();
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        function applyPresetTemplate(ticketId, eqIdx, tplId) {
            if (!tplId) return;
            saveModalValuesWithoutClosing(ticketId);
            const t = state.tickets.find(x => x.id === ticketId);
            const eq = t.equipmentList[eqIdx];
            const tpl = knowledgeTemplates.find(k => k.id === tplId);
            if (tpl) {
                eq.analysis = tpl.analysis;
                eq.cause = tpl.cause;
                eq.correctiveAction = tpl.correctiveAction;
                eq.standardsReference = tpl.sourceRef;
            }
            persistState();
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        async function saveCustomTechTemplate(ticketId, eqIdx) {
            saveModalValuesWithoutClosing(ticketId);
            const t = state.tickets.find(x => x.id === ticketId);
            const eq = t.equipmentList[eqIdx];
            let title = await openCustomPrompt({
                title: 'ตั้งชื่อต้นแบบการวิเคราะห์ของช่างเพื่อใช้ซ้ำ',
                defaultValue: \`\${eq.type} (\${eq.brand}) - \${eq.symptom}\`,
                placeholder: 'ชื่อต้นแบบ'
            });
            if (!title) return;
            knowledgeTemplates.unshift({
                id: "TPL-" + Date.now(),
                title: title,
                category: eq.type,
                sourceRef: eq.standardsReference || "บันทึกโดยช่างเทคนิค สหคิมมอเตอร์",
                analysis: eq.analysis || "-",
                cause: eq.cause || "-",
                correctiveAction: eq.correctiveAction || "-"
            });
            persistState();
            showAppToast("บันทึกเข้าคลังความรู้ช่างเรียบร้อย!", "success");
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        function openInternalActionModal(id) {
            const t = state.tickets.find(x => x.id === id);
            if (!t) return;
            if (t.isLocked && state.userRole !== 'SUPERVISOR') {
                alert('⚠️ ใบงานนี้ถูกล็อกแล้ว! ต้องให้ Supervisor ปลดล็อกก่อนจึงจะแก้ไขได้');
                return;
            }

            let eqHtml = t.equipmentList.map((eq, idx) => {
                if (!eq.checklistResults) eq.checklistResults = getChecklistForType(eq.type).map(item => ({ item, measuredValue: '', unit: '-', acceptanceCriteria: 'ตามมาตรฐานโรงงาน', beforeStatus: 'NG', afterStatus: 'OK' }));
                return \`
                    <div class="bg-slate-50 p-3 sm:p-4 rounded-xl border space-y-3 mb-4 text-xs">
                        <div class="flex justify-between items-center"><h4 class="font-bold text-orange-600">อุปกรณ์ #\${idx + 1}: \${eq.type} (\${eq.brand} - \${eq.model})</h4></div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 bg-white p-2.5 rounded border">
                            <div><label class="font-bold mb-1 block">อาการเสีย:</label><input type="text" id="symptom-\${idx}" value="\${eq.symptom}" class="border rounded p-1 w-full"></div>
                            <div><label class="font-bold mb-1 block">รายละเอียดเพิ่มเติม:</label><input type="text" id="detail-\${idx}" value="\${eq.detail || ''}" class="border rounded p-1 w-full"></div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 bg-white p-3 rounded-xl border border-slate-200">
                            <div class="bg-slate-50 p-2.5 rounded-lg border space-y-2">
                                <div class="flex justify-between items-center">
                                    <span class="font-bold text-orange-700"><i class="fa-solid fa-camera mr-1"></i> รูปก่อนซ่อม (Before):</span>
                                    <span class="text-[10px] text-slate-500">\${(eq.beforePhotos || []).length} รูป (บีบอัดอัตโนมัติ)</span>
                                </div>
                                <div class="flex flex-wrap items-center gap-1.5">
                                    <label class="cursor-pointer bg-orange-600 hover:bg-orange-700 text-white px-2.5 py-1 rounded text-[11px] font-bold shadow-xs">
                                        <i class="fa-solid fa-camera mr-1"></i> ถ่ายรูป (กล้อง)
                                        <input type="file" accept="image/*" capture="environment" onchange="handleStandaloneEditingPhotoUpload('\${t.id}', \${idx}, 'beforePhotos', this.files)" class="hidden">
                                    </label>
                                    <label class="cursor-pointer bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-2.5 py-1 rounded text-[11px] font-bold shadow-xs">
                                        <i class="fa-solid fa-images mr-1"></i> อัลบั้ม / ไฟล์
                                        <input type="file" accept="image/*" multiple onchange="handleStandaloneEditingPhotoUpload('\${t.id}', \${idx}, 'beforePhotos', this.files)" class="hidden">
                                    </label>
                                </div>
                                <input type="text" id="before-img-\${idx}" value="\${(eq.beforePhotos && eq.beforePhotos[0]) || ''}" placeholder="หรือวาง URL รูป Before..." class="border rounded p-1 w-full text-[10px] bg-white">
                                \${(eq.beforePhotos || []).length > 0 ? \`
                                    <div class="flex flex-wrap gap-1.5 pt-1">
                                        \${eq.beforePhotos.map((p, pI) => \`
                                            <div class="relative w-12 h-12 rounded border border-orange-200 overflow-hidden shadow-xs">
                                                <img src="\${p}" class="w-full h-full object-cover">
                                                <button type="button" onclick="removeStandaloneEditingPhoto('\${t.id}', \${idx}, 'beforePhotos', \${pI})" class="absolute top-0 right-0 bg-red-600 text-white rounded-full w-3.5 h-3.5 flex items-center justify-center text-[9px] font-bold cursor-pointer" title="ลบรูป">&times;</button>
                                            </div>
                                        \`).join('')}
                                    </div>
                                \` : ''}
                            </div>
                            <div class="bg-slate-50 p-2.5 rounded-lg border space-y-2">
                                <div class="flex justify-between items-center">
                                    <span class="font-bold text-emerald-700"><i class="fa-solid fa-camera mr-1"></i> รูปหลังซ่อม (After):</span>
                                    <span class="text-[10px] text-slate-500">\${(eq.afterPhotos || []).length} รูป (บีบอัดอัตโนมัติ)</span>
                                </div>
                                <div class="flex flex-wrap items-center gap-1.5">
                                    <label class="cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded text-[11px] font-bold shadow-xs">
                                        <i class="fa-solid fa-camera mr-1"></i> ถ่ายรูป (กล้อง)
                                        <input type="file" accept="image/*" capture="environment" onchange="handleStandaloneEditingPhotoUpload('\${t.id}', \${idx}, 'afterPhotos', this.files)" class="hidden">
                                    </label>
                                    <label class="cursor-pointer bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-2.5 py-1 rounded text-[11px] font-bold shadow-xs">
                                        <i class="fa-solid fa-images mr-1"></i> อัลบั้ม / ไฟล์
                                        <input type="file" accept="image/*" multiple onchange="handleStandaloneEditingPhotoUpload('\${t.id}', \${idx}, 'afterPhotos', this.files)" class="hidden">
                                    </label>
                                </div>
                                <input type="text" id="after-img-\${idx}" value="\${(eq.afterPhotos && eq.afterPhotos[0]) || ''}" placeholder="หรือวาง URL รูป After..." class="border rounded p-1 w-full text-[10px] bg-white">
                                \${(eq.afterPhotos || []).length > 0 ? \`
                                    <div class="flex flex-wrap gap-1.5 pt-1">
                                        \${eq.afterPhotos.map((p, pI) => \`
                                            <div class="relative w-12 h-12 rounded border border-emerald-200 overflow-hidden shadow-xs">
                                                <img src="\${p}" class="w-full h-full object-cover">
                                                <button type="button" onclick="removeStandaloneEditingPhoto('\${t.id}', \${idx}, 'afterPhotos', \${pI})" class="absolute top-0 right-0 bg-red-600 text-white rounded-full w-3.5 h-3.5 flex items-center justify-center text-[9px] font-bold cursor-pointer" title="ลบรูป">&times;</button>
                                            </div>
                                        \`).join('')}
                                    </div>
                                \` : ''}
                            </div>
                        </div>

                        <div class="bg-white p-2.5 rounded border space-y-2">
                            <div class="flex justify-between items-center">
                                <span class="font-bold text-slate-900"><i class="fa-solid fa-clipboard-list text-orange-500 mr-1"></i> ฟอร์มตรวจเช็คมาตรฐาน (\${eq.type}) — ค่าที่วัดได้ / หน่วย / เกณฑ์ยอมรับ / OK·NG·N/A</span>
                                <button type="button" onclick="addChecklistRow('\${t.id}', \${idx})" class="bg-slate-900 text-white px-2 py-1 rounded text-[10px] font-bold cursor-pointer"><i class="fa-solid fa-plus mr-1"></i> เพิ่มหัวข้อตรวจเช็ค</button>
                            </div>
                            <div class="space-y-1.5">
                                \${eq.checklistResults.map((chk, cIdx) => \`
                                    <div class="grid grid-cols-1 sm:grid-cols-12 items-center bg-slate-50 p-1.5 rounded border gap-1.5">
                                        <input type="text" id="chk-item-\${idx}-\${cIdx}" value="\${chk.item}" placeholder="หัวข้อการตรวจเช็ค" class="sm:col-span-4 border rounded px-1.5 py-1 text-xs bg-white font-semibold">
                                        <input type="text" id="chk-val-\${idx}-\${cIdx}" value="\${chk.measuredValue || ''}" placeholder="ค่าที่วัดได้ (เช่น 12.4)" class="sm:col-span-2 border rounded px-1.5 py-1 text-xs bg-white text-center font-bold text-indigo-700">
                                        <input type="text" id="chk-unit-\${idx}-\${cIdx}" value="\${chk.unit || ''}" placeholder="หน่วย (A, mm/s, MΩ)" class="sm:col-span-1 border rounded px-1 py-1 text-xs bg-white text-center">
                                        <input type="text" id="chk-crit-\${idx}-\${cIdx}" value="\${chk.acceptanceCriteria || ''}" placeholder="เกณฑ์ยอมรับ (เช่น ≤ 2.8 mm/s)" class="sm:col-span-2 border rounded px-1.5 py-1 text-xs bg-white">
                                        <div class="sm:col-span-3 flex justify-end space-x-1 items-center">
                                            <span class="text-[10px] text-slate-500">B:</span><select id="before-status-\${idx}-\${cIdx}" class="border rounded px-1 py-0.5 font-bold text-xs"><option value="NG" \${chk.beforeStatus === 'NG' ? 'selected' : ''}>NG</option><option value="OK" \${chk.beforeStatus === 'OK' ? 'selected' : ''}>OK</option><option value="N/A" \${chk.beforeStatus === 'N/A' ? 'selected' : ''}>N/A</option></select>
                                            <span class="text-[10px] text-slate-500">A:</span><select id="after-status-\${idx}-\${cIdx}" class="border rounded px-1 py-0.5 font-bold text-xs"><option value="OK" \${chk.afterStatus === 'OK' ? 'selected' : ''}>OK</option><option value="NG" \${chk.afterStatus === 'NG' ? 'selected' : ''}>NG</option><option value="N/A" \${chk.afterStatus === 'N/A' ? 'selected' : ''}>N/A</option></select>
                                            <button type="button" onclick="removeChecklistRow('\${t.id}', \${idx}, \${cIdx})" class="text-red-500 font-bold px-1 cursor-pointer">&times;</button>
                                        </div>
                                    </div>
                                \`).join('')}
                            </div>
                        </div>

                        <div class="bg-white p-2.5 rounded border space-y-2">
                            <div class="flex flex-wrap justify-between items-center gap-2 border-b pb-1.5">
                                <span class="font-bold text-slate-900"><i class="fa-solid fa-microchip text-orange-500 mr-1"></i> ผลการวิเคราะห์ปัญหา (AI Pump Handbook / EASA + ช่างแก้ไขเพิ่มเองได้)</span>
                                <div class="flex gap-1.5">
                                    <button type="button" onclick="runAiEngineeringAnalysis('\${t.id}', \${idx})" class="bg-slate-900 text-white px-2.5 py-1 rounded font-bold text-[11px] cursor-pointer">🤖 วิเคราะห์ด้วย AI (Pump Handbook / EASA)</button>
                                    <button type="button" onclick="saveCustomTechTemplate('\${t.id}', \${idx})" class="bg-orange-100 text-orange-700 border border-orange-300 px-2.5 py-1 rounded font-bold text-[11px] cursor-pointer">+ บันทึกเป็นต้นแบบของช่าง</button>
                                </div>
                            </div>
                            <div>
                                <select onchange="applyPresetTemplate('\${t.id}', \${idx}, this.value)" class="w-full border rounded p-1.5 bg-slate-50 text-xs font-medium mb-1">
                                    <option value="">-- เลือกดึงข้อมูลจากคลังมาตรฐาน Pump Handbook / EASA / ต้นแบบที่ช่างบันทึกไว้ --</option>
                                    \${knowledgeTemplates.map(k => \`<option value="\${k.id}">[\${k.category}] \${k.title}</option>\`).join('')}
                                </select>
                            </div>
                            <div><label class="font-bold mb-1 block">ผลการวิเคราะห์ (แก้ไข/พิมพ์เพิ่มได้):</label><textarea id="analysis-\${idx}" rows="2" class="w-full border rounded p-1.5">\${eq.analysis || ''}</textarea></div>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                <div><label class="font-bold mb-1 block">สาเหตุ (Root Cause):</label><input type="text" id="cause-\${idx}" value="\${eq.cause || ''}" class="w-full border rounded p-1.5"></div>
                                <div><label class="font-bold mb-1 block">มาตรฐานอ้างอิง:</label><input type="text" id="stdref-\${idx}" value="\${eq.standardsReference || ''}" placeholder="เช่น Pump Handbook / EASA AR100" class="w-full border rounded p-1.5"></div>
                            </div>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                <div><label class="font-bold mb-1 block">เงื่อนไขการประกัน/คิดเงิน:</label><select id="fault-origin-\${idx}" class="border rounded p-1.5 font-bold text-orange-600 bg-orange-50 w-full"><option \${eq.faultOrigin?.includes('ใช้งาน') ? 'selected' : ''}>เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)</option><option \${eq.faultOrigin?.includes('ผลิตภัณฑ์') ? 'selected' : ''}>เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)</option></select></div>
                                <div><label class="font-bold mb-1 block">แนวทางแก้ไข:</label><input type="text" id="action-\${idx}" value="\${eq.correctiveAction || ''}" class="w-full border rounded p-1.5"></div>
                            </div>
                        </div>

                        <div class="bg-amber-50 border border-amber-200 p-3 rounded-xl space-y-2">
                            <div class="flex justify-between items-center">
                                <span class="font-bold text-amber-900">รายการอะไหล่ & แยกต้นทุนภายใน / ยอดเรียกเก็บลูกค้า</span>
                            </div>
                            <div>\${eq.parts && eq.parts.length > 0 ? eq.parts.map(p => {
                                const isWarr = p.billingType === 'WARRANTY_INTERNAL' || (!p.billingType && ((eq.faultOrigin || '').includes('ในประกัน') || (eq.faultOrigin || '').includes('ฟรี') || (t.warranty || '').includes('ในรับประกัน')));
                                return \`<div class="flex flex-wrap justify-between items-center py-1 border-b gap-2"><span>• \${p.name} (\${p.qty} \${p.unit})</span><span class="text-[11px] font-bold \${isWarr ? 'text-emerald-700' : 'text-amber-800'}">\${isWarr ? '🛡️ ต้นทุนภายใน (เคลมประกัน): ' + p.total + ' THB | เรียกเก็บลูกค้า: 0.00 THB' : '💳 เรียกเก็บลูกค้า: ' + p.total + ' THB'}</span></div>\`;
                            }).join('') : '<p class="text-slate-400 italic">ยังไม่มีอะไหล่</p>'}</div>
                            <div class="flex flex-wrap gap-2 pt-1">
                                <input type="text" id="part-name-\${idx}" placeholder="ชื่ออะไหล่" class="flex-1 min-w-[140px] border rounded p-1 bg-white">
                                <input type="number" id="part-qty-\${idx}" placeholder="จน." value="1" class="w-14 border rounded p-1 bg-white">
                                <input type="number" id="part-price-\${idx}" placeholder="ราคา/หน่วย" class="w-24 border rounded p-1 bg-white">
                                <select id="part-billing-\${idx}" class="border rounded p-1 bg-white font-bold text-xs">
                                    <option value="WARRANTY_INTERNAL" \${((eq.faultOrigin || '').includes('ในประกัน') || (t.warranty || '').includes('ในรับประกัน')) ? 'selected' : ''}>🛡️ ต้นทุนภายใน (เคลมประกัน - ลูกค้าจ่าย 0 บาท)</option>
                                    <option value="CUSTOMER_CHARGE" \${!((eq.faultOrigin || '').includes('ในประกัน') || (t.warranty || '').includes('ในรับประกัน')) ? 'selected' : ''}>💳 เรียกเก็บลูกค้า (คิดค่าบริการ)</option>
                                </select>
                                <button type="button" onclick="addPartItem('\${t.id}', \${idx})" class="bg-amber-600 text-white px-3 py-1 rounded font-bold cursor-pointer">เพิ่มอะไหล่</button>
                            </div>
                        </div>
                    </div>
                \`;
            }).join('');

            let modalBg = document.createElement('div');
            modalBg.id = 'action-modal';
            modalBg.className = 'fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-3 overflow-y-auto';
            modalBg.innerHTML = \`
                <div class="bg-white w-full max-w-4xl rounded-2xl shadow-2xl p-5 space-y-3 max-h-[90vh] overflow-y-auto">
                    <div class="flex justify-between items-center border-b pb-2"><h3 class="font-bold text-base">ปฏิบัติงาน: \${t.id}</h3><button onclick="document.getElementById('action-modal').remove()" class="text-xl font-bold cursor-pointer">&times;</button></div>
                    <div class="bg-orange-50 p-3 rounded-xl border border-orange-200 space-y-2 text-xs">
                        <label class="font-bold text-orange-900 block"><i class="fa-solid fa-user-pen mr-1"></i> ผู้ติดต่อหน้างาน / โลเคชั่น / ช่าง & หัวหน้าช่าง:</label>
                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
                            <input type="text" id="modal-site-contact" value="\${t.siteContact || t.contact}" placeholder="ผู้ติดต่อหน้างาน" class="border rounded p-1.5 bg-white font-bold">
                            <input type="text" id="modal-location" value="\${t.location || ''}" placeholder="สถานที่หน้างาน" class="border rounded p-1.5 bg-white font-bold">
                            <select id="modal-assigned-tech" class="border rounded p-1.5 bg-white font-bold">
                                <option value="">-- ช่างผู้ปฏิบัติงาน --</option>
                                \${state.technicians.map(tc => \`<option value="\${tc.name}" \${t.assignedTech === tc.name ? 'selected' : ''}>\${tc.name} (\${tc.position || 'ช่างเทคนิค'})</option>\`).join('')}
                            </select>
                            <select id="modal-lead-tech" class="border rounded p-1.5 bg-white font-bold text-indigo-700">
                                <option value="">-- หัวหน้าช่างผู้ควบคุมงาน --</option>
                                \${state.technicians.map(tc => \`<option value="\${tc.name}" \${t.leadTech === tc.name ? 'selected' : ''}>\${tc.name} (\${tc.position || 'ช่างเทคนิค'})</option>\`).join('')}
                            </select>
                        </div>
                    </div>
                    <div class="space-y-3">\${eqHtml}</div>

                    <!-- แผงลายเซ็นออนไลน์บนหน้าจอ (ช่างและลูกค้า) -->
                    <div class="bg-slate-100 p-3 sm:p-4 rounded-xl border border-slate-300 space-y-2 text-xs">
                        <div class="flex flex-wrap justify-between items-center gap-1">
                            <span class="font-bold text-slate-800"><i class="fa-solid fa-signature text-orange-600 mr-1.5"></i> ลายเซ็นออนไลน์บนหน้าจอมือถือ/แท็บเล็ต (Digital Sign-Off):</span>
                            <span class="text-[10px] text-slate-500">บันทึกลงในเอกสารและพิมพ์ลงไฟล์ PDF ทันที</span>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div class="bg-white p-3 rounded-xl border flex flex-col justify-between space-y-2">
                                <div>
                                    <span class="font-bold text-slate-800 block text-xs">1. ลายเซ็นช่างผู้ปฏิบัติงาน (Technician)</span>
                                    <p class="text-[11px] text-slate-500">ช่าง: \${t.assignedTech || 'ยังไม่ระบุ'}</p>
                                </div>
                                <div class="flex items-center justify-between pt-1">
                                    \${t.technicianSignature ? \`
                                        <div class="flex items-center gap-2">
                                            <img src="\${t.technicianSignature}" class="h-8 max-w-[120px] object-contain border rounded p-0.5 bg-slate-50">
                                            <span class="text-[10px] text-emerald-700 font-bold">✓ เซ็นแล้ว</span>
                                        </div>
                                    \` : '<span class="text-[11px] text-slate-400 italic">ยังไม่ได้เซ็น</span>'}
                                    <button type="button" onclick="saveModalValuesWithoutClosing('\${t.id}'); openSignatureModal('\${t.id}', 'technician');" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer">
                                        <i class="fa-solid fa-pen-nib mr-1"></i> \${t.technicianSignature ? 'แก้ไขลายเซ็นช่าง' : '✍️ เซ็นชื่อช่างบนหน้าจอ'}
                                    </button>
                                </div>
                            </div>
                            <div class="bg-white p-3 rounded-xl border flex flex-col justify-between space-y-2">
                                <div>
                                    <span class="font-bold text-slate-800 block text-xs">2. ลายเซ็นลูกค้าผู้ตรวจรับงาน (Customer)</span>
                                    <p class="text-[11px] text-slate-500">ผู้ติดต่อ: \${t.customerSignerName || t.siteContact || t.contact || '-'}</p>
                                </div>
                                <div class="flex items-center justify-between pt-1">
                                    \${t.customerSignature ? \`
                                        <div class="flex items-center gap-2">
                                            <img src="\${t.customerSignature}" class="h-8 max-w-[120px] object-contain border rounded p-0.5 bg-slate-50">
                                            <span class="text-[10px] text-emerald-700 font-bold">✓ เซ็นแล้ว</span>
                                        </div>
                                    \` : '<span class="text-[11px] text-slate-400 italic">ยังไม่ได้เซ็น</span>'}
                                    <button type="button" onclick="saveModalValuesWithoutClosing('\${t.id}'); openSignatureModal('\${t.id}', 'customer');" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer">
                                        <i class="fa-solid fa-signature mr-1"></i> \${t.customerSignature ? 'แก้ไขลายเซ็นลูกค้า' : '✍️ เซ็นชื่อลูกค้าบนหน้าจอ'}
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="flex justify-end space-x-2 pt-2 border-t text-xs">
                        <button onclick="document.getElementById('action-modal').remove()" class="bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold cursor-pointer">ปิด</button>
                        <button onclick="saveModalChanges('\${t.id}')" class="bg-orange-500 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">บันทึกข้อมูล</button>
                        <button onclick="submitForApproval('\${t.id}')" class="bg-purple-600 text-white px-5 py-2 rounded-xl font-bold shadow cursor-pointer">ส่ง Supervisor อนุมัติรายงาน</button>
                    </div>
                </div>
            \`;
            document.body.appendChild(modalBg);
        }

        function addChecklistRow(ticketId, eqIdx) {
            const t = state.tickets.find(x => x.id === ticketId);
            saveModalValuesWithoutClosing(ticketId);
            t.equipmentList[eqIdx].checklistResults.push({ item: "หัวข้อตรวจสอบเพิ่มเติม", measuredValue: "", unit: "-", acceptanceCriteria: "ตามมาตรฐานโรงงาน", beforeStatus: "NG", afterStatus: "OK" });
            persistState();
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        function removeChecklistRow(ticketId, eqIdx, chkIdx) {
            const t = state.tickets.find(x => x.id === ticketId);
            saveModalValuesWithoutClosing(ticketId);
            t.equipmentList[eqIdx].checklistResults.splice(chkIdx, 1);
            persistState();
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        function addPartItem(ticketId, eqIdx) {
            const t = state.tickets.find(x => x.id === ticketId);
            saveModalValuesWithoutClosing(ticketId);
            const name = document.getElementById(\`part-name-\${eqIdx}\`).value;
            const qty = parseFloat(document.getElementById(\`part-qty-\${eqIdx}\`).value) || 1;
            const price = parseFloat(document.getElementById(\`part-price-\${eqIdx}\`).value) || 0;
            const billingEl = document.getElementById(\`part-billing-\${eqIdx}\`);
            const billingType = billingEl ? billingEl.value : 'CUSTOMER_CHARGE';
            if (!name) { alert('กรุณากรอกชื่ออะไหล่'); return; }
            if (!t.equipmentList[eqIdx].parts) t.equipmentList[eqIdx].parts = [];
            t.equipmentList[eqIdx].parts.push({ name, qty, unit: 'PCS', price, total: qty * price, billingType });
            t.status = "PENDING_PARTS_APPROVAL";
            const lineMsg = formatTicketLineMessage(t, '⏳ [ขออนุมัติเบิกอะไหล่ด่วน! - รอ Supervisor อนุมัติ]');
            persistState(lineMsg);
            document.getElementById('action-modal').remove();
            openInternalActionModal(ticketId);
        }

        function submitForApproval(ticketId) {
            saveModalValuesWithoutClosing(ticketId);
            const t = state.tickets.find(x => x.id === ticketId);
            t.status = "PENDING_SUPERVISOR_APPROVAL";
            const lineMsg = formatTicketLineMessage(t, '✍️ [ช่างส่งรายงานซ่อมแล้ว! - รอ Supervisor อนุมัติ]');
            persistState(lineMsg);
            alert('ส่งรายงานให้ Supervisor อนุมัติ และแจ้งเตือนผ่าน LINE เรียบร้อย!');
            document.getElementById('action-modal').remove();
            render();
        }

        function saveModalValuesWithoutClosing(ticketId) {
            const t = state.tickets.find(x => x.id === ticketId);
            let siteContactInput = document.getElementById('modal-site-contact');
            let locationInput = document.getElementById('modal-location');
            let assignedTechInput = document.getElementById('modal-assigned-tech');
            let leadTechInput = document.getElementById('modal-lead-tech');
            if (siteContactInput) t.siteContact = siteContactInput.value;
            if (locationInput) t.location = locationInput.value;
            if (assignedTechInput) t.assignedTech = assignedTechInput.value;
            if (leadTechInput) t.leadTech = leadTechInput.value;

            t.equipmentList.forEach((eq, idx) => {
                let sym = document.getElementById(\`symptom-\${idx}\`);
                let det = document.getElementById(\`detail-\${idx}\`);
                let an = document.getElementById(\`analysis-\${idx}\`);
                let cs = document.getElementById(\`cause-\${idx}\`);
                let sr = document.getElementById(\`stdref-\${idx}\`);
                let fo = document.getElementById(\`fault-origin-\${idx}\`);
                let ac = document.getElementById(\`action-\${idx}\`);
                let bFile = document.getElementById(\`before-file-\${idx}\`);
                let aFile = document.getElementById(\`after-file-\${idx}\`);
                let bImg = document.getElementById(\`before-img-\${idx}\`);
                let aImg = document.getElementById(\`after-img-\${idx}\`);

                if (sym) eq.symptom = sym.value;
                if (det) eq.detail = det.value;
                if (an) eq.analysis = an.value;
                if (cs) eq.cause = cs.value;
                if (sr) eq.standardsReference = sr.value;
                if (fo) eq.faultOrigin = fo.value;
                if (ac) eq.correctiveAction = ac.value;

                if (bImg && bImg.value) {
                    eq.beforePhotos = [bImg.value];
                } else if (bFile && bFile.files && bFile.files[0]) {
                    eq.beforePhotos = [URL.createObjectURL(bFile.files[0])];
                }
                if (aImg && aImg.value) {
                    eq.afterPhotos = [aImg.value];
                } else if (aFile && aFile.files && aFile.files[0]) {
                    eq.afterPhotos = [URL.createObjectURL(aFile.files[0])];
                }

                if (eq.checklistResults) {
                    eq.checklistResults.forEach((chk, cIdx) => {
                        let itemInput = document.getElementById(\`chk-item-\${idx}-\${cIdx}\`);
                        let valInput = document.getElementById(\`chk-val-\${idx}-\${cIdx}\`);
                        let unitInput = document.getElementById(\`chk-unit-\${idx}-\${cIdx}\`);
                        let critInput = document.getElementById(\`chk-crit-\${idx}-\${cIdx}\`);
                        let bSel = document.getElementById(\`before-status-\${idx}-\${cIdx}\`);
                        let aSel = document.getElementById(\`after-status-\${idx}-\${cIdx}\`);
                        if (itemInput) chk.item = itemInput.value;
                        if (valInput) chk.measuredValue = valInput.value;
                        if (unitInput) chk.unit = unitInput.value;
                        if (critInput) chk.acceptanceCriteria = critInput.value;
                        if (bSel) chk.beforeStatus = bSel.value;
                        if (aSel) chk.afterStatus = aSel.value;
                    });
                }
            });
        }

        function saveModalChanges(ticketId) {
            saveModalValuesWithoutClosing(ticketId);
            persistState();
            alert('บันทึกข้อมูลเรียบร้อย!');
            document.getElementById('action-modal').remove();
            render();
        }

        function openServiceReport(id) {
            state.activeReportId = id;
            state.currentView = 'report_view';
            render();
        }

        function renderReportView(id) {
            const t = state.tickets.find(x => x.id === id) || state.tickets[0];
            if (!t) return '';
            let internalWarrantyCost = 0;
            let customerBilledCost = 0;
            (t.equipmentList || []).forEach(eq => {
                (eq.parts || []).forEach(p => {
                    const isWarr = p.billingType === 'WARRANTY_INTERNAL' || (!p.billingType && ((eq.faultOrigin || '').includes('ในประกัน') || (eq.faultOrigin || '').includes('ฟรี') || (t.warranty || '').includes('ในรับประกัน')));
                    if (isWarr) internalWarrantyCost += (p.total || 0);
                    else customerBilledCost += (p.total || 0);
                });
            });
            return \`
                <div class="max-w-3xl mx-auto mb-3 flex flex-wrap justify-between items-center gap-2 bg-white p-3.5 rounded-xl border no-print">
                    <button onclick="state.currentView = state.userRole === 'CUSTOMER' ? 'customer_portal' : 'internal_dashboard'; render();" class="bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold px-3.5 py-2 rounded-xl text-xs cursor-pointer"><i class="fa-solid fa-arrow-left mr-1"></i> ย้อนกลับ</button>
                    <div class="flex flex-wrap items-center gap-2">
                        <button onclick="openSignatureModal('\${t.id}', 'technician')" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-2 rounded-xl text-xs shadow cursor-pointer"><i class="fa-solid fa-pen-nib mr-1"></i> \${t.technicianSignature ? '✓ ลายเซ็นช่าง (แก้ไข)' : '✍️ ช่างเซ็นชื่อบนหน้าจอ'}</button>
                        <button onclick="openSignatureModal('\${t.id}', 'customer')" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3 py-2 rounded-xl text-xs shadow cursor-pointer"><i class="fa-solid fa-signature mr-1"></i> \${t.customerSignature ? '✓ ลายเซ็นลูกค้า (แก้ไข)' : '✍️ ลูกค้าเซ็นรับงานบนจอ'}</button>
                        <button onclick="downloadPdfById('standalone-service-report-sheet', 'Service-Report-\${t.id}.pdf')" class="bg-orange-500 hover:bg-orange-600 text-white font-bold px-4 py-2 rounded-xl text-xs shadow cursor-pointer"><i class="fa-solid fa-file-pdf mr-1"></i> บันทึกเป็นไฟล์ PDF (Save PDF)</button>
                        <button onclick="window.print()" class="bg-slate-800 hover:bg-slate-900 text-white font-bold px-3.5 py-2 rounded-xl text-xs cursor-pointer"><i class="fa-solid fa-print mr-1"></i> สั่งพิมพ์ (Print A4)</button>
                    </div>
                </div>
                <div id="standalone-service-report-sheet" class="max-w-3xl mx-auto bg-white p-6 sm:p-8 rounded-xl shadow border space-y-4 text-xs sm:text-sm print-container">
                    <div class="flex justify-between items-center border-b-4 border-orange-500 pb-3">
                        <div>
                            <h1 class="text-base sm:text-lg font-extrabold text-slate-900">บริษัท สหคิมมอเตอร์ จำกัด</h1>
                            <p class="text-[11px] text-orange-600 font-bold">SAHA KIM MOTOR CO., LTD. (เชี่ยวชาญงาน ใส่ใจบริการ SKM SERVICE)</p>
                            <p class="text-[10px] text-slate-500">บริการซ่อม บำรุงรักษา และจำหน่ายปั๊มน้ำ มอเตอร์ Blower อุตสาหกรรม</p>
                        </div>
                        <div class="text-right bg-slate-50 p-2.5 rounded-lg border">
                            <h2 class="text-xs sm:text-sm font-extrabold text-orange-600">SERVICE REPORT</h2>
                            <p class="text-[11px] font-bold">เลขที่ใบงาน: \${t.id}</p>
                        </div>
                    </div>

                    <div class="bg-blue-50 border border-blue-200 p-3 rounded-lg space-y-1 text-xs">
                        <p class="font-bold text-blue-900 text-sm border-b border-blue-200 pb-1">ส่วนที่ 1: ข้อมูลลูกค้าผู้แจ้งซ่อม (Customer Information)</p>
                        <p><b>บริษัท:</b> \${t.company} (TAX ID: \${t.taxId || '-'}) | <b>เลขที่เอกสาร/ประกัน:</b> \${t.docNo || '-'}</p>
                        <p><b>ผู้ติดต่อตอนแจ้ง:</b> \${t.contact}</p>
                        <p><b>ประเภทบริการ:</b> \${t.serviceType} | <b>เงื่อนไข:</b> <span class="text-indigo-600 font-bold">\${t.warranty}</span> | <b>ความเร่งด่วน:</b> <span class="text-red-600 font-bold">\${t.urgency}</span></p>
                        <p><b>สถานที่ปฏิบัติงาน:</b> \${t.location || '-'}</p>
                    </div>

                    <div class="bg-slate-50 border border-slate-200 p-3 rounded-lg space-y-1 text-xs">
                        <p class="font-bold text-slate-900 text-sm border-b border-slate-200 pb-1">ส่วนที่ 2: ข้อมูลการดำเนินงานโดยทีมงานบริษัท (Company Operations)</p>
                        <p><b>ช่างผู้ปฏิบัติงาน:</b> <span class="font-bold text-slate-800">\${t.assignedTech || 'ยังไม่ระบุ'}</span> | <b>หัวหน้าช่างผู้ควบคุมงาน:</b> <span class="font-bold text-indigo-700">\${t.leadTech || (state.technicians.find(tc => (tc.position || '').includes('หัวหน้าช่าง'))?.name || '-')}</span></p>
                        <p><b>ผู้รับหน้างานจริง:</b> <span class="font-bold text-indigo-600">\${t.siteContact || t.contact}</span></p>
                        <p><b>วันที่นัดหมาย:</b> <span class="text-emerald-600 font-bold">\${t.confirmedDate || t.desiredDate}</span> | <b>ผู้อนุมัติ (Supervisor):</b> <span class="font-bold text-emerald-700">\${t.approvedBySupervisor || 'รออนุมัติ'}</span></p>
                        <p><b>สถานะใบงาน:</b> <span class="font-bold text-orange-600">\${t.status}</span></p>
                    </div>

                    <div class="space-y-3">
                        <h3 class="font-bold text-xs sm:text-sm border-b pb-1">ส่วนที่ 3: รายการอุปกรณ์และการตรวจเช็คมาตรฐาน</h3>
                        \${t.equipmentList.map((eq, idx) => \`
                            <div class="border rounded-lg p-3 space-y-2 bg-slate-50 text-xs">
                                <p class="font-bold text-orange-600">อุปกรณ์ #\${idx + 1}: \${eq.type} (\${eq.brand} - \${eq.model}) [Tag: \${eq.tag || 'N/A'}]</p>
                                <p><b>อาการเสีย:</b> \${eq.symptom} (\${eq.detail || '-'})</p>
                                <div class="bg-white p-2.5 rounded border space-y-1">
                                    <p class="font-bold text-indigo-700">ผลการวิเคราะห์ปัญหา:</p>
                                    <p><b>วิเคราะห์:</b> \${eq.analysis || '-'}</p>
                                    <p><b>สาเหตุ:</b> \${eq.cause || '-'}</p>
                                    <p><b>มาตรฐานอ้างอิง:</b> \${eq.standardsReference || '-'}</p>
                                    <p><b>เงื่อนไข:</b> <span class="font-bold text-orange-600">\${eq.faultOrigin || '-'}</span></p>
                                    <p><b>แนวทางแก้ไข:</b> \${eq.correctiveAction || '-'}</p>
                                </div>
                                \${eq.checklistResults ? \`
                                    <div>
                                        <p class="font-bold mb-1">ผลการตรวจเช็คมาตรฐาน (Checklist - ค่าที่วัดได้จริง & เกณฑ์ยอมรับ):</p>
                                        <table class="w-full border-collapse bg-white rounded border text-[11px]">
                                            <tr class="bg-slate-100 border-b">
                                                <th class="p-1 text-left">หัวข้อการตรวจสอบ</th>
                                                <th class="p-1 text-center w-20">ค่าที่วัดได้</th>
                                                <th class="p-1 text-center w-14">หน่วย</th>
                                                <th class="p-1 text-left w-28">เกณฑ์ยอมรับ</th>
                                                <th class="p-1 text-center w-14">Before</th>
                                                <th class="p-1 text-center w-14">After</th>
                                            </tr>
                                            \${eq.checklistResults.map(chk => \`<tr class="border-b">
                                                <td class="p-1 font-medium">\${chk.item}</td>
                                                <td class="p-1 text-center font-bold text-indigo-700">\${chk.measuredValue || '-'}</td>
                                                <td class="p-1 text-center text-slate-600">\${chk.unit || '-'}</td>
                                                <td class="p-1 text-slate-600">\${chk.acceptanceCriteria || 'ตามมาตรฐานโรงงาน'}</td>
                                                <td class="p-1 text-center font-bold \${chk.beforeStatus==='NG'?'text-red-600':chk.beforeStatus==='N/A'?'text-slate-500':'text-emerald-600'}">\${chk.beforeStatus}</td>
                                                <td class="p-1 text-center font-bold \${chk.afterStatus==='OK'?'text-emerald-600':chk.afterStatus==='N/A'?'text-slate-500':'text-red-600'}">\${chk.afterStatus}</td>
                                            </tr>\`).join('')}
                                        </table>
                                    </div>
                                \` : ''}
                                \${eq.parts && eq.parts.length > 0 ? \`
                                    <div>
                                        <p class="font-bold mb-1">อะไหล่ที่เบิกใช้ (แยกต้นทุนภายใน vs ยอดเรียกเก็บลูกค้า):</p>
                                        <table class="w-full border-collapse bg-white rounded border text-[11px]">
                                            <tr class="bg-slate-100 border-b">
                                                <th class="p-1 text-left">รายการ</th>
                                                <th class="p-1 text-center w-16">จน.</th>
                                                <th class="p-1 text-center w-32">เงื่อนไขค่าใช้จ่าย</th>
                                                <th class="p-1 text-right w-28">ต้นทุนภายใน</th>
                                                <th class="p-1 text-right w-28">ยอดเรียกเก็บลูกค้า</th>
                                            </tr>
                                            \${eq.parts.map(p => {
                                                const isWarr = p.billingType === 'WARRANTY_INTERNAL' || (!p.billingType && ((eq.faultOrigin || '').includes('ในประกัน') || (eq.faultOrigin || '').includes('ฟรี') || (t.warranty || '').includes('ในรับประกัน')));
                                                return \`<tr class="border-b">
                                                    <td class="p-1">\${p.name}</td>
                                                    <td class="p-1 text-center">\${p.qty} \${p.unit}</td>
                                                    <td class="p-1 text-center font-bold \${isWarr ? 'text-emerald-700' : 'text-amber-700'}">\${isWarr ? 'เคลมในประกัน (FOC)' : 'คิดค่าบริการ'}</td>
                                                    <td class="p-1 text-right text-slate-500">\${isWarr ? p.total.toLocaleString() : '-'}</td>
                                                    <td class="p-1 text-right font-bold \${isWarr ? 'text-emerald-700' : 'text-slate-900'}">\${isWarr ? '0.00 (ฟรี)' : p.total.toLocaleString()}</td>
                                                </tr>\`;
                                            }).join('')}
                                        </table>
                                    </div>
                                \` : ''}
                            </div>
                        \`).join('')}
                    </div>

                    <div class="border-2 rounded-xl p-3.5 text-xs \${customerBilledCost === 0 ? 'bg-emerald-50 border-emerald-300' : 'bg-amber-50 border-amber-300'}">
                        <p class="font-extrabold text-sm mb-2">ส่วนที่ 4: สรุปแยกต้นทุนภายในบริษัท และ ยอดเรียกเก็บลูกค้าสุทธิ (Cost & Warranty Summary)</p>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div class="bg-white p-2.5 rounded-lg border border-slate-200">
                                <p class="text-[11px] font-bold text-slate-500">ต้นทุนภายในบริษัท (Internal Warranty Cost - บริษัทรับภาระ):</p>
                                <p class="text-base font-extrabold text-slate-700">\${internalWarrantyCost.toLocaleString()} THB <span class="text-[10px] text-emerald-700 font-bold">(ไม่เรียกเก็บลูกค้า)</span></p>
                            </div>
                            <div class="bg-white p-2.5 rounded-lg border-2 \${customerBilledCost === 0 ? 'border-emerald-500' : 'border-orange-500'}">
                                <p class="text-[11px] font-bold text-slate-700">ยอดเรียกเก็บลูกค้าสุทธิ (Net Customer Payable Amount):</p>
                                <p class="text-lg font-extrabold \${customerBilledCost === 0 ? 'text-emerald-700' : 'text-orange-600'}">\${customerBilledCost === 0 ? '0.00 THB (อยู่ในประกัน ฟรี / ไม่มียอดเรียกเก็บ)' : customerBilledCost.toLocaleString() + ' THB'}</p>
                            </div>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-6 pt-6 border-t text-center text-xs">
                        <div class="flex flex-col items-center justify-end min-h-[110px]">
                            \${t.technicianSignature ? \`
                                <img src="\${t.technicianSignature}" class="h-14 w-auto object-contain mx-auto mb-1">
                                <button type="button" onclick="openSignatureModal('\${t.id}', 'technician')" class="no-print text-[10px] text-blue-600 hover:underline cursor-pointer mb-1">✎ แก้ไขลายเซ็นช่าง</button>
                            \` : \`
                                <button type="button" onclick="openSignatureModal('\${t.id}', 'technician')" class="no-print mb-2 px-3 py-1 bg-blue-50 text-blue-800 border border-blue-200 rounded-lg text-xs font-bold hover:bg-blue-600 hover:text-white cursor-pointer transition">✍️ เซ็นชื่อช่างบนหน้าจอ</button>
                            \`}
                            <p class="text-slate-400">___________________________</p>
                            <p class="mt-1 font-bold">(\${t.assignedTech || 'ช่างผู้ปฏิบัติงาน'})</p>
                            <p class="text-slate-500">ช่างผู้ปฏิบัติงาน (Technician)</p>
                            \${t.technicianSignedAt ? \`<p class="text-[10px] font-mono text-slate-400 mt-0.5">เซ็นเมื่อ: \${t.technicianSignedAt}</p>\` : ''}
                        </div>
                        <div class="flex flex-col items-center justify-end min-h-[110px]">
                            \${t.approvedBySupervisor ? \`
                                <span class="inline-block bg-emerald-100 border border-emerald-300 text-emerald-800 font-bold text-[10px] px-2.5 py-1 rounded-md mb-2">✓ APPROVED DIGITAL STAMP</span>
                            \` : \`
                                <div class="h-8 mb-2"></div>
                            \`}
                            <p class="text-slate-400">___________________________</p>
                            <p class="mt-1 font-bold">(\${t.approvedBySupervisor || t.leadTech || 'หัวหน้าช่าง / Supervisor'})</p>
                            <p class="text-slate-500">หัวหน้าช่าง / Supervisor ผู้อนุมัติ</p>
                        </div>
                        <div class="flex flex-col items-center justify-end min-h-[110px]">
                            \${t.customerSignature ? \`
                                <img src="\${t.customerSignature}" class="h-14 w-auto object-contain mx-auto mb-1">
                                <button type="button" onclick="openSignatureModal('\${t.id}', 'customer')" class="no-print text-[10px] text-emerald-700 hover:underline cursor-pointer mb-1">✎ แก้ไขลายเซ็นลูกค้า</button>
                            \` : \`
                                <button type="button" onclick="openSignatureModal('\${t.id}', 'customer')" class="no-print mb-2 px-3 py-1 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-lg text-xs font-bold hover:bg-emerald-600 hover:text-white cursor-pointer transition">✍️ เซ็นรับงานบนหน้าจอ</button>
                            \`}
                            <p class="text-slate-400">___________________________</p>
                            <p class="mt-1 font-bold">(\${t.customerSignerName || t.siteContact || t.contact || 'ลูกค้าผู้ตรวจรับงาน'})</p>
                            <p class="text-slate-500">ลูกค้าผู้ตรวจรับงาน (Customer Acceptance)</p>
                            \${t.customerSignedAt ? \`<p class="text-[10px] font-mono text-slate-400 mt-0.5">เซ็นเมื่อ: \${t.customerSignedAt}</p>\` : ''}
                        </div>
                    </div>

                    <div class="bg-emerald-50 border border-emerald-200 p-2.5 rounded-lg text-center font-bold text-emerald-800 text-xs">
                        Supervisor / หัวหน้าช่าง ผู้อนุมัติ: \${t.approvedBySupervisor || 'รออนุมัติ'}
                    </div>

                    <div class="flex flex-wrap justify-center gap-2 pt-2 no-print">
                        <button onclick="downloadPdfById('standalone-service-report-sheet', 'Service-Report-\${t.id}.pdf')" class="bg-orange-500 hover:bg-orange-600 text-white font-bold px-6 py-2 rounded-xl text-xs shadow cursor-pointer"><i class="fa-solid fa-file-pdf mr-1"></i> บันทึกเป็นไฟล์ PDF (Save PDF)</button>
                        <button onclick="window.print()" class="bg-slate-800 hover:bg-slate-900 text-white font-bold px-5 py-2 rounded-xl text-xs shadow cursor-pointer"><i class="fa-solid fa-print mr-1"></i> สั่งพิมพ์ (Print A4)</button>
                    </div>
                </div>
            \`;
        }

        render();
        pullFromCloudSyncServer(false);
        if (gasWebAppUrl) loadFromGoogleSheet(false);
    <\/script>
</body>
</html>`;

  return htmlContent;
}

export function downloadStandaloneHtml(
  currentTickets: ServiceTicket[],
  currentTechnicians: Technician[],
  currentTemplates: TroubleshootingTemplate[],
  customEquipTypes: string[],
  customBrands: string[],
  supPassword: string,
  servicePin = '1234',
  gasWebAppUrl = '',
  lineChannelToken = '',
  lineTargetId = '',
  currentSupervisors: string[] = [
    'หัวหน้าวิศวกรบริการ (Supervisor)',
    'ผู้จัดการฝ่ายบริการ (Service Manager)'
  ]
) {
  const htmlContent = generateStandaloneHtmlString(
    currentTickets,
    currentTechnicians,
    currentTemplates,
    customEquipTypes,
    customBrands,
    supPassword,
    servicePin,
    gasWebAppUrl,
    lineChannelToken,
    lineTargetId,
    currentSupervisors
  );
  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'index.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

