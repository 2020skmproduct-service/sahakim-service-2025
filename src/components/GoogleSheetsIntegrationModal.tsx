import React, { useState, useEffect } from 'react';
import {
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  ExternalLink,
  RefreshCw,
  PlusCircle,
  FolderOpen,
  Link as LinkIcon,
  DownloadCloud,
  UploadCloud,
  LogOut,
  X,
  Check,
  ShieldCheck,
  Clock,
  Sparkles,
  Zap,
  Trash2
} from 'lucide-react';
import { User } from 'firebase/auth';
import { ServiceTicket } from '../types';
import {
  googleSignIn,
  logoutGoogle,
  getAccessToken,
  initAuth
} from '../services/googleAuth';
import {
  createSahakimSpreadsheet,
  syncTicketsToSpreadsheet,
  loadTicketsFromSpreadsheet,
  loadTechniciansFromSpreadsheet,
  listUserSpreadsheets,
  extractSpreadsheetId,
  SpreadsheetFileInfo
} from '../services/googleSheets';

interface GoogleSheetsIntegrationModalProps {
  isOpen: boolean;
  onClose: () => void;
  tickets: ServiceTicket[];
  onTicketsImported: (importedTickets: ServiceTicket[]) => void;
  onTechniciansImported: (technicians: import('../types').Technician[]) => void;
  connectedSheetId: string;
  onSetConnectedSheetId: (id: string, name?: string) => void;
  connectedSheetName: string;
  autoSyncEnabled: boolean;
  onToggleAutoSync: (enabled: boolean) => void;
  lastSyncedAt: string | null;
  onSetLastSyncedAt: (timestamp: string) => void;
  showToast: (msg: string, type?: 'success' | 'warning' | 'error' | 'info') => void;
}

export const GoogleSheetsIntegrationModal: React.FC<GoogleSheetsIntegrationModalProps> = ({
  isOpen,
  onClose,
  tickets,
  onTicketsImported,
  onTechniciansImported,
  connectedSheetId,
  onSetConnectedSheetId,
  connectedSheetName,
  autoSyncEnabled,
  onToggleAutoSync,
  lastSyncedAt,
  onSetLastSyncedAt,
  showToast
}) => {
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [cachedToken, setCachedToken] = useState<string | null>(null);
  const [isAuthenticating, setIsAuthenticating] = useState(false);
  const [isOperating, setIsOperating] = useState(false);
  const [operationProgress, setOperationProgress] = useState<string>('');

  // Drive sheets list
  const [driveFiles, setDriveFiles] = useState<SpreadsheetFileInfo[]>([]);
  const [isLoadingDriveFiles, setIsLoadingDriveFiles] = useState(false);
  const [customSheetInput, setCustomSheetInput] = useState('');

  // Confirmation dialogs for destructive operations (Mandatory per Workspace Skill)
  const [confirmSyncOpen, setConfirmSyncOpen] = useState(false);
  const [confirmImportOpen, setConfirmImportOpen] = useState(false);
  const [activeTab, setActiveTab] = useState<'direct' | 'drive' | 'custom'>('direct');

  useEffect(() => {
    const unsubscribe = initAuth(
      (user, token) => {
        setCurrentUser(user);
        setCachedToken(token);
      },
      () => {
        setCurrentUser(null);
        setCachedToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleSignIn = async () => {
    setIsAuthenticating(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setCurrentUser(result.user);
        setCachedToken(result.accessToken);
        showToast(`เข้าสู่ระบบ Google สำเร็จ (${result.user.email})`, 'success');
        // Fetch drive files automatically
        fetchDriveFiles(result.accessToken);
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ Google';
      showToast(msg, 'error');
    } finally {
      setIsAuthenticating(false);
    }
  };

  const handleSignOut = async () => {
    try {
      await logoutGoogle();
      setCurrentUser(null);
      setCachedToken(null);
      setDriveFiles([]);
      showToast('ออกจากระบบ Google เรียบร้อยแล้ว', 'info');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'เกิดข้อผิดพลาดในการออกจากระบบ';
      showToast(msg, 'error');
    }
  };

  const fetchDriveFiles = async (tokenOverride?: string) => {
    const token = tokenOverride || cachedToken || (await getAccessToken());
    if (!token) return;
    setIsLoadingDriveFiles(true);
    try {
      const files = await listUserSpreadsheets(token);
      setDriveFiles(files);
    } catch (err: unknown) {
      console.warn('Could not list drive files:', err);
    } finally {
      setIsLoadingDriveFiles(false);
    }
  };

  // 1. Create Brand New Spreadsheet
  const handleCreateNewSpreadsheet = async () => {
    let token = cachedToken || (await getAccessToken());
    if (!token) {
      showToast('กรุณาเข้าสู่ระบบ Google ก่อนสร้างตาราง', 'warning');
      return;
    }
    setIsOperating(true);
    setOperationProgress('กำลังสร้างตาราง Google Sheet ใน Google Drive ของคุณ...');
    try {
      const result = await createSahakimSpreadsheet(token, tickets);
      onSetConnectedSheetId(result.spreadsheetId, result.title);
      const nowStr = new Date().toLocaleString('th-TH');
      onSetLastSyncedAt(nowStr);
      showToast(`สร้างตาราง "${result.title}" ใน Google Drive และซิงก์ข้อมูล ${tickets.length} รายการสำเร็จ!`, 'success');
      fetchDriveFiles(token);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'สร้างตารางไม่สำเร็จ';
      showToast(msg, 'error');
    } finally {
      setIsOperating(false);
      setOperationProgress('');
    }
  };

  // 2. Perform Sync / Overwrite (After explicit confirmation dialog)
  const handleExecuteSync = async () => {
    setConfirmSyncOpen(false);
    let token = cachedToken || (await getAccessToken());
    if (!token) {
      showToast('กรุณาเข้าสู่ระบบ Google ก่อนซิงก์ข้อมูล', 'warning');
      return;
    }
    if (!connectedSheetId) {
      showToast('ยังไม่ได้ระบุตาราง Google Sheet ที่จะซิงก์', 'warning');
      return;
    }
    if (tickets.length === 0) {
      showToast('ในระบบปัจจุบันมี 0 ใบงาน หากต้องการดึงข้อมูลที่มีอยู่ใน Google Sheet มาแสดง ให้กดปุ่ม "ดึงข้อมูลจาก Google Sheet ลงระบบ" แทนครับ', 'warning');
      return;
    }

    setIsOperating(true);
    setOperationProgress(`กำลังบันทึกข้อมูล ${tickets.length} ใบงานลงใน Google Sheet...`);
    try {
      await syncTicketsToSpreadsheet(token, connectedSheetId, tickets);
      const nowStr = new Date().toLocaleString('th-TH');
      onSetLastSyncedAt(nowStr);
      showToast(`ซิงก์ข้อมูล ${tickets.length} ใบงานลงใน Google Sheet สำเร็จเรียบร้อย!`, 'success');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'ซิงก์ข้อมูลไม่สำเร็จ';
      showToast(msg, 'error');
    } finally {
      setIsOperating(false);
      setOperationProgress('');
    }
  };

  // Read-only import of the authoritative technician master sheet
  const handleImportTechnicians = async () => {
    const token = cachedToken || (await getAccessToken());
    if (!token) {
      showToast('กรุณาเข้าสู่ระบบ Google ก่อนนำเข้าทะเบียนช่าง', 'warning');
      return;
    }
    if (!connectedSheetId) {
      showToast('ยังไม่ได้เชื่อมต่อ Google Sheet', 'warning');
      return;
    }
    setIsOperating(true);
    setOperationProgress('กำลังอ่านทะเบียนช่างจากรายชื่อช่าง_Technicians (ไม่แก้ไขชีตต้นทาง)...');
    try {
      const imported = await loadTechniciansFromSpreadsheet(token, connectedSheetId);
      if (imported.length === 0) {
        showToast('ไม่พบรายชื่อช่างในแท็บรายชื่อช่าง_Technicians', 'warning');
        return;
      }
      onTechniciansImported(imported);
      showToast(`นำเข้ารายชื่อช่างจริง ${imported.length} คนสำเร็จ (อ่านอย่างเดียว ไม่เขียนทับชีต)`, 'success');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'อ่านทะเบียนช่างไม่สำเร็จ';
      showToast(msg, 'error');
    } finally {
      setIsOperating(false);
      setOperationProgress('');
    }
  };

  // 3. Perform Import / Load (After explicit confirmation dialog)
  const handleExecuteImport = async () => {
    setConfirmImportOpen(false);
    let token = cachedToken || (await getAccessToken());
    if (!token) {
      showToast('กรุณาเข้าสู่ระบบ Google ก่อนดึงข้อมูล', 'warning');
      return;
    }
    if (!connectedSheetId) {
      showToast('ยังไม่ได้ระบุตาราง Google Sheet ที่จะดึงข้อมูล', 'warning');
      return;
    }

    setIsOperating(true);
    setOperationProgress('กำลังอ่านข้อมูลใบงานจาก Google Sheet...');
    try {
      const imported = await loadTicketsFromSpreadsheet(token, connectedSheetId);
      if (imported.length === 0) {
        showToast('ไม่พบข้อมูลใบงานในตาราง หรือตารางยังว่างอยู่', 'warning');
        return;
      }
      onTicketsImported(imported);
      const nowStr = new Date().toLocaleString('th-TH');
      onSetLastSyncedAt(nowStr);
      showToast(`ดึงข้อมูลใบงานสำเร็จ (${imported.length} ใบงาน)`, 'success');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'ดึงข้อมูลไม่สำเร็จ';
      showToast(msg, 'error');
    } finally {
      setIsOperating(false);
      setOperationProgress('');
    }
  };

  const handleConnectCustomSheet = () => {
    const id = extractSpreadsheetId(customSheetInput);
    if (!id) {
      showToast('กรุณากรอก Spreadsheet URL หรือ ID ที่ถูกต้อง', 'warning');
      return;
    }
    onSetConnectedSheetId(id, `ตาราง ${id.substring(0, 8)}...`);
    setCustomSheetInput('');
    showToast('เชื่อมต่อตาราง Google Sheet สำเร็จ!', 'success');
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/70 backdrop-blur-xs overflow-y-auto">
      <div className="relative w-full max-w-2xl bg-white rounded-3xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-6 py-5 bg-gradient-to-r from-emerald-600 via-teal-700 to-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-white/10 backdrop-blur-xs flex items-center justify-center border border-white/20 shadow-inner">
              <FileSpreadsheet className="w-5 h-5 text-emerald-300" />
            </div>
            <div>
              <h3 className="font-bold text-base sm:text-lg flex items-center gap-2">
                เชื่อมต่อ Google Sheets โดยตรง
                <span className="text-[10px] bg-emerald-400 text-slate-900 font-extrabold px-2 py-0.5 rounded-full">
                  Google Workspace
                </span>
              </h3>
              <p className="text-xs text-emerald-100">
                ซิงก์ข้อมูลใบงาน บันทึกซ่อม และสถิติลงใน Google Drive และ Google Sheets แบบเรียลไทม์
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-white/10 hover:bg-white/20 flex items-center justify-center text-white transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 sm:p-6 space-y-6 overflow-y-auto">
          {/* User Auth Section */}
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 sm:p-5">
            {!currentUser ? (
              <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                <div className="space-y-1 text-center sm:text-left">
                  <h4 className="font-bold text-slate-800 text-sm flex items-center gap-2 justify-center sm:justify-start">
                    <ShieldCheck className="w-4 h-4 text-emerald-600" />
                    เข้าสู่ระบบ Google เพื่อเปิดใช้งานการซิงก์ตาราง
                  </h4>
                  <p className="text-xs text-slate-600">
                    ระบบจะขอสิทธิ์เข้าถึง Google Sheets และ Google Drive ในการสร้างและอัปเดตข้อมูลใบงาน โดยได้รับอนุญาตจากคุณ
                  </p>
                </div>
                <div>
                  <button
                    onClick={handleSignIn}
                    disabled={isAuthenticating}
                    className="gsi-material-button shadow-xs hover:shadow-md cursor-pointer transition-all"
                  >
                    <div className="gsi-material-button-state"></div>
                    <div className="gsi-material-button-content-wrapper">
                      <div className="gsi-material-button-icon">
                        <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
                          <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                          <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                          <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                          <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                          <path fill="none" d="M0 0h48v48H0z"></path>
                        </svg>
                      </div>
                      <span className="gsi-material-button-contents">
                        {isAuthenticating ? 'กำลังเชื่อมต่อ...' : 'เข้าสู่ระบบด้วย Google'}
                      </span>
                    </div>
                  </button>
                </div>
              </div>
            ) : (
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  {currentUser.photoURL ? (
                    <img
                      src={currentUser.photoURL}
                      alt={currentUser.displayName || 'Google User'}
                      className="w-10 h-10 rounded-full border border-slate-200"
                    />
                  ) : (
                    <div className="w-10 h-10 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center text-sm">
                      {(currentUser.displayName || currentUser.email || 'G')[0].toUpperCase()}
                    </div>
                  )}
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="font-bold text-slate-800 text-sm">
                        {currentUser.displayName || 'Google Account'}
                      </span>
                      <span className="text-[11px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full font-semibold flex items-center gap-1">
                        <CheckCircle2 className="w-3 h-3 text-emerald-600" /> เชื่อมต่อแล้ว
                      </span>
                    </div>
                    <p className="text-xs text-slate-500">{currentUser.email}</p>
                  </div>
                </div>
                <button
                  onClick={handleSignOut}
                  className="text-xs text-rose-600 hover:text-rose-700 hover:bg-rose-50 px-3 py-1.5 rounded-xl border border-rose-200 font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
                >
                  <LogOut className="w-3.5 h-3.5" />
                  ออกจากระบบ Google
                </button>
              </div>
            )}
          </div>

          {/* Connected Spreadsheet Box */}
          {connectedSheetId ? (
            <div className="bg-gradient-to-br from-emerald-50 via-teal-50 to-slate-50 border border-emerald-200 rounded-2xl p-5 space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                    <span className="text-xs font-bold text-emerald-800 uppercase tracking-wider">
                      ตารางที่เชื่อมต่ออยู่
                    </span>
                  </div>
                  <h4 className="font-bold text-slate-900 text-base mt-1">
                    {connectedSheetName || 'SAHA KIM MOTOR - Service Database'}
                  </h4>
                  <p className="font-mono text-xs text-slate-500 break-all">
                    ID: {connectedSheetId}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <a
                    href={`https://docs.google.com/spreadsheets/d/${connectedSheetId}`}
                    target="_blank"
                    rel="noreferrer"
                    className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3.5 py-2 rounded-xl text-xs shadow-xs transition-colors"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    เปิดใน Google Sheets
                  </a>
                  <button
                    onClick={() => onSetConnectedSheetId('', '')}
                    title="ยกเลิกการเชื่อมต่อตารางนี้"
                    className="p-2 rounded-xl text-slate-400 hover:text-rose-600 hover:bg-white transition-colors"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>

              {/* Status and Last Synced Info */}
              <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-emerald-100 text-xs text-slate-600">
                <div className="flex items-center gap-1.5">
                  <Clock className="w-3.5 h-3.5 text-slate-400" />
                  <span>
                    ซิงก์ล่าสุด: {lastSyncedAt ? lastSyncedAt : 'ยังไม่มีการซิงก์ในรอบนี้'}
                  </span>
                </div>
                <label className="flex items-center gap-2 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={autoSyncEnabled}
                    onChange={(e) => onToggleAutoSync(e.target.checked)}
                    className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                  />
                  <span className="font-semibold text-slate-700 flex items-center gap-1">
                    <Zap className="w-3 h-3 text-amber-500" />
                    ซิงก์อัตโนมัติเมื่อบันทึกใบงาน
                  </span>
                </label>
              </div>

              {/* Sync Action Buttons */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                <button
                  type="button"
                  onClick={() => setConfirmSyncOpen(true)}
                  disabled={isOperating || !currentUser}
                  className="flex items-center justify-center gap-2 bg-slate-900 hover:bg-slate-800 text-white font-bold py-2.5 px-4 rounded-xl text-xs shadow-xs transition-all cursor-pointer disabled:opacity-50"
                >
                  <UploadCloud className="w-4 h-4 text-emerald-400" />
                  ส่งข้อมูลไป Google Sheet ({tickets.length} ใบงาน)
                </button>

                <button
                  type="button"
                  onClick={() => setConfirmImportOpen(true)}
                  disabled={isOperating || !currentUser}
                  className="flex items-center justify-center gap-2 bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 font-bold py-2.5 px-4 rounded-xl text-xs shadow-xs transition-all cursor-pointer disabled:opacity-50"
                >
                  <DownloadCloud className="w-4 h-4 text-teal-600" />
                  ดึงข้อมูลจาก Google Sheet ลงระบบ
                </button>
                <button
                  type="button"
                  onClick={handleImportTechnicians}
                  disabled={isOperating || !currentUser}
                  className="sm:col-span-2 flex items-center justify-center gap-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 font-bold py-2.5 px-4 rounded-xl text-xs shadow-xs transition-all cursor-pointer disabled:opacity-50"
                >
                  <DownloadCloud className="w-4 h-4" />
                  นำเข้าทะเบียนช่างจากรายชื่อช่าง_Technicians (อ่านอย่างเดียว)
                </button>
              </div>
            </div>
          ) : (
            /* Connect or Create New Spreadsheet Box */
            <div className="space-y-4">
              <div className="flex border-b border-slate-200 gap-2">
                <button
                  type="button"
                  onClick={() => setActiveTab('direct')}
                  className={`pb-2.5 text-xs font-bold border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
                    activeTab === 'direct'
                      ? 'border-emerald-600 text-emerald-700'
                      : 'border-transparent text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  สร้างตารางใหม่อัตโนมัติ (แนะนำ)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setActiveTab('drive');
                    if (currentUser && driveFiles.length === 0) fetchDriveFiles();
                  }}
                  className={`pb-2.5 text-xs font-bold border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
                    activeTab === 'drive'
                      ? 'border-emerald-600 text-emerald-700'
                      : 'border-transparent text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <FolderOpen className="w-3.5 h-3.5" />
                  เลือกจาก Google Drive
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('custom')}
                  className={`pb-2.5 text-xs font-bold border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
                    activeTab === 'custom'
                      ? 'border-emerald-600 text-emerald-700'
                      : 'border-transparent text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <LinkIcon className="w-3.5 h-3.5" />
                  วางลิงก์ตารางที่มีอยู่
                </button>
              </div>

              {/* Tab 1: Create New */}
              {activeTab === 'direct' && (
                <div className="bg-emerald-50/70 border border-emerald-200 rounded-2xl p-5 space-y-3">
                  <div className="flex items-start gap-3">
                    <div className="w-8 h-8 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-xs">
                      <Sparkles className="w-4 h-4" />
                    </div>
                    <div>
                      <h4 className="font-bold text-slate-900 text-sm">
                        สร้างตาราง Google Sheet ของสหคิมมอเตอร์ใหม่อัตโนมัติ
                      </h4>
                      <p className="text-xs text-slate-600 mt-1 leading-relaxed">
                        ระบบจะสร้างตารางชื่อ <code className="font-mono bg-white px-1.5 py-0.5 rounded border border-emerald-200 font-bold text-emerald-900">SAHA KIM MOTOR - Service &amp; Repair Database</code> บน Google Drive ของคุณทันที พร้อมจัดฟอร์แมตหัวตารางสีกรมท่า ปรับความกว้างคอลัมน์อัตโนมัติ และซิงก์ข้อมูลใบงานปัจจุบันทั้งหมด {tickets.length} ใบงาน
                      </p>
                    </div>
                  </div>
                  <div className="pt-2">
                    <button
                      type="button"
                      onClick={handleCreateNewSpreadsheet}
                      disabled={isOperating || !currentUser}
                      className="w-full sm:w-auto inline-flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded-xl text-xs shadow-md transition-all cursor-pointer disabled:opacity-50"
                    >
                      <PlusCircle className="w-4 h-4" />
                      คลิกสร้างตารางใหม่และเชื่อมต่อทันที
                    </button>
                  </div>
                </div>
              )}

              {/* Tab 2: Drive Browser */}
              {activeTab === 'drive' && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-slate-600">
                      ตาราง Google Sheet ล่าสุดใน Google Drive ของคุณ:
                    </span>
                    <button
                      type="button"
                      onClick={() => fetchDriveFiles()}
                      disabled={isLoadingDriveFiles || !currentUser}
                      className="text-xs text-emerald-700 hover:text-emerald-800 flex items-center gap-1 font-semibold cursor-pointer"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isLoadingDriveFiles ? 'animate-spin' : ''}`} />
                      รีเฟรชรายการ
                    </button>
                  </div>

                  {!currentUser ? (
                    <div className="p-6 text-center bg-slate-50 rounded-2xl border border-slate-200 text-xs text-slate-500">
                      กรุณาเข้าสู่ระบบ Google ด้านบนก่อนเพื่อค้นหาตารางใน Google Drive
                    </div>
                  ) : isLoadingDriveFiles ? (
                    <div className="p-8 text-center text-xs text-slate-500 space-y-2">
                      <RefreshCw className="w-5 h-5 animate-spin mx-auto text-emerald-600" />
                      <span>กำลังโหลดตารางจาก Google Drive...</span>
                    </div>
                  ) : driveFiles.length === 0 ? (
                    <div className="p-6 text-center bg-slate-50 rounded-2xl border border-slate-200 text-xs text-slate-500 space-y-2">
                      <p>ไม่พบไฟล์ Google Sheet ใน Google Drive ของคุณ</p>
                      <button
                        type="button"
                        onClick={handleCreateNewSpreadsheet}
                        className="text-emerald-700 font-bold hover:underline"
                      >
                        กดที่นี่เพื่อสร้างตารางใหม่ทันที
                      </button>
                    </div>
                  ) : (
                    <div className="max-h-60 overflow-y-auto space-y-2 pr-1">
                      {driveFiles.map((f) => (
                        <div
                          key={f.id}
                          className="flex items-center justify-between p-3 rounded-xl border border-slate-200 hover:border-emerald-400 bg-white hover:bg-emerald-50/40 transition-all text-xs"
                        >
                          <div className="truncate mr-2">
                            <span className="font-bold text-slate-800 truncate block">
                              {f.name}
                            </span>
                            <span className="text-[11px] text-slate-400 font-mono">
                              แก้ไขล่าสุด:{' '}
                              {f.modifiedTime
                                ? new Date(f.modifiedTime).toLocaleDateString('th-TH')
                                : '-'}
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => {
                              onSetConnectedSheetId(f.id, f.name);
                              showToast(`เชื่อมต่อตาราง "${f.name}" สำเร็จ!`, 'success');
                            }}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3 py-1.5 rounded-lg text-xs shrink-0 cursor-pointer shadow-xs"
                          >
                            เลือกตารางนี้
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* Tab 3: Custom URL / ID */}
              {activeTab === 'custom' && (
                <div className="space-y-3 bg-slate-50 p-4 rounded-2xl border border-slate-200">
                  <label className="block text-xs font-bold text-slate-700">
                    วางลิงก์ Google Sheet URL หรือระบุ Spreadsheet ID
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={customSheetInput}
                      onChange={(e) => setCustomSheetInput(e.target.value)}
                      placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKv.../edit"
                      className="flex-1 bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-mono focus:outline-emerald-600"
                    />
                    <button
                      type="button"
                      onClick={handleConnectCustomSheet}
                      className="bg-slate-900 hover:bg-slate-800 text-white font-bold px-4 py-2 rounded-xl text-xs shrink-0 cursor-pointer transition-colors"
                    >
                      เชื่อมต่อ
                    </button>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    💡 คุณสามารถคัดลอก URL หน้าเว็บ Google Sheet ของคุณจากแถบ Address Bar ของเบราว์เซอร์มาวางได้โดยตรง
                  </p>
                </div>
              )}
            </div>
          )}

          {/* Operation Indicator */}
          {isOperating && (
            <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-200 flex items-center gap-3 animate-pulse">
              <RefreshCw className="w-5 h-5 text-emerald-600 animate-spin" />
              <span className="text-xs font-bold text-emerald-900">
                {operationProgress || 'กำลังดำเนินการ...'}
              </span>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <span>ใบงานในระบบปัจจุบัน: {tickets.length} รายการ</span>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl font-bold transition-colors cursor-pointer"
          >
            ปิดหน้าต่าง
          </button>
        </div>
      </div>

      {/* Confirmation Dialog 1: Export / Overwrite Data (MANDATORY per Workspace Skill) */}
      {confirmSyncOpen && (
        <div className="fixed inset-0 z-60 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
          <div className="w-full max-w-md bg-white rounded-3xl p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center gap-3 text-amber-600">
              <div className="w-10 h-10 rounded-2xl bg-amber-100 flex items-center justify-center">
                <AlertCircle className="w-6 h-6" />
              </div>
              <div>
                <h4 className="font-bold text-slate-900 text-sm sm:text-base">
                  ยืนยันการบันทึกข้อมูลไปยัง Google Sheet
                </h4>
                <p className="text-xs text-slate-500">การเขียนทับข้อมูลใน Google Sheets</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 leading-relaxed">
              การดำเนินการนี้จะอัปเดตและเขียนข้อมูลใบงานจำนวน{' '}
              <strong className="text-slate-900 font-bold">{tickets.length} รายการ</strong> ลงในแท็บ
              <code className="bg-slate-100 px-1 py-0.5 rounded text-slate-800 font-bold mx-1">
                Tickets
              </code>{' '}
              ของ Google Sheet:
              <br />
              <strong className="block mt-1 font-semibold text-slate-800">
                {connectedSheetName || connectedSheetId}
              </strong>
              คุณต้องการดำเนินการต่อหรือไม่?
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setConfirmSyncOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
              >
                ยกเลิก (Cancel)
              </button>
              <button
                type="button"
                onClick={handleExecuteSync}
                className="px-4 py-2 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs transition-colors cursor-pointer"
              >
                ยืนยันการบันทึก (Confirm)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Dialog 2: Import Data (MANDATORY per Workspace Skill) */}
      {confirmImportOpen && (
        <div className="fixed inset-0 z-60 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
          <div className="w-full max-w-md bg-white rounded-3xl p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center gap-3 text-teal-600">
              <div className="w-10 h-10 rounded-2xl bg-teal-100 flex items-center justify-center">
                <DownloadCloud className="w-6 h-6" />
              </div>
              <div>
                <h4 className="font-bold text-slate-900 text-sm sm:text-base">
                  ยืนยันการดึงข้อมูลจาก Google Sheet
                </h4>
                <p className="text-xs text-slate-500">นำเข้าข้อมูลใบงานจาก Google Drive</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 leading-relaxed">
              การดำเนินการนี้จะดึงข้อมูลใบงานทั้งหมดจากแท็บ{' '}
              <code className="bg-slate-100 px-1 py-0.5 rounded text-slate-800 font-bold">
                Tickets
              </code>{' '}
              ในตาราง Google Sheet มาอัปเดตและแสดงในระบบ คุณต้องการดำเนินการต่อหรือไม่?
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setConfirmImportOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
              >
                ยกเลิก (Cancel)
              </button>
              <button
                type="button"
                onClick={handleExecuteImport}
                className="px-4 py-2 rounded-xl text-xs font-bold bg-teal-600 hover:bg-teal-700 text-white shadow-xs transition-colors cursor-pointer"
              >
                ยืนยันการดึงข้อมูล (Confirm Import)
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
