import React, { useState, useEffect } from 'react';
import {
  Bell,
  CheckCircle2,
  AlertTriangle,
  ExternalLink,
  RefreshCw,
  Send,
  Users,
  MessageSquare,
  ShieldCheck,
  Check,
  Copy,
  Info,
  X,
  Smartphone,
  QrCode
} from 'lucide-react';
import { ServiceTicket } from '../types';

interface LineNotificationModalProps {
  isOpen: boolean;
  onClose: () => void;
  tickets: ServiceTicket[];
  lineTargetId: string;
  onSetLineTargetId: (id: string) => void;
  showToast: (msg: string, type?: 'success' | 'warning' | 'error' | 'info') => void;
}

export const LineNotificationModal: React.FC<LineNotificationModalProps> = ({
  isOpen,
  onClose,
  tickets,
  lineTargetId,
  onSetLineTargetId,
  showToast
}) => {
  const [botConfig, setBotConfig] = useState<{
    status: string;
    botBasicId: string;
    hasValidToken: boolean;
    targetId: string;
    groupIds: string[];
    userIds: string[];
    updatedAt: string | null;
  } | null>(null);

  const [isLoadingConfig, setIsLoadingConfig] = useState(false);
  const [isSendingTest, setIsSendingTest] = useState(false);
  const [testMessage, setTestMessage] = useState(
    '🔔 [ทดสอบการแจ้งเตือน LINE] บริษัท สหคิมมอเตอร์ จำกัด\nระบบบริการหลังการขายและจัดการงานซ่อม (Sahakim Service Portal)\nสถานะ: บอทเชื่อมต่อสำเร็จ พร้อมส่งแจ้งเตือนใบงานแบบเรียลไทม์!'
  );
  const [testResult, setTestResult] = useState<{
    ok: boolean;
    text: string;
    time: string;
  } | null>(null);

  const [customTargetInput, setCustomTargetInput] = useState('');

  const fetchConfig = async () => {
    setIsLoadingConfig(true);
    try {
      const res = await fetch('/api/line/config');
      if (res.ok) {
        const data = await res.json();
        setBotConfig(data);
        if (data.targetId && !lineTargetId) {
          onSetLineTargetId(data.targetId);
        }
      }
    } catch (err) {
      console.warn('Failed to fetch LINE config:', err);
    } finally {
      setIsLoadingConfig(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      fetchConfig();
    }
  }, [isOpen]);

  const handleSendTestMessage = async () => {
    if (!testMessage.trim()) {
      showToast('กรุณาระบุข้อความที่ต้องการทดสอบส่ง', 'warning');
      return;
    }
    setIsSendingTest(true);
    setTestResult(null);
    try {
      const res = await fetch('/api/line/notify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: testMessage.trim(),
          targetId: customTargetInput.trim() || lineTargetId.trim()
        })
      });

      const data = await res.json();
      const nowStr = new Date().toLocaleTimeString('th-TH');

      if (data?.quotaExceeded || data?.status === 'quota_exceeded') {
        const msg = data.message || 'โควต้าส่งข้อความ LINE รายเดือนเต็มแล้ว (429)';
        setTestResult({ ok: false, text: msg, time: nowStr });
        showToast(msg, 'warning');
      } else if (data?.status === 'ok') {
        const pushedTarget = data.pushedTo?.join(', ') || 'กลุ่ม LINE';
        const msg = `ส่งข้อความทดสอบเข้า LINE สำเร็จเรียบร้อย (${pushedTarget})`;
        setTestResult({ ok: true, text: msg, time: nowStr });
        showToast(msg, 'success');
        fetchConfig();
      } else {
        const errMsg = data?.error || data?.message || 'ส่งข้อความไม่สำเร็จ';
        setTestResult({ ok: false, text: errMsg, time: nowStr });
        showToast(`ส่งไม่สำเร็จ: ${errMsg}`, 'error');
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์';
      setTestResult({ ok: false, text: msg, time: new Date().toLocaleTimeString('th-TH') });
      showToast(msg, 'error');
    } finally {
      setIsSendingTest(false);
    }
  };

  const handleSaveCustomTarget = () => {
    if (!customTargetInput.trim()) {
      showToast('กรุณาระบุ Group ID หรือ User ID', 'warning');
      return;
    }
    onSetLineTargetId(customTargetInput.trim());
    showToast(`บันทึกเป้าหมายการแจ้งเตือน: ${customTargetInput.trim()}`, 'success');
  };

  if (!isOpen) return null;

  const currentActiveTarget = customTargetInput.trim() || lineTargetId || botConfig?.targetId || '';

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/70 backdrop-blur-xs overflow-y-auto">
      <div className="relative w-full max-w-2xl bg-white rounded-3xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="px-6 py-5 bg-gradient-to-r from-emerald-600 via-green-600 to-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-white/10 backdrop-blur-xs flex items-center justify-center border border-white/20 shadow-inner">
              <MessageSquare className="w-5 h-5 text-emerald-300" />
            </div>
            <div>
              <h3 className="font-bold text-base sm:text-lg flex items-center gap-2">
                ระบบแจ้งเตือนผ่าน LINE อัตโนมัติ
                <span className="text-[10px] bg-emerald-400 text-slate-950 font-extrabold px-2 py-0.5 rounded-full">
                  @389fyhbl
                </span>
              </h3>
              <p className="text-xs text-emerald-100">
                แจ้งเตือนงานใหม่ มอบหมายช่าง ขออนุมัติอะไหล่ และปิดจบงานเข้ากลุ่ม LINE แบบเรียลไทม์
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-white/10 hover:bg-white/20 flex items-center justify-center text-white transition-colors cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="p-5 sm:p-6 space-y-5 overflow-y-auto">
          {/* Bot Card */}
          <div className="bg-gradient-to-r from-emerald-50 via-green-50 to-slate-50 border border-emerald-200 rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3.5">
              <div className="w-12 h-12 rounded-2xl bg-[#06C755] text-white flex items-center justify-center font-bold text-xl shadow-md shrink-0">
                LINE
              </div>
              <div className="space-y-0.5">
                <div className="flex items-center gap-2">
                  <h4 className="font-bold text-slate-900 text-sm sm:text-base">
                    Sahakim.service (บอททางการ สหคิมมอเตอร์)
                  </h4>
                  <span className="flex items-center gap-1 text-[11px] bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded-full">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    ออนไลน์พร้อมใช้งาน
                  </span>
                </div>
                <p className="text-xs text-slate-600 flex items-center gap-2">
                  LINE ID: <strong className="text-slate-900 font-mono">@389fyhbl</strong>
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 w-full sm:w-auto">
              <a
                href="https://line.me/R/ti/p/@389fyhbl"
                target="_blank"
                rel="noreferrer"
                className="w-full sm:w-auto inline-flex items-center justify-center gap-1.5 bg-[#06C755] hover:bg-[#05b34c] text-white font-bold px-4 py-2 rounded-xl text-xs shadow-sm transition-colors cursor-pointer"
              >
                <Smartphone className="w-4 h-4" />
                <span>เพิ่มเพื่อน / เชิญเข้ากลุ่ม ↗</span>
              </a>
            </div>
          </div>

          {/* 3 Easy Steps to Connect */}
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 sm:p-5 space-y-3">
            <h4 className="font-bold text-slate-800 text-xs sm:text-sm flex items-center gap-2">
              <Info className="w-4 h-4 text-emerald-600" />
              วิธีตั้งค่าให้ระบบแจ้งเตือนเข้ากลุ่ม LINE ของคุณ (3 ขั้นตอนง่ายๆ):
            </h4>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs text-slate-700">
              <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-1 shadow-xs">
                <span className="w-5 h-5 rounded-full bg-emerald-600 text-white font-bold text-[11px] inline-flex items-center justify-center mb-1">
                  1
                </span>
                <p className="font-bold text-slate-900">เพิ่มเพื่อนหรือเชิญเข้ากลุ่ม</p>
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  ค้นหาไอดี <code className="font-mono bg-slate-100 px-1 py-0.5 rounded font-bold text-emerald-800">@389fyhbl</code> หรือกดปุ่มด้านบน แล้วเชิญบอทเข้ากลุ่ม LINE ทีมงานช่าง
                </p>
              </div>

              <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-1 shadow-xs">
                <span className="w-5 h-5 rounded-full bg-emerald-600 text-white font-bold text-[11px] inline-flex items-center justify-center mb-1">
                  2
                </span>
                <p className="font-bold text-slate-900">พิมพ์ทักทายในกลุ่ม</p>
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  พิมพ์คำว่า <code className="bg-slate-100 px-1 rounded font-bold">สวัสดี</code> หรือ <code className="bg-slate-100 px-1 rounded font-bold">/id</code> ในกลุ่ม บอทจะตอบกลับและบันทึก Group ID อัตโนมัติทันที
                </p>
              </div>

              <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-1 shadow-xs">
                <span className="w-5 h-5 rounded-full bg-emerald-600 text-white font-bold text-[11px] inline-flex items-center justify-center mb-1">
                  3
                </span>
                <p className="font-bold text-slate-900">พร้อมรับแจ้งเตือนทันที</p>
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  เมื่อมีการแจ้งซ่อมใหม่ มอบหมายงาน หรือเปลี่ยนสถานะ ระบบจะส่งแจ้งเตือนเข้ากลุ่มทันทีอัตโนมัติ
                </p>
              </div>
            </div>
          </div>

          {/* Current Target Status */}
          <div className="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="space-y-0.5">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                  สถานะกลุ่ม / ผู้รับการแจ้งเตือน
                </span>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-slate-900 text-sm">
                    {currentActiveTarget ? (
                      <span className="text-emerald-700 flex items-center gap-1.5 font-mono">
                        <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                        {currentActiveTarget.startsWith('C')
                          ? `กลุ่ม LINE (Group ID: ${currentActiveTarget.substring(0, 10)}...)`
                          : `ผู้ใช้ LINE (ID: ${currentActiveTarget.substring(0, 10)}...)`}
                      </span>
                    ) : (
                      <span className="text-amber-600 flex items-center gap-1.5">
                        <AlertTriangle className="w-4 h-4" />
                        ยังไม่มีกลุ่มผูกไว้ (ระบบจะบรอดแคสต์หรือรอจับคู่กลุ่มเมื่อเชิญบอท)
                      </span>
                    )}
                  </span>
                </div>
              </div>

              <button
                type="button"
                onClick={fetchConfig}
                disabled={isLoadingConfig}
                className="text-xs text-slate-600 hover:text-slate-900 flex items-center gap-1.5 border border-slate-200 px-3 py-1.5 rounded-xl hover:bg-slate-50 transition-colors cursor-pointer"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isLoadingConfig ? 'animate-spin' : ''}`} />
                <span>รีเฟรชข้อมูล</span>
              </button>
            </div>

            {botConfig?.groupIds && botConfig.groupIds.length > 0 && (
              <div className="pt-2 border-t border-slate-100">
                <p className="text-xs text-slate-600 font-semibold mb-1.5">
                  กลุ่ม LINE ที่บอทเคยตรวจพบ ({botConfig.groupIds.length} กลุ่ม):
                </p>
                <div className="flex flex-wrap gap-2">
                  {botConfig.groupIds.map((gid, idx) => (
                    <button
                      key={gid}
                      type="button"
                      onClick={() => {
                        onSetLineTargetId(gid);
                        setCustomTargetInput(gid);
                        showToast(`เลือกกลุ่ม ${gid.substring(0, 10)}... สำเร็จ`, 'success');
                      }}
                      className={`text-xs px-2.5 py-1 rounded-lg border font-mono transition-colors cursor-pointer ${
                        currentActiveTarget === gid
                          ? 'bg-emerald-100 border-emerald-400 text-emerald-900 font-bold'
                          : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      กลุ่ม {idx + 1}: {gid.substring(0, 12)}...
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Test Notification Section */}
          <div className="bg-emerald-50/60 border border-emerald-200 rounded-2xl p-4 sm:p-5 space-y-3">
            <div className="flex items-center justify-between">
              <label className="block text-xs font-bold text-slate-900">
                🚀 ทดสอบส่งข้อความแจ้งเตือนเข้า LINE ตอนนี้:
              </label>
              {testResult && (
                <span
                  className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full ${
                    testResult.ok ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                  }`}
                >
                  {testResult.ok ? '✅ ' : '⚠️ '} {testResult.text} ({testResult.time})
                </span>
              )}
            </div>

            <textarea
              rows={3}
              value={testMessage}
              onChange={(e) => setTestMessage(e.target.value)}
              className="w-full bg-white border border-slate-300 rounded-xl p-2.5 text-xs focus:outline-emerald-600 leading-relaxed font-sans"
              placeholder="ข้อความที่ต้องการทดสอบส่งเข้า LINE..."
            />

            <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
              <p className="text-[11px] text-slate-500">
                💡 ข้อความจะถูกส่งไปยังกลุ่ม LINE หรือห้องแชทที่เชื่อมต่อไว้ทันที
              </p>
              <button
                type="button"
                onClick={handleSendTestMessage}
                disabled={isSendingTest}
                className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold px-5 py-2 rounded-xl text-xs shadow-sm inline-flex items-center gap-1.5 cursor-pointer transition-all"
              >
                {isSendingTest ? (
                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Send className="w-3.5 h-3.5" />
                )}
                <span>{isSendingTest ? 'กำลังส่ง...' : 'ส่งข้อความทดสอบเข้า LINE'}</span>
              </button>
            </div>
          </div>

          {/* Automated Events Summary */}
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 text-xs space-y-2">
            <h5 className="font-bold text-slate-800">
              📌 รายการที่ระบบจะส่งแจ้งเตือนเข้า LINE อัตโนมัติ:
            </h5>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-slate-600">
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อลูกค้าแจ้งซ่อมเปิดงานใหม่ (NEW)
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อมอบหมายช่างและกำหนดวันนัด
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อช่างกดตอบรับงานและเข้าตรวจสอบ
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อส่งขออนุมัติอะไหล่ / รายงานช่าง
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อช่างหรือลูกค้าลงลายเซ็นรับรอง
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-emerald-600 font-bold">✓</span> เมื่อตรวจรับงานและปิดจบใบงาน (CLOSE)
              </div>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <span>ใบงานในระบบ: {tickets.length} รายการ</span>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl font-bold transition-colors cursor-pointer"
          >
            ปิดหน้าต่าง
          </button>
        </div>
      </div>
    </div>
  );
};
