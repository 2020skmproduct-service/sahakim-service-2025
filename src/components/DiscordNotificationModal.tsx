import React, { useState, useEffect } from 'react';
import {
  Bell,
  CheckCircle2,
  AlertTriangle,
  ExternalLink,
  Send,
  Check,
  Copy,
  Info,
  X,
  Sparkles,
  Layers,
  ArrowRight
} from 'lucide-react';
import { sendDiscordNotification } from '../services/discordNotify';
import { ServiceTicket } from '../types';

interface DiscordNotificationModalProps {
  isOpen: boolean;
  onClose: () => void;
  discordWebhookUrl: string;
  onSaveWebhookUrl: (url: string) => void;
  sampleTicket?: ServiceTicket;
  showToast: (msg: string, type?: 'success' | 'warning' | 'error' | 'info') => void;
}

export const DiscordNotificationModal: React.FC<DiscordNotificationModalProps> = ({
  isOpen,
  onClose,
  discordWebhookUrl,
  onSaveWebhookUrl,
  sampleTicket,
  showToast
}) => {
  const [inputUrl, setInputUrl] = useState(discordWebhookUrl || '');
  const [isTesting, setIsTesting] = useState(false);
  const [isCopied, setIsCopied] = useState(false);
  const [testResult, setTestResult] = useState<{
    ok: boolean;
    text: string;
    time: string;
  } | null>(null);

  useEffect(() => {
    if (isOpen) {
      setInputUrl(discordWebhookUrl || localStorage.getItem('sahakim_discord_webhook_url') || '');
      setTestResult(null);
    }
  }, [isOpen, discordWebhookUrl]);

  if (!isOpen) return null;

  const handleSave = () => {
    const trimmed = inputUrl.trim();
    if (trimmed && !trimmed.startsWith('https://discord.com/api/webhooks/')) {
      showToast('กรุณากรอก Discord Webhook URL ให้ถูกต้อง (ขึ้นต้นด้วย https://discord.com/api/webhooks/...)', 'warning');
      return;
    }
    onSaveWebhookUrl(trimmed);
    showToast('บันทึกการตั้งค่า Discord Webhook เรียบร้อยแล้ว', 'success');
  };

  const handleTestSend = async () => {
    const trimmed = inputUrl.trim();
    if (!trimmed) {
      showToast('กรุณากรอก Discord Webhook URL ก่อนกดทดสอบ', 'warning');
      return;
    }
    if (!trimmed.startsWith('https://discord.com/api/webhooks/')) {
      showToast('URL ไม่ถูกต้อง ต้องขึ้นต้นด้วย https://discord.com/api/webhooks/...', 'warning');
      return;
    }

    setIsTesting(true);
    setTestResult(null);
    try {
      const nowStr = new Date().toLocaleTimeString('th-TH');
      let result;

      if (sampleTicket) {
        result = await sendDiscordNotification(trimmed, {
          ticket: sampleTicket,
          eventHeading: '🧪 [ทดสอบระบบ] ใบงานแจ้งซ่อมใหม่ - บริษัท สหคิมมอเตอร์ จำกัด'
        });
      } else {
        result = await sendDiscordNotification(trimmed, {
          message: `🔔 **[ทดสอบการเชื่อมต่อ Discord Webhook สำเร็จ]**\nระบบบริหารงานบริการหลังการขายและงานซ่อม (Sahakim Service Portal V12.1)\n• เวลาทดสอบ: ${nowStr}\n• พร้อมรับการแจ้งเตือนใบงานใหม่และการเปลี่ยนสถานะแบบเรียลไทม์ทันที!`
        });
      }

      if (result.ok) {
        setTestResult({
          ok: true,
          text: 'ส่งข้อความทดสอบเข้า Discord สำเร็จ! ตรวจสอบห้องแชทใน Discord ได้เลย',
          time: nowStr
        });
        showToast('ส่งข้อความทดสอบเข้า Discord สำเร็จ!', 'success');
        onSaveWebhookUrl(trimmed);
      } else {
        setTestResult({
          ok: false,
          text: `ส่งข้อความไม่สำเร็จ: ${result.message || 'โปรดตรวจสอบความถูกต้องของ URL'}`,
          time: nowStr
        });
        showToast(`ส่งไม่สำเร็จ: ${result.message || 'ตรวจสอบ URL'}`, 'error');
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'เกิดข้อผิดพลาดในการเชื่อมต่อ';
      setTestResult({
        ok: false,
        text: msg,
        time: new Date().toLocaleTimeString('th-TH')
      });
      showToast(msg, 'error');
    } finally {
      setIsTesting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[999] flex items-center justify-center p-3 sm:p-4 bg-slate-950/70 backdrop-blur-sm animate-fadeIn">
      <div className="relative w-full max-w-2xl bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-800 max-h-[92vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="px-5 py-4 bg-gradient-to-r from-[#5865F2] to-[#4752C4] text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-white/20 backdrop-blur flex items-center justify-center shadow-inner">
              <Bell className="w-5 h-5 text-white animate-bounce" />
            </div>
            <div>
              <h3 className="text-lg font-bold flex items-center gap-2">
                ตั้งค่าการแจ้งเตือน Discord Webhook
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-emerald-400 text-slate-900 font-semibold tracking-wide uppercase">
                  ฟรี ไม่จำกัดจำนวน
                </span>
              </h3>
              <p className="text-xs text-indigo-100">
                แจ้งเตือนงานซ่อมและสถานะใบงานแบบเรียลไทม์ แทน LINE ที่ติดโควต้า
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-white/80 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Content */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Advantage Banner */}
          <div className="p-3.5 rounded-xl bg-indigo-50 dark:bg-indigo-950/40 border border-indigo-200 dark:border-indigo-800 text-xs text-indigo-900 dark:text-indigo-200 flex items-start gap-3">
            <Sparkles className="w-5 h-5 text-[#5865F2] shrink-0 mt-0.5" />
            <div>
              <div className="font-bold text-sm mb-1 text-[#5865F2] dark:text-indigo-300">
                ทำไม Discord Webhook ถึงสะดวกและดีกว่า?
              </div>
              <ul className="list-disc list-inside space-y-0.5 text-slate-600 dark:text-slate-300">
                <li><strong className="text-slate-900 dark:text-white">ฟรี 100% ไม่มีลิมิต:</strong> ส่งแจ้งเตือนได้ไม่จำกัด (LINE บัญชีฟรีจำกัด 300 ข้อความ/เดือน)</li>
                <li><strong className="text-slate-900 dark:text-white">ติดตั้งง่ายที่สุด:</strong> ไม่ต้องสร้าง Bot Token หรือแอดเพื่อน แค่ก๊อปปี้ Webhook URL มาวาง</li>
                <li><strong className="text-slate-900 dark:text-white">รองรับ GitHub Pages ทันที:</strong> กดบันทึกงานซ่อมแล้วแจ้งเตือนเด้งเข้าห้องแชทของช่างทันที</li>
              </ul>
            </div>
          </div>

          {/* 3 Simple Setup Steps */}
          <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/80 space-y-3">
            <div className="font-bold text-slate-800 dark:text-slate-200 text-sm flex items-center gap-2">
              <Layers className="w-4 h-4 text-[#5865F2]" />
              วิธีสร้าง Discord Webhook ใน 3 ขั้นตอน (ใช้เวลาไม่ถึง 1 นาที):
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 text-xs">
              <div className="p-2.5 rounded-lg bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700">
                <div className="font-bold text-[#5865F2] mb-1 flex items-center gap-1">
                  <span className="w-4 h-4 rounded-full bg-[#5865F2] text-white flex items-center justify-center text-[10px]">1</span>
                  เข้าห้องแชท
                </div>
                <p className="text-slate-600 dark:text-slate-300">
                  เปิด Discord ไปที่ Server ของทีม แล้วกดรูปเฟือง ⚙️ ข้างชื่อห้องแชทที่ต้องการรับแจ้งเตือน
                </p>
              </div>

              <div className="p-2.5 rounded-lg bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700">
                <div className="font-bold text-[#5865F2] mb-1 flex items-center gap-1">
                  <span className="w-4 h-4 rounded-full bg-[#5865F2] text-white flex items-center justify-center text-[10px]">2</span>
                  สร้าง Webhook
                </div>
                <p className="text-slate-600 dark:text-slate-300">
                  คลิกแท็บ <strong>Integrations (การรวมระบบ)</strong> &gt; <strong>Webhooks</strong> &gt; <strong>New Webhook</strong>
                </p>
              </div>

              <div className="p-2.5 rounded-lg bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700">
                <div className="font-bold text-[#5865F2] mb-1 flex items-center gap-1">
                  <span className="w-4 h-4 rounded-full bg-[#5865F2] text-white flex items-center justify-center text-[10px]">3</span>
                  คัดลอก URL
                </div>
                <p className="text-slate-600 dark:text-slate-300">
                  กดปุ่ม <strong>Copy Webhook URL</strong> แล้วนำลิงก์มาวางในช่องด้านล่างนี้
                </p>
              </div>
            </div>
          </div>

          {/* URL Input Form */}
          <div className="space-y-2">
            <label className="block text-xs font-bold text-slate-700 dark:text-slate-300">
              วาง Discord Webhook URL:
            </label>
            <div className="relative">
              <input
                type="url"
                value={inputUrl}
                onChange={(e) => setInputUrl(e.target.value)}
                placeholder="https://discord.com/api/webhooks/1234567890/abcdefg..."
                className="w-full px-3.5 py-2.5 text-xs rounded-xl border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 font-mono focus:ring-2 focus:ring-[#5865F2] focus:border-transparent outline-none pr-10"
              />
              {inputUrl && (
                <button
                  type="button"
                  onClick={() => setInputUrl('')}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
            <p className="text-[11px] text-slate-500 dark:text-slate-400 flex items-center gap-1">
              <Info className="w-3.5 h-3.5" />
              URL นี้จะถูกบันทึกไว้ในเบราว์เซอร์และเซิร์ฟเวอร์อัตโนมัติ สำหรับส่งแจ้งเตือนทุกครั้งที่บันทึกงาน
            </p>
          </div>

          {/* Test Status Banner */}
          {testResult && (
            <div
              className={`p-3 rounded-xl border text-xs flex items-start gap-2.5 ${
                testResult.ok
                  ? 'bg-emerald-50 dark:bg-emerald-950/40 border-emerald-200 dark:border-emerald-800 text-emerald-800 dark:text-emerald-200'
                  : 'bg-rose-50 dark:bg-rose-950/40 border-rose-200 dark:border-rose-800 text-rose-800 dark:text-rose-200'
              }`}
            >
              {testResult.ok ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
              ) : (
                <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
              )}
              <div className="flex-1">
                <div className="font-semibold">{testResult.text}</div>
                <div className="text-[10px] opacity-75 mt-0.5">เวลา: {testResult.time}</div>
              </div>
            </div>
          )}

          {/* Ticket Preview Example */}
          <div className="p-3.5 rounded-xl bg-slate-900 text-slate-200 border border-slate-800 text-xs space-y-2">
            <div className="text-[11px] font-semibold text-slate-400 flex items-center justify-between">
              <span>ตัวอย่างหน้าตาการแจ้งเตือนใน Discord (Rich Embed Card):</span>
              <span className="text-[10px] text-emerald-400">● มีสีแยกตามสถานะ</span>
            </div>
            <div className="border-l-4 border-amber-500 pl-3 py-1 space-y-1 bg-slate-800/50 rounded-r-lg p-2">
              <div className="font-bold text-white text-sm">
                📋 ใบงาน: {sampleTicket?.id || 'SKM-2026-0001'} ({sampleTicket?.docNo || 'SR-2026-01'})
              </div>
              <div className="text-slate-300 text-xs">
                🔧 <span className="font-semibold">รายการ:</span> PUMP ปั๊มน้ำหอยโข่ง (BULL POWER) - อาการ: มอเตอร์ร้อนจัด เสียงดัง
              </div>
              <div className="grid grid-cols-2 gap-x-2 gap-y-0.5 text-[11px] text-slate-300 pt-1">
                <div>📌 สถานะ: <span className="bg-amber-900/60 text-amber-200 px-1 rounded">REPAIRING</span></div>
                <div>🚨 ความเร่งด่วน: ด่วนมาก</div>
                <div>🏢 ลูกค้า: บริษัท สหคิมมอเตอร์ จำกัด</div>
                <div>📍 หน้างาน: โรงงานนิคมบางชัน</div>
                <div>👷 ช่าง: นพดล แสนสุข</div>
                <div>💰 ยอดอะไหล่: 4,500 THB</div>
              </div>
            </div>
          </div>
        </div>

        {/* Footer Actions */}
        <div className="px-5 py-3.5 bg-slate-50 dark:bg-slate-800/80 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-colors"
          >
            ปิดหน้าต่าง
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              disabled={isTesting || !inputUrl.trim()}
              onClick={handleTestSend}
              className="px-4 py-2 text-xs font-semibold rounded-xl bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center gap-1.5 disabled:opacity-50"
            >
              {isTesting ? (
                <>
                  <div className="w-3.5 h-3.5 border-2 border-[#5865F2] border-t-transparent rounded-full animate-spin" />
                  กำลังทดสอบ...
                </>
              ) : (
                <>
                  <Send className="w-3.5 h-3.5 text-[#5865F2]" />
                  ทดสอบส่งเข้า Discord
                </>
              )}
            </button>

            <button
              type="button"
              onClick={handleSave}
              className="px-5 py-2 text-xs font-bold rounded-xl bg-[#5865F2] hover:bg-[#4752C4] text-white shadow-sm transition-colors flex items-center gap-1.5"
            >
              <Check className="w-4 h-4" />
              บันทึก Webhook
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
