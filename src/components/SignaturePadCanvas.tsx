import React, { useRef, useState, useEffect, useCallback } from 'react';
import { Eraser, Check, PenTool, RotateCcw } from 'lucide-react';

export interface SignaturePadCanvasProps {
  label: string;
  sublabel?: string;
  signerNameLabel?: string;
  signerNameValue?: string;
  onSignerNameChange?: (name: string) => void;
  signerNamePlaceholder?: string;
  initialSignature?: string;
  signedAt?: string;
  onSaveSignature: (signatureDataUrl: string, signedAtTimestamp: string) => void;
  onClearSignature?: () => void;
  strokeColor?: string;
  compact?: boolean;
}

export const SignaturePadCanvas: React.FC<SignaturePadCanvasProps> = ({
  label,
  sublabel,
  signerNameLabel,
  signerNameValue,
  onSignerNameChange,
  signerNamePlaceholder = 'ระบุชื่อ-นามสกุล ผู้ลงนาม...',
  initialSignature,
  signedAt,
  onSaveSignature,
  onClearSignature,
  strokeColor = '#0f172a',
  compact = false
}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const isDrawingRef = useRef(false);
  const lastPointRef = useRef<{ x: number; y: number } | null>(null);

  const [hasDrawnStrokes, setHasDrawnStrokes] = useState(false);
  const [isEditingExisting, setIsEditingExisting] = useState(!initialSignature);

  useEffect(() => {
    setIsEditingExisting(!initialSignature);
  }, [initialSignature]);

  const setupCanvasContext = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return null;
    const rect = canvas.getBoundingClientRect();
    const dpr = Math.min(window.devicePixelRatio || 1, 2);

    const targetWidth = Math.max(280, Math.floor(rect.width * dpr));
    const targetHeight = Math.max(120, Math.floor(rect.height * dpr));

    if (canvas.width !== targetWidth || canvas.height !== targetHeight) {
      canvas.width = targetWidth;
      canvas.height = targetHeight;
    }

    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = strokeColor;
    ctx.lineWidth = 2.4 * dpr;
    return ctx;
  }, [strokeColor]);

  useEffect(() => {
    if (isEditingExisting) {
      setupCanvasContext();
    }
  }, [isEditingExisting, setupCanvasContext]);

  const getCanvasCoordinates = (
    e: React.PointerEvent<HTMLCanvasElement> | TouchEvent | MouseEvent
  ): { x: number; y: number } | null => {
    const canvas = canvasRef.current;
    if (!canvas) return null;
    const rect = canvas.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return null;

    let clientX = 0;
    let clientY = 0;

    if ('touches' in e) {
      if (e.touches.length === 0) return null;
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }

    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY
    };
  };

  const exportCompactSignatureDataUrl = (): string => {
    const canvas = canvasRef.current;
    if (!canvas) return '';

    // Export into a clean, lightweight 400x140 PNG so it renders crisply in A4 PDF & fits Google Sheets cells
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = 400;
    exportCanvas.height = 140;
    const exportCtx = exportCanvas.getContext('2d');
    if (!exportCtx) return canvas.toDataURL('image/png');

    exportCtx.clearRect(0, 0, exportCanvas.width, exportCanvas.height);
    exportCtx.drawImage(canvas, 0, 0, exportCanvas.width, exportCanvas.height);
    return exportCanvas.toDataURL('image/png');
  };

  const commitCurrentCanvasSignature = () => {
    if (!hasDrawnStrokes) return;
    const dataUrl = exportCompactSignatureDataUrl();
    if (!dataUrl) return;
    const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 16);
    onSaveSignature(dataUrl, nowStr);
  };

  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    e.stopPropagation();
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch {
      // Ignore if pointer capture is not supported
    }

    const ctx = setupCanvasContext();
    const pt = getCanvasCoordinates(e);
    if (!ctx || !pt) return;

    isDrawingRef.current = true;
    lastPointRef.current = pt;
    setHasDrawnStrokes(true);

    ctx.beginPath();
    ctx.arc(pt.x, pt.y, ctx.lineWidth / 2, 0, Math.PI * 2);
    ctx.fillStyle = strokeColor;
    ctx.fill();
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!isDrawingRef.current) return;
    e.preventDefault();
    e.stopPropagation();

    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    const pt = getCanvasCoordinates(e);
    const lastPt = lastPointRef.current;
    if (!ctx || !pt || !lastPt) return;

    ctx.beginPath();
    ctx.moveTo(lastPt.x, lastPt.y);
    // Smooth quadratic midpoint
    const midX = (lastPt.x + pt.x) / 2;
    const midY = (lastPt.y + pt.y) / 2;
    ctx.quadraticCurveTo(lastPt.x, lastPt.y, midX, midY);
    ctx.lineTo(pt.x, pt.y);
    ctx.stroke();

    lastPointRef.current = pt;
  };

  const handlePointerUpOrCancel = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!isDrawingRef.current) return;
    e.preventDefault();
    e.stopPropagation();
    isDrawingRef.current = false;
    lastPointRef.current = null;
    try {
      e.currentTarget.releasePointerCapture(e.pointerId);
    } catch {
      // Ignore
    }
    // Auto-save stroke to parent state so users on mobile don't lose their signature if they forget to tap a secondary button
    const dataUrl = exportCompactSignatureDataUrl();
    if (dataUrl) {
      const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 16);
      onSaveSignature(dataUrl, nowStr);
    }
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
    }
    setHasDrawnStrokes(false);
    setIsEditingExisting(true);
    if (onClearSignature) {
      onClearSignature();
    }
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-3.5 space-y-2.5 text-xs">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <p className="font-bold text-slate-900 flex items-center gap-1.5">
            <PenTool className="w-3.5 h-3.5 text-orange-600 shrink-0" />
            <span>{label}</span>
          </p>
          {sublabel && <p className="text-[11px] text-slate-500 mt-0.5">{sublabel}</p>}
        </div>
        {signedAt && (
          <span className="inline-flex items-center gap-1 bg-emerald-50 border border-emerald-300 text-emerald-800 px-2.5 py-0.5 rounded-full text-[10px] font-bold font-mono">
            ✓ เซ็นแล้วเมื่อ {signedAt}
          </span>
        )}
      </div>

      {signerNameLabel && onSignerNameChange && (
        <div>
          <label className="block text-[11px] font-semibold text-slate-700 mb-1">
            {signerNameLabel}
          </label>
          <input
            type="text"
            value={signerNameValue || ''}
            onChange={(e) => onSignerNameChange(e.target.value)}
            placeholder={signerNamePlaceholder}
            className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-orange-500"
          />
        </div>
      )}

      {!isEditingExisting && initialSignature ? (
        <div className="space-y-2">
          <div className="relative bg-slate-50 border-2 border-emerald-400 rounded-xl p-3 flex flex-col items-center justify-center min-h-[120px]">
            <img
              src={initialSignature}
              alt="Captured Signature"
              className="max-h-24 w-auto object-contain select-none"
            />
            <div className="w-4/5 border-b border-dashed border-slate-300 mt-1" />
            <span className="text-[10px] text-emerald-700 font-semibold mt-1">
              ลายมือชื่อดิจิทัลที่บันทึกไว้ (พร้อมแสดงในใบรายงานและ PDF)
            </span>
          </div>
          <div className="flex flex-wrap items-center justify-end gap-2">
            <button
              type="button"
              onClick={() => {
                setIsEditingExisting(true);
                setHasDrawnStrokes(false);
                if (onClearSignature) onClearSignature();
              }}
              className="bg-slate-100 hover:bg-slate-200 text-slate-800 px-3 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>เซ็นใหม่ / แก้ไขลายเซ็น</span>
            </button>
          </div>
        </div>
      ) : (
        <div className="space-y-2">
          <div className="relative rounded-xl border-2 border-dashed border-slate-300 bg-slate-50/70 hover:border-orange-400 transition-colors overflow-hidden">
            <canvas
              ref={canvasRef}
              onPointerDown={handlePointerDown}
              onPointerMove={handlePointerMove}
              onPointerUp={handlePointerUpOrCancel}
              onPointerCancel={handlePointerUpOrCancel}
              style={{ touchAction: 'none' }}
              className={`w-full ${
                compact ? 'h-28' : 'h-36 sm:h-40'
              } cursor-crosshair block touch-none select-none`}
            />
            {/* Signature Guide Line & Watermark */}
            <div className="pointer-events-none absolute inset-x-6 bottom-5 border-b border-slate-300/90 flex justify-between items-end pb-0.5">
              <span className="text-[10px] font-mono text-slate-400">SIGN HERE ✍️</span>
              <span className="text-[10px] text-slate-400">
                ใช้นิ้วหรือปากกาเซ็นชื่อบนหน้าจอได้เลย
              </span>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-[11px] text-slate-500">
              * รองรับการใช้นิ้วเขียนบนจอมือถือ / แท็บเล็ต และเมาส์ (บันทึกอัตโนมัติทันทีที่เซ็น)
            </span>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={clearCanvas}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1 cursor-pointer"
              >
                <Eraser className="w-3.5 h-3.5" />
                <span>ล้างลายเซ็น</span>
              </button>
              {hasDrawnStrokes && (
                <button
                  type="button"
                  onClick={() => {
                    commitCurrentCanvasSignature();
                    setIsEditingExisting(false);
                  }}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1 cursor-pointer"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>ล็อกลายเซ็น</span>
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
