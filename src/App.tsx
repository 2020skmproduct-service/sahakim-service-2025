import React, { useState, useEffect, useMemo } from 'react';
import {
  Wrench,
  Users,
  UserCog,
  Printer,
  FileSpreadsheet,
  Plus,
  Trash2,
  UserMinus,
  CheckCircle2,
  XCircle,
  Lock,
  Unlock,
  KeyRound,
  Camera,
  ClipboardCheck,
  Search,
  ArrowLeft,
  Bell,
  ShieldCheck,
  LogOut,
  Edit3,
  AlertTriangle,
  Volume2,
  VolumeX,
  Cpu,
  BookOpen,
  BookmarkPlus,
  Loader2,
  Check,
  Download,
  Copy,
  Send,
  PieChart as PieChartIcon,
  BarChart3,
  PenTool,
  RotateCcw,
  ImageIcon,
  Upload
} from 'lucide-react';
import {
  PieChart,
  Pie,
  Cell,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer
} from 'recharts';
import {
  UserRole,
  TicketStatus,
  ChecklistItem,
  PartItem,
  EquipmentItem,
  ServiceTicket,
  Technician,
  TroubleshootingTemplate
} from './types';
import { SignaturePadCanvas } from './components/SignaturePadCanvas';
import { GoogleSheetsIntegrationModal } from './components/GoogleSheetsIntegrationModal';
import { getAccessToken } from './services/googleAuth';
import { syncTicketsToSpreadsheet } from './services/googleSheets';
import {
  downloadStandaloneHtml,
  generateStandaloneHtmlString,
  GOOGLE_APPS_SCRIPT_TEMPLATE
} from './standaloneGenerator';
import {
  EMBEDDED_SAHAKIM_LOGO,
  EMBEDDED_SKM_MASCOT
} from './assets/embeddedImages';
import pumpBeforeImg from './assets/images/pump_before_repair_1793148158779.jpg';
import pumpAfterImg from './assets/images/pump_after_repair_1793148158800.jpg';

const SKM_SERVICE_MASCOT_LOGO = EMBEDDED_SKM_MASCOT;
const DEFAULT_SAHAKIM_LOGO = EMBEDDED_SAHAKIM_LOGO;

export const DEFAULT_EQUIPMENT_MASTERS = [
  'MOTOR',
  'END SUCTION PUMP',
  'VERTICAL PUMP',
  'BOOSTER PUMP',
  'TRANSFER PUMP',
  'BLOWER',
  'FIRE PUMP',
  'GEN SET',
  'SEWAGE',
  'EJECTOR',
  'AIR PUMP',
  'AERATOR',
  'CONTROL',
  'ETC. (กรอกเพิ่มเองได้)'
];

export const DEFAULT_BRANDS = [
  'BULL POWER',
  'BROOK CROMPTON',
  'STAC',
  'UNOMACH',
  'VARISCO',
  'FYBROC',
  'EBARA',
  'GRUNDFOS',
  'SIEMENS',
  'ABB',
  'ETC. (กรอกเพิ่มเองได้)'
];

export const CHECKLIST_TEMPLATES: Record<string, string[]> = {
  MOTOR: [
    'Stator Winding Insulation Resistance Test (Megger IEEE 43 / EASA AR100)',
    'Phase-to-Phase Winding Resistance Balance Check',
    'Bearing Housing Vibration & Temperature Measurement (ISO 10816)',
    'Terminal Box, Cable Lug Tightness & Grounding Inspection',
    'No-Load / Full-Load Current & Voltage Unbalance Check',
    'Cooling Fan, Shroud & Shaft Runout (TIR) Check'
  ],
  PUMP: [
    'Pump-Motor Laser / Dial Alignment & Coupling Inspection (Pump Handbook)',
    'Insulation Resistance Test (Megger Motor Winding)',
    'Mechanical Seal / Gland Packing Leakage & Flushing Line Check',
    'Suction & Discharge Pressure Gauge (NPSHa vs NPSHr / Cavitation Check)',
    'Impeller Wear Ring Clearance & Bearing Temperature Check'
  ],
  BLOWER: [
    'V-Belt Tension & Sheave Pulley Alignment Check',
    'Rotary Lobe / Impeller Timing Clearance & Housing Inspection',
    'Inlet Filter Differential Pressure & Safety Relief Valve Test',
    'Gearbox Oil Level & Bearing Greasing Inspection',
    'Motor Running Current & Vibration Check'
  ],
  'GEN SET': [
    'Engine Oil Pressure, Coolant Temperature & Radiator Core Check',
    'Battery Charger Voltage, Specific Gravity & Cranking Voltage Test',
    'Alternator Winding Insulation & AVR Voltage Regulation Check',
    'Fuel Filter, Water Separator & Governor Response Test',
    'ATS Transfer Switch & Auto-Start Simulation Test'
  ],
  SEWAGE: [
    'Submersible Cable Entry Seal & Moisture Sensor Probe Check',
    'Non-Clog Impeller / Cutter Clearance & Wear Plate Inspection',
    'Mechanical Seal Oil Chamber Inspection (Water Emulsion Check)',
    'Motor Winding Insulation Resistance (Megger @ 500/1000 VDC)',
    'Guide Rail, Duckfoot Bend & Float Switch Operation Check'
  ],
  EJECTOR: [
    'Jet Nozzle & Venturi Diffuser Throat Wear Inspection',
    'Motive Pump Discharge Pressure & Suction Vacuum Check',
    'Check Valve (Non-Return Valve) Sealing & Spring Inspection',
    'Motor Current & Insulation Resistance Check'
  ],
  'AIR PUMP': [
    'Diaphragm / Piston Valve & Air Distribution Valve Inspection',
    'Air Filter Regulator & Lubricator (FRL) Pressure Check',
    'Discharge Manifold & Muffler Exhaust Inspection',
    'Operating Cycle Rate & Leakage Test'
  ],
  AERATOR: [
    'Impeller / Propeller Balance & Shaft Sleeve Wear Check',
    'Gearbox / Reducer Oil Seal & Lubrication Check',
    'Float Pontoon Mooring & Water Submergence Depth Check',
    'Motor Winding Insulation & Running Current Balance'
  ],
  CONTROL: [
    'Main Circuit Breaker (MCCB), Magnetic Contactor & Thermal Overload Check',
    'VFD / Soft Starter Parameter, Fault Log & Cooling Fan Check',
    'Phase Protection Relay, Timer & Float/Pressure Switch Test',
    'Power & Control Terminal Torque Tightness & Thermoscan Check',
    'Pilot Lamp, Voltmeter & Ammeter Calibration Check'
  ]
};

export function getChecklistForType(
  type: string,
  templates: Record<string, string[]> = CHECKLIST_TEMPLATES
): string[] {
  const upper = type.toUpperCase();
  if (templates[upper]) return [...templates[upper]];
  if (upper.includes('MOTOR') && templates.MOTOR) return [...templates.MOTOR];
  if (upper.includes('BLOWER') && templates.BLOWER) return [...templates.BLOWER];
  if ((upper.includes('GEN SET') || upper.includes('GENSET')) && templates['GEN SET'])
    return [...templates['GEN SET']];
  if ((upper.includes('SEWAGE') || upper.includes('SUBMERSIBLE')) && templates.SEWAGE)
    return [...templates.SEWAGE];
  if (upper.includes('EJECTOR') && templates.EJECTOR) return [...templates.EJECTOR];
  if (upper.includes('AIR PUMP') && templates['AIR PUMP']) return [...templates['AIR PUMP']];
  if (upper.includes('AERATOR') && templates.AERATOR) return [...templates.AERATOR];
  if ((upper.includes('CONTROL') || upper.includes('PANEL')) && templates.CONTROL)
    return [...templates.CONTROL];
  return [...(templates.PUMP || CHECKLIST_TEMPLATES.PUMP)];
}

export function getChecklistMetaForTitle(itemText: string): {
  measuredValue: string;
  unit: string;
  acceptanceCriteria: string;
} {
  const lower = (itemText || '').toLowerCase();
  if (lower.includes('insulation') || lower.includes('megger') || lower.includes('dielectric')) {
    return {
      measuredValue: 'ก่อน: 0.5 / หลัง: 550',
      unit: 'MΩ',
      acceptanceCriteria: '≥ 100 MΩ (@500VDC IEEE 43)'
    };
  }
  if (lower.includes('alignment') || lower.includes('runout') || lower.includes('tir')) {
    return {
      measuredValue: 'ก่อน: 0.18 / หลัง: 0.03',
      unit: 'mm',
      acceptanceCriteria: '≤ 0.05 mm (Laser/Dial)'
    };
  }
  if (lower.includes('vibration') && lower.includes('temperature')) {
    return {
      measuredValue: '1.8 mm/s / 62 °C',
      unit: 'mm/s · °C',
      acceptanceCriteria: '≤ 4.5 mm/s (ISO 10816) · ≤ 80 °C'
    };
  }
  if (lower.includes('vibration')) {
    return {
      measuredValue: 'ก่อน: 6.8 / หลัง: 1.8',
      unit: 'mm/s',
      acceptanceCriteria: '≤ 4.5 mm/s (ISO 10816)'
    };
  }
  if (lower.includes('temperature') || lower.includes('thermoscan')) {
    return {
      measuredValue: 'ก่อน: 88 / หลัง: 62',
      unit: '°C',
      acceptanceCriteria: '≤ 80 °C'
    };
  }
  if (lower.includes('clearance') || lower.includes('wear ring')) {
    return {
      measuredValue: '0.25 / 64 °C',
      unit: 'mm · °C',
      acceptanceCriteria: '0.20-0.35 mm · ≤ 80 °C'
    };
  }
  if (lower.includes('pressure') || lower.includes('npsh') || lower.includes('vacuum') || lower.includes('frl')) {
    return {
      measuredValue: 'ก่อน: 1.8 / หลัง: 4.2',
      unit: 'bar',
      acceptanceCriteria: 'ตามพิกัด Nameplate / ไม่เกิด Cavitation'
    };
  }
  if (lower.includes('current') || lower.includes('voltage') || lower.includes('avr') || lower.includes('fla')) {
    return {
      measuredValue: '382V / 14.2A',
      unit: 'V / A',
      acceptanceCriteria: '380V ±5% · กระแส ≤ FLA Nameplate'
    };
  }
  if (lower.includes('winding resistance')) {
    return {
      measuredValue: '1.24 / 1.24 / 1.25',
      unit: 'Ω',
      acceptanceCriteria: 'Phase Unbalance ≤ 2%'
    };
  }
  return {
    measuredValue: 'ปกติ (ไม่รั่วซึม)',
    unit: '-',
    acceptanceCriteria: 'ปกติ / ไม่รั่วซึม / ได้มาตรฐาน'
  };
}

/**
 * Client-side canvas image compressor to ensure photos uploaded from mobile/album
 * are lightweight (~50-90KB) without overloading localStorage, memory, or network.
 */
export function compressImage(
  file: File,
  maxWidth = 1000,
  maxHeight = 1000,
  quality = 0.72
): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onerror = () => resolve('');
    reader.onload = (e) => {
      const img = new Image();
      img.onerror = () => resolve(String(e.target?.result || ''));
      img.onload = () => {
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
        if (!ctx) {
          resolve(String(e.target?.result || ''));
          return;
        }
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.src = String(e.target?.result || '');
    };
    reader.readAsDataURL(file);
  });
}

export function enrichChecklistItem(chk: ChecklistItem): ChecklistItem {
  const defaults = getChecklistMetaForTitle(chk.item);
  return {
    ...chk,
    measuredValue: chk.measuredValue !== undefined ? chk.measuredValue : defaults.measuredValue,
    unit: chk.unit !== undefined && chk.unit !== '' ? chk.unit : defaults.unit,
    acceptanceCriteria:
      chk.acceptanceCriteria !== undefined && chk.acceptanceCriteria !== ''
        ? chk.acceptanceCriteria
        : defaults.acceptanceCriteria,
    beforeStatus: chk.beforeStatus || 'NG',
    afterStatus: chk.afterStatus || 'OK'
  };
}

export function buildDefaultChecklistItems(
  type: string,
  templates: Record<string, string[]> = CHECKLIST_TEMPLATES
): ChecklistItem[] {
  return getChecklistForType(type, templates).map((item) => {
    const meta = getChecklistMetaForTitle(item);
    return {
      item,
      measuredValue: meta.measuredValue,
      unit: meta.unit,
      acceptanceCriteria: meta.acceptanceCriteria,
      beforeStatus: 'NG',
      afterStatus: 'OK'
    };
  });
}

export function isPartInternalWarranty(
  part: PartItem,
  eq: EquipmentItem,
  ticket: ServiceTicket
): boolean {
  if (part.billingType === 'WARRANTY_INTERNAL') return true;
  if (part.billingType === 'CUSTOMER_CHARGE') return false;
  if (eq.faultOrigin && eq.faultOrigin.includes('อยู่ในประกัน')) return true;
  if (eq.faultOrigin && eq.faultOrigin.includes('หมดรับประกัน')) return false;
  return Boolean(ticket.warranty && ticket.warranty.includes('In-Warranty'));
}

export function getTicketCostSummary(ticket: ServiceTicket): {
  isTicketInWarranty: boolean;
  internalCostTotal: number;
  customerBilledTotal: number;
  grossPartsTotal: number;
} {
  const isTicketInWarranty = Boolean(ticket.warranty && ticket.warranty.includes('In-Warranty'));
  let internalCostTotal = 0;
  let customerBilledTotal = 0;
  (ticket.equipmentList || []).forEach((eq) => {
    (eq.parts || []).forEach((p) => {
      const amount = Number(p.total) || 0;
      if (isPartInternalWarranty(p, eq, ticket)) {
        internalCostTotal += amount;
      } else {
        customerBilledTotal += amount;
      }
    });
  });
  return {
    isTicketInWarranty,
    internalCostTotal,
    customerBilledTotal,
    grossPartsTotal: internalCostTotal + customerBilledTotal
  };
}

export const DEFAULT_TROUBLESHOOTING_TEMPLATES: TroubleshootingTemplate[] = [
  {
    id: 'TPL-EASA-01',
    title: 'มอเตอร์ขดลวดไหม้ / ค่าฉนวนต่ำ (EASA AR100 & IEEE 43)',
    category: 'MOTOR',
    sourceRef: 'EASA AR100 / IEEE Std 43 / BULL POWER & BROOK CROMPTON Manual',
    analysis:
      'วัดค่าความต้านทานฉนวน (Insulation Resistance) ด้วย Megger 500VDC ได้ต่ำกว่าเกณฑ์มาตรฐาน IEEE 43 (< 100 MΩ) และพบกลิ่นไหม้ที่ขดลวดสเตเตอร์เนื่องจากความร้อนสะสมเกินพิกัด (Thermal Overload Class F/H)',
    cause:
      'แรงดันไฟฟ้าไม่สมดุล (Voltage Unbalance > 1%) หรือเกิดภาวะ Single Phasing / ทำงานเกินพิกัดกระแสต่อเนื่อง (Overload) ร่วมกับการระบายความร้อนไม่เพียงพอ',
    correctiveAction:
      'พันขดลวดสเตเตอร์ใหม่ด้วยลวดทองแดงเคลือบน้ำยา Class H อาบวานิชอบแห้งตามมาตรฐาน EASA AR100 เปลี่ยนตลับลูกปืนทั้งด้าน DE/NDE และทดสอบเดินตัวเปล่า (No-Load Test) วัดกระแสทั้ง 3 เฟสให้สมดุล',
    recommendedChecklist: [
      'Measure 3-Phase Supply Voltage Unbalance (< 1% per NEMA MG1)',
      'Polarization Index (PI) & Megger Insulation Test at 500/1000 VDC'
    ]
  },
  {
    id: 'TPL-PUMP-01',
    title: 'ปั๊มเกิดโพรงอากาศ / เสียงดังคล้ายกรวดและแรงดันตก (Cavitation)',
    category: 'PUMP',
    sourceRef: "Karassik's Pump Handbook (4th Ed.) / Hydraulic Institute / STAC & UNOMACH Guide",
    analysis:
      'เกิดเสียงดังรัวคล้ายเม็ดกรวดกระทบภายในเรือนปั๊ม (Volute Casing) ร่วมกับค่าความสั่นสะเทือนสูงผิดปกติ และเข็มเกจวัดแรงดันด้านส่ง (Discharge Pressure) สั่นแกว่งเนื่องจากเกิดภาวะโพรงอากาศ (Suction Cavitation: NPSHa < NPSHr)',
    cause:
      'ท่อทางดูด (Suction Line) หรือ Strainer อุดตันทำให้เกิด Pressure Drop สูง หรือระดับน้ำฝั่งดูดต่ำกว่าค่าออกแบบ (Insufficient Submergence / High Suction Lift)',
    correctiveAction:
      'ถอดล้างทำความสะอาด Y-Strainer และตรวจสอบวาล์วทางดูด ตรวจสอบรอยรั่วอากาศฝั่งท่อดูด (Suction Air Ingress) ตรวจสอบสภาพใบพัด (Impeller) หากสึกกร่อนมากให้เปลี่ยนใบพัดและ Wear Ring ใหม่พร้อมทำ Dynamic Balancing ตามมาตรฐาน ISO 1940 G6.3',
    recommendedChecklist: [
      'Suction Compound Gauge Vacuum Reading & Strainer Differential Check',
      'Impeller Vane Pitting & Wear Ring Radial Clearance Measurement'
    ]
  },
  {
    id: 'TPL-PUMP-02',
    title: 'ซีลกันรั่วซึม / ตลับลูกปืนร้อนและสั่นสะเทือนสูง (Misalignment & Seal Failure)',
    category: 'PUMP',
    sourceRef: "Karassik's Pump Handbook / VARISCO & FYBROC Maintenance Manual / ISO 10816",
    analysis:
      'พบการรั่วซึมของของเหลวออกทางชุด Mechanical Seal อย่างต่อเนื่อง และอุณหภูมิเสื้อลูกปืน (Bearing Housing) สูงเกิน 85°C ร่วมกับค่า Vibration เกินเกณฑ์มาตรฐาน ISO 10816',
    cause:
      'เกิดจากการเยื้องศูนย์ของเพลาปั๊มและมอเตอร์ (Angular / Parallel Shaft Misalignment) หรือเกิดภาวะเดินปั๊มตัวเปล่าโดยไม่มีน้ำหล่อเลี้ยงซีล (Dry Running / Thermal Shock)',
    correctiveAction:
      'ถอดเปลี่ยนชุด Mechanical Seal (SiC/SiC/Viton) และตลับลูกปืนใหม่ ตรวจสอบค่าความคดของเพลา (Shaft Runout TIR ≤ 0.05 mm) และทำ Laser Shaft Alignment หน้างานให้ได้ค่าความคลาดเคลื่อนไม่เกิน ±0.05 mm พร้อมตรวจสอบความเครียดท่อ (Pipe Strain)',
    recommendedChecklist: [
      'Shaft Runout (TIR) & Coupling Angular/Parallel Alignment Check',
      'Pipe Strain Check on Suction & Discharge Flanges'
    ]
  },
  {
    id: 'TPL-SEWAGE-01',
    title: 'ปั๊มจุ่มน้ำเสียตัด Overload / ใบพัดติดขัดและน้ำเข้าห้องน้ำมัน',
    category: 'SEWAGE',
    sourceRef: 'Submersible Sewage Pump Handbook / UNOMACH & VARISCO Troubleshooting Guide',
    analysis:
      'เบรกเกอร์และ Overload Relay ทริปทันทีเมื่อสั่งสตาร์ท ตรวจพบเศษขยะพันติดใบพัด (Impeller Jamming) และพบน้ำปนเปื้อนในห้องน้ำมันซีล (Oil Chamber Emulsified) ทำให้ซีลตัวล่างสึกหรอ',
    cause:
      'ขยะเส้นใยยาวหรือของแข็งเกินขนาดผ่านเข้าใบพัดทำให้เพลาล็อก (Locked Rotor Current สูง 6-7 เท่า) และทำลายผิวหน้าสัมผัส Mechanical Seal',
    correctiveAction:
      'ถอดล้างสิ่งอุดตันในเรือนปั๊มและใบพัด เปลี่ยนชุด Double Mechanical Seal และถ่ายเปลี่ยนน้ำมันหล่อลื่น Food Grade / Turbine Oil VG32 ในห้องซีล อบไล่ความชื้นขดลวดและอัดอีพ็อกซี่ที่ขั้วสายไฟ (Cable Entry) ใหม่',
    recommendedChecklist: [
      'Oil Chamber Dielectric & Moisture Inspection',
      'Submersible Cable Insulation & Float Switch Level Check'
    ]
  },
  {
    id: 'TPL-CTRL-01',
    title: 'ตู้คอนโทรลทริป / Magnetic Contactor หน้าสัมผัสไหม้ / VFD Fault',
    category: 'CONTROL',
    sourceRef: 'IEC 60947-4-1 / EASA Electrical Testing & Control Panel Troubleshooting',
    analysis:
      'ตรวจพบรอยอาร์คและหน้าสัมผัส (Contact Tips) ของ Magnetic Contactor สึกหรอไม่เท่ากัน ส่งผลให้แรงดันไฟฟ้าตกคร่อมเฟสและทำให้ Thermal Overload หรือ VFD แจ้งเตือน Overcurrent / Phase Loss',
    cause:
      'เกิดจากการสตาร์ท-หยุดบ่อยครั้ง (Frequent Jogging/Starting) หรือขั้วต่อสายไฟหลวมทำให้เกิดความร้อนสูงที่จุดต่อ (Loose Terminal Hot Spot)',
    correctiveAction:
      'เปลี่ยนชุด Magnetic Contactor และ Overload Relay ให้ตรงตามพิกัดกระแส (AC-3 Duty) ขันกวดขั้วต่อสายกำลังทั้งหมดตามค่า Torque มาตรฐาน ตั้งค่าพารามิเตอร์ Ramp-up/Ramp-down และ Current Limit ใน VFD/Soft Starter ใหม่',
    recommendedChecklist: [
      'Contactor Contact Voltage Drop & Phase Balance Test',
      'Overload Relay Trip Current Setting vs Motor Nameplate FLA'
    ]
  }
];

const DEFAULT_BEFORE_IMAGE = pumpBeforeImg;
const DEFAULT_AFTER_IMAGE = pumpAfterImg;

export const TECH_POSITIONS = [
  'ช่างเทคนิค (Technician)',
  'หัวหน้าช่าง (Lead Technician)',
  'วิศวกรบริการ (Service Engineer)'
];

export const DEFAULT_SUPERVISORS = [
  'หัวหน้าวิศวกรบริการ (Supervisor)',
  'ผู้จัดการฝ่ายบริการ (Service Manager)'
];

const INITIAL_TECHS: Technician[] = [];

export const isDemoTechnician = (tc: { name?: string; id?: string }): boolean => {
  if (!tc || !tc.name) return false;
  const n = tc.name.trim();
  return (
    n.includes('วิชัย') ||
    n.includes('สมชาย') ||
    n.includes('ประเสริฐ')
  );
};

export const isDemoTicket = (t: { id?: string; company?: string }): boolean => {
  if (!t || !t.id) return false;
  return (
    t.id === 'SKM-69-001' ||
    t.id === 'SKM-69-002' ||
    t.id === 'SKM-69-003' ||
    (t.company === 'บริษัท สยามอุตสาหกรรม จำกัด' && t.id.startsWith('SKM-69-00')) ||
    (t.company === 'บริษัท ไทยเคมีคอล เอ็นจิเนียริ่ง จำกัด' && t.id.startsWith('SKM-69-00'))
  );
};

// Production Ready: Clean start without mock demo tickets
const INITIAL_TICKETS: ServiceTicket[] = [];

function playNotificationChime(enabled: boolean) {
  if (!enabled) return;
  try {
    const AudioCtx =
      window.AudioContext ||
      (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sine';
    osc.frequency.setValueAtTime(880, ctx.currentTime);
    osc.frequency.setValueAtTime(1174.66, ctx.currentTime + 0.11);
    gain.gain.setValueAtTime(0.08, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.35);
  } catch {
    // Ignore audio context restriction before user gesture
  }
}

interface ResilientImageProps {
  src: string;
  alt: string;
  className?: string;
}

const ResilientImage: React.FC<ResilientImageProps> = ({ src, alt, className = '' }) => {
  const [hasError, setHasError] = useState(false);

  useEffect(() => {
    setHasError(false);
  }, [src]);

  if (!src || hasError) {
    return (
      <div
        className={`flex flex-col items-center justify-center bg-slate-100 border border-slate-200 text-slate-500 p-2 text-center ${className}`}
      >
        <Wrench className="w-5 h-5 text-slate-400 mb-1" />
        <span className="text-[10px] leading-tight line-clamp-2">{alt || 'ไม่มีรูปภาพ'}</span>
      </div>
    );
  }

  return (
    <img
      src={src}
      alt={alt}
      referrerPolicy="no-referrer"
      onError={() => setHasError(true)}
      className={className}
    />
  );
};

interface NewEquipmentDraft {
  type: string;
  customType: string;
  brand: string;
  customBrand: string;
  model: string;
  tag: string;
  symptom: string;
  detail: string;
  beforePhotos: string[];
  afterPhotos: string[];
}

export function App() {
  const [currentView, setCurrentView] = useState<
    'home' | 'customer_portal' | 'internal_login' | 'internal_dashboard' | 'service_report'
  >('home');

  const [previousView, setPreviousView] = useState<'home' | 'customer_portal' | 'internal_dashboard'>('home');
  const [userRole, setUserRole] = useState<UserRole>(null);
  const [customerCompany, setCustomerCompany] = useState<string>('บริษัท สยามอุตสาหกรรม จำกัด');
  const [soundEnabled, setSoundEnabled] = useState<boolean>(true);

  // Editable Master Data: Equipment Types, Brands, Checklist Templates & Company Profile
  const [masterEquipmentTypes, setMasterEquipmentTypes] = useState<string[]>(() => {
    const defaultBase = DEFAULT_EQUIPMENT_MASTERS.filter((x) => !x.startsWith('ETC.'));
    const savedMaster = localStorage.getItem('sahakim_master_equip_types');
    const savedCustom = localStorage.getItem('sahakim_custom_equip_types');
    let extraCustom: string[] = [];
    if (savedCustom) {
      try {
        extraCustom = JSON.parse(savedCustom);
      } catch {
        extraCustom = [];
      }
    }
    if (savedMaster) {
      try {
        const parsed = JSON.parse(savedMaster);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return Array.from(new Set([...parsed, ...extraCustom]));
        }
      } catch {
        // fallback
      }
    }
    return Array.from(new Set([...defaultBase, ...extraCustom]));
  });

  const [masterBrands, setMasterBrands] = useState<string[]>(() => {
    const defaultBase = DEFAULT_BRANDS.filter((x) => !x.startsWith('ETC.'));
    const savedMaster = localStorage.getItem('sahakim_master_brands');
    const savedCustom = localStorage.getItem('sahakim_custom_brands');
    let extraCustom: string[] = [];
    if (savedCustom) {
      try {
        extraCustom = JSON.parse(savedCustom);
      } catch {
        extraCustom = [];
      }
    }
    if (savedMaster) {
      try {
        const parsed = JSON.parse(savedMaster);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return Array.from(new Set([...parsed, ...extraCustom]));
        }
      } catch {
        // fallback
      }
    }
    return Array.from(new Set([...defaultBase, ...extraCustom]));
  });

  const [masterChecklists, setMasterChecklists] = useState<Record<string, string[]>>(() => {
    const saved = localStorage.getItem('sahakim_master_checklists');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          return { ...CHECKLIST_TEMPLATES, ...parsed };
        }
      } catch {
        // fallback
      }
    }
    return { ...CHECKLIST_TEMPLATES };
  });

  const [companyProfile, setCompanyProfile] = useState<{
    nameTh: string;
    nameEn: string;
    phone: string;
    tagline: string;
    logoUrl: string;
    staffMascotUrl: string;
  }>(() => {
    const defaultProfile = {
      nameTh: 'บริษัท สหคิมมอเตอร์ จำกัด',
      nameEn: 'SAHA KIM MOTOR CO., LTD.',
      phone: '02-XXX-XXXX',
      tagline:
        'บริการซ่อม บำรุงรักษา และจำหน่ายปั๊มน้ำ มอเตอร์ Blower, Gen Set, Sewage, Aerator และระบบควบคุมอุตสาหกรรม',
      logoUrl: DEFAULT_SAHAKIM_LOGO,
      staffMascotUrl: SKM_SERVICE_MASCOT_LOGO
    };
    const saved = localStorage.getItem('sahakim_company_profile');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return {
          ...defaultProfile,
          ...parsed,
          logoUrl: parsed.logoUrl || DEFAULT_SAHAKIM_LOGO,
          staffMascotUrl: parsed.staffMascotUrl || SKM_SERVICE_MASCOT_LOGO
        };
      } catch {
        return defaultProfile;
      }
    }
    return defaultProfile;
  });

  const equipmentOptions = useMemo(() => {
    return [...masterEquipmentTypes, 'ETC. (กรอกเพิ่มเองได้)'];
  }, [masterEquipmentTypes]);

  const brandOptions = useMemo(() => {
    return [...masterBrands, 'ETC. (กรอกเพิ่มเองได้)'];
  }, [masterBrands]);

  useEffect(() => {
    localStorage.setItem('sahakim_master_equip_types', JSON.stringify(masterEquipmentTypes));
  }, [masterEquipmentTypes]);

  useEffect(() => {
    localStorage.setItem('sahakim_master_brands', JSON.stringify(masterBrands));
  }, [masterBrands]);

  useEffect(() => {
    localStorage.setItem('sahakim_master_checklists', JSON.stringify(masterChecklists));
  }, [masterChecklists]);

  useEffect(() => {
    localStorage.setItem('sahakim_company_profile', JSON.stringify(companyProfile));
  }, [companyProfile]);

  // Custom Knowledge Base Templates saved by Technicians
  const [knowledgeTemplates, setKnowledgeTemplates] = useState<TroubleshootingTemplate[]>(() => {
    const saved = localStorage.getItem('sahakim_knowledge_templates');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch {
        return DEFAULT_TROUBLESHOOTING_TEMPLATES;
      }
    }
    return DEFAULT_TROUBLESHOOTING_TEMPLATES;
  });

  useEffect(() => {
    localStorage.setItem('sahakim_knowledge_templates', JSON.stringify(knowledgeTemplates));
  }, [knowledgeTemplates]);

  const [supPassword, setSupPassword] = useState<string>(() => {
    return localStorage.getItem('sahakim_sup_pass') || 'admin123';
  });

  const [servicePin, setServicePin] = useState<string>(() => {
    return localStorage.getItem('sahakim_service_pin') || '1234';
  });

  const [gasWebAppUrl, setGasWebAppUrl] = useState<string>(() => {
    return localStorage.getItem('sahakim_gas_url') || '';
  });
  const [lineChannelToken, setLineChannelToken] = useState<string>(() => {
    try {
      localStorage.removeItem('sahakim_line_token');
    } catch {}
    return '';
  });
  const [lineTargetId, setLineTargetId] = useState<string>(() => {
    try {
      localStorage.removeItem('sahakim_line_target');
    } catch {}
    return '';
  });
  const [gasModalOpen, setGasModalOpen] = useState(false);
  const [isGoogleSheetsModalOpen, setIsGoogleSheetsModalOpen] = useState(false);
  const [googleSheetId, setGoogleSheetId] = useState<string>(() => {
    return localStorage.getItem('sahakim_google_sheet_id') || '';
  });
  const [googleSheetName, setGoogleSheetName] = useState<string>(() => {
    return localStorage.getItem('sahakim_google_sheet_name') || '';
  });
  const [googleAutoSync, setGoogleAutoSync] = useState<boolean>(() => {
    const s = localStorage.getItem('sahakim_google_autosync');
    return s === null ? true : s === 'true';
  });
  const [googleLastSyncedAt, setGoogleLastSyncedAt] = useState<string | null>(() => {
    return localStorage.getItem('sahakim_google_last_synced') || null;
  });

  const handleSetConnectedSheetId = (id: string, name?: string) => {
    setGoogleSheetId(id);
    if (id) {
      localStorage.setItem('sahakim_google_sheet_id', id);
      if (name) {
        setGoogleSheetName(name);
        localStorage.setItem('sahakim_google_sheet_name', name);
      }
    } else {
      localStorage.removeItem('sahakim_google_sheet_id');
      localStorage.removeItem('sahakim_google_sheet_name');
      setGoogleSheetName('');
    }
  };

  const handleToggleGoogleAutoSync = (enabled: boolean) => {
    setGoogleAutoSync(enabled);
    localStorage.setItem('sahakim_google_autosync', String(enabled));
  };

  const handleSetGoogleLastSyncedAt = (timestamp: string) => {
    setGoogleLastSyncedAt(timestamp);
    localStorage.setItem('sahakim_google_last_synced', timestamp);
  };

  const handleTicketsImportedFromGoogleSheets = (imported: ServiceTicket[]) => {
    if (Array.isArray(imported) && imported.length > 0) {
      setTickets(imported);
      localStorage.setItem('sahakim_tickets', JSON.stringify(imported));
      pushOnlineSharedState(imported, userRole || 'SUPERVISOR');
      showToast(`นำเข้าใบงานจาก Google Sheet ทั้งหมด ${imported.length} รายการเรียบร้อยแล้ว`, 'success');
    }
  };

  const [technicians, setTechnicians] = useState<Technician[]>(() => {
    const saved = localStorage.getItem('sahakim_technicians');
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as Technician[];
        const clean = Array.isArray(parsed)
          ? parsed.filter((tc) => !isDemoTechnician(tc))
          : [];
        return clean.map((tc, i) => ({
          ...tc,
          position:
            tc.position ||
            (i === 0 ? 'หัวหน้าช่าง (Lead Technician)' : 'ช่างเทคนิค (Technician)')
        }));
      } catch {
        return [];
      }
    }
    return [];
  });

  const [supervisors, setSupervisors] = useState<string[]>(() => {
    const saved = localStorage.getItem('sahakim_supervisors');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const clean = parsed.filter((s) => typeof s === 'string' && !s.includes('วิชัย'));
          return clean.length > 0 ? clean : DEFAULT_SUPERVISORS;
        }
      } catch {
        // ignore
      }
    }
    return DEFAULT_SUPERVISORS;
  });

  useEffect(() => {
    localStorage.setItem('sahakim_supervisors', JSON.stringify(supervisors));
  }, [supervisors]);

  // Current logged in technician identity (for Technicians viewing only their own jobs)
  const [loggedInTech, setLoggedInTech] = useState<Technician | null>(() => {
    const saved = localStorage.getItem('sahakim_logged_in_tech');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && !isDemoTechnician(parsed)) return parsed;
        localStorage.removeItem('sahakim_logged_in_tech');
        return null;
      } catch {
        return null;
      }
    }
    return null;
  });
  const [selectedTechNameForLogin, setSelectedTechNameForLogin] = useState<string>('');
  const [techWorkFilter, setTechWorkFilter] = useState<'MY_JOBS' | 'AVAILABLE_JOBS'>('MY_JOBS');
  const [techIdToDeleteConfirm, setTechIdToDeleteConfirm] = useState<string | null>(null);

  const [tickets, setTickets] = useState<ServiceTicket[]>(() => {
    const saved = localStorage.getItem('sahakim_tickets_v121');
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as ServiceTicket[];
        const listToUse = Array.isArray(parsed)
          ? parsed.filter((t) => !isDemoTicket(t))
          : [];
        return listToUse.map((t) => ({
          ...t,
          equipmentList: (t.equipmentList || []).map((eq) => ({
            ...eq,
            beforePhotos: eq.beforePhotos || [],
            afterPhotos: eq.afterPhotos || [],
            checklistResults: (eq.checklistResults || []).map((chk) => enrichChecklistItem(chk)),
            parts: (eq.parts || []).map((p) => ({
              ...p,
              billingType:
                p.billingType ||
                (isPartInternalWarranty(p, eq, t) ? 'WARRANTY_INTERNAL' : 'CUSTOMER_CHARGE')
            }))
          }))
        }));
      } catch {
        return [];
      }
    }
    return [];
  });

  useEffect(() => {
    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(tickets));
  }, [tickets]);

  useEffect(() => {
    localStorage.setItem('sahakim_technicians', JSON.stringify(technicians));
  }, [technicians]);

  useEffect(() => {
    localStorage.setItem('sahakim_sup_pass', supPassword);
  }, [supPassword]);

  useEffect(() => {
    localStorage.setItem('sahakim_service_pin', servicePin);
  }, [servicePin]);

  useEffect(() => {
    localStorage.setItem('sahakim_gas_url', gasWebAppUrl);
  }, [gasWebAppUrl]);



  const [sheetSyncStatus, setSheetSyncStatus] = useState<{
    ok: boolean;
    text: string;
    time: string;
  } | null>(null);
  const [isSyncingSheet, setIsSyncingSheet] = useState(false);

  // Real-Time Online Multi-User Sync State
  const onlineClientId = useMemo(
    () => 'skm-' + Math.random().toString(36).slice(2, 10),
    []
  );
  const [onlineSyncStatus, setOnlineSyncStatus] = useState<{
    connected: boolean;
    version: number;
    lastSyncTime: string;
    onlineClients: number;
    isSyncing: boolean;
  }>({
    connected: true,
    version: 0,
    lastSyncTime: new Date().toLocaleTimeString('th-TH'),
    onlineClients: 1,
    isSyncing: false
  });

  const sanitizeTicketsForSheet = (list: ServiceTicket[]): ServiceTicket[] => {
    return list.map((t) => ({
      ...t,
      equipmentList: (t.equipmentList || []).map((eq) => ({
        ...eq,
        beforePhotos: (eq.beforePhotos || []).map((p) =>
          typeof p === 'string' && (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800)
            ? 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=600'
            : p
        ),
        afterPhotos: (eq.afterPhotos || []).map((p) =>
          typeof p === 'string' && (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800)
            ? 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=600'
            : p
        )
      }))
    }));
  };

  const fetchBackendLineConfig = async () => {
    try {
      const res = await fetch('/api/line/config');
      if (res.ok) {
        const data = await res.json();
        if (data?.gasWebAppUrl && !gasWebAppUrl.trim()) {
          setGasWebAppUrl(data.gasWebAppUrl);
        }
        if (data?.targetId) {
          if (
            !lineTargetId.trim() ||
            (String(data.targetId).startsWith('C') && !lineTargetId.trim().startsWith('C'))
          ) {
            setLineTargetId(data.targetId);
          }
          return data;
        }
      }
    } catch {
      // ignore backend error and try GAS fallback
    }
    if (gasWebAppUrl.trim() && !gasWebAppUrl.includes('docs.google.com/spreadsheets')) {
      try {
        const gasRes = await fetch(gasWebAppUrl.trim(), { redirect: 'follow' });
        if (gasRes.ok) {
          const gasData = await gasRes.json();
          const detectedId = gasData?.lineGroupId || gasData?.targetId || '';
          if (detectedId) {
            if (
              !lineTargetId.trim() ||
              (String(detectedId).startsWith('C') && !lineTargetId.trim().startsWith('C'))
            ) {
              setLineTargetId(detectedId);
            }
            return { targetId: detectedId };
          }
        }
      } catch {
        // ignore
      }
    }
    return null;
  };

  useEffect(() => {
    fetchBackendLineConfig();
    if (!gasModalOpen) return;
    const timer = setInterval(() => {
      fetchBackendLineConfig();
    }, 4000);
    return () => clearInterval(timer);
  }, [gasModalOpen]);

  const formatTicketLineMessage = (t: ServiceTicket, eventHeading?: string) => {
    const portalUrl = 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
    const eqSummary = (t.equipmentList || [])
      .map((eq) => `${eq.type} (${eq.brand} ${eq.model || '-'}) อาการ: ${eq.symptom}`)
      .join(', ');
    const costInfo = getTicketCostSummary(t);
    const partsList: string[] = [];
    (t.equipmentList || []).forEach((eq) => {
      (eq.parts || []).forEach((p) => {
        const isInternal = isPartInternalWarranty(p, eq, t);
        partsList.push(
          `${p.name} x${p.qty} (${p.total.toLocaleString()}฿ ${
            isInternal ? '[เคลมในประกัน-ไม่คิดเงินลูกค้า]' : '[คิดเงินลูกค้า]'
          })`
        );
      });
    });

    const lines = [
      eventHeading || '🔔 [อัปเดตงานบริการ - สหคิมมอเตอร์]',
      '━━━━━━━━━━━━━━',
      `📋 เลขที่ใบงาน: ${t.id}`,
      `🏢 บริษัท: ${t.company}`,
      `📌 สถานะ: ${t.status}`,
      `🛡️ เงื่อนไขประกัน: ${t.warranty || '-'}`,
      `🚨 ความเร่งด่วน: ${t.urgency || '-'}`,
      `🔧 อุปกรณ์: ${eqSummary || '-'}`,
      `📍 สถานที่: ${t.location || '-'}`,
      `📞 ผู้ติดต่อ: ${t.siteContact || t.contact || '-'}`
    ];
    if (t.assignedTech) lines.push(`👷 ช่างผู้ปฏิบัติงาน: ${t.assignedTech}`);
    if (t.leadTech) lines.push(`👨‍🔧 หัวหน้าช่างผู้ควบคุมงาน: ${t.leadTech}`);
    if (t.approvedBySupervisor) {
      lines.push(`✅ ผู้อนุมัติ (Supervisor/หัวหน้าช่าง): ${t.approvedBySupervisor}`);
    }
    if (t.confirmedDate || t.desiredDate) {
      lines.push(`📅 วันนัดหมาย: ${t.confirmedDate || t.desiredDate}`);
    }
    if (partsList.length > 0) {
      lines.push(`⚙️ รายการอะไหล่: ${partsList.join(', ')}`);
      lines.push(
        `💰 สรุปค่าใช้จ่าย: ต้นทุนภายในบริษัท ${costInfo.internalCostTotal.toLocaleString()} THB | ยอดเรียกเก็บลูกค้าสุทธิ ${costInfo.customerBilledTotal.toLocaleString()} THB${
          costInfo.customerBilledTotal === 0 ? ' (ฟรี/อยู่ในประกัน ไม่เรียกเก็บลูกค้า)' : ''
        }`
      );
    }
    if (t.rejectReason && t.status === 'REJECTED') {
      lines.push(`❌ เหตุผลส่งกลับแก้ไข: ${t.rejectReason}`);
    }
    lines.push('━━━━━━━━━━━━━━');
    lines.push(`🔗 เปิดระบบ/อนุมัติงาน: ${portalUrl}`);
    return lines.join('\n');
  };

  const getTodayDateStr = () => {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };

  const extractTicketDueDate = (t: ServiceTicket) => {
    const raw = (t.confirmedDate || t.desiredDate || '').trim();
    const match = raw.match(/\d{4}-\d{2}-\d{2}/);
    return match ? match[0] : raw;
  };

  const isTicketOverdue = (t: ServiceTicket) => {
    if (!t || t.status === 'CLOSE') return false;
    const dueStr = extractTicketDueDate(t);
    if (!dueStr || !/^\d{4}-\d{2}-\d{2}$/.test(dueStr)) return false;
    return dueStr < getTodayDateStr();
  };

  const getOverdueDays = (t: ServiceTicket) => {
    const dueStr = extractTicketDueDate(t);
    const todayStr = getTodayDateStr();
    if (!dueStr || dueStr >= todayStr) return 0;
    const diffMs = new Date(todayStr).getTime() - new Date(dueStr).getTime();
    return Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)));
  };

  const formatOverdueSummaryLineMessage = (overdueList: ServiceTicket[]) => {
    const portalUrl = 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
    const todayStr = getTodayDateStr();
    const lines = [
      '🚨 [แจ้งเตือนงานค้างเกินกำหนดนัดหมาย (OVERDUE) - สหคิมมอเตอร์]',
      `📅 ข้อมูล ณ วันที่: ${todayStr}`,
      `⚠️ พบงานค้างเกินกำหนดนัดหมาย จำนวน ${overdueList.length} ใบงาน:`,
      '━━━━━━━━━━━━━━'
    ];
    overdueList.forEach((t, idx) => {
      const dueStr = extractTicketDueDate(t);
      const daysLate = getOverdueDays(t);
      const eqStr = (t.equipmentList || [])
        .map((eq) => `${eq.type} (${eq.brand})`)
        .join(', ');
      lines.push(`${idx + 1}) 📋 ${t.id} | ${t.company}`);
      lines.push(`   📌 สถานะ: ${t.status}`);
      lines.push(`   🗓️ วันนัดหมาย: ${dueStr} (เกินกำหนด ${daysLate} วัน)`);
      lines.push(`   👷 ช่างผู้รับผิดชอบ: ${t.assignedTech || 'ยังไม่ระบุช่าง'}`);
      lines.push(`   🔧 อุปกรณ์: ${eqStr || '-'}`);
      lines.push(`   📍 สถานที่: ${t.location || '-'}`);
      lines.push('━━━━━━━━━━━━━━');
    });
    lines.push(`🔗 กดเพื่อตรวจสอบและเร่งรัดงาน: ${portalUrl}`);
    return lines.join('\n');
  };

  const sendLineDirectViaServer = async (
    messageText: string
  ): Promise<{ ok: boolean; quotaExceeded?: boolean; message?: string }> => {
    if (!messageText.trim()) return { ok: false };
    try {
      const res = await fetch('/api/line/notify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: messageText,
          targetId: lineTargetId.trim()
        })
      });
      if (res.ok) {
        const data = await res.json();
        if (data?.quotaExceeded || data?.status === 'quota_exceeded') {
          return { ok: false, quotaExceeded: true, message: data.message };
        }
        if (data?.targetId && data.targetId !== 'BROADCAST' && data.targetId !== lineTargetId) {
          setLineTargetId(data.targetId);
        }
        if (data?.status === 'ok') {
          return { ok: true };
        }
      }
    } catch {
      // ignore network fetch error
    }
    return { ok: false };
  };

  const formatAllTicketsSummaryLineMessage = (allList: ServiceTicket[]) => {
    const portalUrl = 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
    const todayStr = getTodayDateStr();
    const overdueCount = allList.filter((t) => isTicketOverdue(t)).length;
    const closedCount = allList.filter((t) => t.status === 'CLOSE').length;
    const activeCount = allList.length - closedCount;
    const lines = [
      '📊 [อัปเดตสรุปข้อมูลใบงานทั้งหมด - บริษัท สหคิมมอเตอร์ จำกัด]',
      `📅 ข้อมูล ณ วันที่: ${todayStr}`,
      `📋 งานทั้งหมด: ${allList.length} ใบงาน (กำลังดำเนินการ ${activeCount} | ปิดงานแล้ว ${closedCount} | เกินกำหนด ${overdueCount})`,
      '━━━━━━━━━━━━━━'
    ];
    allList.forEach((t, idx) => {
      const dueStr = extractTicketDueDate(t);
      const overdue = isTicketOverdue(t);
      const eqStr = (t.equipmentList || [])
        .map((eq) => `${eq.type} (${eq.brand} ${eq.model || '-'})`)
        .join(', ');
      lines.push(`${idx + 1}) 📋 ${t.id} — ${t.company}`);
      lines.push(`   📌 สถานะ: ${t.status}${overdue ? ' [⚠️ เกินกำหนดนัด]' : ''}`);
      lines.push(`   🔧 อุปกรณ์: ${eqStr || '-'}`);
      lines.push(
        `   👷 ช่าง: ${t.assignedTech || 'ยังไม่ระบุ'} | 🗓️ นัด: ${dueStr || '-'}`
      );
      lines.push('━━━━━━━━━━━━━━');
    });
    lines.push(`🔗 เปิดระบบจัดการใบงาน: ${portalUrl}`);
    return lines.join('\n');
  };

  // Push current state to central server (/api/online/state) so all online users stay in sync automatically
  const pushOnlineSharedState = async (
    overrideTickets?: ServiceTicket[],
    updatedByRole?: string,
    overrideTechs?: Technician[],
    overrideSups?: string[]
  ): Promise<boolean> => {
    try {
      const payloadTickets = overrideTickets || tickets;
      const payloadTechs = overrideTechs || technicians;
      const payloadSups = overrideSups || supervisors;
      const res = await fetch('/api/online/state', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          clientId: onlineClientId,
          updatedBy: updatedByRole || userRole || 'ONLINE_USER',
          tickets: payloadTickets,
          technicians: payloadTechs,
          supervisors: payloadSups,
          masterBrands,
          masterEquipmentTypes,
          masterChecklists,
          companyProfile
        })
      });
      if (res.ok) {
        const data = await res.json();
        setOnlineSyncStatus((prev) => ({
          ...prev,
          connected: true,
          version: Number(data.version) || prev.version + 1,
          lastSyncTime: new Date().toLocaleTimeString('th-TH'),
          onlineClients: Number(data.onlineClients) || prev.onlineClients
        }));
        return true;
      }
    } catch {
      // Ignore if running as offline standalone HTML
    }
    return false;
  };

  // Helper to persist technicians and push to server immediately
  const handleSaveTechniciansList = async (
    updatedList: Technician[],
    successMessage?: string
  ) => {
    setTechnicians(updatedList);
    try {
      localStorage.setItem('sahakim_technicians', JSON.stringify(updatedList));
    } catch {
      // ignore
    }
    await pushOnlineSharedState(tickets, userRole || 'SUPERVISOR', updatedList, supervisors);
    if (successMessage) {
      showToast(successMessage, 'success');
    }
  };

  // Helper to persist supervisors and push to server immediately
  const handleSaveSupervisorsList = async (
    updatedList: string[],
    successMessage?: string
  ) => {
    setSupervisors(updatedList);
    try {
      localStorage.setItem('sahakim_supervisors', JSON.stringify(updatedList));
    } catch {
      // ignore
    }
    await pushOnlineSharedState(tickets, userRole || 'SUPERVISOR', technicians, updatedList);
    if (successMessage) {
      showToast(successMessage, 'success');
    }
  };

  // Pull latest shared state from central server (/api/online/state)
  const pullOnlineSharedState = async (showNotification = false): Promise<boolean> => {
    try {
      const res = await fetch('/api/online/state');
      if (!res.ok) return false;
      const data = await res.json();
      if (data && data.status === 'ok') {
        const serverVer = Number(data.version) || 0;
        if (serverVer === 0 && tickets.length > 0) {
          // Server is fresh — seed it with our current data
          await pushOnlineSharedState(tickets, 'INITIAL_SEED');
          return true;
        }
        const rawServerTickets = Array.isArray(data.tickets)
          ? data.tickets
          : data.data?.tickets || [];
        const cleanServerTickets = (Array.isArray(rawServerTickets) ? rawServerTickets : [])
          .filter((t: ServiceTicket) => !isDemoTicket(t));
        setTickets(cleanServerTickets);
        localStorage.setItem('sahakim_tickets_v121', JSON.stringify(cleanServerTickets));

        if (Array.isArray(data.technicians)) {
          const cleanTechs = data.technicians.filter((tc: any) => !isDemoTechnician(tc));
          setTechnicians(cleanTechs);
          localStorage.setItem('sahakim_technicians', JSON.stringify(cleanTechs));
        }
        if (Array.isArray(data.supervisors)) {
          const cleanSups = data.supervisors.filter((s: string) => typeof s === 'string' && !s.includes('วิชัย'));
          setSupervisors(cleanSups);
          localStorage.setItem('sahakim_supervisors', JSON.stringify(cleanSups));
        }
        if (Array.isArray(data.masterBrands) && data.masterBrands.length > 0) {
          setMasterBrands(data.masterBrands);
        }
        if (Array.isArray(data.masterEquipmentTypes) && data.masterEquipmentTypes.length > 0) {
          setMasterEquipmentTypes(data.masterEquipmentTypes);
        }
        if (data.masterChecklists && typeof data.masterChecklists === 'object') {
          setMasterChecklists((prev) => ({ ...prev, ...data.masterChecklists }));
        }
        if (data.companyProfile && typeof data.companyProfile === 'object') {
          setCompanyProfile((prev) => ({ ...prev, ...data.companyProfile }));
        }
        setOnlineSyncStatus((prev) => ({
          ...prev,
          connected: true,
          version: serverVer,
          lastSyncTime: new Date().toLocaleTimeString('th-TH'),
          onlineClients: Number(data.onlineClients) || 1
        }));
        if (showNotification) {
          showToast(
            `🔄 อัปเดตข้อมูลออนไลน์ล่าสุดสำเร็จ (${cleanServerTickets.length} ใบงานจริงในระบบ)`,
            'success'
          );
        }
        return true;
      }
    } catch {
      // Ignore if offline
    }
    return false;
  };

  // Subscribe to Server-Sent Events (/api/online/events) for instant multi-user updates
  useEffect(() => {
    pullOnlineSharedState(false);

    let es: EventSource | null = null;
    try {
      es = new EventSource('/api/online/events');
      es.onmessage = (ev) => {
        try {
          const msg = JSON.parse(ev.data);
          if (msg.type === 'state_updated' && msg.clientId !== onlineClientId) {
            pullOnlineSharedState(false);
          }
        } catch {
          // ignore
        }
      };
      es.onerror = () => {
        setOnlineSyncStatus((prev) => ({ ...prev, connected: false }));
      };
    } catch {
      // EventSource not supported or offline
    }

    const pollInterval = setInterval(() => {
      pullOnlineSharedState(false);
    }, 12000);

    return () => {
      if (es) es.close();
      clearInterval(pollInterval);
    };
  }, []);

  // 1-Click Online Update Button Handler (Syncs both Online Server & Google Sheets if configured)
  const handleOneClickOnlineUpdate = async () => {
    setOnlineSyncStatus((prev) => ({ ...prev, isSyncing: true }));
    try {
      await pullOnlineSharedState(true);
      await pushOnlineSharedState(undefined, userRole || 'MANUAL_SYNC');
      if (gasWebAppUrl.trim() && !gasWebAppUrl.includes('docs.google.com/spreadsheets')) {
        await syncToGoogleSheetAndLine(tickets);
      }
    } finally {
      setOnlineSyncStatus((prev) => ({
        ...prev,
        isSyncing: false,
        lastSyncTime: new Date().toLocaleTimeString('th-TH')
      }));
    }
  };

  // Export all app data as a portable JSON file for easy sharing/backup
  const exportFullSystemDataJson = () => {
    const payload = {
      exportedAt: new Date().toISOString(),
      companyProfile,
      tickets,
      technicians,
      supervisors,
      masterBrands,
      masterEquipmentTypes,
      masterChecklists
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: 'application/json;charset=utf-8;'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `SKM_Service_Online_Data_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('ดาวน์โหลดไฟล์ข้อมูลระบบ (.json) สำหรับสำรองหรือส่งต่อให้เครื่องอื่นเรียบร้อย', 'success');
  };

  // GitHub index.html Download & Copy Assistant Modal State & Handlers
  const [githubExportModalOpen, setGithubExportModalOpen] = useState(false);
  const [generatedHtmlCode, setGeneratedHtmlCode] = useState('');
  const [copiedHtmlSuccess, setCopiedHtmlSuccess] = useState(false);

  const buildCurrentStandaloneHtml = () => {
    return generateStandaloneHtmlString(
      tickets,
      technicians,
      knowledgeTemplates,
      masterEquipmentTypes,
      masterBrands,
      supPassword,
      servicePin,
      gasWebAppUrl,
      lineChannelToken,
      lineTargetId,
      supervisors
    );
  };

  const handleOpenGithubExportModal = () => {
    const htmlStr = buildCurrentStandaloneHtml();
    setGeneratedHtmlCode(htmlStr);
    setCopiedHtmlSuccess(false);
    setGithubExportModalOpen(true);

    // Trigger automatic download to browser's Downloads folder as well
    downloadStandaloneHtml(
      tickets,
      technicians,
      knowledgeTemplates,
      masterEquipmentTypes,
      masterBrands,
      supPassword,
      servicePin,
      gasWebAppUrl,
      lineChannelToken,
      lineTargetId,
      supervisors
    );
    showToast(
      'ดาวน์โหลดไฟล์ index.html ลงโฟลเดอร์ Downloads แล้ว! (หรือกดคัดลอกโค้ดในหน้าต่างนี้ได้เลย)',
      'success'
    );
  };

  const handleSaveHtmlWithFilePicker = async () => {
    const htmlStr = generatedHtmlCode || buildCurrentStandaloneHtml();
    const anyWin = window as unknown as {
      showSaveFilePicker?: (opts: unknown) => Promise<{
        createWritable: () => Promise<{
          write: (data: string) => Promise<void>;
          close: () => Promise<void>;
        }>;
      }>;
    };
    if (typeof anyWin.showSaveFilePicker === 'function') {
      try {
        const handle = await anyWin.showSaveFilePicker({
          suggestedName: 'index.html',
          types: [
            {
              description: 'HTML File (สำหรับอัปเดต GitHub)',
              accept: { 'text/html': ['.html'] }
            }
          ]
        });
        const writable = await handle.createWritable();
        await writable.write(htmlStr);
        await writable.close();
        showToast('บันทึกไฟล์ index.html ลงในโฟลเดอร์ที่คุณเลือกเรียบร้อยแล้ว!', 'success');
        return;
      } catch (err: unknown) {
        if (err && typeof err === 'object' && 'name' in err && (err as { name: string }).name === 'AbortError') {
          return;
        }
        // Fallback to standard download if iframe blocks showSaveFilePicker
      }
    }
    downloadStandaloneHtml(
      tickets,
      technicians,
      knowledgeTemplates,
      masterEquipmentTypes,
      masterBrands,
      supPassword,
      servicePin,
      gasWebAppUrl,
      lineChannelToken,
      lineTargetId,
      supervisors
    );
    showToast('ดาวน์โหลดไฟล์ index.html ลงในโฟลเดอร์ Downloads (กด Ctrl + J เพื่อดูไฟล์) เรียบร้อย!', 'success');
  };

  const handleCopyStandaloneHtmlToClipboard = async () => {
    const htmlStr = generatedHtmlCode || buildCurrentStandaloneHtml();
    try {
      await navigator.clipboard.writeText(htmlStr);
      setCopiedHtmlSuccess(true);
      showToast('คัดลอกโค้ด index.html ทั้งหมดเรียบร้อย! นำไปกด Ctrl+V วางใน GitHub ได้เลย', 'success');
    } catch {
      const ta = document.createElement('textarea');
      ta.value = htmlStr;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      setCopiedHtmlSuccess(true);
      showToast('คัดลอกโค้ด index.html ทั้งหมดเรียบร้อย! นำไปกด Ctrl+V วางใน GitHub ได้เลย', 'success');
    }
  };

  // Import full app data from a JSON file and broadcast to all online users
  const handleImportFullSystemDataJson = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(String(reader.result || '{}'));
        if (Array.isArray(parsed.tickets)) setTickets(parsed.tickets);
        if (Array.isArray(parsed.technicians)) setTechnicians(parsed.technicians);
        if (Array.isArray(parsed.supervisors)) setSupervisors(parsed.supervisors);
        if (Array.isArray(parsed.masterBrands)) setMasterBrands(parsed.masterBrands);
        if (Array.isArray(parsed.masterEquipmentTypes))
          setMasterEquipmentTypes(parsed.masterEquipmentTypes);
        if (parsed.masterChecklists && typeof parsed.masterChecklists === 'object') {
          setMasterChecklists(parsed.masterChecklists);
        }
        if (parsed.companyProfile && typeof parsed.companyProfile === 'object') {
          setCompanyProfile((prev) => ({ ...prev, ...parsed.companyProfile }));
        }
        await pushOnlineSharedState(parsed.tickets || tickets, 'JSON_IMPORT');
        showToast('นำเข้าข้อมูลและอัปเดตออนไลน์ไปยังทุกเครื่องเรียบร้อยแล้ว!', 'success');
      } catch {
        showToast('ไฟล์ JSON ไม่ถูกต้อง กรุณาเลือกไฟล์ที่ส่งออกจากระบบ SKM Service', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const syncToGoogleSheetAndLine = async (
    updatedTickets: ServiceTicket[],
    lineNotifyText?: string,
    explicitUrl?: string
  ): Promise<{ ok: boolean; message: string }> => {
    // 1. Always sync to built-in Online Server first so all online users see the change immediately!
    pushOnlineSharedState(updatedTickets, userRole || 'CUSTOMER');

    // 2. Sync directly to connected Google Sheets via Google Workspace OAuth if connected & auto-sync enabled
    if (googleSheetId && googleAutoSync) {
      getAccessToken()
        .then((token) => {
          if (token) {
            syncTicketsToSpreadsheet(token, googleSheetId, updatedTickets)
              .then(() => {
                const nowStr = new Date().toLocaleString('th-TH');
                handleSetGoogleLastSyncedAt(nowStr);
              })
              .catch((err) => {
                console.warn('Google Sheets background auto-sync warning:', err);
              });
          }
        })
        .catch(() => {});
    }

    let sentByBackend = false;
    if (lineNotifyText && lineNotifyText.trim()) {
      const lr = await sendLineDirectViaServer(lineNotifyText);
      sentByBackend = lr.ok;
    }

    const targetUrl = (explicitUrl !== undefined ? explicitUrl : gasWebAppUrl).trim();
    if (!targetUrl) {
      return {
        ok: true,
        message: 'อัปเดตข้อมูลออนไลน์ไปยังเซิร์ฟเวอร์กลางเรียบร้อยแล้ว'
      };
    }

    if (
      targetUrl.includes('docs.google.com/spreadsheets') &&
      !targetUrl.includes('script.google.com')
    ) {
      const errMsg =
        'ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet (docs.google.com) — กรุณาใช้ลิงก์ Web App URL จากเมนู ส่วนขยาย > Apps Script > การทำให้ใช้งานได้ (Deploy) ที่ลงท้ายด้วย /exec';
      setSheetSyncStatus({
        ok: false,
        text: errMsg,
        time: new Date().toLocaleTimeString('th-TH')
      });
      return { ok: false, message: errMsg };
    }

    const cleanTickets = sanitizeTicketsForSheet(updatedTickets);
    const nowTime = new Date().toLocaleTimeString('th-TH');

    try {
      const apiRes = await fetch('/api/gsheet/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          gasWebAppUrl: targetUrl,
          tickets: cleanTickets,
          lineNotifyText: sentByBackend ? '' : lineNotifyText || '',
          lineChannelToken: lineChannelToken.trim(),
          lineTargetId: lineTargetId.trim()
        })
      });
      const apiData = await apiRes.json().catch(() => null);
      if (apiRes.ok && apiData?.status === 'ok') {
        const msg = `อัปเดตข้อมูล ${cleanTickets.length} ใบงานลงแท็บ Tickets ใน Google Sheet สำเร็จ`;
        setSheetSyncStatus({ ok: true, text: msg, time: nowTime });
        return { ok: true, message: msg };
      }
      if (apiData?.message) {
        setSheetSyncStatus({ ok: false, text: apiData.message, time: nowTime });
        return { ok: false, message: apiData.message };
      }
    } catch {
      // Fallback to direct browser fetch below
    }

    try {
      const payloadStr = JSON.stringify({
        tickets: cleanTickets,
        lineNotifyText: sentByBackend ? '' : lineNotifyText || '',
        lineChannelToken: lineChannelToken.trim(),
        lineTargetId: lineTargetId.trim()
      });
      await fetch(targetUrl, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: payloadStr
      }).catch(() =>
        fetch(targetUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: payloadStr
        })
      );
      const msg = `ส่งข้อมูล ${cleanTickets.length} ใบงานไปยัง Google Sheet เรียบร้อยแล้ว`;
      setSheetSyncStatus({ ok: true, text: msg, time: nowTime });
      return { ok: true, message: msg };
    } catch {
      const errMsg = 'ไม่สามารถเชื่อมต่อ Google Apps Script URL ได้ กรุณาตรวจสอบลิงก์ /exec';
      setSheetSyncStatus({ ok: false, text: errMsg, time: nowTime });
      return { ok: false, message: errMsg };
    }
  };

  const forceSyncToGoogleSheetNow = async () => {
    if (!gasWebAppUrl.trim()) {
      setGasModalOpen(true);
      showToast(
        'กรุณาวางลิงก์ Google Apps Script Web App URL (ลงท้ายด้วย /exec) ในขั้นตอนที่ 2 เพื่อเชื่อมต่อ Google Sheet',
        'warning'
      );
      return;
    }
    setIsSyncingSheet(true);
    const res = await syncToGoogleSheetAndLine(tickets);
    setIsSyncingSheet(false);
    if (res.ok) {
      showToast(`✅ ${res.message}`, 'success');
    } else {
      showToast(`⚠️ ${res.message}`, 'error');
    }
  };

  const loadTicketsFromGoogleSheet = async (showNotification = true) => {
    if (!gasWebAppUrl.trim()) {
      if (showNotification) {
        showToast('กรุณาวางลิงก์ Google Apps Script Web App URL (/exec) ก่อนดึงข้อมูล', 'warning');
      }
      return;
    }
    if (
      gasWebAppUrl.includes('docs.google.com/spreadsheets') &&
      !gasWebAppUrl.includes('script.google.com')
    ) {
      if (showNotification) {
        showToast(
          'ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet กรุณาใช้ลิงก์ Web App URL จาก Apps Script ที่ลงท้ายด้วย /exec',
          'error'
        );
      }
      return;
    }
    setIsSyncingSheet(true);
    try {
      const res = await fetch(`/api/gsheet/load?url=${encodeURIComponent(gasWebAppUrl.trim())}`);
      const data = await res.json();
      if (res.ok && data && Array.isArray(data.tickets)) {
        if (data.tickets.length > 0) {
          setTickets(data.tickets);
          setSheetSyncStatus({
            ok: true,
            text: `ดึงข้อมูลล่าสุดจาก Google Sheet สำเร็จ (${data.tickets.length} ใบงาน)`,
            time: new Date().toLocaleTimeString('th-TH')
          });
          if (showNotification) {
            showToast(`ดึงข้อมูลจาก Google Sheet สำเร็จ (${data.tickets.length} ใบงาน)`, 'success');
          }
        } else {
          await syncToGoogleSheetAndLine(tickets);
          if (showNotification) {
            showToast(
              `ใน Google Sheet ยังไม่มีใบงาน ระบบจึงส่งข้อมูลปัจจุบัน (${tickets.length} ใบงาน) ขึ้นไปบันทึกให้ทันที!`,
              'success'
            );
          }
        }
      } else if (showNotification) {
        showToast(data?.message || 'ไม่สามารถดึงข้อมูลจาก Google Sheet ได้', 'error');
      }
    } catch {
      if (showNotification) {
        showToast('ไม่สามารถดึงข้อมูลจาก Google Sheet ได้ กรุณาตรวจสอบสิทธิ์ Deploy (Anyone)', 'error');
      }
    } finally {
      setIsSyncingSheet(false);
    }
  };

  const syncAllTicketsUpdateToLine = async () => {
    const msg = formatAllTicketsSummaryLineMessage(tickets);
    showToast(
      `กำลังส่งสรุปข้อมูลใบงานทั้งหมด (${tickets.length} ใบงาน) เข้า LINE (@389fyhbl)...`,
      'success'
    );
    const sentDirect = await sendLineDirectViaServer(msg);
    if (sentDirect.quotaExceeded) {
      try {
        await navigator.clipboard.writeText(msg);
      } catch {
        // ignore
      }
      showToast(
        '⚠️ โควต้าบอท LINE รายเดือนเต็มแล้ว (429) — คัดลอกข้อความสรุปงานลงคลิปบอร์ดแล้ว สามารถวางใน LINE ได้ทันที!',
        'warning'
      );
      if (gasWebAppUrl.trim()) {
        await syncToGoogleSheetAndLine(tickets);
      }
      return;
    }
    if (gasWebAppUrl.trim()) {
      const sheetRes = await syncToGoogleSheetAndLine(tickets, sentDirect.ok ? '' : msg);
      if (sheetRes.ok) {
        showToast(
          `อัปเดตข้อมูลใบงานทั้งหมด (${tickets.length} ใบงาน) เข้า LINE (@389fyhbl) และ Google Sheet สำเร็จ!`,
          'success'
        );
      } else {
        showToast(
          `ส่งเข้า LINE สำเร็จแล้ว แต่ Google Sheet แจ้งว่า: ${sheetRes.message}`,
          'warning'
        );
      }
    } else {
      showToast(
        `ส่งสรุปใบงานทั้งหมดเข้า LINE (@389fyhbl) สำเร็จ! (หากต้องการลง Google Sheet กรุณาตั้งค่า Web App URL)`,
        'success'
      );
    }
  };

  const notifyOverdueTicketsToLine = async () => {
    const overdueList = tickets.filter((t) => isTicketOverdue(t));
    if (overdueList.length === 0) {
      showToast('ไม่มีงานค้างที่เกินกำหนดนัดหมายในขณะนี้ (ทุกใบงานอยู่ในกำหนดเวลา)', 'success');
      return;
    }
    const msg = formatOverdueSummaryLineMessage(overdueList);
    showToast(
      `กำลังส่งแจ้งเตือนงานค้าง (${overdueList.length} ใบงาน) เข้า LINE (@389fyhbl)...`,
      'success'
    );
    const sentDirect = await sendLineDirectViaServer(msg);
    if (sentDirect.quotaExceeded) {
      try {
        await navigator.clipboard.writeText(msg);
      } catch {
        // ignore
      }
      showToast(
        `⚠️ โควต้าบอท LINE รายเดือนเต็มแล้ว (429) — คัดลอกข้อความงานค้าง (${overdueList.length} งาน) ลงคลิปบอร์ดแล้ว สามารถวางใน LINE ได้ทันที!`,
        'warning'
      );
      if (gasWebAppUrl.trim()) {
        await syncToGoogleSheetAndLine(tickets);
      }
      return;
    }
    if (gasWebAppUrl.trim()) {
      await syncToGoogleSheetAndLine(tickets, sentDirect.ok ? '' : msg);
    }
    showToast(
      `ส่งแจ้งเตือนงานค้างเกินกำหนดนัดหมาย (${overdueList.length} ใบงาน) เข้า LINE (@389fyhbl) เรียบร้อยแล้ว!`,
      'success'
    );
  };

  const shareTicketToLine = async (t: ServiceTicket, customHeading?: string) => {
    const msg = formatTicketLineMessage(
      t,
      customHeading || '📢 [แจ้งเตือนใบงาน - บริษัท สหคิมมอเตอร์ จำกัด]'
    );
    const sentDirect = await sendLineDirectViaServer(msg);
    if (sentDirect.quotaExceeded) {
      try {
        await navigator.clipboard.writeText(msg);
      } catch {
        // ignore
      }
      showToast(
        `⚠️ โควต้าบอท LINE รายเดือนเต็มแล้ว (429) — ระบบคัดลอกข้อความใบงาน ${t.id} ลงคลิปบอร์ดให้แล้ว สามารถกดวาง (Paste) ส่งใน LINE ได้ทันที!`,
        'warning'
      );
      if (gasWebAppUrl.trim()) {
        await syncToGoogleSheetAndLine(tickets);
      }
      return;
    }
    if (sentDirect.ok) {
      showToast(
        `ส่งแจ้งเตือนใบงาน ${t.id} ผ่าน LINE bot (@389fyhbl) เรียบร้อยแล้ว!`,
        'success'
      );
    } else {
      try {
        await navigator.clipboard.writeText(msg);
      } catch {
        // ignore
      }
      showToast(
        `คัดลอกข้อความใบงาน ${t.id} ลงคลิปบอร์ดแล้ว สามารถนำไปวางส่งใน LINE ได้ทันที`,
        'warning'
      );
    }
    if (gasWebAppUrl.trim()) {
      await syncToGoogleSheetAndLine(tickets, sentDirect.ok ? '' : msg);
    }
  };

  // Toast notifications
  const [toastMessage, setToastMessage] = useState<{
    text: string;
    type: 'success' | 'warning' | 'error';
  } | null>(null);

  const showToast = (text: string, type: 'success' | 'warning' | 'error' = 'success') => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage((prev) => (prev?.text === text ? null : prev));
    }, 3500);
  };

  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);

  const downloadElementAsPdf = async (elementId: string, filename: string) => {
    const sourceEl = document.getElementById(elementId);
    if (!sourceEl) {
      showToast('ไม่พบส่วนเอกสารที่ต้องการบันทึกเป็น PDF', 'error');
      return;
    }

    setIsGeneratingPdf(true);
    showToast('กำลังสร้างไฟล์ PDF กรุณารอสักครู่...', 'warning');

    const restoredStyles: Array<{ el: HTMLElement | SVGElement; cssText: string }> = [];
    const darkThemeEls: Element[] = [];

    try {
      // Ensure html2pdf is loaded
      let html2pdfLib = (window as unknown as { html2pdf?: any }).html2pdf;
      if (!html2pdfLib) {
        await new Promise<void>((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
          script.onload = () => resolve();
          script.onerror = () => reject(new Error('Failed to load html2pdf.js'));
          document.head.appendChild(script);
        });
        html2pdfLib = (window as unknown as { html2pdf?: any }).html2pdf;
      }

      // Helper to convert Tailwind v4 oklch/oklab/color-mix/color() into standard rgb()/rgba() for html2canvas
      const cv = document.createElement('canvas');
      cv.width = 1;
      cv.height = 1;
      const ctx = cv.getContext('2d', { willReadFrequently: true });
      const toRgbColor = (cssColor: string): string => {
        if (!ctx) return '#1e293b';
        try {
          ctx.clearRect(0, 0, 1, 1);
          ctx.fillStyle = '#0f172a';
          ctx.fillStyle = cssColor.trim();
          ctx.fillRect(0, 0, 1, 1);
          const d = ctx.getImageData(0, 0, 1, 1).data;
          return d[3] < 255
            ? `rgba(${d[0]}, ${d[1]}, ${d[2]}, ${(d[3] / 255).toFixed(2)})`
            : `rgb(${d[0]}, ${d[1]}, ${d[2]})`;
        } catch {
          return '#1e293b';
        }
      };

      const hasModernColorFn = (val: string): boolean =>
        /oklch|oklab|color-mix|lch\(|lab\(|color\(/i.test(val);

      const sanitizeModernColors = (cssText: string): string => {
        if (!cssText || !hasModernColorFn(cssText)) return cssText;
        let out = cssText
          // Disable @supports (color: color-mix(...)) blocks cleanly without breaking balanced parentheses
          .replace(
            /@supports\s*\(\s*color\s*:\s*color-mix\([^)]*\)\s*\)/gi,
            '@supports (color: never-match-pdf)'
          )
          // Remove modern gradient interpolation spaces unsupported by html2canvas
          .replace(/\s+in\s+(?:oklab|oklch|srgb|srgb-linear|lab|lch|hsl|hwb)\b/gi, '');

        // Iteratively replace inner-most oklch(...), oklab(...), color-mix(...), color(...)
        for (let pass = 0; pass < 3; pass++) {
          out = out
            .replace(/oklch\([^()]+\)/gi, (m) => toRgbColor(m))
            .replace(/oklab\([^()]+\)/gi, (m) => toRgbColor(m))
            .replace(/color\([^()]+\)/gi, (m) => toRgbColor(m))
            .replace(/color-mix\([^()]+\)/gi, (m) => toRgbColor(m));
        }
        return out;
      };

      const COLOR_PROPS = [
        'color',
        'background-color',
        'border-top-color',
        'border-right-color',
        'border-bottom-color',
        'border-left-color',
        'outline-color',
        'text-decoration-color',
        '-webkit-text-stroke-color',
        'column-rule-color',
        'fill',
        'stroke',
        'stop-color',
        'flood-color',
        'lighting-color',
        'box-shadow',
        'text-shadow',
        'background-image'
      ];

      const sanitizeDocumentStylesheets = (doc: Document) => {
        try {
          const extraCss: string[] = [];
          Array.from(doc.styleSheets).forEach((sheet) => {
            try {
              const owner = sheet.ownerNode as Element | null;
              if (owner && owner.tagName === 'LINK') {
                const rules = sheet.cssRules;
                if (rules) {
                  for (let i = 0; i < rules.length; i++) {
                    extraCss.push(rules[i].cssText);
                  }
                }
              }
            } catch {
              // Ignore cross-origin font stylesheet
            }
          });

          if (extraCss.length > 0) {
            doc.querySelectorAll('link[rel="stylesheet"]').forEach((lnk) => {
              const href = lnk.getAttribute('href') || '';
              if (!href.includes('fonts.googleapis.com') && !href.includes('font-awesome')) {
                lnk.remove();
              }
            });
            const injectedStyle = doc.createElement('style');
            injectedStyle.setAttribute('data-pdf-sanitized', 'true');
            injectedStyle.textContent = sanitizeModernColors(extraCss.join('\n'));
            doc.head.appendChild(injectedStyle);
          }
        } catch {
          // Ignore stylesheet inspection error
        }

        doc.querySelectorAll('style').forEach((styleEl) => {
          if (styleEl.textContent && hasModernColorFn(styleEl.textContent)) {
            styleEl.textContent = sanitizeModernColors(styleEl.textContent);
          }
        });
      };

      const inlineSafeComputedColors = (
        doc: Document,
        elements: Array<Element>,
        recordRestore = false
      ) => {
        const win = doc.defaultView || window;
        elements.forEach((node) => {
          if (!(node instanceof HTMLElement || node instanceof SVGElement)) return;
          const cs = win.getComputedStyle(node);
          let savedBefore = false;

          COLOR_PROPS.forEach((prop) => {
            const val = cs.getPropertyValue(prop);
            if (val && hasModernColorFn(val)) {
              if (recordRestore && !savedBefore) {
                restoredStyles.push({ el: node, cssText: node.style.cssText });
                savedBefore = true;
              }
              const cleaned = sanitizeModernColors(val);
              if (hasModernColorFn(cleaned)) {
                node.style.setProperty(
                  prop,
                  prop.includes('shadow') || prop === 'background-image' ? 'none' : '#0f172a',
                  'important'
                );
              } else {
                node.style.setProperty(prop, cleaned, 'important');
              }
            }
          });
        });
      };

      // 1. Sanitize live document stylesheets & temporarily force light mode before html2pdf clones sourceEl
      document.querySelectorAll('.dark-theme').forEach((el) => {
        darkThemeEls.push(el);
        el.classList.remove('dark-theme');
      });
      sanitizeDocumentStylesheets(document);

      const ancestors: Element[] = [];
      let curr: Element | null = sourceEl.parentElement;
      while (curr) {
        ancestors.push(curr);
        curr = curr.parentElement;
      }
      const liveNodesToSanitize = [
        ...ancestors,
        sourceEl,
        ...Array.from(sourceEl.querySelectorAll('*'))
      ];
      inlineSafeComputedColors(document, liveNodesToSanitize, true);

      const opt = {
        margin: [8, 8, 8, 8],
        filename: filename.endsWith('.pdf') ? filename : `${filename}.pdf`,
        image: { type: 'jpeg', quality: 0.96 },
        html2canvas: {
          scale: 2,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false,
          onclone: (clonedDoc: Document) => {
            clonedDoc.querySelectorAll('.dark-theme').forEach((el) => {
              el.classList.remove('dark-theme');
            });
            sanitizeDocumentStylesheets(clonedDoc);

            clonedDoc.querySelectorAll('.no-print').forEach((el) => el.remove());
            clonedDoc.querySelectorAll(`[id="${elementId}"]`).forEach((el) => {
              if (el instanceof HTMLElement) {
                el.style.maxHeight = 'none';
                el.style.height = 'auto';
                el.style.overflow = 'visible';
                el.style.boxShadow = 'none';
                el.style.margin = '0 auto';
              }
            });

            const allClonedNodes = [
              clonedDoc.documentElement,
              clonedDoc.body,
              ...Array.from(clonedDoc.querySelectorAll('*'))
            ].filter(Boolean) as Element[];
            inlineSafeComputedColors(clonedDoc, allClonedNodes, false);
          }
        },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
      };

      await html2pdfLib().set(opt).from(sourceEl).save();
      showToast(`ดาวน์โหลดไฟล์ ${filename} สำเร็จแล้ว!`, 'success');
    } catch (err) {
      console.error('PDF export error:', err);
      showToast('ระบบกำลังเปิดหน้าต่างพิมพ์เพื่อบันทึกเป็น PDF แทน', 'warning');
      window.print();
    } finally {
      for (let i = restoredStyles.length - 1; i >= 0; i--) {
        restoredStyles[i].el.style.cssText = restoredStyles[i].cssText;
      }
      darkThemeEls.forEach((el) => el.classList.add('dark-theme'));
      setIsGeneratingPdf(false);
    }
  };

  // Screen Theme Mode (Light = ขาว / Dark = ดำ) & Role-Based Color Identity
  const [themeMode, setThemeMode] = useState<'light' | 'dark'>(() => {
    try {
      const saved = localStorage.getItem('sahakim_theme_mode');
      return saved === 'dark' ? 'dark' : 'light';
    } catch {
      return 'light';
    }
  });
  const [roleColorIdentityEnabled, setRoleColorIdentityEnabled] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('sahakim_role_color');
      return saved === null ? true : saved === 'true';
    } catch {
      return true;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('sahakim_theme_mode', themeMode);
    } catch {
      // Ignore storage error
    }
  }, [themeMode]);

  useEffect(() => {
    try {
      localStorage.setItem('sahakim_role_color', String(roleColorIdentityEnabled));
    } catch {
      // Ignore storage error
    }
  }, [roleColorIdentityEnabled]);

  // Internal Login State
  const [loginRole, setLoginRole] = useState<'SERVICE TEAM' | 'TECHNICIAN' | 'SUPERVISOR'>('SUPERVISOR');
  const [loginPin, setLoginPin] = useState('');

  const roleThemeConfig = useMemo(() => {
    if (!roleColorIdentityEnabled) {
      return {
        headerBg: 'bg-slate-900 border-slate-800',
        bannerBg: 'bg-slate-900 border-l-4 border-orange-500',
        badgeBg: 'bg-orange-500 text-white',
        roleLabel: userRole || 'SAHA KIM MOTOR',
        roleColorName: 'สีมาตรฐานสหคิม (Slate & Orange)'
      };
    }
    if (currentView === 'customer_portal') {
      return {
        headerBg: 'bg-gradient-to-r from-slate-950 via-orange-950 to-slate-900 border-orange-500/60',
        bannerBg: 'bg-gradient-to-r from-orange-950 via-slate-900 to-amber-950 border-l-8 border-orange-500',
        badgeBg: 'bg-orange-500 text-white',
        roleLabel: 'CUSTOMER · SERVICE REQUEST',
        roleColorName: 'โทนสีส้ม-ทอง (Customer Brand Theme)'
      };
    }
    if (currentView === 'internal_dashboard' || currentView === 'internal_login') {
      const activeRole = currentView === 'internal_login' ? loginRole : userRole;
      if (activeRole === 'SERVICE TEAM') {
        return {
          headerBg: 'bg-gradient-to-r from-blue-950 via-slate-900 to-sky-950 border-blue-500/60',
          bannerBg: 'bg-gradient-to-r from-blue-950 via-slate-900 to-sky-950 border-l-8 border-blue-400',
          badgeBg: 'bg-blue-600 text-white',
          roleLabel: 'SERVICE SUPPORT (ฝ่ายประสานงานและจัดคิวช่าง)',
          roleColorName: 'โทนสีน้ำเงิน-ฟ้า (Blue Dispatch Center)'
        };
      }
      if (activeRole === 'TECHNICIAN') {
        return {
          headerBg: 'bg-gradient-to-r from-teal-950 via-slate-900 to-emerald-950 border-teal-500/60',
          bannerBg: 'bg-gradient-to-r from-teal-950 via-slate-900 to-emerald-950 border-l-8 border-teal-400',
          badgeBg: 'bg-teal-600 text-white',
          roleLabel: 'TECHNICIAN (ช่างเทคนิคภาคสนาม)',
          roleColorName: 'โทนสีเขียวมรกต-ทีล (Teal Field Engineering)'
        };
      }
      if (activeRole === 'SUPERVISOR') {
        return {
          headerBg: 'bg-gradient-to-r from-indigo-950 via-slate-900 to-purple-950 border-indigo-500/60',
          bannerBg: 'bg-gradient-to-r from-indigo-950 via-slate-900 to-purple-950 border-l-8 border-indigo-400',
          badgeBg: 'bg-indigo-600 text-white',
          roleLabel: 'SUPERVISOR / ADMIN (หัวหน้าวิศวกรและผู้อนุมัติ)',
          roleColorName: 'โทนสีม่วงคราม-ทอง (Indigo Executive Approval)'
        };
      }
    }
    return {
      headerBg: 'bg-slate-900 border-slate-800',
      bannerBg: 'bg-slate-900 border-l-4 border-orange-500',
      badgeBg: 'bg-orange-500 text-white',
      roleLabel: 'SAHA KIM MOTOR PORTAL',
      roleColorName: 'โทนสีสหคิมมอเตอร์ (Slate & Orange)'
    };
  }, [roleColorIdentityEnabled, currentView, userRole, loginRole]);

  // Customer Portal & Internal Portal Menu Tabs State
  const [custTab, setCustTab] = useState<'list' | 'new' | 'schedule'>('list');
  const [internalTab, setInternalTab] = useState<
    'tickets' | 'new_ticket' | 'schedule' | 'dashboard' | 'all'
  >('tickets');
  const [pendingInternalTab, setPendingInternalTab] = useState<
    'tickets' | 'new_ticket' | 'schedule' | 'dashboard' | 'all' | null
  >(null);
  const [companyPickerOpen, setCompanyPickerOpen] = useState(false);
  const [companyInputDraft, setCompanyInputDraft] = useState(customerCompany);
  const [customerLookupQuery, setCustomerLookupQuery] = useState(customerCompany);
  const [scheduleDayOffset, setScheduleDayOffset] = useState(0);
  const [supportScheduleOffset, setSupportScheduleOffset] = useState(0);
  const [supportSelectedDate, setSupportSelectedDate] = useState(() =>
    new Date().toISOString().slice(0, 10)
  );

  // New Ticket Form Fields
  const [newCompany, setNewCompany] = useState(customerCompany);
  const [newTaxId, setNewTaxId] = useState('0105552012345');
  const [newDocNo, setNewDocNo] = useState('');
  const [newContact, setNewContact] = useState('');
  const [newWarranty, setNewWarranty] = useState('In-Warranty (อยู่ในประกัน - ฟรี / FOC)');
  const [newServiceType, setNewServiceType] = useState('ซ่อมหน้างาน');
  const [newUrgency, setNewUrgency] = useState('ปกติ (Normal)');
  const [newLocation, setNewLocation] = useState('');
  const [newDesiredDate, setNewDesiredDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [newUploadedPhoto, setNewUploadedPhoto] = useState<string>('');
  const [tempEquipmentList, setTempEquipmentList] = useState<NewEquipmentDraft[]>([
    {
      type: 'END SUCTION PUMP',
      customType: '',
      brand: 'STAC',
      customBrand: '',
      model: '',
      tag: '',
      symptom: '',
      detail: '',
      beforePhotos: [],
      afterPhotos: []
    }
  ]);

  // Reset demo tickets modal state
  const [resetConfirmModalOpen, setResetConfirmModalOpen] = useState(false);

  // Internal Dashboard Filter State
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');

  // Modals State
  const [activeReportTicketId, setActiveReportTicketId] = useState<string | null>(null);
  const [editingTicket, setEditingTicket] = useState<ServiceTicket | null>(null);
  const [techManagerOpen, setTechManagerOpen] = useState(false);
  const [userManualOpen, setUserManualOpen] = useState(false);
  const [newTechName, setNewTechName] = useState('');
  const [newTechPosition, setNewTechPosition] = useState('ช่างเทคนิค (Technician)');
  const [newTechSpec, setNewTechSpec] = useState('');
  const [newTechPin, setNewTechPin] = useState('');
  const [newSupervisorName, setNewSupervisorName] = useState('');

  // AI Analysis Loading & Results State per equipment index inside Action Modal
  const [aiLoadingByIdx, setAiLoadingByIdx] = useState<Record<number, boolean>>({});
  const [aiSuggestionsByIdx, setAiSuggestionsByIdx] = useState<
    Record<
      number,
      {
        recommendedChecklist: string[];
        recommendedParts: Array<{ name: string; qty: number; unit: string }>;
      }
    >
  >({});

  // Custom Equipment/Brand inline edit inside Action Modal
  const [modalCustomTypeByIdx, setModalCustomTypeByIdx] = useState<Record<number, boolean>>({});
  const [modalCustomBrandByIdx, setModalCustomBrandByIdx] = useState<Record<number, boolean>>({});

  // Save custom knowledge template modal
  const [saveTemplateForIdx, setSaveTemplateForIdx] = useState<number | null>(null);
  const [newTemplateTitle, setNewTemplateTitle] = useState('');

  // Workflow Interactive Dialogs
  const [assignModalTicket, setAssignModalTicket] = useState<ServiceTicket | null>(null);
  const [assignDateInput, setAssignDateInput] = useState('');
  const [assignTechInput, setAssignTechInput] = useState('');
  const [assignLeadTechInput, setAssignLeadTechInput] = useState('');

  const [approveReportTicket, setApproveReportTicket] = useState<ServiceTicket | null>(null);
  const [approveMode, setApproveMode] = useState<'REPORT' | 'PARTS'>('REPORT');
  const [supervisorSignName, setSupervisorSignName] = useState('หัวหน้าวิศวกรบริการ (Supervisor)');
  const [customSupervisorInput, setCustomSupervisorInput] = useState('');

  const [rejectModalTicket, setRejectModalTicket] = useState<ServiceTicket | null>(null);
  const [rejectReasonInput, setRejectReasonInput] = useState('');

  const [unlockModalTicket, setUnlockModalTicket] = useState<ServiceTicket | null>(null);
  const [unlockPassInput, setUnlockPassInput] = useState('');

  // Mobile Signature Capture States (for Service Report & Customer/Technician Sign-Off Modal)
  const [reportSignaturePadOpen, setReportSignaturePadOpen] = useState(false);
  const [customerSignOffModalTicket, setCustomerSignOffModalTicket] = useState<ServiceTicket | null>(null);
  const [technicianSignModalTicket, setTechnicianSignModalTicket] = useState<ServiceTicket | null>(null);

  // New part draft per equipment row inside Action Modal
  const [partDrafts, setPartDrafts] = useState<
    Record<
      number,
      {
        name: string;
        qty: string;
        unit?: string;
        price: string;
        billingType?: 'WARRANTY_INTERNAL' | 'CUSTOMER_CHARGE';
      }
    >
  >({});

  // Master Data Manager Modal & Sub-tab States
  const [masterDataModalOpen, setMasterDataModalOpen] = useState(false);
  const [masterDataTab, setMasterDataTab] = useState<
    'brands' | 'equipment' | 'checklists' | 'company' | 'staff'
  >('brands');
  const [newMasterBrandInput, setNewMasterBrandInput] = useState('');
  const [newMasterEquipInput, setNewMasterEquipInput] = useState('');
  const [selectedChecklistCategory, setSelectedChecklistCategory] = useState<string>('PUMP');
  const [newChecklistCategoryInput, setNewChecklistCategoryInput] = useState('');
  const [newChecklistItemInput, setNewChecklistItemInput] = useState('');

  // Toggle whether to show internal cost numbers on the printable report (default true for clear breakdown, can hide for customer copy)
  const [showInternalCostInReport, setShowInternalCostInReport] = useState<boolean>(true);

  const urgentCount = useMemo(
    () =>
      tickets.filter(
        (t) =>
          t.status === 'NEW' ||
          t.status === 'PENDING_SUPERVISOR_APPROVAL' ||
          t.status === 'PENDING_PARTS_APPROVAL'
      ).length,
    [tickets]
  );

  const distinctCompanies = useMemo(() => {
    const set = new Set<string>(['บริษัท สยามอุตสาหกรรม จำกัด']);
    tickets.forEach((t) => {
      if (t.company) set.add(t.company);
    });
    return Array.from(set);
  }, [tickets]);

  const customerMatchedTickets = useMemo(() => {
    const q = (customerLookupQuery || customerCompany || '').trim().toLowerCase();
    if (!q) return [];
    return tickets.filter(
      (t) =>
        t.id.toLowerCase().includes(q) ||
        t.company.toLowerCase().includes(q)
    );
  }, [tickets, customerLookupQuery, customerCompany]);

  const buildScheduleDays = (baseOffset: number) => {
    const thaiDays = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
    const thaiMonths = [
      'ม.ค.',
      'ก.พ.',
      'มี.ค.',
      'เม.ย.',
      'พ.ค.',
      'มิ.ย.',
      'ก.ค.',
      'ส.ค.',
      'ก.ย.',
      'ต.ค.',
      'พ.ย.',
      'ธ.ค.'
    ];
    const baseToday = new Date();
    baseToday.setHours(0, 0, 0, 0);

    return Array.from({ length: 14 }, (_, idx) => {
      const offset = baseOffset + idx;
      const d = new Date(baseToday);
      d.setDate(baseToday.getDate() + offset);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      const dateStr = `${yyyy}-${mm}-${dd}`;
      const dayLabel = `${thaiDays[d.getDay()]} ${d.getDate()} ${thaiMonths[d.getMonth()]}`;

      const ticketsOnDay = tickets.filter(
        (t) => t.status !== 'CLOSE' && extractTicketDueDate(t) === dateStr
      );

      // Total active jobs on this date (based strictly on real active tickets, green when empty)
      const totalJobs = ticketsOnDay.length;
      const level: 'RED' | 'YELLOW' | 'GREEN' =
        totalJobs >= 3 ? 'RED' : totalJobs >= 1 ? 'YELLOW' : 'GREEN';

      const techAvailability = technicians.map((tc) => {
        const assignedTickets = ticketsOnDay.filter((t) => t.assignedTech === tc.name);
        const jobCount = assignedTickets.length;
        const techLevel: 'RED' | 'YELLOW' | 'GREEN' =
          jobCount >= 2 ? 'RED' : jobCount === 1 ? 'YELLOW' : 'GREEN';
        return {
          ...tc,
          jobCount,
          assignedTickets,
          techLevel,
          availabilityLabel:
            techLevel === 'RED'
              ? `คิวแน่น (${jobCount} งาน)`
              : techLevel === 'YELLOW'
              ? `รับงานเพิ่มได้ (มี ${jobCount} งาน)`
              : 'ว่าง (พร้อมรับงาน)'
        };
      });

      const freeTechCount = techAvailability.filter((tc) => tc.jobCount === 0).length;

      return {
        dateStr,
        dayLabel,
        offset,
        totalJobs,
        realTicketCount: ticketsOnDay.length,
        ticketsOnDay,
        techAvailability,
        freeTechCount,
        level,
        statusText:
          level === 'RED'
            ? 'เต็มมาก'
            : level === 'YELLOW'
            ? 'รับงานได้'
            : 'ว่าง',
        detailText:
          level === 'RED'
            ? `คิวแน่น (${totalJobs} งาน)`
            : level === 'YELLOW'
            ? `รับงานได้ (${totalJobs} งาน)`
            : 'คิวช่างว่าง'
      };
    });
  };

  const customerScheduleDays = useMemo(
    () => buildScheduleDays(scheduleDayOffset),
    [tickets, technicians, scheduleDayOffset]
  );

  const supportScheduleDays = useMemo(
    () => buildScheduleDays(supportScheduleOffset),
    [tickets, technicians, supportScheduleOffset]
  );

  const getScheduleInfoForDate = (targetDateStr: string) => {
    const existing =
      supportScheduleDays.find((d) => d.dateStr === targetDateStr) ||
      customerScheduleDays.find((d) => d.dateStr === targetDateStr);
    if (existing) return existing;

    const ticketsOnDay = tickets.filter(
      (t) => t.status !== 'CLOSE' && extractTicketDueDate(t) === targetDateStr
    );
    const totalJobs = ticketsOnDay.length;
    const level: 'RED' | 'YELLOW' | 'GREEN' =
      totalJobs >= 3 ? 'RED' : totalJobs >= 1 ? 'YELLOW' : 'GREEN';
    const techAvailability = technicians.map((tc) => {
      const assignedTickets = ticketsOnDay.filter((t) => t.assignedTech === tc.name);
      const jobCount = assignedTickets.length;
      const techLevel: 'RED' | 'YELLOW' | 'GREEN' =
        jobCount >= 2 ? 'RED' : jobCount === 1 ? 'YELLOW' : 'GREEN';
      return {
        ...tc,
        jobCount,
        assignedTickets,
        techLevel,
        availabilityLabel:
          techLevel === 'RED'
            ? `คิวแน่น (${jobCount} งาน)`
            : techLevel === 'YELLOW'
            ? `รับงานเพิ่มได้ (มี ${jobCount} งาน)`
            : 'ว่าง (พร้อมรับงาน)'
      };
    });
    return {
      dateStr: targetDateStr,
      dayLabel: targetDateStr,
      offset: 0,
      totalJobs,
      realTicketCount: ticketsOnDay.length,
      ticketsOnDay,
      techAvailability,
      freeTechCount: techAvailability.filter((tc) => tc.jobCount === 0).length,
      level,
      statusText: level === 'RED' ? 'เต็มมาก' : level === 'YELLOW' ? 'รับงานได้' : 'ว่าง',
      detailText:
        level === 'RED'
          ? `คิวแน่น (${totalJobs} งาน)`
          : level === 'YELLOW'
          ? `รับงานได้ (${totalJobs} งาน)`
          : 'คิวช่างว่าง'
    };
  };

  const applyCustomerLookup = (rawInput: string) => {
    const trimmed = rawInput.trim();
    if (!trimmed) return;
    setCustomerLookupQuery(trimmed);
    const matchedByTicketId = tickets.find(
      (t) =>
        t.id.toLowerCase() === trimmed.toLowerCase() ||
        t.id.toLowerCase().includes(trimmed.toLowerCase())
    );
    if (matchedByTicketId) {
      setCustomerCompany(matchedByTicketId.company);
      setNewCompany(matchedByTicketId.company);
    } else {
      const matchedByCompany = tickets.find((t) =>
        t.company.toLowerCase().includes(trimmed.toLowerCase())
      );
      const resolvedCompany = matchedByCompany ? matchedByCompany.company : trimmed;
      setCustomerCompany(resolvedCompany);
      setNewCompany(resolvedCompany);
    }
  };

  const openCustomerPortalPrompt = (targetTab?: 'list' | 'new' | 'schedule') => {
    if (targetTab) {
      setCustTab(targetTab);
    }
    setCompanyInputDraft(customerLookupQuery || customerCompany || 'บริษัท สยามอุตสาหกรรม จำกัด');
    setCompanyPickerOpen(true);
  };

  const openInternalSection = (
    targetTab: 'tickets' | 'new_ticket' | 'schedule' | 'dashboard' | 'all'
  ) => {
    setInternalTab(targetTab);
    if (userRole) {
      setCurrentView('internal_dashboard');
    } else {
      setPendingInternalTab(targetTab);
      setCurrentView('internal_login');
    }
  };

  const confirmCustomerPortalCompany = (e: React.FormEvent) => {
    e.preventDefault();
    if (!companyInputDraft.trim()) return;
    applyCustomerLookup(companyInputDraft);
    setCompanyPickerOpen(false);
    setCurrentView('customer_portal');
  };

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginRole === 'SUPERVISOR') {
      if (loginPin !== supPassword && loginPin !== 'admin123') {
        showToast('รหัสผ่าน Supervisor (Admin) ไม่ถูกต้อง (รหัสเริ่มต้นคือ admin123)', 'error');
        return;
      }
      if (loginPin === 'admin123' && supPassword !== 'admin123') {
        setSupPassword('admin123');
        localStorage.setItem('sahakim_sup_pass', 'admin123');
      }
    }
    if (
      loginRole === 'SERVICE TEAM' &&
      loginPin !== servicePin &&
      loginPin !== supPassword &&
      loginPin !== 'admin123' &&
      loginPin !== '1234'
    ) {
      showToast('รหัสผ่านเจ้าหน้าที่รับเรื่อง (Service Support) ไม่ถูกต้อง', 'error');
      return;
    }
    if (loginRole === 'TECHNICIAN') {
      let matchedTech: Technician | undefined = undefined;
      if (selectedTechNameForLogin) {
        matchedTech = technicians.find((tc) => tc.name === selectedTechNameForLogin);
      }
      if (!matchedTech) {
        matchedTech = technicians.find((tc) => tc.pin === loginPin);
      }

      const isMaster = loginPin === supPassword || loginPin === 'admin123';
      if (!matchedTech && !isMaster) {
        showToast('รหัสพนักงานช่าง (PIN) หรือชื่อช่างไม่ถูกต้อง กรุณาเลือกชื่อช่างหรือตรวจสอบรหัส PIN', 'error');
        return;
      }
      if (matchedTech && matchedTech.pin !== loginPin && !isMaster) {
        showToast(`รหัส PIN ของ ${matchedTech.name} ไม่ถูกต้อง กรุณากรอกรหัส PIN ให้ถูกต้อง`, 'error');
        return;
      }

      const finalTech = matchedTech || technicians[0];
      if (finalTech) {
        setLoggedInTech(finalTech);
        localStorage.setItem('sahakim_logged_in_tech', JSON.stringify(finalTech));
      }
      setUserRole('TECHNICIAN');
      setLoginPin('');
      setInternalTab('tickets');
      setTechWorkFilter('MY_JOBS');
      setCurrentView('internal_dashboard');
      if (urgentCount > 0) {
        playNotificationChime(soundEnabled);
      }
      showToast(
        `ยินดีต้อนรับ ${finalTech ? finalTech.name : 'ช่างเทคนิค'} (${finalTech?.position || 'TECHNICIAN'}) เข้าสู่ระบบ`,
        'success'
      );
      return;
    }
    setUserRole(loginRole);
    setLoginPin('');
    if (pendingInternalTab) {
      setInternalTab(pendingInternalTab);
      setPendingInternalTab(null);
    } else if (loginRole === 'SERVICE TEAM') {
      setInternalTab('schedule');
    } else if (loginRole === 'SUPERVISOR') {
      setInternalTab('dashboard');
    } else {
      setInternalTab('tickets');
    }
    setCurrentView('internal_dashboard');
    if (urgentCount > 0) {
      playNotificationChime(soundEnabled);
    }
    showToast(`เข้าสู่ระบบในฐานะ ${loginRole} สำเร็จ`, 'success');
  };

  const exportToExcel = () => {
    let csv =
      '\uFEFFTicketID,Company,TaxID,DocNo,Contact,SiteContact,Location,ServiceType,Warranty,Urgency,Status,AssignedTech,EquipmentSummary,CreatedAt\n';
    tickets.forEach((t) => {
      const equipSummary = t.equipmentList
        .map((eq) => `${eq.type} (${eq.brand} ${eq.model}): ${eq.symptom}`)
        .join(' | ');
      csv += `"${t.id}","${t.company}","${t.taxId || ''}","${t.docNo || ''}","${t.contact}","${
        t.siteContact || ''
      }","${t.location || ''}","${t.serviceType}","${t.warranty}","${t.urgency}","${t.status}","${
        t.assignedTech || ''
      }","${equipSummary.replace(/"/g, '""')}","${t.createdAt}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `SahaKim_Service_Report_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('ส่งออกไฟล์ CSV สำหรับ Excel เรียบร้อยแล้ว', 'success');
  };

  const openAssignOrRescheduleModal = (
    t: ServiceTicket,
    preferredDate?: string,
    preferredTech?: string
  ) => {
    if (t.isLocked && userRole !== 'SUPERVISOR') {
      showToast('ใบงานนี้ถูกล็อกแล้ว ต้องให้ Supervisor ปลดล็อกเท่านั้น', 'warning');
      return;
    }
    const defaultLead =
      t.leadTech ||
      technicians.find((tc) => tc.position?.includes('หัวหน้าช่าง'))?.name ||
      '';
    const targetDate =
      preferredDate ||
      extractTicketDueDate(t) ||
      t.desiredDate ||
      new Date().toISOString().slice(0, 10);
    const dateInfo = getScheduleInfoForDate(targetDate);
    const firstFreeTech = dateInfo.techAvailability.find((tc) => tc.jobCount === 0)?.name;

    setAssignModalTicket(t);
    setAssignDateInput(targetDate);
    setAssignTechInput(
      preferredTech || t.assignedTech || firstFreeTech || technicians[0]?.name || ''
    );
    setAssignLeadTechInput(defaultLead);
  };

  const supportAcceptOnBehalfOfTech = (
    t: ServiceTicket,
    overrideTech?: string,
    overrideDate?: string
  ) => {
    if (t.isLocked && userRole !== 'SUPERVISOR') {
      showToast('ใบงานนี้ถูกล็อกแล้ว ต้องให้ Supervisor ปลดล็อกเท่านั้น', 'warning');
      return;
    }
    const dateToUse =
      overrideDate ||
      extractTicketDueDate(t) ||
      t.desiredDate ||
      new Date().toISOString().slice(0, 10);
    const dateInfo = getScheduleInfoForDate(dateToUse);
    const chosenTech =
      overrideTech ||
      t.assignedTech ||
      dateInfo.techAvailability.find((tc) => tc.jobCount === 0)?.name ||
      technicians[0]?.name ||
      'ช่างเทคนิค';
    const chosenLead =
      t.leadTech ||
      technicians.find((tc) => tc.position?.includes('หัวหน้าช่าง'))?.name ||
      '';

    const updatedList = tickets.map((item) =>
      item.id === t.id
        ? {
            ...item,
            confirmedDate: `${dateToUse} (Service Support รับงานแทนช่าง ${chosenTech} แล้ว)`,
            assignedTech: chosenTech,
            leadTech: chosenLead,
            status: 'REPAIRING' as TicketStatus
          }
        : item
    );
    setTickets(updatedList);
    const targetTicket = updatedList.find((item) => item.id === t.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(
          targetTicket,
          `✅ [Service Support รับงานแทนช่าง (${chosenTech}) เรียบร้อยแล้ว - เริ่มดำเนินการ]`
        )
      );
    }
    showToast(
      `Service Support กดรับงาน ${t.id} แทนช่าง (${chosenTech}) ในวันที่ ${dateToUse} เรียบร้อยแล้ว!`,
      'success'
    );
  };

  const techAcceptJob = (t: ServiceTicket) => {
    const dateToUse =
      extractTicketDueDate(t) || t.desiredDate || new Date().toISOString().slice(0, 10);
    const techName = loggedInTech?.name || t.assignedTech || technicians[0]?.name || 'ช่างเทคนิค';
    const updatedList = tickets.map((item) =>
      item.id === t.id
        ? {
            ...item,
            confirmedDate: `${dateToUse} (ช่าง ${techName} กดตอบรับงานแล้ว)`,
            assignedTech: techName,
            status: 'REPAIRING' as TicketStatus
          }
        : item
    );
    setTickets(updatedList);
    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(updatedList));
    pushOnlineSharedState(updatedList, 'TECHNICIAN_ACCEPT');
    const targetTicket = updatedList.find((item) => item.id === t.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(targetTicket, `👷 [ช่าง ${techName} กดตอบรับงานแล้ว - เริ่มดำเนินการ]`)
      );
    }
    showToast(`ช่าง ${techName} กดตอบรับใบงาน ${t.id} เรียบร้อยแล้ว (บรรจุเข้าหมวด "งานของฉัน" ทันที)`, 'success');
  };

  const handleStatusButtonClick = (t: ServiceTicket, targetStatus: TicketStatus) => {
    if (t.isLocked && userRole !== 'SUPERVISOR') {
      showToast('ใบงานนี้ถูกล็อกแล้ว ต้องให้ Supervisor ปลดล็อกเท่านั้น', 'warning');
      return;
    }
    if (targetStatus === 'ASSIGNED') {
      openAssignOrRescheduleModal(t);
      return;
    }
    const updatedList = tickets.map((item) =>
      item.id === t.id ? { ...item, status: targetStatus } : item
    );
    setTickets(updatedList);
    const updatedTicket = updatedList.find((item) => item.id === t.id);
    if (updatedTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(updatedTicket, `🔄 [อัปเดตสถานะใบงานเป็น ${targetStatus} - สหคิมมอเตอร์]`)
      );
    }
    showToast(`อัปเดตสถานะใบงาน ${t.id} เป็น ${targetStatus}`, 'success');
  };

  const handleConfirmAssign = (e: React.FormEvent) => {
    e.preventDefault();
    if (!assignModalTicket || !assignDateInput || !assignTechInput) return;
    const updatedList = tickets.map((item) =>
      item.id === assignModalTicket.id
        ? {
            ...item,
            desiredDate: assignDateInput,
            confirmedDate: `${assignDateInput} (ยืนยันโดย Service Support - รอช่างตอบรับ)`,
            assignedTech: assignTechInput,
            leadTech: assignLeadTechInput,
            status: 'ASSIGNED' as TicketStatus
          }
        : item
    );
    setTickets(updatedList);
    const targetTicket = updatedList.find((item) => item.id === assignModalTicket.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(targetTicket, '👷 [กำหนดวันนัดหมายและมอบหมายช่างแล้ว - รอช่างตอบรับ]')
      );
    }
    setSupportSelectedDate(assignDateInput);
    showToast(
      `กำหนดวันนัด ${assignDateInput} และมอบหมายงาน ${assignModalTicket.id} ให้ ${assignTechInput} เรียบร้อย`,
      'success'
    );
    setAssignModalTicket(null);
  };

  const handleConfirmAssignAndAcceptImmediate = () => {
    if (!assignModalTicket || !assignDateInput || !assignTechInput) return;
    const updatedList = tickets.map((item) =>
      item.id === assignModalTicket.id
        ? {
            ...item,
            desiredDate: assignDateInput,
            confirmedDate: `${assignDateInput} (Service Support รับงานแทนช่าง ${assignTechInput} แล้ว)`,
            assignedTech: assignTechInput,
            leadTech: assignLeadTechInput,
            status: 'REPAIRING' as TicketStatus
          }
        : item
    );
    setTickets(updatedList);
    const targetTicket = updatedList.find((item) => item.id === assignModalTicket.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(
          targetTicket,
          `✅ [Service Support กำหนดวันนัดและรับงานแทนช่าง (${assignTechInput}) ทันที]`
        )
      );
    }
    setSupportSelectedDate(assignDateInput);
    showToast(
      `กำหนดวัน ${assignDateInput} มอบหมาย ${assignTechInput} และกดรับงานแทนช่างเรียบร้อย!`,
      'success'
    );
    setAssignModalTicket(null);
  };

  const availableApproverOptions = useMemo(() => {
    const leadTechNames = technicians
      .filter((tc) => tc.position?.includes('หัวหน้าช่าง'))
      .map((tc) => `${tc.name} (หัวหน้าช่าง)`);
    return Array.from(new Set([...supervisors, ...leadTechNames]));
  }, [supervisors, technicians]);

  const openApproveModal = (t: ServiceTicket, mode: 'PARTS' | 'REPORT') => {
    setApproveReportTicket(t);
    setApproveMode(mode);
    const defaultApprover =
      t.approvedBySupervisor ||
      availableApproverOptions[0] ||
      'หัวหน้าวิศวกรบริการ (Supervisor)';
    setSupervisorSignName(defaultApprover);
    setCustomSupervisorInput('');
  };

  const handleConfirmSupervisorApproveReport = (e: React.FormEvent) => {
    e.preventDefault();
    if (!approveReportTicket) return;
    const finalSign =
      supervisorSignName === '__CUSTOM__'
        ? customSupervisorInput.trim() || 'หัวหน้าวิศวกรบริการ (Supervisor)'
        : supervisorSignName.trim() || 'หัวหน้าวิศวกรบริการ (Supervisor)';

    if (
      supervisorSignName === '__CUSTOM__' &&
      customSupervisorInput.trim() &&
      !supervisors.includes(customSupervisorInput.trim())
    ) {
      setSupervisors((prev) => [...prev, customSupervisorInput.trim()]);
    }

    if (approveMode === 'PARTS') {
      const updatedList = tickets.map((t) =>
        t.id === approveReportTicket.id
          ? {
              ...t,
              approvedBySupervisor: finalSign,
              status: 'REPAIRING' as TicketStatus,
              equipmentList: t.equipmentList.map((eq) => ({ ...eq, partsApproved: true }))
            }
          : t
      );
      setTickets(updatedList);
      const targetTicket = updatedList.find((t) => t.id === approveReportTicket.id);
      if (targetTicket) {
        syncToGoogleSheetAndLine(
          updatedList,
          formatTicketLineMessage(
            targetTicket,
            `✅ [อนุมัติอะไหล่แล้วโดย ${finalSign} - ดำเนินการซ่อมได้]`
          )
        );
      }
      showToast(`อนุมัติรายการอะไหล่ใบงาน ${approveReportTicket.id} โดย ${finalSign} สำเร็จ`, 'success');
      setApproveReportTicket(null);
      return;
    }

    const updatedList = tickets.map((t) =>
      t.id === approveReportTicket.id
        ? {
            ...t,
            approvedBySupervisor: finalSign,
            status: 'CUSTOMER_ACCEPTANCE' as TicketStatus,
            isLocked: true
          }
        : t
    );
    setTickets(updatedList);
    const targetTicket = updatedList.find((t) => t.id === approveReportTicket.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(
          targetTicket,
          `✅ [อนุมัติใบรายงานแล้วโดย ${finalSign} - รอลูกค้าตรวจรับงาน]`
        )
      );
    }
    showToast(`อนุมัติรายงาน ${approveReportTicket.id} โดย ${finalSign} และล็อกข้อมูลสำเร็จ`, 'success');
    setApproveReportTicket(null);
  };

  const handleConfirmReject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!rejectModalTicket || !rejectReasonInput.trim()) return;
    const updatedList = tickets.map((t) =>
      t.id === rejectModalTicket.id
        ? {
            ...t,
            rejectReason: rejectReasonInput.trim(),
            status: 'REJECTED' as TicketStatus
          }
        : t
    );
    setTickets(updatedList);
    const targetTicket = updatedList.find((t) => t.id === rejectModalTicket.id);
    if (targetTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(targetTicket, '❌ [Supervisor ส่งกลับให้แก้ไขข้อมูลใบงาน]')
      );
    }
    showToast(`ส่งกลับใบงาน ${rejectModalTicket.id} เพื่อแก้ไขแล้ว`, 'warning');
    setRejectModalTicket(null);
    setRejectReasonInput('');
  };

  const handleConfirmUnlock = (e: React.FormEvent) => {
    e.preventDefault();
    if (!unlockModalTicket) return;
    if (unlockPassInput !== supPassword && unlockPassInput !== 'admin123') {
      showToast('รหัสผ่าน Supervisor ไม่ถูกต้อง (รหัสเริ่มต้นคือ admin123)', 'error');
      return;
    }
    const updatedList = tickets.map((t) =>
      t.id === unlockModalTicket.id ? { ...t, isLocked: false } : t
    );
    setTickets(updatedList);
    syncToGoogleSheetAndLine(updatedList);
    showToast(`ปลดล็อกใบงาน ${unlockModalTicket.id} สำเร็จ`, 'success');
    setUnlockModalTicket(null);
    setUnlockPassInput('');
  };

  const customerAcceptJob = (id: string) => {
    const target = tickets.find((t) => t.id === id);
    if (target) {
      setCustomerSignOffModalTicket({
        ...target,
        customerSignerName: target.customerSignerName || target.siteContact || target.contact
      });
      return;
    }
    const updatedList = tickets.map((t) =>
      t.id === id ? { ...t, status: 'CLOSE' as TicketStatus, isLocked: true } : t
    );
    setTickets(updatedList);
    const closedTicket = updatedList.find((t) => t.id === id);
    if (closedTicket) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(closedTicket, '🎉 [ลูกค้าตรวจรับงานและปิดจบใบงาน (CLOSE) เรียบร้อยแล้ว]')
      );
    }
    showToast(`ตรวจรับงานใบงาน ${id} และปิดจบเรียบร้อย`, 'success');
  };

  const updateTicketSignatureFields = (
    ticketId: string,
    patch: Partial<
      Pick<
        ServiceTicket,
        | 'customerSignature'
        | 'customerSignerName'
        | 'customerSignedAt'
        | 'technicianSignature'
        | 'technicianSignedAt'
        | 'status'
        | 'isLocked'
      >
    >,
    notifyHeading?: string
  ) => {
    const updatedList = tickets.map((t) => (t.id === ticketId ? { ...t, ...patch } : t));
    setTickets(updatedList);
    const updatedTicket = updatedList.find((t) => t.id === ticketId);
    if (updatedTicket && notifyHeading) {
      syncToGoogleSheetAndLine(
        updatedList,
        formatTicketLineMessage(updatedTicket, notifyHeading)
      );
    } else {
      syncToGoogleSheetAndLine(updatedList);
    }
  };

  const handleNewImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const compressed = await compressImage(e.target.files[0]);
      if (compressed) {
        setNewUploadedPhoto(compressed);
        showToast('บีบอัดรูปภาพและพร้อมใช้งานแล้ว', 'success');
      }
    }
  };

  // Upload and auto-compress photos per equipment draft item (from album or mobile camera)
  const handleEquipmentPhotoUpload = async (
    eqIdx: number,
    field: 'beforePhotos' | 'afterPhotos',
    source: React.ChangeEvent<HTMLInputElement> | FileList | null
  ) => {
    let fileList: FileList | null = null;
    if (source && 'target' in source) {
      fileList = (source as React.ChangeEvent<HTMLInputElement>).target.files;
    } else {
      fileList = source as FileList | null;
    }
    if (fileList && fileList.length > 0) {
      const files: File[] = Array.from(fileList);
      const compressedList: string[] = [];
      for (const f of files) {
        const compressed = await compressImage(f);
        if (compressed) compressedList.push(compressed);
      }
      if (compressedList.length > 0) {
        setTempEquipmentList((prev) =>
          prev.map((item, i) =>
            i === eqIdx
              ? {
                  ...item,
                  [field]: [...(item[field] || []), ...compressedList]
                }
              : item
          )
        );
        showToast(`อัปโหลดและบีบอัดรูปภาพ ${compressedList.length} รูปเรียบร้อย!`, 'success');
      }
      if (source && 'target' in source && (source as React.ChangeEvent<HTMLInputElement>).target) {
        (source as React.ChangeEvent<HTMLInputElement>).target.value = '';
      }
    }
  };

  const handleRemoveEquipmentPhoto = (
    eqIdx: number,
    field: 'beforePhotos' | 'afterPhotos',
    photoIdx: number
  ) => {
    setTempEquipmentList((prev) =>
      prev.map((item, i) =>
        i === eqIdx
          ? {
              ...item,
              [field]: (item[field] || []).filter((_, pI) => pI !== photoIdx)
            }
          : item
      )
    );
  };

  // Upload and auto-compress photos when technician is editing an active ticket
  const handleEditingEquipmentPhotoUpload = async (
    eqIdx: number,
    field: 'beforePhotos' | 'afterPhotos',
    source: React.ChangeEvent<HTMLInputElement> | FileList | null
  ) => {
    let fileList: FileList | null = null;
    if (source && 'target' in source) {
      fileList = (source as React.ChangeEvent<HTMLInputElement>).target.files;
    } else {
      fileList = source as FileList | null;
    }
    if (fileList && fileList.length > 0 && editingTicket) {
      const files: File[] = Array.from(fileList);
      const compressedList: string[] = [];
      for (const f of files) {
        const compressed = await compressImage(f);
        if (compressed) compressedList.push(compressed);
      }
      if (compressedList.length > 0) {
        setEditingTicket({
          ...editingTicket,
          equipmentList: editingTicket.equipmentList.map((item, i) =>
            i === eqIdx
              ? {
                  ...item,
                  [field]: [...(item[field] || []), ...compressedList]
                }
              : item
          )
        });
        showToast(`บันทึกรูปภาพ ${compressedList.length} รูปในใบงานเรียบร้อย!`, 'success');
      }
      if (source && 'target' in source && (source as React.ChangeEvent<HTMLInputElement>).target) {
        (source as React.ChangeEvent<HTMLInputElement>).target.value = '';
      }
    }
  };

  const handleRemoveEditingEquipmentPhoto = (
    eqIdx: number,
    field: 'beforePhotos' | 'afterPhotos',
    photoIdx: number
  ) => {
    if (!editingTicket) return;
    setEditingTicket({
      ...editingTicket,
      equipmentList: editingTicket.equipmentList.map((item, i) =>
        i === eqIdx
          ? {
              ...item,
              [field]: (item[field] || []).filter((_, pI) => pI !== photoIdx)
            }
          : item
      )
    });
  };

  // Delete ticket completely from system and sync
  const handleDeleteTicket = async (ticketId: string) => {
    const target = tickets.find((t) => t.id === ticketId);
    if (!target) return;
    if (
      !window.confirm(
        `⚠️ ยืนยันการลบใบงาน ${target.id} (${target.company}) ออกจากระบบ?\n\n• ข้อมูลใบงานและคิวงานช่างจะถูกลบออกจากระบบทันที\n• ระบบจะซิงค์การลบข้อมูลไปยังทุกเครื่องออนไลน์`
      )
    ) {
      return;
    }
    const updated = tickets.filter((t) => t.id !== ticketId);
    setTickets(updated);
    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(updated));
    await pushOnlineSharedState(updated, userRole || 'DELETE_TICKET');
    if (gasWebAppUrl.trim()) {
      await syncToGoogleSheetAndLine(
        updated,
        `🗑️ [ลบใบงาน ${target.id} (${target.company}) ออกจากระบบเรียบร้อยแล้ว]`
      );
    }
    if (assignModalTicket?.id === ticketId) {
      setAssignModalTicket(null);
    }
    showToast(`ลบใบงาน ${ticketId} ออกจากระบบเรียบร้อยแล้ว`, 'warning');
  };

  // Remove/Unassign job from technician's queue and return to NEW
  const handleUnassignTicketFromTech = async (ticketId: string) => {
    const target = tickets.find((t) => t.id === ticketId);
    if (!target) return;
    const oldTech = target.assignedTech || 'ช่าง';
    if (
      !window.confirm(
        `ต้องการลบงาน ${target.id} ออกจากคิวของ ${oldTech} ใช่หรือไม่?\n\n• ใบงานจะกลับไปอยู่ในสถานะ NEW (รอจัดคิวใหม่)\n• คิวของ ${oldTech} จะกลับมาว่างทันที`
      )
    ) {
      return;
    }
    const updated = tickets.map((item) =>
      item.id === ticketId
        ? {
            ...item,
            assignedTech: '',
            confirmedDate: '',
            status: 'NEW' as TicketStatus
          }
        : item
    );
    setTickets(updated);
    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(updated));
    await pushOnlineSharedState(updated, userRole || 'UNASSIGN_TECH');
    if (gasWebAppUrl.trim()) {
      await syncToGoogleSheetAndLine(
        updated,
        `🔄 [ลบงานออกจากคิวช่าง: ใบงาน ${target.id} (${target.company}) ปลด ${oldTech} ออกและคืนสถานะ NEW เรียบร้อย]`
      );
    }
    if (assignModalTicket?.id === ticketId) {
      setAssignModalTicket(null);
    }
    showToast(
      `ลบงาน ${target.id} ออกจากคิวของ ${oldTech} เรียบร้อยแล้ว (สถานะกลับเป็นงานใหม่ / ตารางช่างว่าง)`,
      'success'
    );
  };

  // Production Reset & Demo Ticket Restore Handlers
  const handleResetToProductionClean = async () => {
    setTickets([]);
    localStorage.setItem('sahakim_tickets_v121', '[]');
    try {
      await pushOnlineSharedState([], userRole || 'RESET_PRODUCTION');
    } catch {
      // ignore
    }
    setResetConfirmModalOpen(false);
    showToast(
      '🎉 ล้างใบงานตัวอย่างเรียบร้อยแล้ว! ตารางงานช่างว่าง (สีเขียว 🟢 ทุกวัน) พร้อมรับงานจริงทันที',
      'success'
    );
  };

  const handleRestoreDemoTickets = async () => {
    try {
      const res = await fetch('/api/online/state');
      if (res.ok) {
        const data = await res.json();
        const serverTickets: ServiceTicket[] = Array.isArray(data.tickets) && data.tickets.length > 0
          ? data.tickets
          : INITIAL_TICKETS;
        setTickets(serverTickets);
        localStorage.setItem('sahakim_tickets_v121', JSON.stringify(serverTickets));
        setResetConfirmModalOpen(false);
        showToast(
          `✅ กู้คืนข้อมูลออนไลน์ล่าสุดสำเร็จ (${serverTickets.length} ใบงาน · รวม SKM-69-003 คุณอาทิตย์ วงศ์สว่าง)`,
          'success'
        );
        return;
      }
    } catch {}
    setTickets(INITIAL_TICKETS);
    localStorage.setItem('sahakim_tickets_v121', JSON.stringify(INITIAL_TICKETS));
    try {
      await pushOnlineSharedState(INITIAL_TICKETS, userRole || 'RESTORE_DEMO');
    } catch {}
    setResetConfirmModalOpen(false);
    showToast('✅ กู้คืนข้อมูลใบงานออนไลน์เรียบร้อยแล้ว (รวม SKM-69-003 คุณอาทิตย์ วงศ์สว่าง)', 'success');
  };

  const handleSubmitNewTicket = (e: React.FormEvent) => {
    e.preventDefault();

    // Persist any newly entered custom equipment types or brands into master lists
    const addedEquipTypes: string[] = [];
    const addedBrands: string[] = [];

    const resolvedEquipment = tempEquipmentList.map((eq) => {
      const finalType =
        eq.type.startsWith('ETC.') && eq.customType.trim()
          ? eq.customType.trim().toUpperCase()
          : eq.type.startsWith('ETC.')
          ? 'OTHER EQUIPMENT'
          : eq.type;

      const finalBrand =
        eq.brand.startsWith('ETC.') && eq.customBrand.trim()
          ? eq.customBrand.trim().toUpperCase()
          : eq.brand.startsWith('ETC.')
          ? 'OTHER BRAND'
          : eq.brand;

      if (eq.type.startsWith('ETC.') && eq.customType.trim()) {
        addedEquipTypes.push(finalType);
      }
      if (eq.brand.startsWith('ETC.') && eq.customBrand.trim()) {
        addedBrands.push(finalBrand);
      }

      const assignedBefore =
        eq.beforePhotos && eq.beforePhotos.length > 0
          ? eq.beforePhotos
          : newUploadedPhoto
          ? [newUploadedPhoto]
          : [DEFAULT_BEFORE_IMAGE];

      return {
        type: finalType,
        brand: finalBrand,
        model: eq.model.trim() || '-',
        tag: eq.tag.trim() || 'N/A',
        symptom: eq.symptom.trim(),
        detail: eq.detail.trim(),
        beforePhotos: assignedBefore,
        afterPhotos: eq.afterPhotos || [],
        analysis: '',
        cause: '',
        standardsReference: '',
        faultOrigin: newWarranty.includes('In-Warranty')
          ? 'เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)'
          : 'เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)',
        correctiveAction: '',
        parts: [],
        partsApproved: false,
        checklistResults: buildDefaultChecklistItems(finalType, masterChecklists)
      };
    });

    if (addedEquipTypes.length > 0) {
      setMasterEquipmentTypes((prev) => Array.from(new Set([...prev, ...addedEquipTypes])));
    }
    if (addedBrands.length > 0) {
      setMasterBrands((prev) => Array.from(new Set([...prev, ...addedBrands])));
    }

    const nextNumber = String(tickets.length + 1).padStart(3, '0');
    const newTicket: ServiceTicket = {
      id: `SKM-69-${nextNumber}`,
      company: newCompany.trim() || customerCompany,
      taxId: newTaxId.trim(),
      docNo: newDocNo.trim(),
      contact: newContact.trim(),
      siteContact: newContact.trim(),
      location: newLocation.trim(),
      warranty: newWarranty,
      serviceType: newServiceType,
      urgency: newUrgency,
      desiredDate: newDesiredDate,
      confirmedDate: '',
      status: 'NEW',
      isLocked: false,
      rejectReason: '',
      createdAt: new Date().toISOString().replace('T', ' ').substring(0, 16),
      equipmentList: resolvedEquipment,
      assignedTech: '',
      approvedBySupervisor: ''
    };

    const updatedTickets = [newTicket, ...tickets];
    setTickets(updatedTickets);
    syncToGoogleSheetAndLine(
      updatedTickets,
      formatTicketLineMessage(newTicket, '🔔 [มีใบแจ้งซ่อมใหม่! (NEW) - สหคิมมอเตอร์]')
    );
    setCustomerCompany(newTicket.company);
    setCustomerLookupQuery(newTicket.company);
    setTempEquipmentList([
      {
        type: 'END SUCTION PUMP',
        customType: '',
        brand: 'STAC',
        customBrand: '',
        model: '',
        tag: '',
        symptom: '',
        detail: ''
      }
    ]);
    setNewDocNo('');
    setNewContact('');
    setNewLocation('');
    setNewUploadedPhoto('');
    setCustTab('list');
    if (currentView === 'internal_dashboard') {
      setInternalTab('tickets');
    }
    playNotificationChime(soundEnabled);
    showToast(
      `สร้างใบแจ้งซ่อม ${newTicket.id} (${resolvedEquipment.length} อุปกรณ์) สำเร็จ และส่งแจ้งเตือนแล้ว`,
      'success'
    );
  };

  const openActionModal = (ticket: ServiceTicket) => {
    if (ticket.isLocked && userRole !== 'SUPERVISOR') {
      showToast('ใบงานนี้ถูกล็อกแล้ว ต้องให้ Supervisor ปลดล็อกก่อนจึงจะแก้ไขได้', 'warning');
      return;
    }
    const cloned: ServiceTicket = JSON.parse(JSON.stringify(ticket));
    cloned.equipmentList.forEach((eq) => {
      if (!eq.checklistResults || eq.checklistResults.length === 0) {
        eq.checklistResults = buildDefaultChecklistItems(eq.type, masterChecklists);
      } else {
        eq.checklistResults = eq.checklistResults.map((chk) => enrichChecklistItem(chk));
      }
      if (eq.parts && eq.parts.length > 0) {
        eq.parts = eq.parts.map((p) => ({
          ...p,
          billingType:
            p.billingType ||
            (isPartInternalWarranty(p, eq, cloned) ? 'WARRANTY_INTERNAL' : 'CUSTOMER_CHARGE')
        }));
      }
    });
    setPartDrafts({});
    setAiSuggestionsByIdx({});
    setModalCustomTypeByIdx({});
    setModalCustomBrandByIdx({});
    setSaveTemplateForIdx(null);
    setEditingTicket(cloned);
  };

  // AI Troubleshoot Handler (calls /api/ai/troubleshoot with fallback to local engineering reasoning)
  const handleRunAiTroubleshoot = async (idx: number) => {
    if (!editingTicket) return;
    const eq = editingTicket.equipmentList[idx];
    if (!eq) return;

    setAiLoadingByIdx((prev) => ({ ...prev, [idx]: true }));

    try {
      const response = await fetch('/api/ai/troubleshoot', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          equipmentType: eq.type,
          brand: eq.brand,
          model: eq.model,
          symptom: eq.symptom,
          detail: eq.detail,
          checklistResults: eq.checklistResults
        })
      });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || `HTTP ${response.status}`);
      }

      const data = await response.json();

      setEditingTicket((prev) => {
        if (!prev) return prev;
        const updatedList = [...prev.equipmentList];
        const target = updatedList[idx];
        updatedList[idx] = {
          ...target,
          analysis: data.analysis || target.analysis,
          cause: data.cause || target.cause,
          correctiveAction: data.correctiveAction || target.correctiveAction,
          standardsReference:
            data.standardsReference ||
            "Karassik's Pump Handbook / EASA AR100 / Manufacturer Troubleshooting Guide"
        };
        return { ...prev, equipmentList: updatedList };
      });

      if (
        (Array.isArray(data.recommendedChecklist) && data.recommendedChecklist.length > 0) ||
        (Array.isArray(data.recommendedParts) && data.recommendedParts.length > 0)
      ) {
        setAiSuggestionsByIdx((prev) => ({
          ...prev,
          [idx]: {
            recommendedChecklist: Array.isArray(data.recommendedChecklist)
              ? data.recommendedChecklist
              : [],
            recommendedParts: Array.isArray(data.recommendedParts) ? data.recommendedParts : []
          }
        }));
      }

      showToast(
        `AI วิเคราะห์อาการของ ${eq.type} (${eq.brand}) สำเร็จ ช่างสามารถปรับแก้ข้อมูลเพิ่มเติมได้ทันที`,
        'success'
      );
    } catch {
      // Fallback to intelligent rule-based engineering synthesis from Pump Handbook & EASA standards
      const upperType = eq.type.toUpperCase();
      const matchedTpl =
        knowledgeTemplates.find((tpl) => upperType.includes(tpl.category)) ||
        DEFAULT_TROUBLESHOOTING_TEMPLATES[1];

      setEditingTicket((prev) => {
        if (!prev) return prev;
        const updatedList = [...prev.equipmentList];
        const target = updatedList[idx];
        updatedList[idx] = {
          ...target,
          analysis: `[อ้างอิง ${matchedTpl.sourceRef}] ตรวจสอบ ${target.type} (${target.brand} ${target.model}) จากอาการ "${target.symptom}": ${matchedTpl.analysis}`,
          cause: matchedTpl.cause,
          correctiveAction: matchedTpl.correctiveAction,
          standardsReference: matchedTpl.sourceRef
        };
        return { ...prev, equipmentList: updatedList };
      });

      if (matchedTpl.recommendedChecklist && matchedTpl.recommendedChecklist.length > 0) {
        setAiSuggestionsByIdx((prev) => ({
          ...prev,
          [idx]: {
            recommendedChecklist: matchedTpl.recommendedChecklist || [],
            recommendedParts: []
          }
        }));
      }

      showToast(
        'ดึงข้อมูลวิเคราะห์เบื้องต้นตามมาตรฐาน Pump Handbook / EASA สำเร็จ (แก้ไขเพิ่มเติมได้)',
        'success'
      );
    } finally {
      setAiLoadingByIdx((prev) => ({ ...prev, [idx]: false }));
    }
  };

  const applyKnowledgeTemplate = (idx: number, tplId: string) => {
    if (!editingTicket || !tplId) return;
    const tpl = knowledgeTemplates.find((x) => x.id === tplId);
    if (!tpl) return;

    const updated = [...editingTicket.equipmentList];
    const eq = updated[idx];
    updated[idx] = {
      ...eq,
      analysis: tpl.analysis,
      cause: tpl.cause,
      correctiveAction: tpl.correctiveAction,
      standardsReference: tpl.sourceRef
    };
    setEditingTicket({ ...editingTicket, equipmentList: updated });
    showToast(`นำเข้าต้นแบบ "${tpl.title}" เรียบร้อย ช่างสามารถแก้ไขต่อได้`, 'success');
  };

  const handleSaveCustomTemplate = (idx: number) => {
    if (!editingTicket) return;
    const eq = editingTicket.equipmentList[idx];
    if (!newTemplateTitle.trim()) {
      showToast('กรุณาตั้งชื่อต้นแบบการวิเคราะห์', 'error');
      return;
    }
    const newTpl: TroubleshootingTemplate = {
      id: `TPL-CUSTOM-${Date.now()}`,
      title: newTemplateTitle.trim(),
      category: eq.type,
      sourceRef: eq.standardsReference || `บันทึกโดยช่าง (${eq.brand} ${eq.type})`,
      analysis: eq.analysis || '-',
      cause: eq.cause || '-',
      correctiveAction: eq.correctiveAction || '-'
    };
    setKnowledgeTemplates((prev) => [newTpl, ...prev]);
    setSaveTemplateForIdx(null);
    setNewTemplateTitle('');
    showToast('บันทึกเข้าคลังความรู้ช่างเรียบร้อย สามารถเลือกใช้ซ้ำในงานถัดไปได้ทันที', 'success');
  };

  const openServiceReportView = (ticketId: string) => {
    if (currentView !== 'service_report') {
      setPreviousView(currentView === 'internal_login' ? 'home' : currentView);
    }
    setActiveReportTicketId(ticketId);
    setCurrentView('service_report');
  };

  // Dashboard Metrics Calculation
  const dashboardStats = useMemo(() => {
    const totalTickets = tickets.length;
    const closedTickets = tickets.filter((t) => t.status === 'CLOSE').length;
    const focCount = tickets.filter((t) => t.warranty && t.warranty.includes('In-Warranty')).length;
    const chargeCount = totalTickets - focCount;

    let totalMoney = 0;
    let internalCostTotal = 0;
    let customerBilledTotal = 0;
    const equipCountMap: Record<string, number> = {};

    tickets.forEach((t) => {
      const costSummary = getTicketCostSummary(t);
      internalCostTotal += costSummary.internalCostTotal;
      customerBilledTotal += costSummary.customerBilledTotal;
      totalMoney += costSummary.grossPartsTotal;
      t.equipmentList.forEach((eq) => {
        equipCountMap[eq.type] = (equipCountMap[eq.type] || 0) + 1;
      });
    });

    const overdueTickets = tickets.filter((t) => isTicketOverdue(t));
    const overdueCount = overdueTickets.length;
    const onTimeCount = Math.max(0, totalTickets - overdueCount);
    const topEquip =
      Object.keys(equipCountMap).sort((a, b) => equipCountMap[b] - equipCountMap[a])[0] ||
      'BOOSTER PUMP';

    const CHART_COLORS = [
      '#f97316', // orange-500
      '#0f172a', // slate-900
      '#2563eb', // blue-600
      '#10b981', // emerald-500
      '#8b5cf6', // violet-500
      '#f59e0b', // amber-500
      '#06b6d4', // cyan-500
      '#ec4899', // pink-500
      '#64748b'  // slate-500
    ];

    const totalEquipItems = Object.values(equipCountMap).reduce((sum, n) => sum + n, 0);
    const equipmentChartData = Object.entries(equipCountMap)
      .sort((a, b) => b[1] - a[1])
      .map(([name, value], index) => ({
        name,
        value,
        percentage: totalEquipItems > 0 ? Math.round((value / totalEquipItems) * 100) : 0,
        color: CHART_COLORS[index % CHART_COLORS.length]
      }));

    const statusDefinitions: {
      key: TicketStatus;
      label: string;
      shortLabel: string;
      filterVal: string;
      color: string;
    }[] = [
      { key: 'NEW', label: 'งานใหม่ (NEW)', shortLabel: 'งานใหม่', filterVal: 'NEW', color: '#2563eb' },
      { key: 'ASSIGNED', label: 'มอบหมายแล้ว (ASSIGNED)', shortLabel: 'มอบหมายแล้ว', filterVal: 'ASSIGNED', color: '#0284c7' },
      { key: 'INSPECTION', label: 'ตรวจเช็ค (INSPECTION)', shortLabel: 'ตรวจเช็ค', filterVal: 'INSPECTION', color: '#0ea5e9' },
      { key: 'REPAIRING', label: 'กำลังซ่อม (REPAIRING)', shortLabel: 'กำลังซ่อม', filterVal: 'REPAIRING', color: '#f59e0b' },
      { key: 'PENDING_PARTS_APPROVAL', label: 'รออนุมัติอะไหล่', shortLabel: 'รออะไหล่', filterVal: 'PENDING_APPROVAL', color: '#f97316' },
      { key: 'PENDING_SUPERVISOR_APPROVAL', label: 'รออนุมัติรายงาน (Supervisor)', shortLabel: 'รออนุมัติ', filterVal: 'PENDING_APPROVAL', color: '#8b5cf6' },
      { key: 'REJECTED', label: 'ส่งกลับแก้ไข (REJECTED)', shortLabel: 'ส่งกลับ', filterVal: 'REJECTED', color: '#ef4444' },
      { key: 'CUSTOMER_ACCEPTANCE', label: 'รอลูกค้าตรวจรับ', shortLabel: 'รอตรวจรับ', filterVal: 'CUSTOMER_ACCEPTANCE', color: '#06b6d4' },
      { key: 'CLOSE', label: 'ปิดงานแล้ว (CLOSE)', shortLabel: 'ปิดงานแล้ว', filterVal: 'CLOSE', color: '#10b981' }
    ];

    const statusCountMap: Record<string, number> = {};
    tickets.forEach((t) => {
      statusCountMap[t.status] = (statusCountMap[t.status] || 0) + 1;
    });

    const statusChartData = statusDefinitions.map((st) => ({
      status: st.key,
      name: st.shortLabel,
      fullLabel: st.label,
      filterVal: st.filterVal,
      count: statusCountMap[st.key] || 0,
      color: st.color
    }));

    return {
      totalTickets,
      closedTickets,
      focCount,
      chargeCount,
      totalMoney,
      internalCostTotal,
      customerBilledTotal,
      onTimeCount,
      overdueCount,
      topEquip,
      totalEquipItems,
      equipmentChartData,
      statusChartData
    };
  }, [tickets]);

  const techJobsCount = useMemo(() => {
    if (userRole !== 'TECHNICIAN') return { myJobs: 0, availableJobs: 0 };
    const myName = loggedInTech?.name || '';
    const myJobs = tickets.filter((t) => myName && t.assignedTech === myName).length;
    const availableJobs = tickets.filter(
      (t) =>
        (!t.assignedTech || t.assignedTech === 'ยังไม่ระบุ' || t.assignedTech === '') &&
        (t.status === 'NEW' || t.status === 'ASSIGNED')
    ).length;
    return { myJobs, availableJobs };
  }, [tickets, userRole, loggedInTech]);

  const filteredInternalTickets = useMemo(() => {
    return tickets.filter((t) => {
      // 1. TECHNICIAN scope filter: see only their own jobs or pool jobs waiting to be accepted
      if (userRole === 'TECHNICIAN') {
        const myName = loggedInTech?.name || '';
        const isMyJob = Boolean(myName && t.assignedTech === myName);
        const isPoolJob =
          (!t.assignedTech || t.assignedTech === 'ยังไม่ระบุ' || t.assignedTech === '') &&
          (t.status === 'NEW' || t.status === 'ASSIGNED');

        if (techWorkFilter === 'AVAILABLE_JOBS') {
          if (!isPoolJob && !(t.status === 'ASSIGNED' && isMyJob)) return false;
        } else {
          // Default MY_JOBS: Strictly only show jobs assigned to this technician!
          if (!isMyJob) return false;
        }
      }

      const matchesStatus =
        statusFilter === 'ALL' ||
        (statusFilter === 'OVERDUE' && isTicketOverdue(t)) ||
        (statusFilter === 'PENDING_APPROVAL' &&
          (t.status === 'PENDING_SUPERVISOR_APPROVAL' || t.status === 'PENDING_PARTS_APPROVAL')) ||
        t.status === statusFilter;

      const q = searchQuery.trim().toLowerCase();
      if (!q) return matchesStatus;
      const matchesQuery =
        t.id.toLowerCase().includes(q) ||
        t.company.toLowerCase().includes(q) ||
        t.location.toLowerCase().includes(q) ||
        (t.assignedTech && t.assignedTech.toLowerCase().includes(q)) ||
        t.equipmentList.some(
          (eq) =>
            eq.type.toLowerCase().includes(q) ||
            eq.brand.toLowerCase().includes(q) ||
            eq.model.toLowerCase().includes(q)
        );
      return matchesStatus && matchesQuery;
    });
  }, [tickets, userRole, loggedInTech, techWorkFilter, statusFilter, searchQuery]);

  const renderStatusControl = (t: ServiceTicket) => {
    const status = t.status;

    if (status === 'NEW') {
      return (
        <div className="flex flex-wrap items-center gap-1.5">
          {userRole === 'TECHNICIAN' && (
            <button
              type="button"
              onClick={() => techAcceptJob(t)}
              className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-bold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer shadow-sm"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>👷 ช่างกดตอบรับงานนี้</span>
            </button>
          )}
          <button
            type="button"
            onClick={() => openAssignOrRescheduleModal(t)}
            className="bg-blue-600 hover:bg-blue-700 text-white text-xs px-3 py-1.5 rounded-lg font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
          >
            <Bell className="w-3.5 h-3.5" />
            <span>NEW: เลือกวัน / เลือกช่างตามคิวว่าง</span>
          </button>
          {(userRole === 'SERVICE TEAM' || userRole === 'SUPERVISOR') && (
            <button
              type="button"
              onClick={() => supportAcceptOnBehalfOfTech(t)}
              className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Check className="w-3.5 h-3.5" />
              <span>รับงานแทนช่างเลย</span>
            </button>
          )}
        </div>
      );
    }

    if (status === 'ASSIGNED') {
      return (
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="bg-amber-100 text-amber-900 border border-amber-300 text-xs px-2.5 py-1 rounded-lg font-bold">
            ⏳ รอช่างตอบรับงาน ({t.assignedTech || 'ยังไม่ระบุช่าง'})
          </span>
          {userRole === 'TECHNICIAN' && (
            <button
              type="button"
              onClick={() => techAcceptJob(t)}
              className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>ช่างกดตอบรับงาน</span>
            </button>
          )}
          {(userRole === 'SERVICE TEAM' || userRole === 'SUPERVISOR') && (
            <>
              <button
                type="button"
                onClick={() => supportAcceptOnBehalfOfTech(t)}
                className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>รับงานแทนช่างเลย</span>
              </button>
              <button
                type="button"
                onClick={() => openAssignOrRescheduleModal(t)}
                className="bg-blue-600 hover:bg-blue-700 text-white text-xs px-2.5 py-1.5 rounded-lg font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <UserCog className="w-3.5 h-3.5" />
                <span>เปลี่ยนช่างใหม่ / เปลี่ยนวัน</span>
              </button>
            </>
          )}
        </div>
      );
    }

    if (status === 'INSPECTION' || status === 'REPAIRING') {
      return (
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="bg-emerald-100 text-emerald-900 border border-emerald-300 text-xs px-2.5 py-1 rounded-lg font-bold">
            🔧 กำลังดำเนินการซ่อม ({t.assignedTech || 'ช่างเทคนิค'})
          </span>
          {(userRole === 'SERVICE TEAM' || userRole === 'SUPERVISOR') && !t.isLocked && (
            <button
              type="button"
              onClick={() => openAssignOrRescheduleModal(t)}
              className="bg-slate-800 hover:bg-slate-900 text-white text-xs px-2.5 py-1 rounded-lg font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
            >
              <UserCog className="w-3 h-3" />
              <span>เปลี่ยนวัน / เปลี่ยนช่าง</span>
            </button>
          )}
        </div>
      );
    }

    if (status === 'PENDING_PARTS_APPROVAL') {
      return (
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-semibold text-amber-700">สถานะ: รออนุมัติอะไหล่</span>
          {userRole === 'SUPERVISOR' && (
            <>
              <button
                type="button"
                onClick={() => openApproveModal(t, 'PARTS')}
                className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1 rounded-md font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>อนุมัติอะไหล่ (เลือกชื่อผู้อนุมัติ)</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setRejectModalTicket(t);
                  setRejectReasonInput('');
                }}
                className="bg-red-600 hover:bg-red-700 text-white text-xs px-2.5 py-1 rounded-md font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <XCircle className="w-3.5 h-3.5" />
                <span>ไม่อนุมัติ</span>
              </button>
            </>
          )}
        </div>
      );
    }

    if (status === 'PENDING_SUPERVISOR_APPROVAL') {
      return (
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-semibold text-purple-700">สถานะ: รออนุมัติรายงาน</span>
          {userRole === 'SUPERVISOR' && (
            <>
              <button
                type="button"
                onClick={() => openApproveModal(t, 'REPORT')}
                className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-2.5 py-1 rounded-md font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>อนุมัติรายงาน (เลือกชื่อผู้อนุมัติ)</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setRejectModalTicket(t);
                  setRejectReasonInput('');
                }}
                className="bg-red-600 hover:bg-red-700 text-white text-xs px-2.5 py-1 rounded-md font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <XCircle className="w-3.5 h-3.5" />
                <span>ส่งกลับ</span>
              </button>
            </>
          )}
        </div>
      );
    }

    if (status === 'REJECTED') {
      return (
        <span className="text-xs font-semibold text-red-600">
          ส่งกลับแก้ไข: {t.rejectReason || 'โปรดตรวจสอบและแก้ไขข้อมูล'}
        </span>
      );
    }

    if (status === 'CUSTOMER_ACCEPTANCE') {
      return (
        <span className="text-xs font-semibold text-cyan-700">
          สถานะ: รอตรวจรับงานจากลูกค้า
        </span>
      );
    }

    if (status === 'CLOSE') {
      return (
        <span className="text-xs font-semibold text-emerald-700">
          สถานะ: ปิดงานสำเร็จ (ล็อกข้อมูล)
        </span>
      );
    }

    return <span className="text-xs font-semibold text-slate-700">สถานะ: {status}</span>;
  };

  const activeReportTicket = useMemo(
    () => tickets.find((t) => t.id === activeReportTicketId) || tickets[0],
    [tickets, activeReportTicketId]
  );

  return (
    <div
      className={`min-h-screen flex flex-col transition-colors duration-200 ${
        themeMode === 'dark' ? 'dark-theme bg-slate-950 text-slate-100' : 'bg-slate-50 text-slate-900'
      }`}
    >
      {/* Toast Notification Banner */}
      {toastMessage && (
        <div className="fixed bottom-5 right-5 z-50 max-w-md bg-slate-900 text-white px-4 py-3 rounded-xl shadow-lg border border-slate-700 flex items-center gap-3 text-xs sm:text-sm no-print">
          {toastMessage.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />}
          {toastMessage.type === 'warning' && <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />}
          {toastMessage.type === 'error' && <XCircle className="w-4 h-4 text-red-400 shrink-0" />}
          <span>{toastMessage.text}</span>
        </div>
      )}

      {/* Top Header following 3-Zone Top Bar Contract + Role Identity Color & Light/Dark Toggle */}
      <header className={`${roleThemeConfig.headerBg} text-white border-b-2 sticky top-0 z-40 no-print transition-colors duration-200`}>
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex items-center justify-between gap-4">
          {/* Zone 1: Brand Wordmark + Optional Company Logo + Role Color Pill */}
          <div className="flex items-center gap-3 shrink-0">
            <button
              type="button"
              onClick={() => setCurrentView('home')}
              className="flex items-center gap-2.5 text-left hover:text-orange-400 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <img
                src={companyProfile.logoUrl || DEFAULT_SAHAKIM_LOGO}
                alt={companyProfile.nameTh}
                className="w-10 h-10 rounded-lg object-contain bg-white p-0.5 border border-slate-700 shadow-xs shrink-0"
              />
              <div>
                <span className="block text-sm sm:text-base font-bold tracking-tight text-white leading-tight">
                  {companyProfile.nameTh}
                </span>
                <span className="block text-[10px] font-semibold text-orange-400">
                  {companyProfile.nameEn}
                </span>
              </div>
            </button>
            {(currentView === 'customer_portal' ||
              currentView === 'internal_dashboard' ||
              currentView === 'internal_login') && (
              <span
                className={`hidden xl:inline-flex items-center px-2.5 py-0.5 rounded-full text-[10px] font-bold tracking-wide ${roleThemeConfig.badgeBg}`}
              >
                {roleThemeConfig.roleColorName}
              </span>
            )}
          </div>

          {/* Zone 2: Navigation Links */}
          <nav className="hidden md:flex items-center gap-5 text-sm font-medium text-slate-300">
            <button
              type="button"
              onClick={() => setCurrentView('home')}
              className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                currentView === 'home' ? 'text-white underline underline-offset-8 decoration-orange-500' : ''
              }`}
            >
              หน้าหลัก
            </button>
            <button
              type="button"
              onClick={() => openCustomerPortalPrompt('list')}
              className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                currentView === 'customer_portal'
                  ? 'text-white underline underline-offset-8 decoration-orange-500'
                  : ''
              }`}
            >
              Service Request
            </button>
            <button
              type="button"
              onClick={() =>
                userRole ? setCurrentView('internal_dashboard') : setCurrentView('internal_login')
              }
              className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                currentView === 'internal_dashboard' || currentView === 'internal_login'
                  ? 'text-white underline underline-offset-8 decoration-orange-500'
                  : ''
              }`}
            >
              เจ้าหน้าที่ภายใน
            </button>
            <button
              type="button"
              onClick={() => openServiceReportView(tickets[0]?.id || 'SKM-69-001')}
              className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                currentView === 'service_report'
                  ? 'text-white underline underline-offset-8 decoration-orange-500'
                  : ''
              }`}
            >
              ตัวอย่างใบรายงาน
            </button>
            {userRole === 'SUPERVISOR' && (
              <button
                type="button"
                onClick={() => setMasterDataModalOpen(true)}
                className="hover:text-orange-400 text-emerald-300 font-semibold transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1 cursor-pointer"
              >
                <Edit3 className="w-3.5 h-3.5" />
                <span>ตั้งค่าข้อมูลตั้งต้น</span>
              </button>
            )}
            <button
              type="button"
              onClick={() => setUserManualOpen(true)}
              className="hover:text-orange-400 text-amber-300 font-semibold transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 cursor-pointer"
            >
              <BookOpen className="w-4 h-4" />
              <span>คู่มือการใช้งาน</span>
            </button>
          </nav>

          {/* Zone 3: Screen Theme Toggle (ขาว / ดำ) + Primary Action */}
          <div className="flex items-center gap-2 shrink-0">
            {/* Light / Dark Screen Theme Switcher */}
            <div className="flex items-center bg-slate-800/90 p-0.5 rounded-lg border border-slate-700 text-xs">
              <button
                type="button"
                onClick={() => {
                  setThemeMode('light');
                  showToast('เปลี่ยนเป็นโหมดหน้าจอสีขาว (Light Mode)', 'success');
                }}
                className={`px-2.5 py-1.5 rounded-md font-bold transition-colors whitespace-nowrap cursor-pointer ${
                  themeMode === 'light'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
                title="โหมดหน้าจอสีขาว (เหมาะกับงานออฟฟิศและตรวจเอกสาร)"
              >
                ☀️ ขาว
              </button>
              <button
                type="button"
                onClick={() => {
                  setThemeMode('dark');
                  showToast('เปลี่ยนเป็นโหมดหน้าจอสีดำ (Dark Mode)', 'success');
                }}
                className={`px-2.5 py-1.5 rounded-md font-bold transition-colors whitespace-nowrap cursor-pointer ${
                  themeMode === 'dark'
                    ? 'bg-orange-500 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
                title="โหมดหน้าจอสีดำ (ถนอมสายตาและประหยัดแบตเตอรี่มือถือหน้างาน)"
              >
                🌙 ดำ
              </button>
            </div>

            <button
              type="button"
              onClick={handleOneClickOnlineUpdate}
              disabled={onlineSyncStatus.isSyncing}
              className="px-3 py-2 text-xs font-bold text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 disabled:opacity-60 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
              title="กดเพื่ออัปเดตและซิงค์ข้อมูลล่าสุดให้ตรงกันทุกเครื่องออนไลน์ทันที"
            >
              {onlineSyncStatus.isSyncing ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <CheckCircle2 className="w-3.5 h-3.5" />
              )}
              <span>🔄 อัปเดตออนไลน์</span>
            </button>

            <button
              type="button"
              onClick={() => setResetConfirmModalOpen(true)}
              className="px-2.5 sm:px-3 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-lg transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1 cursor-pointer shadow-xs"
              title="รีเซ็ตล้างใบงานตัวอย่าง เพื่อเริ่มกรอกใบงานจริงทันที (ตารางช่างจะว่างเขียวทั้งหมด)"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>รีเซ็ตเริ่มงานจริง</span>
            </button>

            <button
              type="button"
              onClick={() => setUserManualOpen(true)}
              className="md:hidden px-2.5 py-2 text-xs font-semibold text-amber-300 bg-slate-800 rounded-lg hover:bg-slate-700 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1 cursor-pointer"
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>คู่มือ</span>
            </button>
            {currentView === 'internal_dashboard' && userRole ? (
              <button
                type="button"
                onClick={() => {
                  setUserRole(null);
                  setCurrentView('home');
                  showToast('ออกจากระบบเรียบร้อยแล้ว', 'success');
                }}
                className="px-3.5 py-2 text-xs font-semibold text-white bg-red-600 rounded-lg hover:bg-red-700 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 cursor-pointer"
              >
                <LogOut className="w-3.5 h-3.5" />
                <span>ออก ({userRole})</span>
              </button>
            ) : (
              <button
                type="button"
                onClick={() =>
                  currentView === 'home' ? openCustomerPortalPrompt('list') : setCurrentView('home')
                }
                className="px-3.5 py-2 text-xs font-semibold text-white bg-orange-500 rounded-lg hover:bg-orange-600 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                {currentView === 'home' ? 'Service Request' : 'กลับหน้าหลัก'}
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 py-6 sm:py-8">
        {/* LIVE ONLINE SYNC & QUICK UPDATE BAR (เฉพาะสิทธิ์ SUPERVISOR เท่านั้น) */}
        {userRole === 'SUPERVISOR' && (
          <div className="mb-6 bg-white border-2 border-emerald-500/70 rounded-2xl p-3 sm:p-4 shadow-xs flex flex-col lg:flex-row items-start lg:items-center justify-between gap-3 no-print">
            <div className="flex flex-wrap items-center gap-2.5 text-xs">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-900 font-extrabold border border-emerald-300">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse inline-block" />
                <span>ONLINE REAL-TIME SYNC</span>
              </span>
              <span className="font-bold text-slate-900">
                ระบบซิงค์ข้อมูลออนไลน์อัตโนมัติ (ใบงาน {tickets.length} รายการ · ช่าง {technicians.length} คน)
              </span>
              <span className="text-slate-500 font-mono text-[11px]">
                (อัปเดตล่าสุด: {onlineSyncStatus.lastSyncTime})
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2 text-xs w-full lg:w-auto">
              <button
                type="button"
                onClick={handleOneClickOnlineUpdate}
                disabled={onlineSyncStatus.isSyncing}
                className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white px-3.5 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
              >
                {onlineSyncStatus.isSyncing ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <CheckCircle2 className="w-3.5 h-3.5" />
                )}
                <span>🔄 กดอัปเดตข้อมูลออนไลน์ทันที (Sync ทุกเครื่อง)</span>
              </button>

              <button
                type="button"
                onClick={() => setIsGoogleSheetsModalOpen(true)}
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
                title="จัดการและเชื่อมต่อ Google Sheets ผ่าน Google Workspace OAuth"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>
                  {googleSheetId ? '📊 Google Sheet (เชื่อมต่อแล้ว)' : '📊 เชื่อมต่อ Google Sheets'}
                </span>
              </button>

              <button
                type="button"
                onClick={() => setMasterDataModalOpen(true)}
                className="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer"
              >
                <Edit3 className="w-3.5 h-3.5" />
                <span>⚙️ แก้ไขข้อมูลตั้งต้น</span>
              </button>

              <button
                type="button"
                onClick={handleOpenGithubExportModal}
                className="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1.5 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
                title="ดาวน์โหลดไฟล์ index.html หรือคัดลอกโค้ดเพื่อนำไปอัปเดตบน GitHub Pages"
              >
                <Download className="w-3.5 h-3.5" />
                <span>⬇️ ดาวน์โหลด index.html (อัปเดต GitHub)</span>
              </button>

              <button
                type="button"
                onClick={exportFullSystemDataJson}
                className="bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 px-2.5 py-1.5 rounded-xl font-semibold cursor-pointer"
                title="ดาวน์โหลดไฟล์ข้อมูลทั้งหมด (.json) เพื่อสำรองหรือส่งต่อ"
              >
                📤 สำรองไฟล์ข้อมูล
              </button>

              <label
                className="bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 px-2.5 py-1.5 rounded-xl font-semibold cursor-pointer inline-flex items-center"
                title="นำเข้าไฟล์ข้อมูล (.json) เพื่ออัปเดตเข้าระบบออนไลน์"
              >
                <span>📥 นำเข้าไฟล์ข้อมูล</span>
                <input
                  type="file"
                  accept=".json,application/json"
                  onChange={handleImportFullSystemDataJson}
                  className="hidden"
                />
              </label>
            </div>
          </div>
        )}
        {/* VIEW 1: HOME */}
        {currentView === 'home' && (
          <div className="py-6 sm:py-10 max-w-5xl mx-auto space-y-8">
            <div className="text-center space-y-2.5 max-w-2xl mx-auto">
              <p className="text-xs font-semibold tracking-wider text-orange-600">
                SAHA KIM MOTOR CO., LTD. · ENTERPRISE SERVICE PORTAL V12.1
              </p>
              <h1
                className="text-2xl sm:text-4xl font-bold text-slate-900 tracking-tight"
                style={{ textWrap: 'balance' }}
              >
                ระบบบริหารงานบริการหลังการขายและรับประกันคุณภาพ
              </h1>
              <p className="text-sm sm:text-base text-slate-600 pt-1">
                รองรับงานซ่อม บำรุงรักษา และวิเคราะห์อาการเสียตามมาตรฐาน Karassik&apos;s Pump Handbook &amp; EASA Motor สำหรับปั๊มน้ำ มอเตอร์ Blower, Gen Set, Sewage, Aerator และชุด Control
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Customer Portal Entry Card */}
              <div
                onClick={() => openCustomerPortalPrompt('list')}
                className="bg-white p-6 sm:p-8 rounded-2xl border border-slate-200 hover:border-orange-500 transition-colors cursor-pointer flex flex-col justify-between group"
              >
                <div className="space-y-4">
                  <div className="w-14 h-14 bg-orange-500 text-white rounded-xl flex items-center justify-center shadow-xs">
                    <Users className="w-7 h-7" />
                  </div>
                  <div className="space-y-1.5">
                    <h2 className="text-xl font-bold text-slate-900">
                      1. Service Request (สำหรับลูกค้า)
                    </h2>
                    <p className="text-slate-600 text-sm leading-relaxed">
                      ติดตามสถานะงานซ่อม แจ้งซ่อมใหม่ และตรวจสอบตารางคิวงานช่าง
                    </p>
                  </div>
                </div>
                <div className="pt-5 mt-5 border-t border-slate-100 flex items-center justify-between text-sm font-semibold text-orange-600">
                  <span>เข้าสู่หน้า Service Request (สำหรับลูกค้า)</span>
                  <span className="group-hover:translate-x-1 transition-transform">→</span>
                </div>
              </div>

              {/* Internal Staff Entry Card */}
              <div
                onClick={() => openInternalSection('tickets')}
                className="bg-white p-6 sm:p-8 rounded-2xl border border-slate-200 hover:border-slate-900 transition-colors cursor-pointer flex flex-col justify-between group"
              >
                <div className="space-y-4">
                  <div className="flex items-center gap-3.5">
                    <img
                      src={companyProfile.staffMascotUrl || SKM_SERVICE_MASCOT_LOGO}
                      alt="เจ้าหน้าที่ภายใน SKM SERVICE เชี่ยวชาญงาน ใส่ใจบริการ"
                      className="w-20 h-20 sm:w-24 sm:h-24 rounded-2xl object-cover border-2 border-orange-500 bg-slate-900 shadow-sm shrink-0"
                    />
                    <div className="space-y-1">
                      <span className="inline-block text-[11px] font-bold text-orange-600 bg-orange-50 border border-orange-200 px-2.5 py-0.5 rounded-md">
                        SKM SERVICE · เชี่ยวชาญงาน ใส่ใจบริการ
                      </span>
                      <h2 className="text-xl font-bold text-slate-900">
                        2. เจ้าหน้าที่ภายใน (Internal Staff)
                      </h2>
                    </div>
                  </div>
                  <div className="space-y-1.5">
                    <p className="text-slate-600 text-sm leading-relaxed">
                      จัดการใบงานซ่อม ปฏิบัติงานช่าง ตารางงานช่าง สรุป Dashboard และระบบตั้งค่าสำหรับเจ้าหน้าที่
                    </p>
                  </div>
                </div>
                <div className="pt-5 mt-5 border-t border-slate-100 flex items-center justify-between text-sm font-semibold text-slate-900">
                  <span>เข้าสู่ระบบเจ้าหน้าที่ภายใน</span>
                  <span className="group-hover:translate-x-1 transition-transform">→</span>
                </div>
              </div>
            </div>

            {/* Summary Metrics Strip */}
            <div className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 grid grid-cols-2 lg:grid-cols-4 gap-6">
              <div>
                <p className="text-xs text-slate-500">ใบงานในระบบทั้งหมด</p>
                <p className="text-2xl font-bold text-slate-900 mt-1 tabular-nums">
                  {dashboardStats.totalTickets} รายการ
                </p>
                <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                  ปิดงานแล้ว {dashboardStats.closedTickets} · รอดำเนินการ{' '}
                  {dashboardStats.totalTickets - dashboardStats.closedTickets}
                </p>
              </div>
              <div>
                <p className="text-xs text-slate-500">สัดส่วนการรับประกัน</p>
                <p className="text-2xl font-bold text-slate-900 mt-1 tabular-nums">
                  FOC {dashboardStats.focCount} · Charge {dashboardStats.chargeCount}
                </p>
                <p className="text-xs text-slate-500 mt-0.5">
                  แยกตามเงื่อนไข In-Warranty และ Out-of-Warranty
                </p>
              </div>
              <div>
                <p className="text-xs text-slate-500">คลังต้นแบบวิเคราะห์ช่าง &amp; AI</p>
                <p className="text-2xl font-bold text-slate-900 mt-1 tabular-nums">
                  {knowledgeTemplates.length} มาตรฐาน
                </p>
                <p className="text-xs text-slate-500 mt-0.5">
                  อ้างอิง Pump Handbook, EASA AR100 &amp; คู่มือผู้ผลิต
                </p>
              </div>
              <div>
                <p className="text-xs text-slate-500">อุปกรณ์ที่แจ้งซ่อมสูงสุด</p>
                <p className="text-lg font-bold text-orange-600 mt-1 truncate">
                  {dashboardStats.topEquip}
                </p>
                <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                  มาตรฐาน SLA ตรงเวลา {dashboardStats.onTimeCount}/{dashboardStats.totalTickets} งาน
                </p>
              </div>
            </div>

            {/* QR Code & Link Card for Customer Equipment Sticker & LINE Sharing */}
            <div
              id="qr-card-print-area"
              className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 flex flex-col md:flex-row items-center justify-between gap-6"
            >
              <div className="space-y-2 flex-1">
                <p className="text-xs font-semibold text-orange-600">
                  QR CODE &amp; DIRECT LINK · สำหรับส่งเข้า LINE และพิมพ์สติกเกอร์ติดหน้าเครื่องจักร/ตู้คอนโทรล
                </p>
                <h3 className="text-base sm:text-lg font-bold text-slate-900">
                  ลิงก์และ QR Code ระบบแจ้งซ่อมออนไลน์ (บริษัท สหคิมมอเตอร์ จำกัด)
                </h3>
                <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                  แนะนำให้ใช้ <b>ควบคู่กันทั้ง 2 แบบ</b>: ส่งเป็น <b>ลิงก์</b> ในกลุ่ม LINE ให้ทีมงานและลูกค้ากดเข้าทันที และพิมพ์ <b>QR Code</b> เป็นสติกเกอร์แปะไว้ที่ตัวปั๊ม มอเตอร์ หรือหน้าตู้คอนโทรลเพื่อให้ช่างหน้างานสแกนแจ้งซ่อมได้ทันที
                </p>
                <div className="pt-2 flex flex-wrap items-center gap-2.5 no-print">
                  <button
                    type="button"
                    onClick={() => {
                      const shareUrl = 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
                      if (navigator.clipboard && navigator.clipboard.writeText) {
                        navigator.clipboard.writeText(shareUrl);
                      }
                      showToast('คัดลอกลิงก์ https://2020skmproduct-service.github.io/sahakim-service-portal/ เรียบร้อย!', 'success');
                    }}
                    className="bg-slate-900 hover:bg-slate-800 text-white px-4 py-2.5 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 cursor-pointer"
                  >
                    <Copy className="w-3.5 h-3.5" />
                    <span>คัดลอกลิงก์ส่ง LINE</span>
                  </button>
                  <button
                    type="button"
                    disabled={isGeneratingPdf}
                    onClick={() =>
                      downloadElementAsPdf('qr-card-print-area', 'Sahakim-QR-Service-Sticker.pdf')
                    }
                    className="bg-orange-500 hover:bg-orange-600 disabled:opacity-60 text-white px-4 py-2.5 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 cursor-pointer"
                  >
                    {isGeneratingPdf ? (
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Download className="w-3.5 h-3.5" />
                    )}
                    <span>Save PDF (ป้าย QR Code)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => window.print()}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-3.5 py-2.5 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 cursor-pointer"
                  >
                    <Printer className="w-3.5 h-3.5" />
                    <span>สั่งพิมพ์</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setUserManualOpen(true)}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 cursor-pointer"
                  >
                    <BookOpen className="w-3.5 h-3.5" />
                    <span>เปิดคู่มือการใช้งานระบบ (User Manual)</span>
                  </button>
                </div>
              </div>
              <div className="border-2 border-slate-900 rounded-2xl p-4 bg-slate-50 text-center shrink-0">
                <p className="text-[10px] font-bold text-orange-600">SAHA KIM MOTOR SERVICE</p>
                <img
                  crossOrigin="anonymous"
                  src={`https://api.qrserver.com/v1/create-qr-code/?size=200x200&margin=8&data=${encodeURIComponent(
                    'https://2020skmproduct-service.github.io/sahakim-service-portal/'
                  )}`}
                  alt="QR Code แจ้งซ่อม สหคิมมอเตอร์"
                  className="w-36 h-36 mx-auto my-2 rounded-lg border border-slate-200 bg-white p-1"
                />
                <p className="text-[11px] font-bold text-slate-800">สแกนแจ้งซ่อม / ติดตามงาน</p>
              </div>
            </div>
          </div>
        )}

        {/* VIEW 2: SERVICE REQUEST (CUSTOMER PORTAL) */}
        {currentView === 'customer_portal' && (
          <div className="space-y-6">
            {/* Top Banner & Search by Company Name or Ticket ID */}
            <div className={`${roleThemeConfig.bannerBg} text-white p-5 sm:p-6 rounded-2xl space-y-4 transition-colors`}>
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                <div className="space-y-1">
                  <p className="text-xs text-orange-400 font-bold tracking-wide uppercase">
                    SERVICE REQUEST · ระบบแจ้งซ่อม จองคิวช่าง และติดตามสถานะงาน
                  </p>
                  <h2 className="text-xl sm:text-2xl font-bold">
                    Service Request — {customerCompany}
                  </h2>
                  <p className="text-xs text-slate-300">
                    ตรวจสอบสถานะงานซ่อมด้วยชื่อบริษัทหรือเลขที่ใบแจ้งซ่อม และเลือกวันขอช่างเบื้องต้นจากตารางคิวงานช่างด้านล่าง
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => openCustomerPortalPrompt('list')}
                    className="bg-white text-slate-900 hover:bg-slate-100 font-semibold px-3.5 py-2 rounded-lg text-xs transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    ค้นหาบริษัทอื่น / เลขใบงานอื่น
                  </button>
                </div>
              </div>

              {/* Direct Search Bar for Company Name or Ticket ID */}
              <div className="bg-slate-800/90 border border-slate-700 rounded-xl p-3.5 space-y-2">
                <label className="block text-xs font-semibold text-orange-300">
                  🔍 ตรวจสอบสถานะงานของตัวเอง (พิมพ์ ชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม เช่น SKM-69-001):
                </label>
                <div className="flex flex-col sm:flex-row gap-2">
                  <div className="relative flex-1">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={customerLookupQuery}
                      onChange={(e) => applyCustomerLookup(e.target.value)}
                      placeholder="พิมพ์ชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม เช่น SKM-69-001..."
                      className="w-full bg-white text-slate-900 rounded-lg pl-9 pr-3 py-2 text-xs sm:text-sm font-medium focus:outline-none focus:ring-2 focus:ring-orange-500"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      applyCustomerLookup(customerLookupQuery);
                      setCustTab('list');
                      showToast(`แสดงสถานะงานสำหรับ "${customerLookupQuery}"`, 'success');
                    }}
                    className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-lg text-xs sm:text-sm font-bold whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    ค้นหาสถานะงาน
                  </button>
                </div>
              </div>
            </div>

            {/* CUSTOMER PORTAL CLICKABLE MENU BAR (เลือกกดเข้าดูทีละหมวด) */}
            <div className="bg-white p-3 sm:p-4 rounded-2xl border-2 border-slate-200 shadow-xs space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2 px-1">
                <p className="text-xs font-bold text-slate-700">
                  📌 เมนูเลือกใช้งาน (กดเลือกหมวดที่ต้องการดู):
                </p>
                <span className="text-[11px] text-slate-500">
                  แยกหน้าจอให้ดูง่าย ไม่ต้องเลื่อนหน้าจอยาว
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                <button
                  type="button"
                  onClick={() => setCustTab('list')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-3 cursor-pointer ${
                    custTab === 'list'
                      ? 'bg-slate-900 border-slate-900 text-white shadow-sm'
                      : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-800'
                  }`}
                >
                  <div className="space-y-0.5">
                    <p className="text-xs sm:text-sm font-bold">1. 🔍 ติดตามสถานะงานซ่อม</p>
                    <p
                      className={`text-[11px] ${
                        custTab === 'list' ? 'text-orange-300' : 'text-slate-500'
                      }`}
                    >
                      ดูความคืบหน้า, ใบรายงาน &amp; เซ็นรับงาน
                    </p>
                  </div>
                  <span
                    className={`px-2.5 py-1 rounded-full text-xs font-mono font-bold shrink-0 ${
                      custTab === 'list'
                        ? 'bg-orange-500 text-white'
                        : 'bg-slate-200 text-slate-800'
                    }`}
                  >
                    {customerMatchedTickets.length} งาน
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setNewCompany(customerCompany);
                    setCustTab('new');
                  }}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-3 cursor-pointer ${
                    custTab === 'new'
                      ? 'bg-orange-500 border-orange-600 text-white shadow-sm'
                      : 'bg-orange-50/60 hover:bg-orange-100/70 border-orange-200 text-slate-900'
                  }`}
                >
                  <div className="space-y-0.5">
                    <p className="text-xs sm:text-sm font-bold">2. ➕ แจ้งซ่อมใหม่ (Service Request)</p>
                    <p
                      className={`text-[11px] ${
                        custTab === 'new' ? 'text-orange-100' : 'text-slate-600'
                      }`}
                    >
                      กรอกอาการเสีย อัปโหลดรูป &amp; ระบุวันนัด
                    </p>
                  </div>
                  <Plus className="w-5 h-5 shrink-0" />
                </button>

                <button
                  type="button"
                  onClick={() => setCustTab('schedule')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-3 cursor-pointer ${
                    custTab === 'schedule'
                      ? 'bg-emerald-700 border-emerald-800 text-white shadow-sm'
                      : 'bg-emerald-50/60 hover:bg-emerald-100/70 border-emerald-200 text-slate-900'
                  }`}
                >
                  <div className="space-y-0.5">
                    <p className="text-xs sm:text-sm font-bold">3. 📅 ตารางงานช่าง (คิวว่าง)</p>
                    <p
                      className={`text-[11px] ${
                        custTab === 'schedule' ? 'text-emerald-100' : 'text-slate-600'
                      }`}
                    >
                      เช็คคิวสีเขียว/เหลือง/แดง เพื่อจองวันขอช่าง
                    </p>
                  </div>
                  <span className="text-xs font-bold shrink-0">🟢🟡🔴</span>
                </button>
              </div>
            </div>

            {/* TAB 3: TECHNICIAN SCHEDULE & AVAILABILITY CALENDAR (แสดงเฉพาะเมื่อกดเมนู "ตารางงานช่าง") */}
            {custTab === 'schedule' && (
              <div className="bg-white p-4 sm:p-6 rounded-2xl border-2 border-slate-200 shadow-xs space-y-4">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 border-b border-slate-200 pb-3">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="bg-orange-100 text-orange-700 text-[11px] font-bold px-2.5 py-0.5 rounded-full">
                        TECHNICIAN QUEUE
                      </span>
                      <h3 className="text-base sm:text-lg font-bold text-slate-900">
                        ตารางงานช่าง (คลิกที่วันเพื่อเลือกระบุวันขอช่างเบื้องต้น)
                      </h3>
                    </div>
                    <p className="text-xs text-slate-600 mt-0.5">
                      ตรวจสอบคิวว่างของทีมช่างสหคิมมอเตอร์ก่อนนัดหมาย — คลิกเลือกวันที่ต้องการเพื่อเปิดฟอร์มแจ้งซ่อมพร้อมใส่วันที่ให้อัตโนมัติ
                    </p>
                  </div>

                  {/* Color Legend: Red / Yellow / Green */}
                  <div className="flex flex-wrap items-center gap-2 text-xs">
                    <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-300 text-emerald-900 font-bold">
                      <span className="w-3 h-3 rounded-full bg-emerald-500 inline-block" />
                      สีเขียว = ว่าง
                    </span>
                    <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-50 border border-amber-300 text-amber-900 font-bold">
                      <span className="w-3 h-3 rounded-full bg-amber-400 inline-block" />
                      สีเหลือง = รับงานได้
                    </span>
                    <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-red-50 border border-red-300 text-red-900 font-bold">
                      <span className="w-3 h-3 rounded-full bg-red-500 inline-block" />
                      สีแดง = เต็มมาก
                    </span>
                  </div>
                </div>

                {/* Week Navigation Controls */}
                <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => setScheduleDayOffset((prev) => Math.max(0, prev - 7))}
                      disabled={scheduleDayOffset === 0}
                      className="px-3 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 font-semibold text-slate-700 cursor-pointer"
                    >
                      ← 7 วันก่อนหน้า
                    </button>
                    <button
                      type="button"
                      onClick={() => setScheduleDayOffset(0)}
                      className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 font-semibold text-slate-800 cursor-pointer"
                    >
                      วันนี้
                    </button>
                    <button
                      type="button"
                      onClick={() => setScheduleDayOffset((prev) => prev + 7)}
                      className="px-3 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 font-semibold text-slate-700 cursor-pointer"
                    >
                      7 วันถัดไป →
                    </button>
                  </div>
                  <span className="text-slate-500 font-medium">
                    แสดงตารางคิวช่าง 14 วันล่วงหน้า (คลิกเลือกวันด้านล่างเพื่อขอคิวช่างเบื้องต้น)
                  </span>
                </div>

                {/* 14-Day Color-Coded Grid */}
                <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5">
                  {customerScheduleDays.map((day) => {
                    const isSelected = newDesiredDate === day.dateStr;
                    const cardColor =
                      day.level === 'RED'
                        ? 'bg-red-50 border-red-400 hover:bg-red-100 text-red-950'
                        : day.level === 'YELLOW'
                        ? 'bg-amber-50 border-amber-400 hover:bg-amber-100 text-amber-950'
                        : 'bg-emerald-50 border-emerald-400 hover:bg-emerald-100 text-emerald-950';
                    const badgeColor =
                      day.level === 'RED'
                        ? 'bg-red-600 text-white'
                        : day.level === 'YELLOW'
                        ? 'bg-amber-500 text-white'
                        : 'bg-emerald-600 text-white';
                    const dotColor =
                      day.level === 'RED'
                        ? 'bg-red-500'
                        : day.level === 'YELLOW'
                        ? 'bg-amber-400'
                        : 'bg-emerald-500';

                    return (
                      <button
                        key={day.dateStr}
                        type="button"
                        onClick={() => {
                          setNewDesiredDate(day.dateStr);
                          setNewCompany(customerCompany);
                          setCustTab('new');
                          if (day.level === 'RED') {
                            showToast(
                              `เลือกวันที่ ${day.dateStr} (สีแดง: คิวเต็มมาก) — เจ้าหน้าที่อาจติดต่อเพื่อนัดหมายเวลาที่เหมาะสมอีกครั้ง`,
                              'warning'
                            );
                          } else if (day.level === 'YELLOW') {
                            showToast(
                              `เลือกวันขอช่างเบื้องต้น ${day.dateStr} (สีเหลือง: รับงานได้) เรียบร้อยแล้ว`,
                              'success'
                            );
                          } else {
                            showToast(
                              `เลือกวันขอช่างเบื้องต้น ${day.dateStr} (สีเขียว: คิวว่าง) เรียบร้อยแล้ว`,
                              'success'
                            );
                          }
                        }}
                        className={`p-3 rounded-xl border-2 text-left transition-all flex flex-col justify-between gap-2 cursor-pointer ${cardColor} ${
                          isSelected ? 'ring-3 ring-slate-900 shadow-md scale-[1.02]' : ''
                        }`}
                      >
                        <div className="flex items-start justify-between gap-1">
                          <div>
                            <p className="text-xs font-extrabold">{day.dayLabel}</p>
                            <p className="text-[10px] font-mono opacity-75">{day.dateStr}</p>
                          </div>
                          <span className={`w-3 h-3 rounded-full shrink-0 mt-0.5 ${dotColor}`} />
                        </div>

                        <div className="space-y-1">
                          <span
                            className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded-md ${badgeColor}`}
                          >
                            {day.statusText}
                          </span>
                          <p className="text-[11px] font-semibold leading-tight">{day.detailText}</p>
                          {isSelected && (
                            <p className="text-[10px] font-extrabold text-slate-900 underline">
                              ✓ วันนัดที่เลือกไว้
                            </p>
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>

                {/* Selected Date Summary Bar */}
                <div className="bg-slate-900 text-white p-3.5 rounded-xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="font-bold text-orange-400">📅 วันขอช่างเบื้องต้นที่เลือก:</span>
                    <span className="font-mono font-bold bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">
                      {newDesiredDate}
                    </span>
                    {(() => {
                      const matchedDay = customerScheduleDays.find((d) => d.dateStr === newDesiredDate);
                      if (!matchedDay) return null;
                      return (
                        <span
                          className={`px-2.5 py-0.5 rounded-full font-bold text-[11px] ${
                            matchedDay.level === 'RED'
                              ? 'bg-red-600 text-white'
                              : matchedDay.level === 'YELLOW'
                              ? 'bg-amber-500 text-white'
                              : 'bg-emerald-600 text-white'
                          }`}
                        >
                          สถานะคิวช่าง: {matchedDay.statusText} ({matchedDay.detailText})
                        </span>
                      );
                    })()}
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setNewCompany(customerCompany);
                      setCustTab('new');
                    }}
                    className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-lg font-bold whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    + ไปหน้าแจ้งซ่อมใหม่ (Service Request) ในวันที่เลือก →
                  </button>
                </div>
              </div>
            )}

            {custTab === 'list' && (
              <div className="space-y-3">
                {customerMatchedTickets.length === 0 ? (
                  <div className="bg-white p-10 rounded-2xl text-center border border-slate-200 space-y-3">
                    <p className="text-slate-600 text-sm">
                      ไม่พบใบแจ้งซ่อมสำหรับคำค้นหา &ldquo;{customerLookupQuery || customerCompany}&rdquo;
                    </p>
                    <p className="text-xs text-slate-500">
                      ท่านสามารถคลิกเลือกวันว่าง (สีเขียว/สีเหลือง) จากตารางงานช่างด้านบน หรือกดปุ่มด้านล่างเพื่อเปิดใบ Service Request ใหม่ได้ทันที
                    </p>
                    <button
                      type="button"
                      onClick={() => {
                        setNewCompany(customerCompany);
                        setCustTab('new');
                      }}
                      className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-lg text-xs font-semibold transition-colors cursor-pointer"
                    >
                      + สร้างใบ Service Request ใหม่
                    </button>
                  </div>
                ) : (
                  customerMatchedTickets.map((t) => {
                    const stepIndex =
                      t.status === 'NEW'
                        ? 1
                        : t.status === 'ASSIGNED'
                        ? 2
                        : t.status === 'CLOSE'
                        ? 4
                        : 3;

                    return (
                      <div
                        key={t.id}
                        className="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 flex flex-col gap-4"
                      >
                        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                          <div className="space-y-1.5 text-xs sm:text-sm">
                            <div className="flex flex-wrap items-center gap-2">
                              <span className="font-mono font-bold text-orange-600 tabular-nums text-base">
                                {t.id}
                              </span>
                              <span className="text-xs font-bold bg-slate-100 text-slate-800 px-2.5 py-0.5 rounded-md">
                                {t.company}
                              </span>
                              <span className="text-xs font-bold px-2.5 py-0.5 rounded-full bg-orange-100 text-orange-800">
                                สถานะ: {t.status}
                              </span>
                              <span className="text-xs text-slate-500 tabular-nums">
                                แจ้งเมื่อ {t.createdAt}
                              </span>
                            </div>
                            <p className="text-slate-800">
                              <span className="font-semibold">ประเภทบริการ:</span> {t.serviceType}
                              <span className="mx-2 text-slate-300">·</span>
                              <span className="font-semibold">การรับประกัน:</span>{' '}
                              <span className="text-slate-900 font-medium">{t.warranty}</span>
                            </p>
                            {(() => {
                              const costSummary = getTicketCostSummary(t);
                              return (
                                <div className="pt-0.5">
                                  {costSummary.isTicketInWarranty ||
                                  (costSummary.customerBilledTotal === 0 &&
                                    costSummary.internalCostTotal > 0) ? (
                                    <span className="inline-flex flex-wrap items-center gap-1.5 bg-emerald-50 border border-emerald-300 text-emerald-900 px-3 py-1 rounded-lg text-xs font-bold">
                                      <span>
                                        🛡️ อยู่ในเงื่อนไขรับประกัน (In-Warranty / FOC) · ยอดเรียกเก็บลูกค้าสุทธิ: 0.00 บาท (ไม่ต้องชำระค่าใช้จ่าย)
                                      </span>
                                      {costSummary.internalCostTotal > 0 && (
                                        <span className="text-[11px] font-medium text-emerald-700">
                                          [ต้นทุนอะไหล่ภายในบริษัทรับภาระ: {costSummary.internalCostTotal.toLocaleString()} THB]
                                        </span>
                                      )}
                                    </span>
                                  ) : (
                                    <span className="inline-flex flex-wrap items-center gap-1.5 bg-amber-50 border border-amber-300 text-amber-900 px-3 py-1 rounded-lg text-xs font-bold">
                                      <span>
                                        💳 งานนอกประกัน (Out-of-Warranty) · ยอดเรียกเก็บลูกค้าสุทธิ:{' '}
                                        {costSummary.customerBilledTotal.toLocaleString()} บาท
                                      </span>
                                    </span>
                                  )}
                                </div>
                              );
                            })()}
                            <p className="text-slate-600 text-xs">
                              วันขอช่าง / วันนัดหมายเข้างาน:{' '}
                              <span className="font-semibold text-emerald-700 tabular-nums">
                                {t.confirmedDate || `${t.desiredDate} (รอนัดหมายยืนยันคิวช่าง)`}
                              </span>
                              <span className="mx-2 text-slate-300">·</span>
                              สถานที่: {t.location || '-'}
                              <span className="mx-2 text-slate-300">·</span>
                              ช่างผู้รับผิดชอบ: {t.assignedTech || 'อยู่ระหว่างจัดสรรช่าง'}
                            </p>
                            <p className="text-slate-500 text-xs">
                              อุปกรณ์ ({t.equipmentList.length} รายการ):{' '}
                              {t.equipmentList
                                .map((eq) => `${eq.type} (${eq.brand} ${eq.model}) - ${eq.symptom}`)
                                .join(' / ')}
                            </p>
                          </div>

                          <div className="flex flex-wrap items-center gap-2 shrink-0">
                            <button
                              type="button"
                              onClick={() => openServiceReportView(t.id)}
                              className="bg-slate-900 hover:bg-slate-800 text-white px-3.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              <Printer className="w-3.5 h-3.5" />
                              <span>เปิดใบรายงาน (Service Report)</span>
                            </button>
                            <button
                              type="button"
                              onClick={() =>
                                setCustomerSignOffModalTicket({
                                  ...t,
                                  customerSignerName:
                                    t.customerSignerName || t.siteContact || t.contact
                                })
                              }
                              className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              <PenTool className="w-3.5 h-3.5" />
                              <span>
                                {t.customerSignature
                                  ? '✓ เซ็นรับงานแล้ว (ดู/แก้ไขลายเซ็น)'
                                  : t.status === 'CUSTOMER_ACCEPTANCE'
                                  ? 'เซ็นชื่อตรวจรับงานและปิดจบ'
                                  : 'เซ็นชื่อตรวจรับงานบนมือถือ'}
                              </span>
                            </button>
                          </div>
                        </div>

                        {/* 4-Step Visual Status Progress Bar for Customer */}
                        <div className="grid grid-cols-4 gap-2 pt-3 border-t border-slate-100 text-[11px]">
                          <div
                            className={`p-2 rounded-lg border text-center font-semibold ${
                              stepIndex >= 1
                                ? 'bg-blue-50 border-blue-300 text-blue-900'
                                : 'bg-slate-50 border-slate-200 text-slate-400'
                            }`}
                          >
                            1. รับแจ้งซ่อม (NEW)
                          </div>
                          <div
                            className={`p-2 rounded-lg border text-center font-semibold ${
                              stepIndex >= 2
                                ? 'bg-amber-50 border-amber-300 text-amber-900'
                                : 'bg-slate-50 border-slate-200 text-slate-400'
                            }`}
                          >
                            2. ยืนยันคิวช่าง (ASSIGNED)
                          </div>
                          <div
                            className={`p-2 rounded-lg border text-center font-semibold ${
                              stepIndex >= 3
                                ? 'bg-orange-50 border-orange-300 text-orange-900'
                                : 'bg-slate-50 border-slate-200 text-slate-400'
                            }`}
                          >
                            3. ดำเนินการซ่อม / ตรวจเช็ค
                          </div>
                          <div
                            className={`p-2 rounded-lg border text-center font-semibold ${
                              stepIndex >= 4
                                ? 'bg-emerald-50 border-emerald-400 text-emerald-900'
                                : t.status === 'CUSTOMER_ACCEPTANCE'
                                ? 'bg-cyan-50 border-cyan-400 text-cyan-900 font-bold'
                                : 'bg-slate-50 border-slate-200 text-slate-400'
                            }`}
                          >
                            {t.status === 'CUSTOMER_ACCEPTANCE'
                              ? '4. รอลูกค้าตรวจรับงาน'
                              : '4. ปิดงานสมบูรณ์ (CLOSE)'}
                          </div>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            )}

            {custTab === 'new' && (
              /* NEW TICKET FORM WITH MULTI-EQUIPMENT & CUSTOM TYPE/BRAND SUPPORT */
              <div className="bg-white p-5 sm:p-6 rounded-2xl border border-slate-200">
                <div className="pb-4 mb-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <h3 className="text-base sm:text-lg font-bold text-slate-900">
                      แบบฟอร์ม Service Request (แจ้งซ่อมและระบุวันขอช่างเบื้องต้น)
                    </h3>
                    <p className="text-xs text-slate-500">
                      กรอกข้อมูลอาการเสีย หรือกดปุ่มดูตารางงานช่างเพื่อเช็คคิวว่าง (สีเขียว/เหลือง/แดง) ก่อนเลือกวันนัดหมาย
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setCustTab('schedule')}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    📅 เปิดดูตารางงานช่าง (เช็คคิวว่าง)
                  </button>
                </div>

                <form onSubmit={handleSubmitNewTicket} className="space-y-5 text-xs sm:text-sm">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ชื่อบริษัท *
                      </label>
                      <input
                        type="text"
                        value={newCompany}
                        onChange={(e) => setNewCompany(e.target.value)}
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-slate-50 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เลขประจำตัวผู้เสียภาษี (TAX ID)
                      </label>
                      <input
                        type="text"
                        value={newTaxId}
                        onChange={(e) => setNewTaxId(e.target.value)}
                        placeholder="0105552012345"
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เลขที่เอกสาร / ใบรับประกัน (ถ้ามี)
                      </label>
                      <input
                        type="text"
                        value={newDocNo}
                        onChange={(e) => setNewDocNo(e.target.value)}
                        placeholder="เช่น DOC-2026-999"
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ชื่อผู้ติดต่อและเบอร์โทรศัพท์ *
                      </label>
                      <input
                        type="text"
                        value={newContact}
                        onChange={(e) => setNewContact(e.target.value)}
                        placeholder="คุณสมชาย (081-234-5678)"
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เงื่อนไขการรับประกัน / คิดค่าบริการ *
                      </label>
                      <select
                        value={newWarranty}
                        onChange={(e) => setNewWarranty(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-orange-600 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>In-Warranty (อยู่ในประกัน - ฟรี / FOC)</option>
                        <option>Out-of-Warranty (หมดประกัน - คิดค่าบริการ / Charge)</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ประเภทบริการ *
                      </label>
                      <select
                        value={newServiceType}
                        onChange={(e) => setNewServiceType(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>ซ่อมหน้างาน</option>
                        <option>รับมาซ่อม</option>
                        <option>งานตรวจเช็ค</option>
                        <option>คอมมิชชั่นนิ่ง (Commissioning)</option>
                        <option>เทสรัน (Test Run)</option>
                        <option>PM (Preventive Maintenance)</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ความเร่งด่วนของงาน *
                      </label>
                      <select
                        value={newUrgency}
                        onChange={(e) => setNewUrgency(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-red-600 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>ปกติ (Normal)</option>
                        <option>ด่วน (High)</option>
                        <option>ด่วนที่สุด (Urgent)</option>
                      </select>
                    </div>
                    <div>
                      <div className="flex flex-wrap items-center justify-between gap-1 mb-1">
                        <label className="block font-semibold text-slate-700">
                          วันที่ขอช่างเบื้องต้น (เลือกจากตารางงานช่างด้านบนได้) *
                        </label>
                        {(() => {
                          const matchedDay = customerScheduleDays.find(
                            (d) => d.dateStr === newDesiredDate
                          );
                          if (!matchedDay) return null;
                          return (
                            <span
                              className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                                matchedDay.level === 'RED'
                                  ? 'bg-red-100 text-red-800'
                                  : matchedDay.level === 'YELLOW'
                                  ? 'bg-amber-100 text-amber-800'
                                  : 'bg-emerald-100 text-emerald-800'
                              }`}
                            >
                              คิวช่าง: {matchedDay.statusText}
                            </span>
                          );
                        })()}
                      </div>
                      <input
                        type="date"
                        value={newDesiredDate}
                        onChange={(e) => setNewDesiredDate(e.target.value)}
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block font-semibold text-slate-700 mb-1">
                        สถานที่ปฏิบัติงาน / โลเคชั่นโรงงาน *
                      </label>
                      <input
                        type="text"
                        value={newLocation}
                        onChange={(e) => setNewLocation(e.target.value)}
                        placeholder="ระบุชื่อโรงงาน นิคมอุตสาหกรรม หรือพิกัดหน้างาน"
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block font-semibold text-slate-700 mb-1">
                        อัปโหลดรูปภาพอาการเสีย (ถ่ายรูปผ่านมือถือหรือเลือกไฟล์ภาพ)
                      </label>
                      <div className="flex flex-col sm:flex-row items-start sm:items-center gap-3">
                        <input
                          type="file"
                          accept="image/*"
                          capture="environment"
                          onChange={handleNewImageUpload}
                          className="w-full border border-slate-300 rounded-lg p-2 bg-slate-50 text-xs cursor-pointer"
                        />
                        {newUploadedPhoto && (
                          <ResilientImage
                            src={newUploadedPhoto}
                            alt="ตัวอย่างรูปที่อัปโหลด"
                            className="w-16 h-16 object-cover rounded-lg border border-slate-300 shrink-0"
                          />
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Dynamic Multi-Equipment Section */}
                  <div className="border-t border-slate-200 pt-4 space-y-3">
                    <div className="flex flex-wrap justify-between items-center gap-2">
                      <div>
                        <h4 className="font-bold text-slate-900 text-xs sm:text-sm">
                          รายการอุปกรณ์ที่ต้องการแจ้งซ่อม ({tempEquipmentList.length} รายการ)
                        </h4>
                        <p className="text-[11px] text-slate-500">
                          เลือก MOTOR, END SUCTION PUMP, VERTICAL PUMP, BOOSTER PUMP, TRANSFER PUMP, BLOWER, FIRE PUMP, GEN SET, SEWAGE, EJECTOR, AIR PUMP, AERATOR, CONTROL หรือเลือก ETC. เพื่อพิมพ์เพิ่มเอง
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() =>
                          setTempEquipmentList((prev) => [
                            ...prev,
                            {
                              type: 'END SUCTION PUMP',
                              customType: '',
                              brand: 'STAC',
                              customBrand: '',
                              model: '',
                              tag: '',
                              symptom: '',
                              detail: ''
                            }
                          ])
                        }
                        className="bg-slate-900 hover:bg-slate-800 text-white text-xs px-3.5 py-2 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>เพิ่มรายการอุปกรณ์</span>
                      </button>
                    </div>

                    <div className="space-y-3">
                      {tempEquipmentList.map((eq, index) => {
                        const isCustomType = eq.type.startsWith('ETC.');
                        const isCustomBrand = eq.brand.startsWith('ETC.');

                        return (
                          <div
                            key={index}
                            className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3 text-xs"
                          >
                            <div className="flex justify-between items-center border-b border-slate-200 pb-2">
                              <span className="font-bold text-orange-600 text-xs sm:text-sm">
                                อุปกรณ์รายการที่ #{index + 1}
                              </span>
                              {tempEquipmentList.length > 1 && (
                                <button
                                  type="button"
                                  onClick={() =>
                                    setTempEquipmentList((prev) =>
                                      prev.filter((_, i) => i !== index)
                                    )
                                  }
                                  className="text-red-600 hover:text-red-700 font-semibold inline-flex items-center gap-1 cursor-pointer"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                  <span>ลบรายการนี้</span>
                                </button>
                              )}
                            </div>

                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                              {/* Equipment Type Selector + Custom Input */}
                              <div className="space-y-1.5">
                                <label className="block font-semibold text-slate-700">
                                  ประเภทอุปกรณ์ (Equipment) *
                                </label>
                                <select
                                  value={eq.type}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, type: val } : item
                                      )
                                    );
                                  }}
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold text-slate-900"
                                >
                                  {equipmentOptions.map((m) => (
                                    <option key={m} value={m}>
                                      {m}
                                    </option>
                                  ))}
                                </select>
                                {isCustomType && (
                                  <input
                                    type="text"
                                    value={eq.customType}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setTempEquipmentList((prev) =>
                                        prev.map((item, i) =>
                                          i === index ? { ...item, customType: val } : item
                                        )
                                      );
                                    }}
                                    placeholder="ระบุประเภทอุปกรณ์เอง (เช่น CHILLER PUMP)..."
                                    required
                                    className="w-full border border-orange-400 rounded-lg p-2 bg-orange-50/40 font-semibold text-slate-900"
                                  />
                                )}
                              </div>

                              {/* Brand Selector + Custom Input */}
                              <div className="space-y-1.5">
                                <label className="block font-semibold text-slate-700">
                                  ยี่ห้อ / แบรนด์ (Brand) *
                                </label>
                                <select
                                  value={eq.brand}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, brand: val } : item
                                      )
                                    );
                                  }}
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold text-slate-900"
                                >
                                  {brandOptions.map((b) => (
                                    <option key={b} value={b}>
                                      {b}
                                    </option>
                                  ))}
                                </select>
                                {isCustomBrand && (
                                  <input
                                    type="text"
                                    value={eq.customBrand}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setTempEquipmentList((prev) =>
                                        prev.map((item, i) =>
                                          i === index ? { ...item, customBrand: val } : item
                                        )
                                      );
                                    }}
                                    placeholder="ระบุชื่อแบรนด์เอง (เช่น KSB, TSURUMI)..."
                                    required
                                    className="w-full border border-orange-400 rounded-lg p-2 bg-orange-50/40 font-semibold text-slate-900"
                                  />
                                )}
                              </div>

                              {/* Model */}
                              <div className="space-y-1.5">
                                <label className="block font-semibold text-slate-700">
                                  รุ่น (Model / kW / HP)
                                </label>
                                <input
                                  type="text"
                                  value={eq.model}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, model: val } : item
                                      )
                                    );
                                  }}
                                  placeholder="เช่น NK 65-200 / 15kW 4P"
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white"
                                />
                              </div>

                              {/* Tag No */}
                              <div className="space-y-1.5">
                                <label className="block font-semibold text-slate-700">
                                  หมายเลขเครื่อง / Tag No.
                                </label>
                                <input
                                  type="text"
                                  value={eq.tag}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, tag: val } : item
                                      )
                                    );
                                  }}
                                  placeholder="เช่น PMP-01A / S/N"
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white font-mono"
                                />
                              </div>
                            </div>

                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                              <div>
                                <label className="block font-semibold text-slate-700 mb-1">
                                  อาการเสียที่พบ (Symptom) *
                                </label>
                                <input
                                  type="text"
                                  value={eq.symptom}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, symptom: val } : item
                                      )
                                    );
                                  }}
                                  placeholder="ระบุอาการเสีย เช่น มอเตอร์ทริป Overload / ซีลรั่ว / เสียงดัง..."
                                  required
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white"
                                />
                              </div>
                              <div>
                                <label className="block font-semibold text-slate-700 mb-1">
                                  รายละเอียดหน้างานเพิ่มเติม (ถ้ามี)
                                </label>
                                <input
                                  type="text"
                                  value={eq.detail}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setTempEquipmentList((prev) =>
                                      prev.map((item, i) =>
                                        i === index ? { ...item, detail: val } : item
                                      )
                                    );
                                  }}
                                  placeholder="เช่น ค่ากระแสที่วัดได้ แรงดันน้ำ หรือสภาวะแวดล้อม"
                                  className="w-full border border-slate-300 rounded-lg p-2 bg-white"
                                />
                              </div>
                            </div>

                            {/* Before & After Photo Upload per Equipment (Camera/Album with auto-compression) */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2.5 border-t border-slate-200">
                              <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2">
                                <div className="flex items-center justify-between">
                                  <label className="text-xs font-bold text-orange-700 flex items-center gap-1.5">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>รูปภาพก่อนซ่อม / สภาพเครื่อง (Before)</span>
                                  </label>
                                  <span className="text-[10px] text-slate-500 font-medium">
                                    {(eq.beforePhotos || []).length} รูป (บีบอัดอัตโนมัติ)
                                  </span>
                                </div>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-orange-600 hover:bg-orange-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>ถ่ายรูป (กล้องมือถือ)</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      capture="environment"
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'beforePhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Upload className="w-3.5 h-3.5" />
                                    <span>เลือกจากอัลบั้ม / ไฟล์</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      multiple
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'beforePhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                </div>
                                {(eq.beforePhotos || []).length > 0 && (
                                  <div className="flex flex-wrap gap-2 pt-1">
                                    {eq.beforePhotos.map((photo, pIdx) => (
                                      <div
                                        key={pIdx}
                                        className="relative group w-16 h-16 rounded-lg overflow-hidden border-2 border-orange-200 bg-white shadow-xs"
                                      >
                                        <img
                                          src={photo}
                                          alt={`Before ${pIdx + 1}`}
                                          className="w-full h-full object-cover"
                                        />
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleRemoveEquipmentPhoto(index, 'beforePhotos', pIdx)
                                          }
                                          className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                          title="ลบรูปนี้"
                                        >
                                          ×
                                        </button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </div>

                              <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2">
                                <div className="flex items-center justify-between">
                                  <label className="text-xs font-bold text-emerald-700 flex items-center gap-1.5">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>รูปภาพหลังซ่อม / ตรวจเช็ค (After)</span>
                                  </label>
                                  <span className="text-[10px] text-slate-500 font-medium">
                                    {(eq.afterPhotos || []).length} รูป (บีบอัดอัตโนมัติ)
                                  </span>
                                </div>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>ถ่ายรูป (กล้องมือถือ)</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      capture="environment"
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'afterPhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Upload className="w-3.5 h-3.5" />
                                    <span>เลือกจากอัลบั้ม / ไฟล์</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      multiple
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'afterPhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                </div>
                                {(eq.afterPhotos || []).length > 0 && (
                                  <div className="flex flex-wrap gap-2 pt-1">
                                    {eq.afterPhotos.map((photo, pIdx) => (
                                      <div
                                        key={pIdx}
                                        className="relative group w-16 h-16 rounded-lg overflow-hidden border-2 border-emerald-200 bg-white shadow-xs"
                                      >
                                        <img
                                          src={photo}
                                          alt={`After ${pIdx + 1}`}
                                          className="w-full h-full object-cover"
                                        />
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleRemoveEquipmentPhoto(index, 'afterPhotos', pIdx)
                                          }
                                          className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                          title="ลบรูปนี้"
                                        >
                                          ×
                                        </button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </div>
                            </div>

                            {/* Before & After Photo Upload per Equipment */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2.5 border-t border-slate-200">
                              <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                                <div className="flex items-center justify-between">
                                  <label className="text-xs font-bold text-orange-700 flex items-center gap-1.5">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>รูปภาพก่อนซ่อม / สภาพเครื่อง (Before)</span>
                                  </label>
                                  <span className="text-[10px] text-slate-500 font-medium">
                                    {(eq.beforePhotos || []).length} รูป
                                  </span>
                                </div>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-orange-600 hover:bg-orange-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>ถ่ายรูป (กล้อง)</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      capture="environment"
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'beforePhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Upload className="w-3.5 h-3.5" />
                                    <span>อัลบั้ม / ไฟล์</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      multiple
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'beforePhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                </div>
                                {(eq.beforePhotos || []).length > 0 && (
                                  <div className="flex flex-wrap gap-2 pt-1">
                                    {eq.beforePhotos.map((photo, pIdx) => (
                                      <div
                                        key={pIdx}
                                        className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-orange-200 bg-white shadow-xs"
                                      >
                                        <img
                                          src={photo}
                                          alt={`Before ${pIdx + 1}`}
                                          className="w-full h-full object-cover"
                                        />
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleRemoveEquipmentPhoto(index, 'beforePhotos', pIdx)
                                          }
                                          className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                          title="ลบรูปนี้"
                                        >
                                          ×
                                        </button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </div>

                              <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                                <div className="flex items-center justify-between">
                                  <label className="text-xs font-bold text-emerald-700 flex items-center gap-1.5">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>รูปภาพหลังซ่อม / ตรวจเช็ค (After)</span>
                                  </label>
                                  <span className="text-[10px] text-slate-500 font-medium">
                                    {(eq.afterPhotos || []).length} รูป
                                  </span>
                                </div>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>ถ่ายรูป (กล้อง)</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      capture="environment"
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'afterPhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                  <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                    <Upload className="w-3.5 h-3.5" />
                                    <span>อัลบั้ม / ไฟล์</span>
                                    <input
                                      type="file"
                                      accept="image/*"
                                      multiple
                                      className="hidden"
                                      onChange={(e) =>
                                        handleEquipmentPhotoUpload(index, 'afterPhotos', e.target.files)
                                      }
                                    />
                                  </label>
                                </div>
                                {(eq.afterPhotos || []).length > 0 && (
                                  <div className="flex flex-wrap gap-2 pt-1">
                                    {eq.afterPhotos.map((photo, pIdx) => (
                                      <div
                                        key={pIdx}
                                        className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-emerald-200 bg-white shadow-xs"
                                      >
                                        <img
                                          src={photo}
                                          alt={`After ${pIdx + 1}`}
                                          className="w-full h-full object-cover"
                                        />
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleRemoveEquipmentPhoto(index, 'afterPhotos', pIdx)
                                          }
                                          className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                          title="ลบรูปนี้"
                                        >
                                          ×
                                        </button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </div>
                            </div>


                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div className="flex justify-end gap-2.5 pt-4 border-t border-slate-200">
                    <button
                      type="button"
                      onClick={() => setCustTab('list')}
                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-semibold transition-colors cursor-pointer"
                    >
                      ยกเลิก
                    </button>
                    <button
                      type="submit"
                      className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-lg font-semibold transition-colors cursor-pointer"
                    >
                      ส่งใบแจ้งซ่อม
                    </button>
                  </div>
                </form>
              </div>
            )}
          </div>
        )}

        {/* VIEW 3: INTERNAL LOGIN */}
        {currentView === 'internal_login' && (
          <div className="max-w-md mx-auto bg-white p-6 sm:p-8 rounded-2xl border border-slate-200 mt-6 space-y-6">
            <div className="text-center space-y-2.5">
              <img
                src={companyProfile.staffMascotUrl || SKM_SERVICE_MASCOT_LOGO}
                alt="เจ้าหน้าที่ภายใน SKM SERVICE"
                className="w-24 h-24 rounded-2xl mx-auto object-cover border-2 border-orange-500 bg-slate-900 shadow-sm"
              />
              <h2 className="text-xl font-bold text-slate-900">เข้าสู่ระบบเจ้าหน้าที่ภายใน</h2>
              <p className="text-xs text-slate-500">{companyProfile.nameTh} · Internal Staff Portal</p>
            </div>

            <form onSubmit={handleLoginSubmit} className="space-y-4 text-xs sm:text-sm">
              <div>
                <label className="block font-semibold text-slate-700 mb-1.5">
                  เลือกบทบาทหน้าที่ (Role) *
                </label>
                <select
                  value={loginRole}
                  onChange={(e) => {
                    const role = e.target.value as 'SERVICE TEAM' | 'TECHNICIAN' | 'SUPERVISOR';
                    setLoginRole(role);
                  }}
                  className="w-full border border-slate-300 rounded-lg px-3 py-2.5 bg-slate-50 font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-orange-500"
                >
                  <option value="SERVICE TEAM">
                    SERVICE TEAM (รับงาน, เลือกช่าง / ยืนยันวันนัดหมาย)
                  </option>
                  <option value="TECHNICIAN">
                    TECHNICIAN (ช่างเทคนิค: วิเคราะห์ด้วย AI, บันทึก Checklist, รูปถ่าย, อะไหล่)
                  </option>
                  <option value="SUPERVISOR">
                    SUPERVISOR (จัดการพนักงาน, วิเคราะห์ Dashboard, อนุมัติใบงาน)
                  </option>
                </select>
              </div>

              {loginRole === 'TECHNICIAN' && (
                <div className="bg-teal-50/80 border border-teal-200 p-3.5 rounded-xl space-y-1.5">
                  <label className="block font-bold text-teal-950">
                    เลือกรายชื่อช่างของคุณ (Technician Name) *
                  </label>
                  <select
                    value={selectedTechNameForLogin}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedTechNameForLogin(val);
                      const match = technicians.find((tc) => tc.name === val);
                      if (match) {
                        setLoggedInTech(match);
                        localStorage.setItem('sahakim_logged_in_tech', JSON.stringify(match));
                      }
                    }}
                    className="w-full border border-teal-300 rounded-lg px-3 py-2 bg-white font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-teal-500"
                  >
                    <option value="">-- คลิกเลือกชื่อช่างของคุณ --</option>
                    {technicians.map((tc) => (
                      <option key={tc.id} value={tc.name}>
                        👷 {tc.name} ({tc.position || 'ช่างเทคนิค'}) - PIN: {tc.pin}
                      </option>
                    ))}
                  </select>
                  <p className="text-[11px] text-teal-700">
                    💡 ช่างแต่ละท่านจะเข้าสู่ระบบและเห็นเฉพาะงานที่มอบหมายให้ตนเองเท่านั้น
                  </p>
                </div>
              )}

              <div>
                <div className="flex justify-between items-center mb-1.5">
                  <label className="block font-semibold text-slate-700">
                    รหัสพนักงาน (PIN) หรือรหัสผ่าน Supervisor *
                  </label>
                  {loginRole === 'SUPERVISOR' && (
                    <button
                      type="button"
                      onClick={() => {
                        setSupPassword('admin123');
                        localStorage.setItem('sahakim_sup_pass', 'admin123');
                        setLoginPin('admin123');
                        showToast(
                          'รีเซ็ตรหัสผ่าน Supervisor เป็น "admin123" เรียบร้อยแล้ว (ใส่รหัสให้เรียบร้อย กดเข้าสู่ระบบได้เลย)',
                          'success'
                        );
                      }}
                      className="text-[11px] font-bold text-orange-600 hover:text-orange-700 underline cursor-pointer"
                    >
                      🔄 รีเซ็ตเป็น admin123
                    </button>
                  )}
                </div>
                <input
                  type="password"
                  value={loginPin}
                  onChange={(e) => setLoginPin(e.target.value)}
                  placeholder="กรอกรหัสผ่าน / PIN เจ้าหน้าที่"
                  required
                  className="w-full border border-slate-300 rounded-lg px-3 py-2.5 text-center font-mono tracking-widest font-bold focus:outline-none focus:ring-2 focus:ring-orange-500"
                />
                <div className="flex flex-col items-center gap-1 mt-2">
                  <p className="text-[11px] text-slate-400 text-center">
                    * เฉพาะเจ้าหน้าที่ภายใน บริษัท สหคิมมอเตอร์ จำกัด เท่านั้น
                  </p>
                  {loginRole === 'SUPERVISOR' && (
                    <div className="flex flex-wrap items-center justify-center gap-1.5 text-[11px]">
                      <span className="text-slate-600 bg-orange-50 text-orange-800 px-2 py-0.5 rounded border border-orange-200 font-semibold">
                        🔑 รหัสเริ่มต้น Supervisor: <b className="font-mono text-orange-900">admin123</b>
                      </span>
                    </div>
                  )}
                  {loginRole === 'SERVICE TEAM' && (
                    <span className="text-slate-600 bg-blue-50 text-blue-800 px-2 py-0.5 rounded border border-blue-200 font-semibold text-[11px]">
                      🔑 รหัสเริ่มต้น Service Team: <b className="font-mono text-blue-900">1234</b>
                    </span>
                  )}
                </div>
              </div>

              <button
                type="submit"
                className="w-full bg-slate-900 hover:bg-slate-800 text-white font-semibold py-2.5 rounded-lg transition-colors cursor-pointer"
              >
                เข้าสู่ระบบ
              </button>
            </form>
          </div>
        )}

        {/* VIEW 4: INTERNAL DASHBOARD */}
        {currentView === 'internal_dashboard' && (
          <div className="space-y-6">
            {/* Internal Top Banner with Role-Specific Color Identity */}
            <div className={`${roleThemeConfig.bannerBg} text-white p-5 sm:p-6 rounded-2xl flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 transition-colors`}>
              <div className="space-y-1.5">
                <div className="flex flex-wrap items-center gap-2 text-xs font-semibold">
                  <span className={`px-2.5 py-0.5 rounded-full text-[11px] font-bold ${roleThemeConfig.badgeBg}`}>
                    {roleThemeConfig.roleColorName}
                  </span>
                  <span className="text-orange-300 flex items-center gap-1">
                    <ShieldCheck className="w-4 h-4" />
                    <span>INTERNAL STAFF PORTAL · SAHA KIM MOTOR</span>
                  </span>
                  {urgentCount > 0 && (
                    <>
                      <span>·</span>
                      <span className="text-amber-300">
                        มีงานรอตรวจสอบ / อนุมัติ {urgentCount} รายการ
                      </span>
                    </>
                  )}
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-xl sm:text-2xl font-bold">
                    บทบาทผู้ใช้งานปัจจุบัน: {userRole}
                  </h2>
                  {userRole === 'TECHNICIAN' && loggedInTech && (
                    <span className="bg-teal-800/90 border border-teal-300 text-teal-100 text-xs px-3 py-1 rounded-full font-bold inline-flex items-center gap-1.5 shadow-sm">
                      <span>👷 ช่างประจำตัว:</span>
                      <strong className="underline text-white font-extrabold">{loggedInTech.name}</strong>
                      <span className="text-[11px] opacity-90">({loggedInTech.position || 'ช่างเทคนิค'})</span>
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-300">
                  {userRole === 'SERVICE TEAM' &&
                    '🔵 โทนสีน้ำเงิน (Dispatch Center): จัดตารางคิวช่าง ระบุวันนัดหมาย กดรับงานแทนช่าง หรือเปลี่ยนช่างใหม่'}
                  {userRole === 'TECHNICIAN' &&
                    '🟢 โทนสีเขียวมรกต-ทีล (Field Engineering): กดตอบรับงาน ตรวจซ่อมหน้างาน วิเคราะห์ด้วย AI และบันทึก Checklist/อะไหล่'}
                  {userRole === 'SUPERVISOR' &&
                    '🟣 โทนสีม่วงคราม (Executive Approval): อนุมัติใบรายงาน/อะไหล่ ปลดล็อกเอกสาร ตั้งค่า LINE/Sheet และดู Dashboard บริหาร'}
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 text-xs">
                {/* Role Switcher with Distinct Role Colors */}
                <div className="flex items-center gap-1 bg-slate-950/60 p-1 rounded-lg border border-white/15">
                  {(['SERVICE TEAM', 'TECHNICIAN', 'SUPERVISOR'] as const).map((r) => {
                    const activeBtnColor =
                      r === 'SERVICE TEAM'
                        ? 'bg-blue-600 text-white'
                        : r === 'TECHNICIAN'
                        ? 'bg-teal-600 text-white'
                        : 'bg-indigo-600 text-white';
                    return (
                      <button
                        key={r}
                        type="button"
                        onClick={() => {
                          setUserRole(r);
                          showToast(`สลับบทบาทเป็น ${r}`, 'success');
                        }}
                        className={`px-2.5 py-1 rounded-md font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          userRole === r ? activeBtnColor : 'text-slate-300 hover:text-white'
                        }`}
                      >
                        {r}
                      </button>
                    );
                  })}
                </div>

                {/* Toggle Role Color Identity On/Off */}
                <button
                  type="button"
                  onClick={() => {
                    setRoleColorIdentityEnabled((prev) => !prev);
                    showToast(
                      !roleColorIdentityEnabled
                        ? 'เปิดการแยกสีแถบหัวข้อตามสิทธิ์ผู้ใช้งานแล้ว'
                        : 'เปลี่ยนเป็นสีมาตรฐานเดียวกันทุกสิทธิ์แล้ว',
                      'success'
                    );
                  }}
                  className="bg-slate-800/90 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-lg font-semibold border border-slate-600 inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <span>🎨 แยกสีตามสิทธิ์: {roleColorIdentityEnabled ? 'เปิด' : 'ปิด'}</span>
                </button>

                <button
                  type="button"
                  onClick={() => setSoundEnabled((prev) => !prev)}
                  className="bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-lg font-semibold border border-slate-700 inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                >
                  {soundEnabled ? <Volume2 className="w-3.5 h-3.5" /> : <VolumeX className="w-3.5 h-3.5" />}
                  <span>{soundEnabled ? 'เสียงเตือน: เปิด' : 'เสียงเตือน: ปิด'}</span>
                </button>

                {userRole === 'SUPERVISOR' && (
                  <>
                    <button
                      type="button"
                      onClick={syncAllTicketsUpdateToLine}
                      className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <Send className="w-3.5 h-3.5" />
                      <span>อัปเดตข้อมูลทั้งหมดเข้า LINE ({tickets.length})</span>
                    </button>
                    <button
                      type="button"
                      onClick={notifyOverdueTicketsToLine}
                      className="bg-rose-600 hover:bg-rose-700 text-white px-3.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <AlertTriangle className="w-3.5 h-3.5 text-amber-200" />
                      <span>แจ้งเตือนงานค้าง ({dashboardStats.overdueCount})</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setIsGoogleSheetsModalOpen(true)}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer shadow-xs"
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-200" />
                      <span>{googleSheetId ? 'Google Sheets (เชื่อมต่อแล้ว)' : 'เชื่อมต่อ Google Sheets & Drive'}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setGasModalOpen(true)}
                      className="bg-teal-700 hover:bg-teal-800 text-white px-3.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <Send className="w-3.5 h-3.5" />
                      <span>แจ้งเตือน LINE &amp; Apps Script</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setTechManagerOpen(true)}
                      className="bg-orange-500 hover:bg-orange-600 text-white px-3.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <UserCog className="w-3.5 h-3.5" />
                      <span>จัดการช่าง / หัวหน้าช่าง / ชื่อ Supervisor</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setMasterDataModalOpen(true)}
                      className="bg-amber-500 hover:bg-amber-600 text-slate-950 px-3.5 py-1.5 rounded-lg font-bold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <Edit3 className="w-3.5 h-3.5" />
                      <span>⚙️ ตั้งค่าข้อมูลตั้งต้น (ยี่ห้อ/เครื่อง/Checklist/โลโก้)</span>
                    </button>
                  </>
                )}

                <button
                  type="button"
                  onClick={exportToExcel}
                  className="bg-slate-800 hover:bg-slate-700 text-white px-3.5 py-1.5 rounded-lg font-semibold border border-slate-700 inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>Export Excel (CSV)</span>
                </button>
              </div>
            </div>

            {/* INTERNAL PORTAL CLICKABLE MENU BAR (เลือกกดเข้าดูทีละหมวด: รายการใบงาน / แจ้งซ่อมใหม่ / ตารางงานช่าง / Dashboard / ตั้งค่าข้อมูลตั้งต้น) */}
            <div className="bg-white p-3.5 sm:p-4 rounded-2xl border-2 border-slate-200 shadow-xs space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2 px-1">
                <p className="text-xs font-bold text-slate-800">
                  📌 เมนูเข้าใช้งานสำหรับเจ้าหน้าที่ภายใน (คลิกเลือกหมวดที่ต้องการดู):
                </p>
                <div className="flex flex-wrap items-center gap-2">
                  {userRole === 'SUPERVISOR' && (
                    <button
                      type="button"
                      onClick={() => setMasterDataModalOpen(true)}
                      className="px-3 py-1 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition-colors cursor-pointer"
                    >
                      ⚙️ เพิ่ม/แก้ไขข้อมูลตั้งต้น (ยี่ห้อ · เครื่องจักร · Checklist · โลโก้/เบอร์โทร)
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() =>
                      setInternalTab((prev) => (prev === 'all' ? 'tickets' : 'all'))
                    }
                    className={`px-3 py-1 rounded-lg text-xs font-bold border transition-colors cursor-pointer ${
                      internalTab === 'all'
                        ? 'bg-slate-900 text-white border-slate-900'
                        : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                    }`}
                  >
                    {internalTab === 'all'
                      ? '✓ กำลังแสดงทุกหมวดในหน้าเดียว (คลิกเพื่อแยกหมวด)'
                      : '👁️ แสดงรวมทุกหมวดในหน้าเดียว'}
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-2 lg:grid-cols-4 gap-2.5">
                <button
                  type="button"
                  onClick={() => setInternalTab('tickets')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-2 cursor-pointer ${
                    internalTab === 'tickets'
                      ? 'bg-slate-900 border-slate-900 text-white shadow-sm'
                      : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-800'
                  }`}
                >
                  <div className="space-y-0.5 min-w-0">
                    <p className="text-xs sm:text-sm font-bold truncate">1. 📋 รายการใบงานซ่อม</p>
                    <p
                      className={`text-[11px] truncate ${
                        internalTab === 'tickets' ? 'text-orange-300' : 'text-slate-500'
                      }`}
                    >
                      รับงาน, วิเคราะห์ AI &amp; อนุมัติงาน
                    </p>
                  </div>
                  <span
                    className={`px-2.5 py-1 rounded-full text-xs font-mono font-bold shrink-0 ${
                      internalTab === 'tickets'
                        ? 'bg-orange-500 text-white'
                        : 'bg-slate-200 text-slate-800'
                    }`}
                  >
                    {tickets.length}
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => setInternalTab('new_ticket')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-2 cursor-pointer ${
                    internalTab === 'new_ticket'
                      ? 'bg-orange-500 border-orange-600 text-white shadow-sm'
                      : 'bg-orange-50/60 hover:bg-orange-100/70 border-orange-200 text-slate-900'
                  }`}
                >
                  <div className="space-y-0.5 min-w-0">
                    <p className="text-xs sm:text-sm font-bold truncate">2. ➕ แจ้งซ่อมใหม่</p>
                    <p
                      className={`text-[11px] truncate ${
                        internalTab === 'new_ticket' ? 'text-orange-100' : 'text-slate-600'
                      }`}
                    >
                      เปิดใบงานแจ้งซ่อมใหม่เข้าระบบ
                    </p>
                  </div>
                  <Plus className="w-5 h-5 shrink-0" />
                </button>

                <button
                  type="button"
                  onClick={() => setInternalTab('schedule')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-2 cursor-pointer ${
                    internalTab === 'schedule'
                      ? 'bg-blue-600 border-blue-700 text-white shadow-sm'
                      : 'bg-blue-50/60 hover:bg-blue-100/70 border-blue-200 text-slate-900'
                  }`}
                >
                  <div className="space-y-0.5 min-w-0">
                    <p className="text-xs sm:text-sm font-bold truncate">3. 📅 ตารางงานช่าง</p>
                    <p
                      className={`text-[11px] truncate ${
                        internalTab === 'schedule' ? 'text-blue-100' : 'text-slate-600'
                      }`}
                    >
                      เช็คคิวว่าง แดง/เหลือง/เขียว &amp; จ่ายงาน
                    </p>
                  </div>
                  <span className="text-xs font-bold shrink-0">🟢🟡🔴</span>
                </button>

                <button
                  type="button"
                  onClick={() => setInternalTab('dashboard')}
                  className={`p-3.5 rounded-xl border-2 text-left transition-all flex items-center justify-between gap-2 cursor-pointer ${
                    internalTab === 'dashboard'
                      ? 'bg-indigo-600 border-indigo-700 text-white shadow-sm'
                      : 'bg-indigo-50/60 hover:bg-indigo-100/70 border-indigo-200 text-slate-900'
                  }`}
                >
                  <div className="space-y-0.5 min-w-0">
                    <p className="text-xs sm:text-sm font-bold truncate">4. 📊 Dashboard</p>
                    <p
                      className={`text-[11px] truncate ${
                        internalTab === 'dashboard' ? 'text-indigo-100' : 'text-slate-600'
                      }`}
                    >
                      สรุปยอดงาน, ต้นทุนประกัน &amp; กราฟ
                    </p>
                  </div>
                  <BarChart3 className="w-5 h-5 shrink-0" />
                </button>
              </div>
            </div>

            {/* INTERNAL SECTION: NEW TICKET FORM (เมื่อกดเมนู "2. แจ้งซ่อมใหม่" ในหน้าของเจ้าหน้าที่ภายใน) */}
            {internalTab === 'new_ticket' && (
              <div className="bg-white p-5 sm:p-6 rounded-2xl border-2 border-orange-300 shadow-xs space-y-5">
                <div className="pb-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="bg-orange-100 text-orange-700 text-[11px] font-bold px-2.5 py-0.5 rounded-full">
                      NEW SERVICE REQUEST · สำหรับเจ้าหน้าที่เปิดใบงานแจ้งซ่อมใหม่
                    </span>
                    <h3 className="text-base sm:text-lg font-bold text-slate-900 mt-1">
                      เปิดใบแจ้งซ่อมใหม่เข้าระบบ (Service Request)
                    </h3>
                    <p className="text-xs text-slate-500">
                      สำหรับ Service Support หรือช่างเทคนิคเปิดใบงานให้ลูกค้า เมื่อกดบันทึกแล้วระบบจะสลับไปหน้ารายการใบงานและแจ้งเตือนเข้า LINE ทันที
                    </p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={() => setInternalTab('schedule')}
                      className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold cursor-pointer"
                    >
                      📅 เช็คตารางงานช่างก่อนนัดวัน
                    </button>
                    <button
                      type="button"
                      onClick={() => setInternalTab('tickets')}
                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3.5 py-2 rounded-xl text-xs font-bold cursor-pointer"
                    >
                      ดูรายการใบงานทั้งหมด ({tickets.length})
                    </button>
                  </div>
                </div>

                <form onSubmit={handleSubmitNewTicket} className="space-y-5 text-xs sm:text-sm">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ชื่อบริษัทลูกค้า *
                      </label>
                      <input
                        type="text"
                        value={newCompany}
                        onChange={(e) => setNewCompany(e.target.value)}
                        placeholder="เช่น บริษัท สยามอุตสาหกรรม จำกัด"
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-slate-50 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เลขประจำตัวผู้เสียภาษี (TAX ID)
                      </label>
                      <input
                        type="text"
                        value={newTaxId}
                        onChange={(e) => setNewTaxId(e.target.value)}
                        placeholder="0105552012345"
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เลขที่เอกสาร / ใบรับประกัน (ถ้ามี)
                      </label>
                      <input
                        type="text"
                        value={newDocNo}
                        onChange={(e) => setNewDocNo(e.target.value)}
                        placeholder="เช่น DOC-2026-999"
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ชื่อผู้ติดต่อและเบอร์โทรศัพท์ *
                      </label>
                      <input
                        type="text"
                        value={newContact}
                        onChange={(e) => setNewContact(e.target.value)}
                        placeholder="คุณสมชาย (081-234-5678)"
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        เงื่อนไขการรับประกัน / คิดค่าบริการ *
                      </label>
                      <select
                        value={newWarranty}
                        onChange={(e) => setNewWarranty(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-orange-600 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>In-Warranty (อยู่ในประกัน - ฟรี / FOC)</option>
                        <option>Out-of-Warranty (หมดประกัน - คิดค่าบริการ / Charge)</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ประเภทบริการ *
                      </label>
                      <select
                        value={newServiceType}
                        onChange={(e) => setNewServiceType(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>ซ่อมหน้างาน</option>
                        <option>รับมาซ่อม</option>
                        <option>งานตรวจเช็ค</option>
                        <option>คอมมิชชั่นนิ่ง (Commissioning)</option>
                        <option>เทสรัน (Test Run)</option>
                        <option>PM (Preventive Maintenance)</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        ความเร่งด่วนของงาน *
                      </label>
                      <select
                        value={newUrgency}
                        onChange={(e) => setNewUrgency(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-red-600 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      >
                        <option>ปกติ (Normal)</option>
                        <option>ด่วน (High)</option>
                        <option>ด่วนที่สุด (Urgent)</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        วันที่ขอช่าง / วันนัดหมายเบื้องต้น *
                      </label>
                      <input
                        type="date"
                        value={newDesiredDate}
                        onChange={(e) => setNewDesiredDate(e.target.value)}
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block font-semibold text-slate-700 mb-1">
                        สถานที่ปฏิบัติงาน / โลเคชั่นโรงงาน *
                      </label>
                      <input
                        type="text"
                        value={newLocation}
                        onChange={(e) => setNewLocation(e.target.value)}
                        placeholder="ระบุชื่อโรงงาน นิคมอุตสาหกรรม หรือพิกัดหน้างาน"
                        required
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                      />
                    </div>
                  </div>

                  {/* Equipment List inside Internal New Ticket Form */}
                  <div className="border-t border-slate-200 pt-4 space-y-3">
                    <div className="flex flex-wrap justify-between items-center gap-2">
                      <h4 className="font-bold text-slate-900 text-xs sm:text-sm">
                        รายการอุปกรณ์ที่แจ้งซ่อม ({tempEquipmentList.length} รายการ)
                      </h4>
                      <button
                        type="button"
                        onClick={() =>
                          setTempEquipmentList((prev) => [
                            ...prev,
                            {
                              type: 'END SUCTION PUMP',
                              customType: '',
                              brand: 'STAC',
                              customBrand: '',
                              model: '',
                              tag: '',
                              symptom: '',
                              detail: ''
                            }
                          ])
                        }
                        className="bg-slate-900 hover:bg-slate-800 text-white text-xs px-3.5 py-2 rounded-lg font-semibold inline-flex items-center gap-1.5 cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>เพิ่มรายการอุปกรณ์</span>
                      </button>
                    </div>

                    <div className="space-y-3">
                      {tempEquipmentList.map((eq, index) => (
                        <div
                          key={index}
                          className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3 text-xs"
                        >
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                            <div>
                              <label className="block font-semibold text-slate-700 mb-1">
                                ประเภทอุปกรณ์ *
                              </label>
                              <select
                                value={eq.type}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setTempEquipmentList((prev) =>
                                    prev.map((item, i) =>
                                      i === index ? { ...item, type: val } : item
                                    )
                                  );
                                }}
                                className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold"
                              >
                                {equipmentOptions.map((m) => (
                                  <option key={m} value={m}>
                                    {m}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div>
                              <label className="block font-semibold text-slate-700 mb-1">
                                ยี่ห้อ / แบรนด์ *
                              </label>
                              <select
                                value={eq.brand}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setTempEquipmentList((prev) =>
                                    prev.map((item, i) =>
                                      i === index ? { ...item, brand: val } : item
                                    )
                                  );
                                }}
                                className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold"
                              >
                                {brandOptions.map((b) => (
                                  <option key={b} value={b}>
                                    {b}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div>
                              <label className="block font-semibold text-slate-700 mb-1">
                                รุ่น (Model)
                              </label>
                              <input
                                type="text"
                                value={eq.model}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setTempEquipmentList((prev) =>
                                    prev.map((item, i) =>
                                      i === index ? { ...item, model: val } : item
                                    )
                                  );
                                }}
                                placeholder="เช่น ISO 65-200 / 15kW"
                                className="w-full border border-slate-300 rounded-lg p-2 bg-white"
                              />
                            </div>
                            <div>
                              <label className="block font-semibold text-slate-700 mb-1">
                                อาการเสียที่พบ *
                              </label>
                              <input
                                type="text"
                                value={eq.symptom}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setTempEquipmentList((prev) =>
                                    prev.map((item, i) =>
                                      i === index ? { ...item, symptom: val } : item
                                    )
                                  );
                                }}
                                placeholder="ระบุอาการเสีย..."
                                required
                                className="w-full border border-slate-300 rounded-lg p-2 bg-white"
                              />
                            </div>
                          </div>

                          {/* Before & After Photo Upload per Equipment (Camera/Album) */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2.5 border-t border-slate-200">
                            <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                              <div className="flex items-center justify-between">
                                <label className="text-xs font-bold text-orange-700 flex items-center gap-1.5">
                                  <Camera className="w-3.5 h-3.5" />
                                  <span>รูปภาพก่อนซ่อม / สภาพเครื่อง (Before)</span>
                                </label>
                                <span className="text-[10px] text-slate-500 font-medium">
                                  {(eq.beforePhotos || []).length} รูป
                                </span>
                              </div>
                              <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                <Upload className="w-3.5 h-3.5" />
                                <span>เลือกรูปจากอัลบั้ม / ถ่ายรูป</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  multiple
                                  className="hidden"
                                  onChange={(e) =>
                                    handleEquipmentPhotoUpload(index, 'beforePhotos', e.target.files)
                                  }
                                />
                              </label>
                              {(eq.beforePhotos || []).length > 0 && (
                                <div className="flex flex-wrap gap-2 pt-1">
                                  {eq.beforePhotos.map((photo, pIdx) => (
                                    <div
                                      key={pIdx}
                                      className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-orange-200 bg-white shadow-xs"
                                    >
                                      <img
                                        src={photo}
                                        alt={`Before ${pIdx + 1}`}
                                        className="w-full h-full object-cover"
                                      />
                                      <button
                                        type="button"
                                        onClick={() =>
                                          handleRemoveEquipmentPhoto(index, 'beforePhotos', pIdx)
                                        }
                                        className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                        title="ลบรูปนี้"
                                      >
                                        ×
                                      </button>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>

                            <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                              <div className="flex items-center justify-between">
                                <label className="text-xs font-bold text-emerald-700 flex items-center gap-1.5">
                                  <Camera className="w-3.5 h-3.5" />
                                  <span>รูปภาพหลังซ่อม / ตรวจเช็ค (After)</span>
                                </label>
                                <span className="text-[10px] text-slate-500 font-medium">
                                  {(eq.afterPhotos || []).length} รูป
                                </span>
                              </div>
                              <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                                <Upload className="w-3.5 h-3.5" />
                                <span>เลือกรูปจากอัลบั้ม / ถ่ายรูป</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  multiple
                                  className="hidden"
                                  onChange={(e) =>
                                    handleEquipmentPhotoUpload(index, 'afterPhotos', e.target.files)
                                  }
                                />
                              </label>
                              {(eq.afterPhotos || []).length > 0 && (
                                <div className="flex flex-wrap gap-2 pt-1">
                                  {eq.afterPhotos.map((photo, pIdx) => (
                                    <div
                                      key={pIdx}
                                      className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-emerald-200 bg-white shadow-xs"
                                    >
                                      <img
                                        src={photo}
                                        alt={`After ${pIdx + 1}`}
                                        className="w-full h-full object-cover"
                                      />
                                      <button
                                        type="button"
                                        onClick={() =>
                                          handleRemoveEquipmentPhoto(index, 'afterPhotos', pIdx)
                                        }
                                        className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                        title="ลบรูปนี้"
                                      >
                                        ×
                                      </button>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="flex justify-end gap-2.5 pt-3 border-t border-slate-200">
                    <button
                      type="button"
                      onClick={() => setInternalTab('tickets')}
                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-semibold cursor-pointer"
                    >
                      ยกเลิก
                    </button>
                    <button
                      type="submit"
                      className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-lg font-bold cursor-pointer"
                    >
                      บันทึกเปิดใบแจ้งซ่อมใหม่
                    </button>
                  </div>
                </form>
              </div>
            )}

            {/* SERVICE SUPPORT: TECHNICIAN AVAILABILITY CALENDAR (แสดงเมื่อเลือกเมนู "ตารางงานช่าง" หรือ "แสดงทั้งหมด") */}
            {(internalTab === 'schedule' || internalTab === 'all') && (
            <div className="bg-white p-5 sm:p-6 rounded-2xl border-2 border-slate-200 shadow-xs space-y-5">
              <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-200 pb-4">
                <div className="space-y-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="bg-orange-500 text-white text-[11px] font-bold px-2.5 py-0.5 rounded-full">
                      SERVICE SUPPORT · TECHNICIAN AVAILABILITY CALENDAR
                    </span>
                    <h3 className="text-base sm:text-lg font-bold text-slate-900">
                      ตารางงานช่างทั้งหมด (เลื่อนดูสัปดาห์ก่อนหน้า/ถัดไปได้) แยกสี แดง / เหลือง / เขียว
                    </h3>
                  </div>
                  <p className="text-xs text-slate-600 leading-relaxed">
                    สิทธิ์ของ <b>Service Support</b>: มีสิทธิ์เปลี่ยนหรือกำหนดวันรับงานช่างได้ เลือกวันและเลือกช่างได้ตามความเหมาะสมและคิวว่าง ·{' '}
                    <span className="text-emerald-700 font-bold">
                      หากช่างยังไม่กดตอบรับงาน สามารถกด &ldquo;รับงานแทนช่างเลย&rdquo; หรือ &ldquo;เปลี่ยนช่างใหม่&rdquo; ได้ทันที
                    </span>
                  </p>
                </div>

                {/* Color Legend with Exact Explanations */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs shrink-0">
                  <div className="px-3 py-2 rounded-xl bg-red-50 border border-red-300 text-red-950">
                    <p className="font-bold flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-red-500 shrink-0" />
                      <span>🔴 สีแดง = เต็มมาก</span>
                    </p>
                    <p className="text-[11px] text-red-800 mt-0.5">
                      คิวช่างแน่นมาก (กรณีเร่งด่วนเลือกเพื่อจัดสรรคิวพิเศษได้)
                    </p>
                  </div>
                  <div className="px-3 py-2 rounded-xl bg-amber-50 border border-amber-300 text-amber-950">
                    <p className="font-bold flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-amber-400 shrink-0" />
                      <span>🟡 สีเหลือง = รับงานได้</span>
                    </p>
                    <p className="text-[11px] text-amber-800 mt-0.5">
                      มีคิวงานบางส่วน ยังสามารถรับงานเพิ่มได้
                    </p>
                  </div>
                  <div className="px-3 py-2 rounded-xl bg-emerald-50 border border-emerald-300 text-emerald-950">
                    <p className="font-bold flex items-center gap-1.5">
                      <span className="w-3 h-3 rounded-full bg-emerald-500 shrink-0" />
                      <span>🟢 สีเขียว = ว่าง</span>
                    </p>
                    <p className="text-[11px] text-emerald-800 mt-0.5">
                      คิวช่างว่าง พร้อมเข้าให้บริการทันที
                    </p>
                  </div>
                </div>
              </div>

              {/* Week Navigation Controls */}
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setSupportScheduleOffset((prev) => prev - 7)}
                    className="px-3.5 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 font-semibold text-slate-700 cursor-pointer"
                  >
                    ← สัปดาห์ก่อนหน้า (7 วัน)
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setSupportScheduleOffset(0);
                      setSupportSelectedDate(new Date().toISOString().slice(0, 10));
                    }}
                    className="px-3.5 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 font-semibold text-slate-900 cursor-pointer"
                  >
                    วันนี้
                  </button>
                  <button
                    type="button"
                    onClick={() => setSupportScheduleOffset((prev) => prev + 7)}
                    className="px-3.5 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 font-semibold text-slate-700 cursor-pointer"
                  >
                    สัปดาห์ถัดไป (7 วัน) →
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <span className="text-slate-600 font-medium">
                    คลิกเลือกวันที่ในตารางด้านล่างเพื่อดูรายชื่อช่างที่ว่าง กำหนดวันรับงาน หรือเปลี่ยนช่างได้ทันที
                  </span>
                  <button
                    type="button"
                    onClick={() => setResetConfirmModalOpen(true)}
                    className="px-3 py-1.5 rounded-lg border border-rose-300 bg-rose-50 hover:bg-rose-100 font-bold text-rose-700 cursor-pointer inline-flex items-center gap-1.5 shadow-xs shrink-0"
                    title="ล้างใบงานตัวอย่างออกทั้งหมด ตารางจะว่างสีเขียว 🟢 ทันที พร้อมเริ่มกรอกงานจริง"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>🔄 รีเซ็ตตารางช่าง (เริ่มกรอกงานจริง / ตารางว่าง 🟢)</span>
                  </button>
                </div>
              </div>

              {/* 14-Day Technician Availability Calendar Grid */}
              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5">
                {supportScheduleDays.map((day) => {
                  const isSelected = supportSelectedDate === day.dateStr;
                  const cardColor =
                    day.level === 'RED'
                      ? 'bg-red-50 border-red-400 hover:bg-red-100 text-red-950'
                      : day.level === 'YELLOW'
                      ? 'bg-amber-50 border-amber-400 hover:bg-amber-100 text-amber-950'
                      : 'bg-emerald-50 border-emerald-400 hover:bg-emerald-100 text-emerald-950';
                  const badgeColor =
                    day.level === 'RED'
                      ? 'bg-red-600 text-white'
                      : day.level === 'YELLOW'
                      ? 'bg-amber-500 text-white'
                      : 'bg-emerald-600 text-white';
                  const dotColor =
                    day.level === 'RED'
                      ? 'bg-red-500'
                      : day.level === 'YELLOW'
                      ? 'bg-amber-400'
                      : 'bg-emerald-500';

                  return (
                    <button
                      key={`sup-cal-${day.dateStr}`}
                      type="button"
                      onClick={() => setSupportSelectedDate(day.dateStr)}
                      className={`p-3 rounded-xl border-2 text-left transition-all flex flex-col justify-between gap-2 cursor-pointer ${cardColor} ${
                        isSelected ? 'ring-3 ring-slate-900 shadow-md scale-[1.02]' : ''
                      }`}
                    >
                      <div className="flex items-start justify-between gap-1">
                        <div>
                          <p className="text-xs font-extrabold">{day.dayLabel}</p>
                          <p className="text-[10px] font-mono opacity-80">{day.dateStr}</p>
                        </div>
                        <span className={`w-3 h-3 rounded-full shrink-0 mt-0.5 ${dotColor}`} />
                      </div>

                      <div className="space-y-1">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <span
                            className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded-md ${badgeColor}`}
                          >
                            {day.statusText}
                          </span>
                          <span className="text-[10px] font-bold text-slate-700">
                            ช่างว่าง {day.freeTechCount}/{technicians.length} คน
                          </span>
                        </div>
                        <p className="text-[11px] font-semibold leading-tight">{day.detailText}</p>
                        {day.ticketsOnDay.length > 0 && (
                          <div className="pt-1 border-t border-black/10 space-y-0.5">
                            {day.ticketsOnDay.slice(0, 2).map((tk) => (
                              <p
                                key={tk.id}
                                className="text-[10px] font-mono font-bold truncate"
                              >
                                • {tk.id} ({tk.assignedTech || 'รอระบุช่าง'})
                              </p>
                            ))}
                            {day.ticketsOnDay.length > 2 && (
                              <p className="text-[10px] font-semibold">
                                + อีก {day.ticketsOnDay.length - 2} ใบงาน
                              </p>
                            )}
                          </div>
                        )}
                      </div>
                    </button>
                  );
                })}
              </div>

              {/* Selected Date Technician Availability & Quick Assignment / Acceptance Panel */}
              {(() => {
                const selectedDayInfo = getScheduleInfoForDate(supportSelectedDate);
                const waitingAcceptanceTickets = tickets.filter(
                  (t) => t.status === 'NEW' || t.status === 'ASSIGNED'
                );

                return (
                  <div className="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800 pb-3">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-xs font-bold text-orange-400">
                          📅 จัดการตารางงานช่างประจำวันที่เลือก:
                        </span>
                        <span className="font-mono font-bold text-sm bg-slate-800 px-3 py-1 rounded-lg border border-slate-700">
                          {selectedDayInfo.dateStr} ({selectedDayInfo.dayLabel})
                        </span>
                        <span
                          className={`px-2.5 py-0.5 rounded-full text-xs font-bold ${
                            selectedDayInfo.level === 'RED'
                              ? 'bg-red-600 text-white'
                              : selectedDayInfo.level === 'YELLOW'
                              ? 'bg-amber-500 text-white'
                              : 'bg-emerald-600 text-white'
                          }`}
                        >
                          สถานะคิวรวม: {selectedDayInfo.statusText} ({selectedDayInfo.detailText})
                        </span>
                      </div>
                      <span className="text-xs text-slate-300">
                        ช่างว่างพร้อมรับงาน {selectedDayInfo.freeTechCount} จาก {technicians.length} คน
                      </span>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                      {/* Left 5 Cols: Technicians Status on Selected Date */}
                      <div className="lg:col-span-5 bg-slate-800/90 border border-slate-700 rounded-xl p-3.5 space-y-2.5">
                        <p className="text-xs font-bold text-orange-300">
                          👷 สถานะคิวช่างแต่ละคนในวันที่ {selectedDayInfo.dateStr} (เลือกช่างตามความเหมาะสมและว่าง):
                        </p>
                        <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                          {selectedDayInfo.techAvailability.map((tc) => (
                            <div
                              key={`sup-tech-${tc.id}`}
                              className="bg-slate-900/90 border border-slate-700 rounded-xl p-3 flex flex-col gap-2 text-xs"
                            >
                              <div className="flex items-center justify-between gap-2">
                                <div className="flex items-center gap-1.5 min-w-0">
                                  <span
                                    className={`w-2.5 h-2.5 rounded-full shrink-0 ${
                                      tc.techLevel === 'RED'
                                        ? 'bg-red-500'
                                        : tc.techLevel === 'YELLOW'
                                        ? 'bg-amber-400'
                                        : 'bg-emerald-500'
                                    }`}
                                  />
                                  <span className="font-bold text-white truncate">{tc.name}</span>
                                  <span className="text-[10px] bg-slate-800 text-slate-300 px-1.5 py-0.5 rounded">
                                    {tc.position || 'ช่างเทคนิค'}
                                  </span>
                                </div>
                                <span
                                  className={`text-[10px] font-bold px-2 py-0.5 rounded-full shrink-0 ${
                                    tc.techLevel === 'RED'
                                      ? 'bg-red-600 text-white'
                                      : tc.techLevel === 'YELLOW'
                                      ? 'bg-amber-500 text-white'
                                      : 'bg-emerald-600 text-white'
                                  }`}
                                >
                                  {tc.availabilityLabel}
                                </span>
                              </div>

                              <p className="text-[11px] text-slate-400">
                                ความเชี่ยวชาญ: {tc.specialization}
                              </p>

                              {/* Assigned Jobs on Selected Date */}
                              {tc.assignedTickets.length > 0 ? (
                                <div className="space-y-1.5 bg-slate-950/70 p-2 rounded-lg border border-slate-800">
                                  <div className="flex items-center justify-between text-[11px]">
                                    <span className="text-amber-300 font-bold">
                                      งานที่ช่างรับในวันนี้ ({tc.assignedTickets.length} งาน):
                                    </span>
                                  </div>
                                  <div className="space-y-1">
                                    {tc.assignedTickets.map((t) => (
                                      <div
                                        key={`tech-sub-job-${t.id}`}
                                        className="bg-slate-900 p-2 rounded border border-slate-700/80 flex flex-wrap items-center justify-between gap-1.5 text-[11px]"
                                      >
                                        <div className="min-w-0">
                                          <span className="font-mono font-bold text-orange-400 mr-1.5">
                                            {t.id}
                                          </span>
                                          <span className="text-white truncate font-medium">
                                            {t.company}
                                          </span>
                                          <span className="text-[10px] text-slate-400 ml-1.5">
                                            ({t.status})
                                          </span>
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          <button
                                            type="button"
                                            onClick={() => handleUnassignTicketFromTech(t.id)}
                                            className="bg-amber-500/20 hover:bg-amber-500 text-amber-300 hover:text-slate-950 border border-amber-500/40 px-2 py-0.5 rounded text-[10px] font-bold transition-colors cursor-pointer inline-flex items-center gap-1"
                                            title="ปลดช่างออกจากงานนี้ คิวงานจะว่างลงทันที"
                                          >
                                            <UserMinus className="w-3 h-3" />
                                            <span>ปลดช่าง</span>
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => handleDeleteTicket(t.id)}
                                            className="bg-red-500/20 hover:bg-red-600 text-red-300 hover:text-white border border-red-500/40 px-2 py-0.5 rounded text-[10px] font-bold transition-colors cursor-pointer inline-flex items-center gap-1"
                                            title="ลบใบงานนี้ออกจากระบบ"
                                          >
                                            <Trash2 className="w-3 h-3" />
                                            <span>ลบงาน</span>
                                          </button>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              ) : (
                                <p className="text-[11px] text-emerald-400 font-medium">
                                  ✓ คิวว่าง ไม่มีงานค้างในวันนี้
                                </p>
                              )}

                              {/* Action Buttons for Technician */}
                              <div className="flex flex-wrap items-center gap-1.5 pt-1 border-t border-slate-800/80">
                                {waitingAcceptanceTickets.length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      openAssignOrRescheduleModal(
                                        waitingAcceptanceTickets[0],
                                        selectedDayInfo.dateStr,
                                        tc.name
                                      )
                                    }
                                    className="bg-orange-600 hover:bg-orange-500 text-white px-2.5 py-1 rounded-lg text-[10px] font-bold cursor-pointer inline-flex items-center gap-1"
                                  >
                                    <Plus className="w-3 h-3" />
                                    <span>+ จ่ายงาน ({waitingAcceptanceTickets[0].id}) ให้ช่างคนนี้</span>
                                  </button>
                                )}
                                <button
                                  type="button"
                                  onClick={() => {
                                    setAssignTechInput(tc.name);
                                    setAssignDateInput(selectedDayInfo.dateStr);
                                    setInternalTab('new_ticket');
                                  }}
                                  className="bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 px-2 py-1 rounded-lg text-[10px] font-bold cursor-pointer inline-flex items-center gap-1"
                                >
                                  <Plus className="w-3 h-3" />
                                  <span>+ เปิดใบแจ้งซ่อมใหม่ระบุช่างนี้</span>
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Right 7 Cols: Jobs on Selected Date + Jobs Waiting for Tech Acceptance (NEW / ASSIGNED) */}
                      <div className="lg:col-span-7 bg-slate-800/90 border border-slate-700 rounded-xl p-3.5 space-y-2.5">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <p className="text-xs font-bold text-orange-300">
                            📋 งานในวันที่ {selectedDayInfo.dateStr} และงานที่ช่างยังไม่กดตอบรับ (กดรับงานแทนช่าง หรือเปลี่ยนช่าง/วันนัดได้ทันที):
                          </p>
                          <span className="text-[11px] text-slate-300">
                            รอจัดคิว/รอช่างตอบรับ {waitingAcceptanceTickets.length} ใบงาน
                          </span>
                        </div>

                        <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                          {tickets.filter(
                            (t) =>
                              t.status !== 'CLOSE' &&
                              (extractTicketDueDate(t) === selectedDayInfo.dateStr ||
                                t.status === 'NEW' ||
                                t.status === 'ASSIGNED')
                          ).length === 0 ? (
                            <div className="text-center py-6 text-xs text-slate-400">
                              ไม่มีใบงานค้างในวันที่ {selectedDayInfo.dateStr} — คิวช่างว่างพร้อมรับงานใหม่
                            </div>
                          ) : (
                            tickets
                              .filter(
                                (t) =>
                                  t.status !== 'CLOSE' &&
                                  (extractTicketDueDate(t) === selectedDayInfo.dateStr ||
                                    t.status === 'NEW' ||
                                    t.status === 'ASSIGNED')
                              )
                              .map((t) => {
                                const isNotAcceptedYet =
                                  t.status === 'NEW' || t.status === 'ASSIGNED';
                                return (
                                  <div
                                    key={`sup-queue-row-${t.id}`}
                                    className="bg-slate-900/95 border border-slate-700 rounded-lg p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-xs"
                                  >
                                    <div className="space-y-0.5">
                                      <div className="flex flex-wrap items-center gap-1.5">
                                        <span className="font-mono font-bold text-orange-400">
                                          {t.id}
                                        </span>
                                        <span className="font-bold text-white">{t.company}</span>
                                        <span
                                          className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                                            t.status === 'NEW'
                                              ? 'bg-blue-600 text-white'
                                              : t.status === 'ASSIGNED'
                                              ? 'bg-amber-500 text-slate-950'
                                              : 'bg-emerald-600 text-white'
                                          }`}
                                        >
                                          {t.status === 'NEW'
                                            ? 'NEW (รอระบุวัน/ช่าง)'
                                            : t.status === 'ASSIGNED'
                                            ? 'ASSIGNED (ช่างยังไม่กดตอบรับ)'
                                            : t.status}
                                        </span>
                                      </div>
                                      <p className="text-[11px] text-slate-300">
                                        วันนัด:{' '}
                                        <span className="font-mono font-semibold text-emerald-400">
                                          {t.confirmedDate || t.desiredDate}
                                        </span>{' '}
                                        · ช่าง:{' '}
                                        <span className="font-semibold text-orange-300">
                                          {t.assignedTech || 'ยังไม่ระบุช่าง'}
                                        </span>
                                      </p>
                                    </div>

                                    <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                                      <button
                                        type="button"
                                        onClick={() =>
                                          openAssignOrRescheduleModal(t, selectedDayInfo.dateStr)
                                        }
                                        className="bg-blue-600 hover:bg-blue-500 text-white px-2.5 py-1.5 rounded-lg text-[11px] font-bold cursor-pointer"
                                      >
                                        📅 เลือกวัน / เปลี่ยนช่าง
                                      </button>
                                      {isNotAcceptedYet && (
                                        <button
                                          type="button"
                                          onClick={() => supportAcceptOnBehalfOfTech(t)}
                                          className="bg-emerald-600 hover:bg-emerald-500 text-white px-2.5 py-1.5 rounded-lg text-[11px] font-bold cursor-pointer"
                                        >
                                          ✅ รับงานแทนช่างเลย
                                        </button>
                                      )}
                                      {t.assignedTech && (
                                        <button
                                          type="button"
                                          onClick={() => handleUnassignTicketFromTech(t.id)}
                                          className="bg-amber-600 hover:bg-amber-500 text-white px-2 py-1.5 rounded-lg text-[11px] font-bold cursor-pointer inline-flex items-center gap-1"
                                          title="ปลดช่างออกจากใบงานนี้ คิวงานจะว่างลงทันที"
                                        >
                                          <UserMinus className="w-3 h-3" />
                                          <span>ปลดช่าง</span>
                                        </button>
                                      )}
                                      <button
                                        type="button"
                                        onClick={() => handleDeleteTicket(t.id)}
                                        className="bg-red-600 hover:bg-red-500 text-white px-2 py-1.5 rounded-lg text-[11px] font-bold cursor-pointer inline-flex items-center gap-1"
                                        title="ลบใบงานนี้ออกจากระบบ"
                                      >
                                        <Trash2 className="w-3 h-3" />
                                        <span>ลบงาน</span>
                                      </button>
                                    </div>
                                  </div>
                                );
                              })
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
            )}

            {/* DASHBOARD KPI & CHARTS SECTION (แสดงเมื่อเลือกเมนู "Dashboard" หรือ "แสดงทั้งหมด") */}
            {(internalTab === 'dashboard' || internalTab === 'all') && (
              <div className="space-y-6">
            {/* 4-Column KPI Analytics Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-xl border border-slate-200">
                <p className="text-xs font-semibold text-slate-500">งานทั้งหมดต่อเดือน</p>
                <p className="text-2xl font-bold text-slate-900 mt-1 tabular-nums">
                  {dashboardStats.totalTickets} งาน{' '}
                  <span className="text-sm font-medium text-emerald-700">
                    (ปิดแล้ว {dashboardStats.closedTickets})
                  </span>
                </p>
                <p className="text-xs text-slate-500 mt-1 tabular-nums">
                  รอดำเนินการ {dashboardStats.totalTickets - dashboardStats.closedTickets} ใบงาน
                </p>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-1">
                <p className="text-xs font-semibold text-slate-500">
                  แยกต้นทุนภายใน (FOC) vs ยอดเรียกเก็บลูกค้า
                </p>
                <p className="text-base font-bold text-slate-900 tabular-nums">
                  ในประกัน (FOC): {dashboardStats.focCount} · คิดเงิน: {dashboardStats.chargeCount}
                </p>
                <div className="text-[11px] font-mono tabular-nums space-y-0.5 pt-0.5">
                  <p className="text-emerald-700 font-bold">
                    🛡️ ต้นทุนภายใน (เคลมในประกัน): {dashboardStats.internalCostTotal.toLocaleString()} THB
                  </p>
                  <p className="text-orange-600 font-bold">
                    💳 ยอดเรียกเก็บลูกค้าสุทธิ: {dashboardStats.customerBilledTotal.toLocaleString()} THB
                  </p>
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200">
                <p className="text-xs font-semibold text-slate-500">สินค้าที่แจ้งซ่อมบ่อยสุด</p>
                <p className="text-lg font-bold text-orange-600 mt-1 truncate">
                  {dashboardStats.topEquip}
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  วิเคราะห์จากรายการอุปกรณ์ในใบงานทั้งหมด
                </p>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200">
                <p className="text-xs font-semibold text-slate-500">เวลาการทำงาน (SLA วันนัดปัจจุบัน)</p>
                <p className="text-lg font-bold text-emerald-700 mt-1 tabular-nums">
                  ตรงเวลา: {dashboardStats.onTimeCount} ·{' '}
                  <span className="text-rose-600">เกินกำหนด: {dashboardStats.overdueCount}</span>
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  คำนวณจริงจากวันนัดหมายเทียบกับวันที่ปัจจุบัน
                </p>
              </div>
            </div>

            {/* Recharts Analytics Summary Dashboard (PieChart & BarChart) */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* Chart 1: PieChart - Equipment Type Proportion */}
              <div className="bg-white p-5 rounded-2xl border border-slate-200 flex flex-col justify-between">
                <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-3 mb-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <PieChartIcon className="w-4 h-4 text-orange-500 shrink-0" />
                      <h3 className="text-sm sm:text-base font-bold text-slate-900">
                        สัดส่วนประเภทอุปกรณ์ที่แจ้งซ่อม (Equipment Breakdown)
                      </h3>
                    </div>
                    <p className="text-xs text-slate-500 mt-0.5">
                      วิเคราะห์จากอุปกรณ์ทั้งหมด {dashboardStats.totalEquipItems} รายการ ใน{' '}
                      {dashboardStats.totalTickets} ใบงาน (คลิกที่รายการเพื่อกรองข้อมูล)
                    </p>
                  </div>
                  {searchQuery && (
                    <button
                      type="button"
                      onClick={() => setSearchQuery('')}
                      className="text-[11px] font-semibold text-orange-600 hover:text-orange-700 bg-orange-50 px-2.5 py-1 rounded-md shrink-0 cursor-pointer"
                    >
                      ล้างตัวกรอง ({searchQuery})
                    </button>
                  )}
                </div>

                {dashboardStats.equipmentChartData.length === 0 ? (
                  <div className="h-64 flex items-center justify-center text-xs text-slate-400">
                    ยังไม่มีข้อมูลอุปกรณ์ที่แจ้งซ่อมในระบบ
                  </div>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-12 gap-4 items-center">
                    <div className="sm:col-span-7 h-64 w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={dashboardStats.equipmentChartData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={52}
                            outerRadius={88}
                            paddingAngle={3}
                            onClick={(entry) => {
                              if (entry && entry.name) {
                                setSearchQuery((prev) => (prev === entry.name ? '' : entry.name));
                              }
                            }}
                            className="cursor-pointer focus:outline-none"
                          >
                            {dashboardStats.equipmentChartData.map((entry, idx) => (
                              <Cell
                                key={`equip-cell-${idx}`}
                                fill={entry.color}
                                stroke="#ffffff"
                                strokeWidth={2}
                              />
                            ))}
                          </Pie>
                          <Tooltip
                            formatter={(value: number, name: string) => [
                              `${value} เครื่อง (${
                                dashboardStats.totalEquipItems > 0
                                  ? Math.round((Number(value) / dashboardStats.totalEquipItems) * 100)
                                  : 0
                              }%)`,
                              name
                            ]}
                            contentStyle={{
                              backgroundColor: '#0f172a',
                              borderColor: '#1e293b',
                              borderRadius: '0.75rem',
                              color: '#f8fafc',
                              fontSize: '12px'
                            }}
                            itemStyle={{ color: '#f8fafc', fontWeight: 600 }}
                          />
                          <Legend
                            verticalAlign="bottom"
                            height={28}
                            iconType="circle"
                            wrapperStyle={{ fontSize: '11px' }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="sm:col-span-5 space-y-2 max-h-60 overflow-y-auto pr-1">
                      {dashboardStats.equipmentChartData.map((item) => {
                        const isSelected = searchQuery.toUpperCase() === item.name.toUpperCase();
                        return (
                          <button
                            key={item.name}
                            type="button"
                            onClick={() =>
                              setSearchQuery((prev) => (prev === item.name ? '' : item.name))
                            }
                            className={`w-full flex items-center justify-between gap-2 px-3 py-2 rounded-lg text-xs border transition-colors cursor-pointer ${
                              isSelected
                                ? 'bg-orange-50 border-orange-400 text-slate-900 font-bold'
                                : 'bg-slate-50 hover:bg-slate-100 border-slate-200/80 text-slate-700'
                            }`}
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              <span
                                className="w-2.5 h-2.5 rounded-full shrink-0"
                                style={{ backgroundColor: item.color }}
                              />
                              <span className="truncate font-semibold">{item.name}</span>
                            </div>
                            <span className="font-mono tabular-nums text-slate-900 shrink-0">
                              {item.value} ({item.percentage}%)
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>

              {/* Chart 2: BarChart - All Work Order Statuses */}
              <div className="bg-white p-5 rounded-2xl border border-slate-200 flex flex-col justify-between">
                <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-3 mb-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <BarChart3 className="w-4 h-4 text-blue-600 shrink-0" />
                      <h3 className="text-sm sm:text-base font-bold text-slate-900">
                        สรุปสถานะใบงานทั้งหมด (Ticket Status Overview)
                      </h3>
                    </div>
                    <p className="text-xs text-slate-500 mt-0.5">
                      แสดงจำนวนใบงานแยกตามสถานะจริงในระบบ (คลิกแท่งกราฟเพื่อกรองสถานะใบงาน)
                    </p>
                  </div>
                  {statusFilter !== 'ALL' && (
                    <button
                      type="button"
                      onClick={() => setStatusFilter('ALL')}
                      className="text-[11px] font-semibold text-blue-600 hover:text-blue-700 bg-blue-50 px-2.5 py-1 rounded-md shrink-0 cursor-pointer"
                    >
                      แสดงทุกสถานะ
                    </button>
                  )}
                </div>

                <div className="h-64 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={dashboardStats.statusChartData}
                      margin={{ top: 10, right: 12, left: -16, bottom: 16 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                      <XAxis
                        dataKey="name"
                        tick={{ fontSize: 11, fill: '#475569', fontWeight: 600 }}
                        axisLine={{ stroke: '#cbd5e1' }}
                        tickLine={false}
                        interval={0}
                        angle={-18}
                        textAnchor="end"
                        height={44}
                      />
                      <YAxis
                        allowDecimals={false}
                        tick={{ fontSize: 11, fill: '#64748b' }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip
                        formatter={(value: number) => [`${value} ใบงาน`, 'จำนวนใบงาน']}
                        labelFormatter={(_, payload) =>
                          payload?.[0]?.payload?.fullLabel || 'สถานะใบงาน'
                        }
                        contentStyle={{
                          backgroundColor: '#0f172a',
                          borderColor: '#1e293b',
                          borderRadius: '0.75rem',
                          color: '#f8fafc',
                          fontSize: '12px'
                        }}
                        itemStyle={{ color: '#f8fafc', fontWeight: 600 }}
                        cursor={{ fill: '#f1f5f9' }}
                      />
                      <Bar
                        dataKey="count"
                        name="จำนวนใบงาน"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={42}
                        onClick={(entry) => {
                          if (entry && entry.filterVal) {
                            setStatusFilter((prev) =>
                              prev === entry.filterVal ? 'ALL' : entry.filterVal
                            );
                          }
                        }}
                        className="cursor-pointer"
                      >
                        {dashboardStats.statusChartData.map((entry, idx) => (
                          <Cell key={`status-bar-${idx}`} fill={entry.color} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Supervisor-Only System Export & GitHub / Google Sheets Admin Panel */}
            {userRole === 'SUPERVISOR' && (
              <div className="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 border border-slate-800">
                <div className="space-y-1">
                  <p className="text-xs font-semibold text-orange-400">
                    SUPERVISOR ADMIN ONLY · จัดการไฟล์ระบบออนไลน์และตั้งค่าแจ้งเตือน LINE
                  </p>
                  <h3 className="text-sm sm:text-base font-bold text-white">
                    คัดลอกโค้ดหรือดาวน์โหลดไฟล์ index.html สำหรับอัปเดตขึ้น GitHub Pages
                  </h3>
                  <p className="text-xs text-slate-300">
                    รองรับการแจ้งเตือนเข้า LINE อัตโนมัติเมื่อมีงานใหม่ (NEW) หรือเมื่อมีงานรออนุมัติอะไหล่/รายงาน พร้อมฝังรหัสผ่านใหม่ที่คุณตั้งไว้ลงในไฟล์ทันที
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => {
                      const code = generateStandaloneHtmlString(
                        tickets,
                        technicians,
                        knowledgeTemplates,
                        masterEquipmentTypes,
                        masterBrands,
                        supPassword,
                        servicePin,
                        gasWebAppUrl,
                        lineChannelToken,
                        lineTargetId,
                        supervisors
                      );
                      if (navigator.clipboard && navigator.clipboard.writeText) {
                        navigator.clipboard.writeText(code).then(() => {
                          showToast(
                            'คัดลอกโค้ด index.html แล้ว! นำไปวางในไฟล์ index.html บน GitHub ได้เลย',
                            'success'
                          );
                        });
                      } else {
                        const ta = document.createElement('textarea');
                        ta.value = code;
                        document.body.appendChild(ta);
                        ta.select();
                        document.execCommand('copy');
                        document.body.removeChild(ta);
                        showToast(
                          'คัดลอกโค้ด index.html แล้ว! นำไปวางในไฟล์ index.html บน GitHub ได้เลย',
                          'success'
                        );
                      }
                    }}
                    className="bg-orange-500 hover:bg-orange-600 text-white px-3.5 py-2 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Copy className="w-3.5 h-3.5" />
                    <span>คัดลอกโค้ด index.html</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      downloadStandaloneHtml(
                        tickets,
                        technicians,
                        knowledgeTemplates,
                        masterEquipmentTypes,
                        masterBrands,
                        supPassword,
                        servicePin,
                        gasWebAppUrl,
                        lineChannelToken,
                        lineTargetId,
                        supervisors
                      );
                      showToast('ดาวน์โหลดไฟล์ index.html สำเร็จ!', 'success');
                    }}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white px-3.5 py-2 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>ดาวน์โหลดไฟล์ index.html</span>
                  </button>
                  <a
                    href="https://github.com/2020skmproduct-service/sahakim-service-portal/edit/main/index.html"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="bg-blue-600 hover:bg-blue-500 text-white px-3.5 py-2 rounded-xl text-xs font-bold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>เปิด GitHub ไปวางโค้ด (index.html)</span>
                  </a>
                </div>
              </div>
            )}
              </div>
            )}

            {/* Workflow Ticket Management Section (แสดงเมื่อเลือกเมนู "รายการใบงานซ่อม" หรือ "แสดงทั้งหมด") */}
            {(internalTab === 'tickets' || internalTab === 'all') && (
            <div className="bg-white p-5 sm:p-6 rounded-2xl border border-slate-200 space-y-4">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-3 border-b border-slate-200 pb-4">
                <div>
                  <h3 className="text-base sm:text-lg font-bold text-slate-900">
                    จัดการรายการคำขอและสถานะใบงาน (Service Workflow)
                  </h3>
                  <p className="text-xs text-slate-500">
                    คลิกปุ่ม &ldquo;ปฏิบัติงาน / ตรวจเช็ค&rdquo; เพื่อใช้ AI ช่วยวิเคราะห์ปัญหาตามมาตรฐาน Pump Handbook / EASA หรือบันทึกผลโดยช่าง
                  </p>
                </div>

                {/* Search & Status Filter Controls */}
                <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
                  {userRole === 'SUPERVISOR' && dashboardStats.overdueCount > 0 && (
                    <button
                      type="button"
                      onClick={notifyOverdueTicketsToLine}
                      className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <AlertTriangle className="w-3.5 h-3.5" />
                      <span>แจ้งเตือนงานค้าง ({dashboardStats.overdueCount})</span>
                    </button>
                  )}
                  <div className="relative flex-grow md:w-64">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="ค้นหาเลขใบงาน, บริษัท, อุปกรณ์..."
                      className="w-full pl-9 pr-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500"
                    />
                  </div>
                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    className="border border-slate-300 rounded-lg px-3 py-1.5 text-xs font-semibold bg-white"
                  >
                    <option value="ALL">ทุกสถานะ ({tickets.length})</option>
                    <option value="OVERDUE">
                      งานค้างเกินกำหนดนัดหมาย ({dashboardStats.overdueCount})
                    </option>
                    <option value="NEW">งานใหม่ (NEW)</option>
                    <option value="ASSIGNED">มอบหมายแล้ว (ASSIGNED)</option>
                    <option value="REPAIRING">กำลังซ่อม (REPAIRING)</option>
                    <option value="PENDING_APPROVAL">รออนุมัติ (Supervisor)</option>
                    <option value="CUSTOMER_ACCEPTANCE">รอตรวจรับ (Customer)</option>
                    <option value="CLOSE">ปิดงานแล้ว (CLOSE)</option>
                  </select>
                </div>
              </div>

              {/* Technician Role View Switcher: "งานของฉัน" vs "รับงานใหม่/งานรอช่าง" */}
              {userRole === 'TECHNICIAN' && (
                <div className="bg-gradient-to-r from-teal-50 via-emerald-50 to-teal-50 border-2 border-teal-300 rounded-2xl p-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-3 shadow-sm">
                  <div className="space-y-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="bg-teal-600 text-white text-xs px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1">
                        👷 ช่างประจำตัว
                      </span>
                      <span className="font-bold text-slate-900 text-sm">
                        {loggedInTech?.name || 'ช่างเทคนิค'}
                      </span>
                      <span className="text-xs text-teal-800 bg-teal-100/90 px-2.5 py-0.5 rounded-md border border-teal-300 font-bold">
                        {loggedInTech?.position || 'ช่างเทคนิค'}
                      </span>
                      <select
                        value={loggedInTech?.name || ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          const found = technicians.find((t) => t.name === val);
                          if (found) {
                            setLoggedInTech(found);
                            localStorage.setItem('sahakim_logged_in_tech', JSON.stringify(found));
                            showToast(`สลับข้อมูลเป็นช่าง: ${found.name}`, 'success');
                          }
                        }}
                        className="text-xs border border-teal-300 rounded-lg px-2 py-1 bg-white font-semibold text-slate-700 cursor-pointer"
                        title="สลับช่างผู้ใช้งาน"
                      >
                        {technicians.map((tc) => (
                          <option key={tc.id} value={tc.name}>
                            สลับเป็น: {tc.name} ({tc.position || 'ช่าง'})
                          </option>
                        ))}
                      </select>
                    </div>
                    <p className="text-xs text-slate-600">
                      {techWorkFilter === 'MY_JOBS'
                        ? '🎯 โหมดแสดงเฉพาะงานที่มอบหมายให้คุณเท่านั้น (ช่างท่านอื่นจะไม่เห็นงานนี้)'
                        : '📥 โหมดรับงานใหม่: แสดงใบงานรอช่างรับงาน คุณสามารถกด "ช่างกดตอบรับงาน" เพื่อนำเข้าสู่งานของตนเองได้ทันที'}
                    </p>
                  </div>
                  <div className="flex items-center gap-2 w-full md:w-auto">
                    <button
                      type="button"
                      onClick={() => setTechWorkFilter('MY_JOBS')}
                      className={`flex-1 md:flex-none px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm ${
                        techWorkFilter === 'MY_JOBS'
                          ? 'bg-teal-600 text-white ring-2 ring-teal-400'
                          : 'bg-white text-slate-700 hover:bg-teal-50 border border-slate-300'
                      }`}
                    >
                      📌 งานของฉัน ({techJobsCount.myJobs} งาน)
                    </button>
                    <button
                      type="button"
                      onClick={() => setTechWorkFilter('AVAILABLE_JOBS')}
                      className={`flex-1 md:flex-none px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm ${
                        techWorkFilter === 'AVAILABLE_JOBS'
                          ? 'bg-amber-600 text-white ring-2 ring-amber-400'
                          : 'bg-white text-slate-700 hover:bg-amber-50 border border-slate-300'
                      }`}
                    >
                      📥 รับงานใหม่ / งานรอช่าง ({techJobsCount.availableJobs} งาน)
                    </button>
                  </div>
                </div>
              )}

              <div className="space-y-3">
                {filteredInternalTickets.length === 0 ? (
                  <div className="text-center py-10 bg-slate-50 border border-dashed border-slate-300 rounded-2xl p-6 space-y-2">
                    <p className="text-sm font-bold text-slate-700">
                      {userRole === 'TECHNICIAN'
                        ? techWorkFilter === 'MY_JOBS'
                          ? `ยังไม่มีใบงานที่มอบหมายให้ช่าง "${loggedInTech?.name || 'ช่างเทคนิค'}" ในขณะนี้`
                          : 'ไม่มีใบงานใหม่ที่รอช่างรับงานในขณะนี้'
                        : 'ไม่พบรายการใบงานที่ตรงกับเงื่อนไขการค้นหา'}
                    </p>
                    {userRole === 'TECHNICIAN' && techWorkFilter === 'MY_JOBS' && techJobsCount.availableJobs > 0 && (
                      <p className="text-xs text-amber-700 font-semibold">
                        💡 มีงานรอช่างรับงาน {techJobsCount.availableJobs} งานในระบบ! คุณสามารถคลิกแท็บ{' '}
                        <button
                          type="button"
                          onClick={() => setTechWorkFilter('AVAILABLE_JOBS')}
                          className="underline font-bold text-amber-900 cursor-pointer"
                        >
                          &ldquo;รับงานใหม่ / งานรอช่าง&rdquo;
                        </button>{' '}
                        เพื่อเลือกรับงานได้ทันที
                      </p>
                    )}
                  </div>
                ) : (
                  filteredInternalTickets.map((t) => {
                    const overdue = isTicketOverdue(t);
                    const dueStr = extractTicketDueDate(t);
                    const daysLate = overdue ? getOverdueDays(t) : 0;
                    return (
                      <div
                        key={t.id}
                        className={`p-4 rounded-xl border flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 ${
                          overdue
                            ? 'bg-rose-50/40 border-rose-300'
                            : 'bg-slate-50 border-slate-200'
                        }`}
                      >
                        <div className="space-y-1.5 text-xs">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className="font-mono font-bold text-sm text-orange-600 tabular-nums">
                              {t.id}
                            </span>
                            <span aria-hidden="true" className="text-slate-300">
                              ·
                            </span>
                            {renderStatusControl(t)}
                            <span aria-hidden="true" className="text-slate-300">
                              ·
                            </span>
                            <span className="font-bold text-slate-900 text-xs sm:text-sm">
                              {t.company}
                            </span>
                            {overdue && (
                              <>
                                <span aria-hidden="true" className="text-slate-300">
                                  ·
                                </span>
                                <span className="bg-rose-600 text-white px-2 py-0.5 rounded font-semibold text-[11px] tabular-nums">
                                  เกินกำหนดนัด {daysLate} วัน (นัด {dueStr})
                                </span>
                              </>
                            )}
                            <span aria-hidden="true" className="text-slate-300">
                              ·
                            </span>
                            {t.isLocked ? (
                              <span className="inline-flex items-center gap-1 text-red-600 font-semibold">
                                <Lock className="w-3 h-3" />
                                <span>ล็อกข้อมูลแล้ว</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold">
                                <Unlock className="w-3 h-3" />
                                <span>เปิดแก้ไข</span>
                              </span>
                            )}
                          </div>

                          <p className="text-slate-600">
                            สถานที่:{' '}
                            <span className="text-slate-900 font-semibold">{t.location || '-'}</span>
                            <span className="mx-2 text-slate-300">·</span>
                            ผู้ติดต่อหน้างาน:{' '}
                            <span className="font-semibold text-slate-900">
                              {t.siteContact || t.contact}
                            </span>
                            <span className="mx-2 text-slate-300">·</span>
                            ช่างผู้ปฏิบัติงาน:{' '}
                            <span className="font-semibold text-orange-600">
                              {t.assignedTech || 'ยังไม่ระบุ'}
                            </span>
                            {t.leadTech && (
                              <>
                                <span className="mx-2 text-slate-300">·</span>
                                หัวหน้าช่าง:{' '}
                                <span className="font-semibold text-indigo-700">{t.leadTech}</span>
                              </>
                            )}
                            {t.approvedBySupervisor && (
                              <>
                                <span className="mx-2 text-slate-300">·</span>
                                ผู้อนุมัติ:{' '}
                                <span className="font-semibold text-emerald-700">
                                  {t.approvedBySupervisor}
                                </span>
                              </>
                            )}
                          </p>

                          <p className="text-slate-500">
                            อุปกรณ์ ({t.equipmentList.length}):{' '}
                            <span className="text-slate-800 font-medium">
                              {t.equipmentList
                                .map((eq) => `${eq.type} [${eq.brand} ${eq.model}]`)
                                .join(', ')}
                            </span>{' '}
                            · วันนัดหมาย:{' '}
                            <span
                              className={`font-mono tabular-nums font-semibold ${
                                overdue ? 'text-rose-600' : 'text-slate-800'
                              }`}
                            >
                              {t.confirmedDate || t.desiredDate}
                            </span>
                          </p>
                        </div>

                      <div className="flex flex-wrap items-center gap-2 shrink-0">
                        {t.isLocked && userRole === 'SUPERVISOR' && (
                          <button
                            type="button"
                            onClick={() => {
                              setUnlockModalTicket(t);
                              setUnlockPassInput('');
                            }}
                            className="bg-purple-600 hover:bg-purple-700 text-white px-3 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <KeyRound className="w-3.5 h-3.5" />
                            <span>ปลดล็อก</span>
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => shareTicketToLine(t)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Send className="w-3.5 h-3.5" />
                          <span>แจ้ง LINE</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => setTechnicianSignModalTicket(t)}
                          className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <PenTool className="w-3.5 h-3.5" />
                          <span>{t.technicianSignature ? '✓ ลายเซ็นช่าง' : '✍️ ลายเซ็นช่าง'}</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => openAssignOrRescheduleModal(t)}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                          title="เลือกวันนัดหมาย หรือเปลี่ยนช่างผู้รับผิดชอบ"
                        >
                          <span>📅 วันนัด/ช่าง</span>
                        </button>
                        {t.assignedTech && (
                          <button
                            type="button"
                            onClick={() => handleUnassignTicketFromTech(t.id)}
                            className="bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300 px-2.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
                            title="ปลดช่างออกจากงานนี้ คิวจะกลับมาว่าง"
                          >
                            <UserMinus className="w-3.5 h-3.5" />
                            <span>ปลดช่าง</span>
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => openActionModal(t)}
                          className="bg-orange-500 hover:bg-orange-600 text-white px-3.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                          <span>ปฏิบัติงาน / วิเคราะห์ AI</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => openServiceReportView(t.id)}
                          className="bg-slate-900 hover:bg-slate-800 text-white px-3.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Printer className="w-3.5 h-3.5" />
                          <span>รายงาน</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteTicket(t.id)}
                          className="bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 px-2.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
                          title="ลบใบงานนี้ออกจากระบบ"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>ลบ</span>
                        </button>
                      </div>
                    </div>
                    );
                  })
                )}
              </div>
            </div>
            )}
          </div>
        )}

        {/* VIEW 5: PRINTABLE A4 SERVICE REPORT */}
        {currentView === 'service_report' && activeReportTicket && (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 no-print">
              <div className="flex items-center gap-3">
                <button
                  type="button"
                  onClick={() => setCurrentView(previousView)}
                  className="bg-slate-100 hover:bg-slate-200 text-slate-800 px-3.5 py-2 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>ย้อนกลับ</span>
                </button>
                <div className="text-xs">
                  <span className="text-slate-500">เลือกใบงานเพื่อดูรายงาน:</span>
                  <select
                    value={activeReportTicket.id}
                    onChange={(e) => setActiveReportTicketId(e.target.value)}
                    className="ml-2 border border-slate-300 rounded-lg px-2.5 py-1.5 font-mono font-bold text-slate-900 bg-white"
                  >
                    {tickets.map((t) => (
                      <option key={t.id} value={t.id}>
                        {t.id} — {t.company}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => setReportSignaturePadOpen((prev) => !prev)}
                  className={`px-3.5 py-2.5 rounded-xl text-xs font-bold border transition-colors inline-flex items-center gap-1.5 cursor-pointer ${
                    reportSignaturePadOpen
                      ? 'bg-emerald-600 border-emerald-700 text-white'
                      : activeReportTicket.customerSignature
                      ? 'bg-emerald-50 border-emerald-400 text-emerald-900 hover:bg-emerald-100'
                      : 'bg-blue-600 border-blue-700 text-white hover:bg-blue-700'
                  }`}
                >
                  <PenTool className="w-3.5 h-3.5" />
                  <span>
                    {activeReportTicket.customerSignature
                      ? '✓ ลูกค้าเซ็นรับงานแล้ว (คลิกเพื่อดู/เซ็นใหม่)'
                      : '✍️ เซ็นชื่อตรวจรับงานบนมือถือ/หน้าจอ'}
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => setShowInternalCostInReport((prev) => !prev)}
                  className={`px-3.5 py-2.5 rounded-xl text-xs font-bold border transition-colors cursor-pointer ${
                    showInternalCostInReport
                      ? 'bg-slate-100 border-slate-300 text-slate-800 hover:bg-slate-200'
                      : 'bg-emerald-50 border-emerald-400 text-emerald-800 hover:bg-emerald-100'
                  }`}
                >
                  {showInternalCostInReport
                    ? '👁️ แสดงตัวเลขต้นทุนภายใน (คลิกเพื่อซ่อนสำหรับส่งลูกค้า)'
                    : '🛡️ ซ่อนตัวเลขต้นทุนภายในแล้ว (โหมดสำหรับส่งลูกค้า)'}
                </button>
                <button
                  type="button"
                  disabled={isGeneratingPdf}
                  onClick={() =>
                    downloadElementAsPdf(
                      'service-report-print-area',
                      `Service-Report-${activeReportTicket.id}.pdf`
                    )
                  }
                  className="bg-orange-500 hover:bg-orange-600 disabled:opacity-60 text-white font-bold px-5 py-2.5 rounded-xl text-xs inline-flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  {isGeneratingPdf ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <Download className="w-4 h-4" />
                  )}
                  <span>บันทึกเป็นไฟล์ PDF (Save PDF)</span>
                </button>
                <button
                  type="button"
                  onClick={() => window.print()}
                  className="bg-slate-800 hover:bg-slate-900 text-white font-semibold px-4 py-2.5 rounded-xl text-xs inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Printer className="w-4 h-4" />
                  <span>สั่งพิมพ์ (Print A4)</span>
                </button>
              </div>
            </div>

            {/* Interactive Mobile Signature Pad Drawer inside Service Report View (no-print) */}
            {reportSignaturePadOpen && (
              <div className="max-w-4xl mx-auto bg-emerald-50/90 border-2 border-emerald-400 rounded-2xl p-4 sm:p-5 space-y-4 no-print shadow-md">
                <div className="flex flex-wrap items-center justify-between gap-2 border-b border-emerald-200 pb-3">
                  <div>
                    <p className="text-xs sm:text-sm font-extrabold text-emerald-950 flex items-center gap-1.5">
                      <PenTool className="w-4 h-4 text-emerald-700" />
                      <span>
                        ลงลายมือชื่อดิจิทัลบนหน้าจอมือถือ / แท็บเล็ต — ใบงาน {activeReportTicket.id} ({activeReportTicket.company})
                      </span>
                    </p>
                    <p className="text-[11px] text-emerald-800 mt-0.5">
                      ใช้นิ้วหรือปากกาเซ็นลงในกรอบด้านล่าง ลายเซ็นจะปรากฏในใบ Service Report และไฟล์ PDF ทันที
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setReportSignaturePadOpen(false)}
                    className="bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer"
                  >
                    พับเก็บแผงเซ็นชื่อ ✕
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Customer Acceptance Signature Pad */}
                  <SignaturePadCanvas
                    label="1. ลายเซ็นลูกค้าผู้ตรวจรับงาน (Customer Acceptance Sign-Off)"
                    sublabel="สำหรับลูกค้าเซ็นรับมอบงานซ่อมบนหน้าจอมือถือ/แท็บเล็ต"
                    signerNameLabel="ชื่อ-นามสกุล ลูกค้าผู้ตรวจรับงาน:"
                    signerNameValue={
                      activeReportTicket.customerSignerName ??
                      (activeReportTicket.siteContact || activeReportTicket.contact)
                    }
                    onSignerNameChange={(val) =>
                      updateTicketSignatureFields(activeReportTicket.id, {
                        customerSignerName: val
                      })
                    }
                    initialSignature={activeReportTicket.customerSignature}
                    signedAt={activeReportTicket.customerSignedAt}
                    onSaveSignature={(dataUrl, timestamp) => {
                      updateTicketSignatureFields(activeReportTicket.id, {
                        customerSignature: dataUrl,
                        customerSignedAt: timestamp,
                        customerSignerName:
                          activeReportTicket.customerSignerName ||
                          activeReportTicket.siteContact ||
                          activeReportTicket.contact
                      });
                    }}
                    onClearSignature={() =>
                      updateTicketSignatureFields(activeReportTicket.id, {
                        customerSignature: '',
                        customerSignedAt: ''
                      })
                    }
                  />

                  {/* Technician Signature Pad */}
                  <SignaturePadCanvas
                    label="2. ลายเซ็นช่างผู้ปฏิบัติงาน (Technician Signature)"
                    sublabel={`ช่างผู้รับผิดชอบ: ${activeReportTicket.assignedTech || 'ช่างเทคนิค SKM'}`}
                    initialSignature={activeReportTicket.technicianSignature}
                    signedAt={activeReportTicket.technicianSignedAt}
                    onSaveSignature={(dataUrl, timestamp) => {
                      updateTicketSignatureFields(activeReportTicket.id, {
                        technicianSignature: dataUrl,
                        technicianSignedAt: timestamp
                      });
                    }}
                    onClearSignature={() =>
                      updateTicketSignatureFields(activeReportTicket.id, {
                        technicianSignature: '',
                        technicianSignedAt: ''
                      })
                    }
                  />
                </div>

                <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-emerald-200 text-xs">
                  <span className="text-emerald-900 font-semibold">
                    💡 เมื่อเซ็นเรียบร้อยแล้ว สามารถกดปุ่ม <b>&ldquo;บันทึกเป็นไฟล์ PDF (Save PDF)&rdquo;</b> ด้านบนเพื่อดาวน์โหลดใบรายงานพร้อมลายเซ็นได้ทันที
                  </span>
                  <div className="flex flex-wrap items-center gap-2">
                    {activeReportTicket.status !== 'CLOSE' && (
                      <button
                        type="button"
                        onClick={() => {
                          if (!activeReportTicket.customerSignature) {
                            showToast('กรุณาลงลายเซ็นลูกค้าในช่องที่ 1 ก่อนกดยืนยันปิดจบงาน', 'warning');
                            return;
                          }
                          updateTicketSignatureFields(
                            activeReportTicket.id,
                            {
                              status: 'CLOSE',
                              isLocked: true
                            },
                            '🎉 [ลูกค้าเซ็นชื่อตรวจรับงานบนมือถือและปิดจบใบงาน (CLOSE) เรียบร้อยแล้ว]'
                          );
                          setReportSignaturePadOpen(false);
                          showToast(
                            `บันทึกลายเซ็นลูกค้าและปิดจบใบงาน ${activeReportTicket.id} เรียบร้อยแล้ว!`,
                            'success'
                          );
                        }}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
                      >
                        <CheckCircle2 className="w-4 h-4" />
                        <span>ยืนยันลายเซ็นลูกค้า &amp; ปิดจบใบงาน (CLOSE)</span>
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={() => {
                        setReportSignaturePadOpen(false);
                        showToast('บันทึกลายเซ็นลงในใบ Service Report เรียบร้อยแล้ว', 'success');
                      }}
                      className="bg-slate-900 hover:bg-slate-800 text-white px-4 py-2 rounded-xl font-bold cursor-pointer"
                    >
                      เสร็จสิ้น (ดูใบรายงาน)
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Printable A4 Sheet */}
            <div
              id="service-report-print-area"
              className="print-sheet max-w-4xl mx-auto bg-white p-6 sm:p-10 rounded-2xl border border-slate-200 space-y-5 text-xs sm:text-sm"
            >
              {/* Company Header */}
              <div className="flex justify-between items-start border-b-2 border-slate-900 pb-4 gap-4">
                <div className="flex items-start gap-3">
                  {companyProfile.logoUrl && (
                    <img
                      src={companyProfile.logoUrl}
                      alt={companyProfile.nameTh}
                      className="w-14 h-14 rounded-xl object-contain border border-slate-200 bg-white p-1 shrink-0"
                    />
                  )}
                  <div className="space-y-0.5">
                    <h1 className="text-lg sm:text-xl font-extrabold text-slate-900">
                      {companyProfile.nameTh}
                    </h1>
                    <p className="text-xs text-orange-600 font-bold tracking-wide">
                      {companyProfile.nameEn}
                    </p>
                    <p className="text-[11px] text-slate-500">
                      {companyProfile.tagline}
                    </p>
                    {companyProfile.phone && (
                      <p className="text-[11px] font-semibold text-slate-700">
                        ติดต่อฝ่ายบริการ: {companyProfile.phone}
                      </p>
                    )}
                  </div>
                </div>
                <div className="text-right bg-slate-50 px-3.5 py-2.5 rounded-lg border border-slate-200 shrink-0">
                  <h2 className="text-xs sm:text-sm font-extrabold text-orange-600 tracking-wide">
                    SERVICE REPORT
                  </h2>
                  <p className="text-xs font-mono font-bold text-slate-900 tabular-nums mt-0.5">
                    เลขที่ใบงาน: {activeReportTicket.id}
                  </p>
                  <p className="text-[10px] font-mono text-slate-500 tabular-nums">
                    วันที่เปิดงาน: {activeReportTicket.createdAt}
                  </p>
                </div>
              </div>

              {/* Section 1: Customer Information */}
              <div className="bg-slate-50 border border-slate-200 p-4 rounded-xl space-y-1.5 text-xs">
                <p className="font-bold text-slate-900 text-sm border-b border-slate-200 pb-1.5 mb-2">
                  ส่วนที่ 1: ข้อมูลลูกค้าผู้แจ้งซ่อม (Customer Information)
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1">
                  <p>
                    <span className="font-semibold text-slate-700">บริษัท:</span>{' '}
                    {activeReportTicket.company}
                  </p>
                  <p className="font-mono tabular-nums">
                    <span className="font-sans font-semibold text-slate-700">TAX ID:</span>{' '}
                    {activeReportTicket.taxId || '-'} ·{' '}
                    <span className="font-sans font-semibold text-slate-700">เลขที่เอกสาร:</span>{' '}
                    {activeReportTicket.docNo || '-'}
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">ผู้ติดต่อตอนแจ้ง:</span>{' '}
                    {activeReportTicket.contact}
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">สถานที่ปฏิบัติงาน:</span>{' '}
                    {activeReportTicket.location || '-'}
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">ประเภทบริการ:</span>{' '}
                    {activeReportTicket.serviceType}
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">เงื่อนไขรับประกัน:</span>{' '}
                    <span className="font-semibold text-slate-900">
                      {activeReportTicket.warranty}
                    </span>{' '}
                    · <span className="font-semibold text-red-600">{activeReportTicket.urgency}</span>
                  </p>
                </div>
              </div>

              {/* Section 2: Company Operations */}
              <div className="bg-slate-50 border border-slate-200 p-4 rounded-xl space-y-1.5 text-xs">
                <p className="font-bold text-slate-900 text-sm border-b border-slate-200 pb-1.5 mb-2">
                  ส่วนที่ 2: ข้อมูลการดำเนินงานโดยทีมงานบริษัท (Company Operations)
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1">
                  <p>
                    <span className="font-semibold text-slate-700">ช่างผู้ปฏิบัติงาน:</span>{' '}
                    <span className="font-bold text-slate-900">
                      {activeReportTicket.assignedTech || 'ยังไม่ระบุ'}
                    </span>
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">หัวหน้าช่างผู้ควบคุมงาน:</span>{' '}
                    <span className="font-bold text-indigo-700">
                      {activeReportTicket.leadTech ||
                        technicians.find((tc) => tc.position?.includes('หัวหน้าช่าง'))?.name ||
                        '-'}
                    </span>
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">ผู้รับหน้างานจริง:</span>{' '}
                    <span className="font-bold text-slate-900">
                      {activeReportTicket.siteContact || activeReportTicket.contact}
                    </span>
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">วันที่นัดหมาย / ปฏิบัติงาน:</span>{' '}
                    <span className="font-mono font-semibold text-emerald-700 tabular-nums">
                      {activeReportTicket.confirmedDate || activeReportTicket.desiredDate}
                    </span>
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">ผู้อนุมัติ (Supervisor):</span>{' '}
                    <span className="font-bold text-emerald-700">
                      {activeReportTicket.approvedBySupervisor || 'รออนุมัติ'}
                    </span>
                  </p>
                  <p>
                    <span className="font-semibold text-slate-700">สถานะใบงาน:</span>{' '}
                    <span className="font-bold text-orange-600">{activeReportTicket.status}</span>
                  </p>
                </div>
              </div>

              {/* Section 3: Equipment & Standard Inspection Checklist */}
              <div className="space-y-4">
                <h3 className="font-bold text-sm text-slate-900 border-b border-slate-200 pb-1.5">
                  ส่วนที่ 3: รายการอุปกรณ์ ผลการวิเคราะห์ทางวิศวกรรม และการตรวจเช็คมาตรฐาน
                </h3>

                {activeReportTicket.equipmentList.map((eq, idx) => (
                  <div
                    key={idx}
                    className="border border-slate-200 rounded-xl p-4 space-y-3 bg-white text-xs"
                  >
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-100 pb-2">
                      <p className="font-bold text-sm text-orange-600">
                        อุปกรณ์ #{idx + 1}: {eq.type} ({eq.brand} - {eq.model}) [Tag:{' '}
                        <span className="font-mono">{eq.tag || 'N/A'}</span>]
                      </p>
                      <p className="text-slate-600">
                        <span className="font-semibold">อาการเสียที่แจ้ง:</span> {eq.symptom}{' '}
                        {eq.detail ? `(${eq.detail})` : ''}
                      </p>
                    </div>

                    {/* Problem Analysis */}
                    <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-1.5">
                      <p>
                        <span className="font-semibold text-slate-700">ผลการวิเคราะห์ปัญหา:</span>{' '}
                        {eq.analysis || '-'}
                      </p>
                      <p>
                        <span className="font-semibold text-slate-700">สาเหตุรากฐาน (Root Cause):</span>{' '}
                        {eq.cause || '-'}
                      </p>
                      {eq.standardsReference && (
                        <p>
                          <span className="font-semibold text-slate-700">มาตรฐานอ้างอิง:</span>{' '}
                          <span className="text-slate-800 font-medium">{eq.standardsReference}</span>
                        </p>
                      )}
                      <p>
                        <span className="font-semibold text-slate-700">เงื่อนไขค่าใช้จ่าย:</span>{' '}
                        <span className="font-semibold text-orange-600">
                          {eq.faultOrigin || '-'}
                        </span>
                      </p>
                      <p>
                        <span className="font-semibold text-slate-700">แนวทางแก้ไข (Corrective Action):</span>{' '}
                        {eq.correctiveAction || '-'}
                      </p>
                    </div>

                    {/* Standard Checklist Table (with Measured Value, Unit, Acceptance Criteria, OK / NG / N/A) */}
                    {eq.checklistResults && eq.checklistResults.length > 0 && (
                      <div className="overflow-x-auto">
                        <p className="font-bold text-slate-800 mb-1.5">
                          ตารางผลการตรวจเช็คและวัดค่าทางวิศวกรรมตามมาตรฐาน (Standard Inspection Checklist):
                        </p>
                        <table className="w-full border-collapse bg-white border border-slate-200 text-xs">
                          <thead>
                            <tr className="bg-slate-100 border-b border-slate-200">
                              <th className="py-1.5 px-2.5 text-left font-semibold text-slate-700">
                                หัวข้อการตรวจสอบ (Inspection Item)
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-32">
                                ค่าที่วัดได้จริง (Measured)
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-16">
                                หน่วย (Unit)
                              </th>
                              <th className="py-1.5 px-2 text-left font-semibold text-slate-700 w-44">
                                เกณฑ์ยอมรับ (Acceptance Criteria)
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-18">
                                Before
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-18">
                                After
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {eq.checklistResults.map((rawChk, cIdx) => {
                              const chk = enrichChecklistItem(rawChk);
                              return (
                                <tr key={cIdx} className="border-b border-slate-200">
                                  <td className="py-1.5 px-2.5 text-slate-800 font-medium">
                                    {chk.item}
                                  </td>
                                  <td className="py-1.5 px-2 text-center font-mono font-bold text-blue-800 bg-blue-50/40">
                                    {chk.measuredValue || '-'}
                                  </td>
                                  <td className="py-1.5 px-2 text-center font-mono text-slate-700">
                                    {chk.unit || '-'}
                                  </td>
                                  <td className="py-1.5 px-2 text-slate-600 text-[11px]">
                                    {chk.acceptanceCriteria || '-'}
                                  </td>
                                  <td
                                    className={`py-1.5 px-2 text-center font-mono font-bold ${
                                      chk.beforeStatus === 'NG'
                                        ? 'text-red-600 bg-red-50/50'
                                        : chk.beforeStatus === 'OK'
                                        ? 'text-emerald-700 bg-emerald-50/50'
                                        : 'text-slate-500 bg-slate-50'
                                    }`}
                                  >
                                    {chk.beforeStatus}
                                  </td>
                                  <td
                                    className={`py-1.5 px-2 text-center font-mono font-bold ${
                                      chk.afterStatus === 'OK'
                                        ? 'text-emerald-700 bg-emerald-50/50'
                                        : chk.afterStatus === 'NG'
                                        ? 'text-red-600 bg-red-50/50'
                                        : 'text-slate-500 bg-slate-50'
                                    }`}
                                  >
                                    {chk.afterStatus}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* Before & After Photos */}
                    <div className="grid grid-cols-2 gap-4 pt-1">
                      <div>
                        <p className="font-semibold text-slate-700 mb-1.5">
                          รูปภาพก่อนปฏิบัติงาน (Before):
                        </p>
                        <div className="flex flex-wrap gap-2">
                          {eq.beforePhotos && eq.beforePhotos.length > 0 ? (
                            eq.beforePhotos.map((img, i) => (
                              <ResilientImage
                                key={i}
                                src={img}
                                alt={`Before ${eq.type}`}
                                className="w-28 h-24 object-cover rounded-lg border border-slate-200"
                              />
                            ))
                          ) : (
                            <span className="text-slate-400 italic">ไม่มีรูป Before</span>
                          )}
                        </div>
                      </div>
                      <div>
                        <p className="font-semibold text-slate-700 mb-1.5">
                          รูปภาพหลังปฏิบัติงาน (After):
                        </p>
                        <div className="flex flex-wrap gap-2">
                          {eq.afterPhotos && eq.afterPhotos.length > 0 ? (
                            eq.afterPhotos.map((img, i) => (
                              <ResilientImage
                                key={i}
                                src={img}
                                alt={`After ${eq.type}`}
                                className="w-28 h-24 object-cover rounded-lg border border-slate-200"
                              />
                            ))
                          ) : (
                            <span className="text-slate-400 italic">ยังไม่มีรูป After</span>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Spare Parts Table (Separating Internal Warranty Cost vs Customer Billed Amount) */}
                    {eq.parts && eq.parts.length > 0 && (
                      <div className="overflow-x-auto">
                        <p className="font-bold text-slate-800 mb-1.5">
                          รายการอะไหล่ที่เบิกใช้ (แยกต้นทุนภายในบริษัท vs ยอดเรียกเก็บลูกค้า):
                        </p>
                        <table className="w-full border-collapse bg-white border border-slate-200 text-xs">
                          <thead>
                            <tr className="bg-slate-100 border-b border-slate-200">
                              <th className="py-1.5 px-2.5 text-left font-semibold text-slate-700">
                                รายการอะไหล่ (Spare Part)
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-20">
                                จำนวน
                              </th>
                              <th className="py-1.5 px-2 text-center font-semibold text-slate-700 w-44">
                                เงื่อนไขค่าใช้จ่าย (Billing Type)
                              </th>
                              {showInternalCostInReport && (
                                <th className="py-1.5 px-2.5 text-right font-semibold text-slate-600 w-32">
                                  ต้นทุนภายใน (บริษัทรับภาระ)
                                </th>
                              )}
                              <th className="py-1.5 px-2.5 text-right font-bold text-slate-900 w-36 bg-emerald-50/70">
                                ยอดเรียกเก็บลูกค้า (THB)
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {eq.parts.map((p, pIdx) => {
                              const isInternal = isPartInternalWarranty(p, eq, activeReportTicket);
                              const internalAmt = isInternal ? p.total : 0;
                              const customerAmt = isInternal ? 0 : p.total;
                              return (
                                <tr key={pIdx} className="border-b border-slate-200">
                                  <td className="py-1.5 px-2.5 text-slate-800 font-medium">
                                    {p.name}
                                  </td>
                                  <td className="py-1.5 px-2 text-center font-mono tabular-nums">
                                    {p.qty} {p.unit}
                                  </td>
                                  <td className="py-1.5 px-2 text-center">
                                    {isInternal ? (
                                      <span className="inline-block bg-emerald-100 text-emerald-800 border border-emerald-300 px-2 py-0.5 rounded-full text-[10px] font-bold">
                                        🛡️ เคลมในประกัน (ไม่คิดเงินลูกค้า)
                                      </span>
                                    ) : (
                                      <span className="inline-block bg-amber-100 text-amber-900 border border-amber-300 px-2 py-0.5 rounded-full text-[10px] font-bold">
                                        💳 คิดค่าบริการ (Charge)
                                      </span>
                                    )}
                                  </td>
                                  {showInternalCostInReport && (
                                    <td className="py-1.5 px-2.5 text-right font-mono tabular-nums text-slate-500">
                                      {isInternal
                                        ? `${internalAmt.toLocaleString()} THB (ต้นทุนภายใน)`
                                        : '-'}
                                    </td>
                                  )}
                                  <td
                                    className={`py-1.5 px-2.5 text-right font-mono font-bold tabular-nums ${
                                      customerAmt === 0
                                        ? 'text-emerald-700 bg-emerald-50/50'
                                        : 'text-slate-900'
                                    }`}
                                  >
                                    {customerAmt === 0
                                      ? '0.00 THB (ฟรีในประกัน)'
                                      : `${customerAmt.toLocaleString()} THB`}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                ))}
              </div>

              {/* Section 4: Clear Separation of Internal Cost vs Net Customer Billed Amount */}
              {(() => {
                const costSummary = getTicketCostSummary(activeReportTicket);
                const isZeroCustomerCharge = costSummary.customerBilledTotal === 0;
                return (
                  <div
                    className={`rounded-xl border-2 p-4 space-y-3 text-xs ${
                      isZeroCustomerCharge
                        ? 'bg-emerald-50/70 border-emerald-400'
                        : 'bg-amber-50/60 border-amber-400'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200/80 pb-2.5">
                      <div>
                        <p className="font-bold text-sm text-slate-900">
                          ส่วนที่ 4: สรุปแยกต้นทุนภายในบริษัท และ ยอดเรียกเก็บลูกค้าสุทธิ (Cost &amp; Warranty Billing Summary)
                        </p>
                        <p className="text-[11px] text-slate-600">
                          แยกแสดงชัดเจนระหว่าง &ldquo;ต้นทุนภายในของบริษัท (กรณีเคลมในประกัน)&rdquo; กับ &ldquo;ยอดที่เรียกเก็บจากลูกค้าจริง&rdquo; เพื่อความโปร่งใสและป้องกันความสับสน
                        </p>
                      </div>
                      <span
                        className={`px-3 py-1 rounded-full text-xs font-extrabold shrink-0 ${
                          isZeroCustomerCharge
                            ? 'bg-emerald-600 text-white'
                            : 'bg-orange-600 text-white'
                        }`}
                      >
                        {isZeroCustomerCharge
                          ? '✅ อยู่ในประกัน (FOC) — ลูกค้าไม่ต้องชำระค่าใช้จ่าย (0.00 บาท)'
                          : `💳 ยอดเรียกเก็บลูกค้าสุทธิ: ${costSummary.customerBilledTotal.toLocaleString()} บาท`}
                      </span>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      {showInternalCostInReport && (
                        <div className="bg-white p-3 rounded-lg border border-slate-200">
                          <p className="text-[11px] font-semibold text-slate-500">
                            1. ต้นทุนอะไหล่/บริการภายใน (บริษัทรับภาระ)
                          </p>
                          <p className="text-base font-mono font-bold text-slate-700 mt-0.5 tabular-nums">
                            {costSummary.internalCostTotal.toLocaleString()} THB
                          </p>
                          <p className="text-[10px] text-slate-500 mt-0.5">
                            * สำหรับบันทึกบัญชีต้นทุนภายในบริษัทเท่านั้น (ไม่เรียกเก็บลูกค้า)
                          </p>
                        </div>
                      )}

                      {showInternalCostInReport && (
                        <div className="bg-white p-3 rounded-lg border border-slate-200">
                          <p className="text-[11px] font-semibold text-emerald-700">
                            2. ส่วนลดความคุ้มครองในประกัน (Warranty FOC)
                          </p>
                          <p className="text-base font-mono font-bold text-emerald-700 mt-0.5 tabular-nums">
                            -{costSummary.internalCostTotal.toLocaleString()} THB (เคลมฟรี 100%)
                          </p>
                          <p className="text-[10px] text-emerald-700 mt-0.5">
                            เงื่อนไข: {activeReportTicket.warranty}
                          </p>
                        </div>
                      )}

                      <div
                        className={`p-3 rounded-lg border-2 ${
                          showInternalCostInReport ? '' : 'sm:col-span-3'
                        } ${
                          isZeroCustomerCharge
                            ? 'bg-emerald-600 text-white border-emerald-700'
                            : 'bg-slate-900 text-white border-orange-500'
                        }`}
                      >
                        <p className="text-[11px] font-bold uppercase tracking-wide opacity-90">
                          ยอดเรียกเก็บลูกค้าสุทธิ (Net Customer Billed Amount)
                        </p>
                        <p className="text-lg sm:text-xl font-mono font-extrabold mt-0.5 tabular-nums">
                          {costSummary.customerBilledTotal === 0
                            ? '0.00 THB (ฟรี / ไม่มียอดเรียกเก็บ)'
                            : `${costSummary.customerBilledTotal.toLocaleString()} THB`}
                        </p>
                        <p className="text-[10px] opacity-95 mt-0.5">
                          {costSummary.customerBilledTotal === 0
                            ? 'รายการนี้อยู่ในเงื่อนไขรับประกัน ลูกค้าไม่ต้องชำระเงินใดๆ ทั้งสิ้น'
                            : 'ยอดรวมค่าอะไหล่และบริการที่เรียกเก็บจากลูกค้าตามจริง'}
                        </p>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* Signatures Block (3 Columns: Technician / Lead Tech & Supervisor / Customer) */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 pt-8 border-t border-slate-200 text-center text-xs">
                {/* Column 1: Technician Signature */}
                <div className="space-y-1 flex flex-col items-center justify-end min-h-[110px]">
                  {activeReportTicket.technicianSignature ? (
                    <img
                      src={activeReportTicket.technicianSignature}
                      alt="Technician Signature"
                      className="h-14 w-auto object-contain mx-auto mb-0.5"
                    />
                  ) : (
                    <div className="h-12 flex items-end justify-center">
                      <button
                        type="button"
                        onClick={() => setReportSignaturePadOpen(true)}
                        className="no-print text-[11px] text-blue-600 hover:text-blue-800 font-semibold bg-blue-50 hover:bg-blue-100 border border-blue-200 px-2.5 py-1 rounded-lg inline-flex items-center gap-1 cursor-pointer mb-1"
                      >
                        <PenTool className="w-3 h-3" />
                        <span>เซ็นชื่อช่างบนจอ</span>
                      </button>
                    </div>
                  )}
                  <p className="text-slate-400">________________________________</p>
                  <p className="font-bold text-slate-900">
                    ({activeReportTicket.assignedTech || 'ช่างผู้ปฏิบัติงาน'})
                  </p>
                  <p className="text-slate-500">ช่างผู้ปฏิบัติงาน (Technician)</p>
                  {activeReportTicket.technicianSignedAt && (
                    <p className="text-[10px] font-mono text-slate-500">
                      วันที่เซ็น: {activeReportTicket.technicianSignedAt}
                    </p>
                  )}
                </div>

                {/* Column 2: Supervisor / Lead Tech Approval */}
                <div className="space-y-1 flex flex-col items-center justify-end min-h-[110px]">
                  <div className="h-12 flex items-end justify-center">
                    {activeReportTicket.approvedBySupervisor ? (
                      <span className="inline-block bg-emerald-50 border border-emerald-300 text-emerald-800 font-bold text-[10px] px-2.5 py-1 rounded-md mb-1">
                        ✓ APPROVED DIGITAL STAMP
                      </span>
                    ) : null}
                  </div>
                  <p className="text-slate-400">________________________________</p>
                  <p className="font-bold text-slate-900">
                    ({activeReportTicket.approvedBySupervisor || activeReportTicket.leadTech || 'หัวหน้าช่าง / Supervisor'})
                  </p>
                  <p className="text-slate-500">หัวหน้าช่าง / Supervisor ผู้อนุมัติ</p>
                </div>

                {/* Column 3: Customer Acceptance Signature */}
                <div className="space-y-1 flex flex-col items-center justify-end min-h-[110px]">
                  {activeReportTicket.customerSignature ? (
                    <div className="flex flex-col items-center">
                      <img
                        src={activeReportTicket.customerSignature}
                        alt="Customer Signature"
                        className="h-14 w-auto object-contain mx-auto mb-0.5"
                      />
                      <button
                        type="button"
                        onClick={() => setReportSignaturePadOpen(true)}
                        className="no-print text-[10px] text-emerald-700 hover:underline font-semibold cursor-pointer"
                      >
                        แก้ไขลายเซ็นลูกค้า
                      </button>
                    </div>
                  ) : (
                    <div className="h-12 flex items-end justify-center">
                      <button
                        type="button"
                        onClick={() => setReportSignaturePadOpen(true)}
                        className="no-print text-[11px] text-emerald-700 hover:text-emerald-900 font-bold bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 px-2.5 py-1 rounded-lg inline-flex items-center gap-1 cursor-pointer mb-1"
                      >
                        <PenTool className="w-3 h-3" />
                        <span>คลิกให้ลูกค้าเซ็นรับงานบนจอ</span>
                      </button>
                    </div>
                  )}
                  <p className="text-slate-400">________________________________</p>
                  <p className="font-bold text-slate-900">
                    ({activeReportTicket.customerSignerName || activeReportTicket.siteContact || activeReportTicket.contact})
                  </p>
                  <p className="text-slate-500">ลูกค้าผู้ตรวจรับงาน (Customer Acceptance)</p>
                  {activeReportTicket.customerSignedAt && (
                    <p className="text-[10px] font-mono text-emerald-700 font-semibold">
                      เซ็นรับมอบเมื่อ: {activeReportTicket.customerSignedAt}
                    </p>
                  )}
                </div>
              </div>

              <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl text-center font-bold text-emerald-800 text-xs">
                Supervisor / หัวหน้าช่าง ผู้อนุมัติ:{' '}
                {activeReportTicket.approvedBySupervisor || 'อยู่ระหว่างรออนุมัติจากหัวหน้าช่าง / Supervisor'}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Quiet Footer */}
      <footer className="bg-slate-900 text-slate-400 text-xs py-4 text-center border-t border-slate-800 mt-auto no-print">
        <div className="max-w-7xl mx-auto px-4">
          <p>
            COPYRIGHT © 2026 บริษัท สหคิมมอเตอร์ จำกัด (SAHA KIM MOTOR CO., LTD.) ALL RIGHTS RESERVED.
          </p>
        </div>
      </footer>

      {/* =====================================================================
          MODALS
      ===================================================================== */}

      {/* 1. Customer Service Request Lookup Modal (Company Name or Ticket ID) */}
      {companyPickerOpen && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 no-print">
          <div className="bg-white w-full max-w-md rounded-2xl border border-slate-200 p-6 space-y-4 shadow-xl">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <div>
                <p className="text-[11px] font-bold text-orange-600 uppercase">SERVICE REQUEST PORTAL</p>
                <h3 className="font-bold text-base text-slate-900">
                  ตรวจสอบสถานะงานซ่อม / จองคิวช่าง (Service Request)
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setCompanyPickerOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>
            <form onSubmit={confirmCustomerPortalCompany} className="space-y-4 text-xs sm:text-sm">
              <div>
                <label className="block font-semibold text-slate-700 mb-1.5">
                  ระบุชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม (เช่น SKM-69-001) *
                </label>
                <input
                  type="text"
                  value={companyInputDraft}
                  onChange={(e) => setCompanyInputDraft(e.target.value)}
                  placeholder="กรอกชื่อบริษัท หรือ เลขที่ใบแจ้งซ่อม เช่น SKM-69-001..."
                  required
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                />
                <p className="text-[11px] text-slate-500 mt-1">
                  * เมื่อกดเข้าสู่ระบบ จะแสดงตารางงานช่าง (สีแดงเต็มมาก / สีเหลืองรับงานได้ / สีเขียวว่าง) และสถานะงานของท่านทันที
                </p>
              </div>
              {tickets.length > 0 && (
                <div className="space-y-1.5">
                  <p className="text-xs text-slate-500">เลือกทดสอบด้วยเลขที่ใบแจ้งซ่อม หรือชื่อบริษัท:</p>
                  <div className="flex flex-wrap gap-1.5">
                    {tickets.slice(0, 4).map((t) => (
                      <button
                        key={`pick-id-${t.id}`}
                        type="button"
                        onClick={() => setCompanyInputDraft(t.id)}
                        className={`px-2.5 py-1 rounded-md text-xs border font-mono transition-colors cursor-pointer ${
                          companyInputDraft === t.id
                            ? 'bg-orange-50 border-orange-500 text-orange-700 font-bold'
                            : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-400'
                        }`}
                      >
                        {t.id}
                      </button>
                    ))}
                    {distinctCompanies.map((comp) => (
                      <button
                        key={comp}
                        type="button"
                        onClick={() => setCompanyInputDraft(comp)}
                        className={`px-2.5 py-1 rounded-md text-xs border transition-colors cursor-pointer ${
                          companyInputDraft === comp
                            ? 'bg-orange-50 border-orange-500 text-orange-700 font-semibold'
                            : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-400'
                        }`}
                      >
                        {comp}
                      </button>
                    ))}
                  </div>
                </div>
              )}
              <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setCompanyPickerOpen(false)}
                  className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-orange-500 hover:bg-orange-600 text-white font-semibold cursor-pointer"
                >
                  เข้าสู่หน้า Service Request
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 2. Assign / Reschedule Date & Technician Modal (Service Support) */}
      {assignModalTicket && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-xl rounded-2xl border border-slate-200 p-6 space-y-4 shadow-xl max-h-[92vh] overflow-y-auto">
            <div className="flex justify-between items-start border-b border-slate-200 pb-3">
              <div>
                <p className="text-[11px] font-bold text-orange-600 uppercase">
                  SERVICE SUPPORT · DISPATCH &amp; SCHEDULE CONTROL
                </p>
                <h3 className="font-bold text-base text-slate-900">
                  กำหนด/เปลี่ยนวันรับงานช่าง และเลือกช่างตามคิวว่าง ({assignModalTicket.id} —{' '}
                  {assignModalTicket.company})
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setAssignModalTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>

            <form onSubmit={handleConfirmAssign} className="space-y-4 text-xs sm:text-sm">
              {/* Quick 7-Day Color-Coded Date Selector inside Modal */}
              <div className="space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-1">
                  <label className="block font-semibold text-slate-700">
                    1. เลือกหรือเปลี่ยนวันรับงานช่าง (ดูสีสถานะคิวว่าง แดง/เหลือง/เขียว) *
                  </label>
                  {(() => {
                    const info = getScheduleInfoForDate(assignDateInput);
                    return (
                      <span
                        className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full ${
                          info.level === 'RED'
                            ? 'bg-red-100 text-red-800 border border-red-300'
                            : info.level === 'YELLOW'
                            ? 'bg-amber-100 text-amber-800 border border-amber-300'
                            : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                        }`}
                      >
                        สถานะวันที่เลือก: {info.statusText} ({info.detailText})
                      </span>
                    );
                  })()}
                </div>

                <input
                  type="date"
                  value={assignDateInput}
                  onChange={(e) => setAssignDateInput(e.target.value)}
                  required
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-orange-500"
                />

                <div className="grid grid-cols-3 sm:grid-cols-7 gap-1.5 pt-1">
                  {supportScheduleDays.slice(0, 7).map((d) => {
                    const selected = assignDateInput === d.dateStr;
                    const cls =
                      d.level === 'RED'
                        ? 'bg-red-50 border-red-300 text-red-900'
                        : d.level === 'YELLOW'
                        ? 'bg-amber-50 border-amber-300 text-amber-900'
                        : 'bg-emerald-50 border-emerald-300 text-emerald-900';
                    return (
                      <button
                        key={`modal-day-${d.dateStr}`}
                        type="button"
                        onClick={() => setAssignDateInput(d.dateStr)}
                        className={`p-1.5 rounded-lg border text-left text-[10px] cursor-pointer ${cls} ${
                          selected ? 'ring-2 ring-slate-900 font-bold' : ''
                        }`}
                      >
                        <p className="font-bold truncate">{d.dayLabel}</p>
                        <p className="font-semibold">{d.statusText}</p>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Technician Selection with Live Availability on Selected Date */}
              <div className="space-y-2">
                <label className="block font-semibold text-slate-700">
                  2. เลือกช่างผู้ปฏิบัติงานตามความเหมาะสมและคิวว่างในวันที่ {assignDateInput} *
                </label>
                {(() => {
                  const dateInfo = getScheduleInfoForDate(assignDateInput);
                  return (
                    <>
                      <select
                        value={assignTechInput}
                        onChange={(e) => setAssignTechInput(e.target.value)}
                        className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold"
                      >
                        {dateInfo.techAvailability.map((tc) => (
                          <option key={tc.id} value={tc.name}>
                            {tc.techLevel === 'GREEN'
                              ? '🟢 [ว่าง] '
                              : tc.techLevel === 'YELLOW'
                              ? `🟡 [มี ${tc.jobCount} งาน] `
                              : `🔴 [คิวแน่น ${tc.jobCount} งาน] `}
                            {tc.name} [{tc.position || 'ช่างเทคนิค'}] ({tc.specialization})
                          </option>
                        ))}
                      </select>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
                        {dateInfo.techAvailability.map((tc) => {
                          const isChosen = assignTechInput === tc.name;
                          return (
                            <button
                              key={`pick-tech-card-${tc.id}`}
                              type="button"
                              onClick={() => setAssignTechInput(tc.name)}
                              className={`p-2.5 rounded-xl border text-left text-xs transition-all cursor-pointer ${
                                isChosen
                                  ? 'border-2 border-orange-500 bg-orange-50/70'
                                  : 'border-slate-200 bg-slate-50 hover:bg-slate-100'
                              }`}
                            >
                              <div className="flex items-center justify-between gap-1">
                                <span className="font-bold text-slate-900">{tc.name}</span>
                                <span
                                  className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                                    tc.techLevel === 'RED'
                                      ? 'bg-red-600 text-white'
                                      : tc.techLevel === 'YELLOW'
                                      ? 'bg-amber-500 text-white'
                                      : 'bg-emerald-600 text-white'
                                  }`}
                                >
                                  {tc.availabilityLabel}
                                </span>
                              </div>
                              <p className="text-[11px] text-slate-500 mt-0.5">
                                {tc.position || 'ช่างเทคนิค'} · {tc.specialization}
                              </p>
                            </button>
                          );
                        })}
                      </div>
                    </>
                  );
                })()}
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  3. เลือกหัวหน้าช่างผู้ควบคุมงาน (Lead Technician)
                </label>
                <select
                  value={assignLeadTechInput}
                  onChange={(e) => setAssignLeadTechInput(e.target.value)}
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-indigo-700"
                >
                  <option value="">-- ไม่ระบุ / ใช้ช่างหลัก --</option>
                  {technicians
                    .filter((tc) => tc.position?.includes('หัวหน้าช่าง'))
                    .map((tc) => (
                      <option key={`lead-${tc.id}`} value={tc.name}>
                        {tc.name} (หัวหน้าช่าง)
                      </option>
                    ))}
                  {technicians
                    .filter((tc) => !tc.position?.includes('หัวหน้าช่าง'))
                    .map((tc) => (
                      <option key={`tech-${tc.id}`} value={tc.name}>
                        {tc.name} ({tc.position || 'ช่างเทคนิค'})
                      </option>
                    ))}
                </select>
              </div>

              <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-xs text-emerald-900 space-y-1">
                <p className="font-bold">
                  💡 สิทธิ์ของ Service Support กรณีช่างยังไม่กดตอบรับงาน:
                </p>
                <p>
                  สามารถกดปุ่มสีเขียว <b>&ldquo;มอบหมาย &amp; รับงานแทนช่างเลย&rdquo;</b> เพื่อรับงานแทนช่างทันทีโดยไม่ต้องรอช่างกดตอบรับ หรือเลือกเปลี่ยนเป็นช่างคนใหม่ที่ว่างได้ตลอดเวลา
                </p>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
                <div className="flex flex-wrap items-center gap-1.5">
                  {assignModalTicket.assignedTech && (
                    <button
                      type="button"
                      onClick={() => {
                        handleUnassignTicketFromTech(assignModalTicket.id);
                        setAssignModalTicket(null);
                      }}
                      className="px-3 py-2 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 text-xs font-bold cursor-pointer inline-flex items-center gap-1"
                      title="ปลดช่างออกจากใบงานนี้ คิวงานจะว่างลงทันที"
                    >
                      <UserMinus className="w-3.5 h-3.5" />
                      <span>ปลดช่างออก (ลบงานจากคิว)</span>
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => {
                      handleDeleteTicket(assignModalTicket.id);
                      setAssignModalTicket(null);
                    }}
                    className="px-3 py-2 rounded-lg bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 text-xs font-bold cursor-pointer inline-flex items-center gap-1"
                    title="ลบใบงานนี้ออกจากระบบ"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>ลบใบงาน</span>
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-2 ml-auto">
                  <button
                    type="button"
                    onClick={() => setAssignModalTicket(null)}
                    className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer text-xs"
                  >
                    ยกเลิก
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-semibold cursor-pointer text-xs"
                  >
                    บันทึกวันนัด / มอบหมายช่าง (รอช่างตอบรับ)
                  </button>
                  <button
                    type="button"
                    onClick={handleConfirmAssignAndAcceptImmediate}
                    className="px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-bold cursor-pointer text-xs"
                  >
                    ✅ มอบหมาย &amp; รับงานแทนช่างเลย
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 3. Supervisor Approve Report / Parts Modal */}
      {approveReportTicket && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 no-print">
          <div className="bg-white w-full max-w-md rounded-2xl border border-slate-200 p-6 space-y-4 shadow-xl">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <h3 className="font-bold text-base text-slate-900">
                {approveMode === 'PARTS'
                  ? `อนุมัติรายการอะไหล่ (${approveReportTicket.id})`
                  : `อนุมัติรายงานและล็อกใบงาน (${approveReportTicket.id})`}
              </h3>
              <button
                type="button"
                onClick={() => setApproveReportTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>
            <form onSubmit={handleConfirmSupervisorApproveReport} className="space-y-4 text-xs sm:text-sm">
              <div className="space-y-2">
                <label className="block font-semibold text-slate-700">
                  เลือกชื่อ Supervisor / หัวหน้าช่าง ผู้อนุมัติ *
                </label>
                <select
                  value={supervisorSignName}
                  onChange={(e) => setSupervisorSignName(e.target.value)}
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 bg-white font-semibold text-slate-900"
                >
                  {availableApproverOptions.map((supName) => (
                    <option key={supName} value={supName}>
                      {supName}
                    </option>
                  ))}
                  <option value="__CUSTOM__">+ พิมพ์เพิ่มชื่อ Supervisor / ผู้อนุมัติใหม่...</option>
                </select>

                {supervisorSignName === '__CUSTOM__' && (
                  <input
                    type="text"
                    value={customSupervisorInput}
                    onChange={(e) => setCustomSupervisorInput(e.target.value)}
                    placeholder="กรอกชื่อ-นามสกุล และตำแหน่งผู้อนุมัติ (ระบบจะจำไว้ให้ครั้งถัดไป)"
                    required
                    className="w-full border border-orange-400 rounded-lg px-3 py-2 bg-orange-50/40 font-semibold"
                  />
                )}
              </div>
              <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setApproveReportTicket(null)}
                  className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold cursor-pointer"
                >
                  {approveMode === 'PARTS' ? 'ยืนยันอนุมัติอะไหล่' : 'ยืนยันอนุมัติและล็อกใบงาน'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 4. Supervisor Reject Modal */}
      {rejectModalTicket && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 no-print">
          <div className="bg-white w-full max-w-md rounded-2xl border border-slate-200 p-6 space-y-4 shadow-xl">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <h3 className="font-bold text-base text-slate-900">
                ส่งกลับแก้ไขใบงาน ({rejectModalTicket.id})
              </h3>
              <button
                type="button"
                onClick={() => setRejectModalTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>
            <form onSubmit={handleConfirmReject} className="space-y-4 text-xs sm:text-sm">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  ระบุเหตุผลที่ไม่อนุมัติ / สิ่งที่ต้องแก้ไข *
                </label>
                <input
                  type="text"
                  value={rejectReasonInput}
                  onChange={(e) => setRejectReasonInput(e.target.value)}
                  placeholder="เช่น กรุณาแนบรูป After และระบุค่า Megger มอเตอร์"
                  required
                  className="w-full border border-slate-300 rounded-lg px-3 py-2"
                />
              </div>
              <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setRejectModalTicket(null)}
                  className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-red-600 hover:bg-red-700 text-white font-semibold cursor-pointer"
                >
                  ยืนยันส่งกลับแก้ไข
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 5. Supervisor Unlock Modal */}
      {unlockModalTicket && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 no-print">
          <div className="bg-white w-full max-w-md rounded-2xl border border-slate-200 p-6 space-y-4 shadow-xl">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <h3 className="font-bold text-base text-slate-900">
                ปลดล็อกใบงาน ({unlockModalTicket.id})
              </h3>
              <button
                type="button"
                onClick={() => setUnlockModalTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>
            <form onSubmit={handleConfirmUnlock} className="space-y-4 text-xs sm:text-sm">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  กรอกรหัสผ่าน Supervisor เพื่อปลดล็อกใบงาน *
                </label>
                <input
                  type="password"
                  value={unlockPassInput}
                  onChange={(e) => setUnlockPassInput(e.target.value)}
                  placeholder="กรอกรหัสผ่าน Supervisor"
                  required
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 font-mono"
                />
              </div>
              <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setUnlockModalTicket(null)}
                  className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-white font-semibold cursor-pointer"
                >
                  ปลดล็อกใบงาน
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 6. Technician & PIN Manager Modal */}
      {techManagerOpen && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-xl rounded-2xl border border-slate-200 p-6 space-y-5 shadow-xl max-h-[92vh] overflow-y-auto">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <div>
                <h3 className="font-bold text-base text-slate-900">
                  ตั้งค่ารหัสผ่านผู้ดูแลระบบ (Admin) และจัดการรหัส PIN เจ้าหน้าที่ภายใน
                </h3>
                <p className="text-[11px] text-slate-500">
                  เปลี่ยนรหัสผ่าน Supervisor, รหัส Service Support และรหัส PIN ของช่างแต่ละคนได้ทันที
                </p>
              </div>
              <button
                type="button"
                onClick={() => setTechManagerOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>

            {/* Section 1: Admin (Supervisor) & Service Support Passwords */}
            <div className="bg-slate-900 text-white p-4 rounded-xl space-y-3 text-xs">
              <p className="font-bold text-orange-400">
                1. ตั้งค่ารหัสผ่านหลัก (Supervisor / Admin &amp; Service Support)
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="bg-slate-800 p-3 rounded-lg border border-slate-700 space-y-1.5">
                  <div className="flex justify-between items-center">
                    <label className="block font-semibold text-slate-200">
                      รหัสผ่าน Supervisor (Admin)
                    </label>
                    <button
                      type="button"
                      onClick={() => {
                        setSupPassword('admin123');
                        localStorage.setItem('sahakim_sup_pass', 'admin123');
                        showToast('รีเซ็ตรหัสผ่าน Supervisor เป็น "admin123" เรียบร้อยแล้ว', 'success');
                      }}
                      className="text-[10px] text-orange-400 hover:text-orange-300 underline font-bold cursor-pointer"
                    >
                      🔄 รีเซ็ตเป็น admin123
                    </button>
                  </div>
                  <div className="flex gap-1.5">
                    <input
                      type="text"
                      value={supPassword}
                      onChange={(e) => setSupPassword(e.target.value)}
                      placeholder="ตั้งรหัส Admin ใหม่"
                      className="w-full bg-slate-950 border border-slate-600 rounded-lg px-2.5 py-1.5 font-mono font-bold text-orange-300"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        if (!supPassword.trim()) {
                          showToast('กรุณากรอกรหัสผ่าน Supervisor', 'error');
                          return;
                        }
                        localStorage.setItem('sahakim_sup_pass', supPassword.trim());
                        showToast('บันทึกรหัสผ่าน Supervisor (Admin) ใหม่เรียบร้อยแล้ว!', 'success');
                      }}
                      className="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1.5 rounded-lg font-bold shrink-0 cursor-pointer"
                    >
                      บันทึก
                    </button>
                  </div>
                </div>

                <div className="bg-slate-800 p-3 rounded-lg border border-slate-700 space-y-1.5">
                  <label className="block font-semibold text-slate-200">
                    รหัส PIN เจ้าหน้าที่รับเรื่อง (SERVICE TEAM)
                  </label>
                  <div className="flex gap-1.5">
                    <input
                      type="text"
                      value={servicePin}
                      onChange={(e) => setServicePin(e.target.value)}
                      placeholder="ตั้งรหัส PIN ใหม่"
                      className="w-full bg-slate-950 border border-slate-600 rounded-lg px-2.5 py-1.5 font-mono font-bold text-emerald-300"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        if (!servicePin.trim()) {
                          showToast('กรุณากรอกรหัส PIN ของ Service Support', 'error');
                          return;
                        }
                        localStorage.setItem('sahakim_service_pin', servicePin.trim());
                        showToast('บันทึกรหัส PIN สำหรับ SERVICE TEAM เรียบร้อยแล้ว!', 'success');
                      }}
                      className="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1.5 rounded-lg font-bold shrink-0 cursor-pointer"
                    >
                      บันทึก
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* Section 2: Supervisor Approver List Management */}
            <div className="bg-emerald-50/70 border border-emerald-200 p-4 rounded-xl space-y-3 text-xs">
              <div>
                <h4 className="font-bold text-emerald-950">
                  2. รายชื่อ Supervisor / หัวหน้าช่าง ผู้อนุมัติงาน (ใช้เลือกตอนกดอนุมัติอะไหล่และรายงาน)
                </h4>
                <p className="text-[11px] text-emerald-800">
                  เพิ่มรายชื่อผู้มีสิทธิ์อนุมัติ เพื่อให้เลือกชื่อได้สะดวกเวลาอนุมัติงานและแสดงในใบ Service Report
                </p>
              </div>
              <div className="flex flex-wrap gap-1.5">
                {supervisors.map((supName, sIdx) => (
                  <span
                    key={sIdx}
                    className="inline-flex items-center gap-1.5 bg-white border border-emerald-300 text-slate-800 px-2.5 py-1 rounded-lg font-semibold"
                  >
                    <span>{supName}</span>
                    {supervisors.length > 1 && (
                      <button
                        type="button"
                        onClick={() => {
                          const updated = supervisors.filter((_, i) => i !== sIdx);
                          handleSaveSupervisorsList(updated, `ลบรายชื่อผู้อนุมัติ "${supName}" เรียบร้อยแล้ว`);
                        }}
                        className="text-red-500 hover:text-red-700 font-bold cursor-pointer"
                      >
                        ×
                      </button>
                    )}
                  </span>
                ))}
              </div>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={newSupervisorName}
                  onChange={(e) => setNewSupervisorName(e.target.value)}
                  placeholder="พิมพ์ชื่อ Supervisor หรือผู้จัดการผู้อนุมัติ เช่น คุณสมศักดิ์ (Supervisor)..."
                  className="flex-grow border border-emerald-300 rounded-lg px-3 py-1.5 bg-white font-semibold"
                />
                <button
                  type="button"
                  onClick={() => {
                    if (!newSupervisorName.trim()) {
                      showToast('กรุณากรอกชื่อ Supervisor ที่ต้องการเพิ่ม', 'error');
                      return;
                    }
                    const updated = Array.from(new Set([...supervisors, newSupervisorName.trim()]));
                    setNewSupervisorName('');
                    handleSaveSupervisorsList(updated, 'เพิ่มรายชื่อ Supervisor ผู้อนุมัติเรียบร้อย!');
                  }}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded-lg font-bold shrink-0 cursor-pointer"
                >
                  + เพิ่มชื่อผู้อนุมัติ
                </button>
              </div>
            </div>

            {/* Section 3: Technician & Lead Technician List & Inline Position/PIN Edit */}
            <div className="space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-1">
                <h4 className="font-bold text-xs text-slate-900">
                  3. รายชื่อช่างเทคนิค &amp; หัวหน้าช่าง (แก้ไขชื่อ ตำแหน่ง ความเชี่ยวชาญ PIN หรือลบช่างได้ทั้งหมด)
                </h4>
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-semibold text-slate-500">
                    รวม {technicians.length} ท่าน
                  </span>
                  {technicians.length > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        const cleanList = technicians.filter((tc) => !isDemoTechnician(tc));
                        if (cleanList.length < technicians.length) {
                          handleSaveTechniciansList(
                            cleanList,
                            'ลบรายชื่อช่าง DEMO ออกเรียบร้อยแล้ว'
                          );
                        } else {
                          handleSaveTechniciansList(
                            [],
                            'ล้างรายชื่อช่างทั้งหมดเรียบร้อยแล้ว (ช่าง = 0 คน พร้อมลงข้อมูลใหม่)'
                          );
                        }
                      }}
                      className="bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 px-2 py-0.5 rounded text-[11px] font-bold cursor-pointer"
                    >
                      🗑️ ล้างรายชื่อช่างทั้งหมด ({technicians.length})
                    </button>
                  )}
                </div>
              </div>

              {technicians.length === 0 ? (
                <div className="text-center py-8 bg-slate-50 border-2 border-dashed border-slate-300 rounded-2xl p-5 space-y-2">
                  <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto text-xl font-bold">
                    ✓
                  </div>
                  <h5 className="font-bold text-sm text-slate-800">
                    ลบรายชื่อช่าง DEMO ออกทั้งหมดเรียบร้อยแล้ว (ปัจจุบันมีช่าง 0 คน)
                  </h5>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    ระบบพร้อมรองรับข้อมูลช่างจริงของบริษัทแล้ว สามารถเริ่มกรอกชื่อช่างคนแรกได้ที่ <b>ข้อ 4. ด้านล่าง</b> (ระบบจะรันรหัสเริ่มต้นเป็น TECH-01 ให้อัตโนมัติ)
                  </p>
                </div>
              ) : (
                <div className="space-y-3 max-h-80 overflow-y-auto pr-1">
                  {technicians.map((tc, idx) => (
                    <div
                      key={tc.id}
                      className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs space-y-2.5 shadow-sm"
                    >
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200/80 pb-2">
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-bold text-slate-600 bg-slate-200 px-2 py-0.5 rounded text-[11px]">
                            {tc.id}
                          </span>
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                              tc.position?.includes('หัวหน้าช่าง')
                                ? 'bg-indigo-600 text-white'
                                : 'bg-teal-700 text-white'
                            }`}
                          >
                            {tc.position || 'ช่างเทคนิค'}
                          </span>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <span className="text-[11px] font-bold text-slate-600">ตำแหน่ง:</span>
                          <select
                            value={tc.position || 'ช่างเทคนิค (Technician)'}
                            onChange={(e) => {
                              const newPos = e.target.value;
                              setTechnicians((prev) =>
                                prev.map((item, i) =>
                                  i === idx ? { ...item, position: newPos } : item
                                )
                              );
                            }}
                            className="border border-slate-300 rounded-lg px-2 py-1 text-[11px] font-semibold bg-white text-slate-800"
                          >
                            {TECH_POSITIONS.map((pos) => (
                              <option key={pos} value={pos}>
                                {pos}
                              </option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">
                            ชื่อ-นามสกุลช่าง *
                          </label>
                          <input
                            type="text"
                            value={tc.name}
                            onChange={(e) => {
                              const val = e.target.value;
                              setTechnicians((prev) =>
                                prev.map((item, i) => (i === idx ? { ...item, name: val } : item))
                              );
                            }}
                            placeholder="ชื่อ-นามสกุลช่าง"
                            className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 font-bold text-slate-900 bg-white focus:outline-none focus:ring-2 focus:ring-orange-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">
                            ความเชี่ยวชาญเฉพาะทาง
                          </label>
                          <input
                            type="text"
                            value={tc.specialization || ''}
                            onChange={(e) => {
                              const val = e.target.value;
                              setTechnicians((prev) =>
                                prev.map((item, i) =>
                                  i === idx ? { ...item, specialization: val } : item
                                )
                              );
                            }}
                            placeholder="เช่น PUMP & MOTOR EXPERT"
                            className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-700 bg-white focus:outline-none focus:ring-2 focus:ring-orange-500"
                          />
                        </div>
                      </div>

                      <div className="flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-slate-200/60">
                        <div className="flex items-center gap-1.5">
                          <span className="text-[11px] font-bold text-slate-700">รหัส PIN ช่าง:</span>
                          <input
                            type="text"
                            value={tc.pin}
                            onChange={(e) => {
                              const val = e.target.value;
                              setTechnicians((prev) =>
                                prev.map((item, i) => (i === idx ? { ...item, pin: val } : item))
                              );
                            }}
                            className="w-20 border border-slate-300 rounded-lg px-2 py-1 font-mono font-bold text-center text-orange-600 bg-white"
                            title="PIN ประจำตัวช่าง (4 หลัก)"
                          />
                        </div>

                        <div className="flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => {
                              if (!tc.name.trim()) {
                                showToast('ชื่อช่างต้องไม่เว้นว่าง', 'error');
                                return;
                              }
                              handleSaveTechniciansList(
                                technicians,
                                `บันทึกข้อมูลของ ${tc.name} เรียบร้อยแล้ว`
                              );
                            }}
                            className="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg font-bold text-xs cursor-pointer shadow-sm transition-colors"
                          >
                            💾 บันทึก
                          </button>

                          {techIdToDeleteConfirm === tc.id ? (
                            <div className="flex items-center gap-1 bg-red-50 border border-red-300 px-2 py-1 rounded-lg">
                              <span className="text-[11px] font-bold text-red-800">ลบจริงไหม?</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const updated = technicians.filter((_, i) => i !== idx);
                                  setTechIdToDeleteConfirm(null);
                                  handleSaveTechniciansList(
                                    updated,
                                    `ลบรายชื่อช่าง ${tc.name} เรียบร้อยแล้ว`
                                  );
                                }}
                                className="bg-red-600 hover:bg-red-700 text-white px-2 py-0.5 rounded text-[11px] font-bold cursor-pointer"
                              >
                                ใช่ ลบเลย
                              </button>
                              <button
                                type="button"
                                onClick={() => setTechIdToDeleteConfirm(null)}
                                className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-2 py-0.5 rounded text-[11px] cursor-pointer"
                              >
                                ยกเลิก
                              </button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => setTechIdToDeleteConfirm(tc.id)}
                              className="bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded-lg font-bold text-xs cursor-pointer shadow-sm transition-colors"
                            >
                              🗑️ ลบ
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Section 4: Add New Technician / Lead Technician */}
            <div className="border-t border-slate-200 pt-4 space-y-3 text-xs">
              <h4 className="font-bold text-slate-900">
                4. เพิ่มพนักงานช่าง / หัวหน้าช่างใหม่
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <input
                  type="text"
                  value={newTechName}
                  onChange={(e) => setNewTechName(e.target.value)}
                  placeholder="ชื่อ-นามสกุลช่าง / หัวหน้าช่าง *"
                  className="border border-slate-300 rounded-lg p-2"
                />
                <select
                  value={newTechPosition}
                  onChange={(e) => setNewTechPosition(e.target.value)}
                  className="border border-slate-300 rounded-lg p-2 bg-white font-semibold text-indigo-700"
                >
                  {TECH_POSITIONS.map((pos) => (
                    <option key={pos} value={pos}>
                      ตำแหน่ง: {pos}
                    </option>
                  ))}
                </select>
                <input
                  type="text"
                  value={newTechSpec}
                  onChange={(e) => setNewTechSpec(e.target.value)}
                  placeholder="ความเชี่ยวชาญ (เช่น PUMP & MOTOR)"
                  className="border border-slate-300 rounded-lg p-2"
                />
                <input
                  type="text"
                  value={newTechPin}
                  onChange={(e) => setNewTechPin(e.target.value)}
                  placeholder="ตั้งรหัส PIN ช่าง (เช่น 4444)"
                  className="border border-slate-300 rounded-lg p-2 font-mono"
                />
              </div>
              <button
                type="button"
                onClick={() => {
                  if (!newTechName.trim()) {
                    showToast('กรุณากรอกชื่อ-นามสกุลช่าง', 'error');
                    return;
                  }
                  let maxNum = 0;
                  technicians.forEach((t) => {
                    const match = t.id.match(/\d+/);
                    if (match) {
                      const num = parseInt(match[0], 10);
                      if (num > maxNum) maxNum = num;
                    }
                  });
                  const nextNum = maxNum + 1;
                  const nextId = `TECH-${nextNum < 10 ? '0' + nextNum : nextNum}`;
                  const newTechItem: Technician = {
                    id: nextId,
                    name: newTechName.trim(),
                    position: newTechPosition,
                    specialization: newTechSpec.trim() || 'GENERAL TECH',
                    status: 'Available',
                    pin: newTechPin.trim() || '1111'
                  };
                  const updated = [...technicians, newTechItem];
                  setNewTechName('');
                  setNewTechSpec('');
                  setNewTechPin('');
                  handleSaveTechniciansList(
                    updated,
                    `เพิ่ม ${newTechItem.name} (${newTechItem.position}) สำเร็จ`
                  );
                }}
                className="w-full bg-orange-500 hover:bg-orange-600 text-white font-semibold py-2 rounded-lg transition-colors cursor-pointer"
              >
                บันทึกเพิ่มพนักงาน / หัวหน้าช่าง
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7. Internal Action / Inspection / AI Engineering Troubleshooter Modal */}
      {editingTicket && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-4xl rounded-2xl border border-slate-200 p-5 sm:p-6 space-y-4 max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <div>
                <h3 className="font-bold text-base sm:text-lg text-slate-900">
                  บันทึกการปฏิบัติงาน วิเคราะห์ปัญหา (AI + ช่าง) และตรวจเช็คมาตรฐาน: {editingTicket.id}
                </h3>
                <p className="text-xs text-slate-500">
                  ลูกค้า: {editingTicket.company} · ช่างผู้รับผิดชอบ:{' '}
                  {editingTicket.assignedTech || 'ยังไม่ระบุ'}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setEditingTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-xl font-bold cursor-pointer"
              >
                ×
              </button>
            </div>

            {/* Site Contact & Location */}
            <div className="bg-orange-50/70 p-3.5 rounded-xl border border-orange-200 space-y-2 text-xs">
              <div className="flex flex-wrap justify-between items-center gap-2">
                <label className="font-bold text-slate-900 block">
                  ข้อมูลผู้รับหน้างานจริง / โลเคชั่นปฏิบัติงาน:
                </label>
                <button
                  type="button"
                  onClick={() => {
                    const nextType = 'END SUCTION PUMP';
                    setEditingTicket({
                      ...editingTicket,
                      equipmentList: [
                        ...editingTicket.equipmentList,
                        {
                          type: nextType,
                          brand: 'STAC',
                          model: '-',
                          tag: `EQ-0${editingTicket.equipmentList.length + 1}`,
                          symptom: 'ตรวจพบอาการเพิ่มเติมหน้างาน',
                          detail: '',
                          beforePhotos: [DEFAULT_BEFORE_IMAGE],
                          afterPhotos: [],
                          analysis: '',
                          cause: '',
                          standardsReference: '',
                          faultOrigin: editingTicket.warranty.includes('In-Warranty')
                            ? 'เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)'
                            : 'เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)',
                          correctiveAction: '',
                          parts: [],
                          partsApproved: false,
                          checklistResults: buildDefaultChecklistItems(nextType, masterChecklists)
                        }
                      ]
                    });
                    showToast('เพิ่มรายการอุปกรณ์ในใบงานนี้แล้ว', 'success');
                  }}
                  className="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1 rounded-lg font-semibold inline-flex items-center gap-1 cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>เพิ่มอุปกรณ์ในใบงานนี้</span>
                </button>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                <div>
                  <span className="text-[11px] text-slate-600 block mb-0.5">
                    ชื่อผู้ติดต่อหน้างานและเบอร์โทร
                  </span>
                  <input
                    type="text"
                    value={editingTicket.siteContact || editingTicket.contact}
                    onChange={(e) =>
                      setEditingTicket({ ...editingTicket, siteContact: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold"
                  />
                </div>
                <div>
                  <span className="text-[11px] text-slate-600 block mb-0.5">
                    สถานที่ / โลเคชั่นหน้างาน
                  </span>
                  <input
                    type="text"
                    value={editingTicket.location || ''}
                    onChange={(e) =>
                      setEditingTicket({ ...editingTicket, location: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold"
                  />
                </div>
                <div>
                  <span className="text-[11px] text-slate-600 block mb-0.5">
                    ช่างผู้ปฏิบัติงาน
                  </span>
                  <select
                    value={editingTicket.assignedTech || ''}
                    onChange={(e) =>
                      setEditingTicket({ ...editingTicket, assignedTech: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold"
                  >
                    <option value="">-- เลือกช่างผู้ปฏิบัติงาน --</option>
                    {technicians.map((tc) => (
                      <option key={tc.id} value={tc.name}>
                        {tc.name} ({tc.position || 'ช่างเทคนิค'})
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <span className="text-[11px] text-slate-600 block mb-0.5">
                    หัวหน้าช่างผู้ควบคุมงาน (Lead Tech)
                  </span>
                  <select
                    value={editingTicket.leadTech || ''}
                    onChange={(e) =>
                      setEditingTicket({ ...editingTicket, leadTech: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg p-2 bg-white font-semibold text-indigo-700"
                  >
                    <option value="">-- เลือกหัวหน้าช่าง --</option>
                    {technicians.map((tc) => (
                      <option key={`modal-lead-${tc.id}`} value={tc.name}>
                        {tc.name} ({tc.position || 'ช่างเทคนิค'})
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {/* Equipment Items */}
            <div className="space-y-5">
              {editingTicket.equipmentList.map((eq, idx) => {
                const draft = partDrafts[idx] || { name: '', qty: '1', price: '' };
                const isAiLoading = !!aiLoadingByIdx[idx];
                const aiSuggestion = aiSuggestionsByIdx[idx];
                const isEditingCustomType = !!modalCustomTypeByIdx[idx];
                const isEditingCustomBrand = !!modalCustomBrandByIdx[idx];

                return (
                  <div
                    key={idx}
                    className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3.5 text-xs"
                  >
                    {/* Equipment Header & Type/Brand Edit Controls */}
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-2.5 gap-2">
                      <h4 className="font-bold text-sm text-orange-600">
                        อุปกรณ์ #{idx + 1}: {eq.type} ({eq.brand} - {eq.model})
                      </h4>
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-slate-500">Tag: {eq.tag || 'N/A'}</span>
                        {editingTicket.equipmentList.length > 1 && (
                          <button
                            type="button"
                            onClick={() => {
                              setEditingTicket({
                                ...editingTicket,
                                equipmentList: editingTicket.equipmentList.filter(
                                  (_, i) => i !== idx
                                )
                              });
                            }}
                            className="text-red-600 hover:text-red-700 font-semibold inline-flex items-center gap-1 cursor-pointer"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            <span>ลบอุปกรณ์</span>
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Editable Equipment Type, Brand, Model, Tag, Symptom & Detail */}
                    <div className="bg-white p-3 rounded-lg border border-slate-200 space-y-2.5">
                      <div className="grid grid-cols-1 sm:grid-cols-4 gap-2">
                        {/* Equipment Type */}
                        <div>
                          <div className="flex justify-between items-center mb-1">
                            <label className="font-semibold text-slate-700">ประเภทอุปกรณ์:</label>
                            <button
                              type="button"
                              onClick={() =>
                                setModalCustomTypeByIdx((prev) => ({
                                  ...prev,
                                  [idx]: !prev[idx]
                                }))
                              }
                              className="text-[10px] text-orange-600 font-semibold hover:underline cursor-pointer"
                            >
                              {isEditingCustomType ? 'เลือกจากรายการ' : '+ พิมพ์เพิ่มเอง'}
                            </button>
                          </div>
                          {isEditingCustomType ? (
                            <input
                              type="text"
                              value={eq.type}
                              onChange={(e) => {
                                const val = e.target.value.toUpperCase();
                                const updated = [...editingTicket.equipmentList];
                                updated[idx] = { ...eq, type: val };
                                setEditingTicket({ ...editingTicket, equipmentList: updated });
                              }}
                              placeholder="พิมพ์ประเภทอุปกรณ์..."
                              className="border border-orange-400 rounded-lg p-1.5 w-full bg-orange-50/30 font-semibold"
                            />
                          ) : (
                            <select
                              value={
                                equipmentOptions.includes(eq.type)
                                  ? eq.type
                                  : 'ETC. (กรอกเพิ่มเองได้)'
                              }
                              onChange={(e) => {
                                const val = e.target.value;
                                if (val.startsWith('ETC.')) {
                                  setModalCustomTypeByIdx((prev) => ({ ...prev, [idx]: true }));
                                  return;
                                }
                                const updated = [...editingTicket.equipmentList];
                                updated[idx] = {
                                  ...eq,
                                  type: val,
                                  checklistResults: buildDefaultChecklistItems(val, masterChecklists)
                                };
                                setEditingTicket({ ...editingTicket, equipmentList: updated });
                              }}
                              className="border border-slate-300 rounded-lg p-1.5 w-full bg-white font-semibold"
                            >
                              {equipmentOptions.map((m) => (
                                <option key={m} value={m}>
                                  {m}
                                </option>
                              ))}
                            </select>
                          )}
                        </div>

                        {/* Brand */}
                        <div>
                          <div className="flex justify-between items-center mb-1">
                            <label className="font-semibold text-slate-700">แบรนด์ (Brand):</label>
                            <button
                              type="button"
                              onClick={() =>
                                setModalCustomBrandByIdx((prev) => ({
                                  ...prev,
                                  [idx]: !prev[idx]
                                }))
                              }
                              className="text-[10px] text-orange-600 font-semibold hover:underline cursor-pointer"
                            >
                              {isEditingCustomBrand ? 'เลือกจากรายการ' : '+ พิมพ์เพิ่มเอง'}
                            </button>
                          </div>
                          {isEditingCustomBrand ? (
                            <input
                              type="text"
                              value={eq.brand}
                              onChange={(e) => {
                                const val = e.target.value.toUpperCase();
                                const updated = [...editingTicket.equipmentList];
                                updated[idx] = { ...eq, brand: val };
                                setEditingTicket({ ...editingTicket, equipmentList: updated });
                              }}
                              placeholder="พิมพ์แบรนด์เอง..."
                              className="border border-orange-400 rounded-lg p-1.5 w-full bg-orange-50/30 font-semibold"
                            />
                          ) : (
                            <select
                              value={
                                brandOptions.includes(eq.brand)
                                  ? eq.brand
                                  : 'ETC. (กรอกเพิ่มเองได้)'
                              }
                              onChange={(e) => {
                                const val = e.target.value;
                                if (val.startsWith('ETC.')) {
                                  setModalCustomBrandByIdx((prev) => ({ ...prev, [idx]: true }));
                                  return;
                                }
                                const updated = [...editingTicket.equipmentList];
                                updated[idx] = { ...eq, brand: val };
                                setEditingTicket({ ...editingTicket, equipmentList: updated });
                              }}
                              className="border border-slate-300 rounded-lg p-1.5 w-full bg-white font-semibold"
                            >
                              {brandOptions.map((b) => (
                                <option key={b} value={b}>
                                  {b}
                                </option>
                              ))}
                            </select>
                          )}
                        </div>

                        {/* Model */}
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">รุ่น (Model):</label>
                          <input
                            type="text"
                            value={eq.model}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, model: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="border border-slate-300 rounded-lg p-1.5 w-full"
                          />
                        </div>

                        {/* Tag */}
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">Tag No.:</label>
                          <input
                            type="text"
                            value={eq.tag || ''}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, tag: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="border border-slate-300 rounded-lg p-1.5 w-full font-mono"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 pt-1">
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            อาการเสียที่พบ:
                          </label>
                          <input
                            type="text"
                            value={eq.symptom}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, symptom: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="border border-slate-300 rounded-lg p-1.5 w-full"
                          />
                        </div>
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            รายละเอียดหน้างานเพิ่มเติม:
                          </label>
                          <input
                            type="text"
                            value={eq.detail || ''}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, detail: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="border border-slate-300 rounded-lg p-1.5 w-full"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Before & After Photo Upload (Camera + Album/Files + Auto Compression) */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-white p-3 rounded-xl border border-slate-200">
                      {/* Before Photos */}
                      <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-bold text-orange-700 flex items-center gap-1.5">
                            <Camera className="w-3.5 h-3.5" />
                            <span>รูปภาพก่อนซ่อม / สภาพเครื่อง (Before)</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-medium">
                            {(eq.beforePhotos || []).length} รูป
                          </span>
                        </div>
                        <div className="flex flex-wrap items-center gap-1.5">
                          <label className="cursor-pointer inline-flex items-center gap-1.5 bg-orange-600 hover:bg-orange-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                            <Camera className="w-3.5 h-3.5" />
                            <span>ถ่ายรูป (กล้อง)</span>
                            <input
                              type="file"
                              accept="image/*"
                              capture="environment"
                              className="hidden"
                              onChange={(e) =>
                                handleEditingEquipmentPhotoUpload(idx, 'beforePhotos', e.target.files)
                              }
                            />
                          </label>
                          <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-orange-50 text-orange-700 border border-orange-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                            <Upload className="w-3.5 h-3.5" />
                            <span>อัลบั้ม / ไฟล์</span>
                            <input
                              type="file"
                              accept="image/*"
                              multiple
                              className="hidden"
                              onChange={(e) =>
                                handleEditingEquipmentPhotoUpload(idx, 'beforePhotos', e.target.files)
                              }
                            />
                          </label>
                        </div>
                        <input
                          type="text"
                          value={eq.beforePhotos[0] || ''}
                          onChange={(e) => {
                            const updated = [...editingTicket.equipmentList];
                            updated[idx] = {
                              ...eq,
                              beforePhotos: e.target.value ? [e.target.value] : []
                            };
                            setEditingTicket({ ...editingTicket, equipmentList: updated });
                          }}
                          placeholder="หรือวาง URL รูป Before..."
                          className="border border-slate-300 rounded-lg p-1.5 w-full text-[11px] bg-white"
                        />
                        {(eq.beforePhotos || []).length > 0 && (
                          <div className="flex flex-wrap gap-2 pt-1">
                            {eq.beforePhotos.map((photo, pIdx) => (
                              <div
                                key={pIdx}
                                className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-orange-200 bg-white shadow-xs"
                              >
                                <img
                                  src={photo}
                                  alt={`Before ${pIdx + 1}`}
                                  className="w-full h-full object-cover"
                                />
                                <button
                                  type="button"
                                  onClick={() =>
                                    handleRemoveEditingEquipmentPhoto(idx, 'beforePhotos', pIdx)
                                  }
                                  className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                  title="ลบรูปนี้"
                                >
                                  ×
                                </button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* After Photos */}
                      <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-2">
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-bold text-emerald-700 flex items-center gap-1.5">
                            <Camera className="w-3.5 h-3.5" />
                            <span>รูปภาพหลังซ่อม / ตรวจเช็ค (After)</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-medium">
                            {(eq.afterPhotos || []).length} รูป
                          </span>
                        </div>
                        <div className="flex flex-wrap items-center gap-1.5">
                          <label className="cursor-pointer inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                            <Camera className="w-3.5 h-3.5" />
                            <span>ถ่ายรูป (กล้อง)</span>
                            <input
                              type="file"
                              accept="image/*"
                              capture="environment"
                              className="hidden"
                              onChange={(e) =>
                                handleEditingEquipmentPhotoUpload(idx, 'afterPhotos', e.target.files)
                              }
                            />
                          </label>
                          <label className="cursor-pointer inline-flex items-center gap-1.5 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-colors shadow-xs">
                            <Upload className="w-3.5 h-3.5" />
                            <span>อัลบั้ม / ไฟล์</span>
                            <input
                              type="file"
                              accept="image/*"
                              multiple
                              className="hidden"
                              onChange={(e) =>
                                handleEditingEquipmentPhotoUpload(idx, 'afterPhotos', e.target.files)
                              }
                            />
                          </label>
                        </div>
                        <input
                          type="text"
                          value={eq.afterPhotos[0] || ''}
                          onChange={(e) => {
                            const updated = [...editingTicket.equipmentList];
                            updated[idx] = {
                              ...eq,
                              afterPhotos: e.target.value ? [e.target.value] : []
                            };
                            setEditingTicket({ ...editingTicket, equipmentList: updated });
                          }}
                          placeholder="หรือวาง URL รูป After..."
                          className="border border-slate-300 rounded-lg p-1.5 w-full text-[11px] bg-white"
                        />
                        {(eq.afterPhotos || []).length > 0 && (
                          <div className="flex flex-wrap gap-2 pt-1">
                            {eq.afterPhotos.map((photo, pIdx) => (
                              <div
                                key={pIdx}
                                className="relative group w-14 h-14 rounded-lg overflow-hidden border-2 border-emerald-200 bg-white shadow-xs"
                              >
                                <img
                                  src={photo}
                                  alt={`After ${pIdx + 1}`}
                                  className="w-full h-full object-cover"
                                />
                                <button
                                  type="button"
                                  onClick={() =>
                                    handleRemoveEditingEquipmentPhoto(idx, 'afterPhotos', pIdx)
                                  }
                                  className="absolute top-0.5 right-0.5 bg-red-600 hover:bg-red-700 text-white rounded-full w-4 h-4 flex items-center justify-center text-[10px] font-bold shadow-xs cursor-pointer"
                                  title="ลบรูปนี้"
                                >
                                  ×
                                </button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Standard Inspection Checklist (with Measured Value, Unit, Acceptance Criteria, OK / NG / N/A) */}
                    <div className="bg-white p-3.5 rounded-lg border border-slate-200 space-y-2.5">
                      <div className="flex flex-wrap justify-between items-center gap-2">
                        <div>
                          <span className="font-bold text-slate-900 flex items-center gap-1.5">
                            <ClipboardCheck className="w-4 h-4 text-orange-500" />
                            <span>
                              ตารางตรวจเช็คและวัดค่าจริงตามมาตรฐาน ({eq.type}) — รองรับ ค่าที่วัดได้ / หน่วย / เกณฑ์ยอมรับ / OK · NG · N/A
                            </span>
                          </span>
                          <p className="text-[11px] text-slate-500 mt-0.5">
                            ระบุค่าที่วัดได้จริงหน้างาน (เช่น Megger, Vibration, Alignment, แรงดัน, กระแส, อุณหภูมิ) หรือเลือก N/A สำหรับรายการที่ไม่เกี่ยวข้อง
                          </p>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = {
                                ...eq,
                                checklistResults: buildDefaultChecklistItems(eq.type, masterChecklists)
                              };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                              showToast(
                                `โหลดหัวข้อตรวจเช็คพร้อมหน่วยและเกณฑ์มาตรฐาน ${eq.type} แล้ว`,
                                'success'
                              );
                            }}
                            className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer"
                          >
                            โหลดหัวข้อมาตรฐาน {eq.type}
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = {
                                ...eq,
                                checklistResults: [
                                  ...eq.checklistResults,
                                  {
                                    item: 'หัวข้อตรวจสอบเพิ่มเติม',
                                    measuredValue: '-',
                                    unit: '-',
                                    acceptanceCriteria: 'ตามมาตรฐานผู้ผลิต',
                                    beforeStatus: 'NG',
                                    afterStatus: 'OK'
                                  }
                                ]
                              };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="bg-slate-900 hover:bg-slate-800 text-white px-2.5 py-1 rounded-md text-[11px] font-semibold inline-flex items-center gap-1 cursor-pointer"
                          >
                            <Plus className="w-3 h-3" />
                            <span>เพิ่มหัวข้อตรวจเช็ค</span>
                          </button>
                        </div>
                      </div>

                      {/* Column Headers on Desktop */}
                      <div className="hidden lg:grid lg:grid-cols-12 gap-2 px-2.5 py-1.5 bg-slate-100 rounded-lg text-[11px] font-bold text-slate-700">
                        <div className="col-span-4">หัวข้อการตรวจสอบ (Inspection Item)</div>
                        <div className="col-span-2">ค่าที่วัดได้จริง (Measured)</div>
                        <div className="col-span-1">หน่วย (Unit)</div>
                        <div className="col-span-3">เกณฑ์ยอมรับ (Acceptance Criteria)</div>
                        <div className="col-span-2 text-center">ผลตรวจ (Before / After)</div>
                      </div>

                      <div className="space-y-2">
                        {eq.checklistResults.map((rawChk, cIdx) => {
                          const chk = enrichChecklistItem(rawChk);
                          return (
                            <div
                              key={cIdx}
                              className="grid grid-cols-1 lg:grid-cols-12 gap-2 items-center bg-slate-50 p-2.5 rounded-lg border border-slate-200"
                            >
                              {/* 1. Inspection Item */}
                              <div className="lg:col-span-4">
                                <label className="lg:hidden text-[10px] font-bold text-slate-500 block mb-0.5">
                                  หัวข้อการตรวจเช็ค:
                                </label>
                                <input
                                  type="text"
                                  value={chk.item}
                                  onChange={(e) => {
                                    const updated = [...editingTicket.equipmentList];
                                    const newChk = [...eq.checklistResults];
                                    newChk[cIdx] = { ...chk, item: e.target.value };
                                    updated[idx] = { ...eq, checklistResults: newChk };
                                    setEditingTicket({ ...editingTicket, equipmentList: updated });
                                  }}
                                  placeholder="หัวข้อการตรวจสอบ..."
                                  className="border border-slate-300 rounded px-2 py-1 w-full bg-white text-xs font-medium"
                                />
                              </div>

                              {/* 2. Measured Value */}
                              <div className="lg:col-span-2">
                                <label className="lg:hidden text-[10px] font-bold text-blue-700 block mb-0.5">
                                  ค่าที่วัดได้จริง (Measured):
                                </label>
                                <input
                                  type="text"
                                  value={chk.measuredValue || ''}
                                  onChange={(e) => {
                                    const updated = [...editingTicket.equipmentList];
                                    const newChk = [...eq.checklistResults];
                                    newChk[cIdx] = { ...chk, measuredValue: e.target.value };
                                    updated[idx] = { ...eq, checklistResults: newChk };
                                    setEditingTicket({ ...editingTicket, equipmentList: updated });
                                  }}
                                  placeholder="เช่น 550 หรือ ก่อน 0.5/หลัง 550"
                                  className="border border-blue-300 rounded px-2 py-1 w-full bg-white text-xs font-mono font-bold text-blue-800"
                                />
                              </div>

                              {/* 3. Unit */}
                              <div className="lg:col-span-1">
                                <label className="lg:hidden text-[10px] font-bold text-slate-500 block mb-0.5">
                                  หน่วย (Unit):
                                </label>
                                <input
                                  type="text"
                                  list="engineering-units-list"
                                  value={chk.unit || ''}
                                  onChange={(e) => {
                                    const updated = [...editingTicket.equipmentList];
                                    const newChk = [...eq.checklistResults];
                                    newChk[cIdx] = { ...chk, unit: e.target.value };
                                    updated[idx] = { ...eq, checklistResults: newChk };
                                    setEditingTicket({ ...editingTicket, equipmentList: updated });
                                  }}
                                  placeholder="หน่วย"
                                  className="border border-slate-300 rounded px-2 py-1 w-full bg-white text-xs font-mono text-center"
                                />
                              </div>

                              {/* 4. Acceptance Criteria */}
                              <div className="lg:col-span-3">
                                <label className="lg:hidden text-[10px] font-bold text-slate-500 block mb-0.5">
                                  เกณฑ์ยอมรับ (Acceptance Criteria):
                                </label>
                                <input
                                  type="text"
                                  value={chk.acceptanceCriteria || ''}
                                  onChange={(e) => {
                                    const updated = [...editingTicket.equipmentList];
                                    const newChk = [...eq.checklistResults];
                                    newChk[cIdx] = { ...chk, acceptanceCriteria: e.target.value };
                                    updated[idx] = { ...eq, checklistResults: newChk };
                                    setEditingTicket({ ...editingTicket, equipmentList: updated });
                                  }}
                                  placeholder="เช่น ≥ 100 MΩ, ≤ 4.5 mm/s"
                                  className="border border-slate-300 rounded px-2 py-1 w-full bg-white text-xs text-slate-700"
                                />
                              </div>

                              {/* 5. Before & After Status (OK / NG / N/A) + Delete */}
                              <div className="lg:col-span-2 flex items-center justify-between lg:justify-end gap-1.5">
                                <div className="flex items-center gap-1">
                                  <span className="text-[10px] text-slate-500">ก่อน:</span>
                                  <select
                                    value={chk.beforeStatus}
                                    onChange={(e) => {
                                      const updated = [...editingTicket.equipmentList];
                                      const newChk = [...eq.checklistResults];
                                      newChk[cIdx] = {
                                        ...chk,
                                        beforeStatus: e.target.value as ChecklistItem['beforeStatus']
                                      };
                                      updated[idx] = { ...eq, checklistResults: newChk };
                                      setEditingTicket({ ...editingTicket, equipmentList: updated });
                                    }}
                                    className={`border rounded px-1.5 py-1 font-mono font-bold text-xs ${
                                      chk.beforeStatus === 'OK'
                                        ? 'bg-emerald-50 border-emerald-400 text-emerald-700'
                                        : chk.beforeStatus === 'NG'
                                        ? 'bg-red-50 border-red-400 text-red-600'
                                        : 'bg-slate-100 border-slate-300 text-slate-600'
                                    }`}
                                  >
                                    <option value="OK">OK</option>
                                    <option value="NG">NG</option>
                                    <option value="N/A">N/A</option>
                                  </select>
                                </div>

                                <div className="flex items-center gap-1">
                                  <span className="text-[10px] text-slate-500">หลัง:</span>
                                  <select
                                    value={chk.afterStatus}
                                    onChange={(e) => {
                                      const updated = [...editingTicket.equipmentList];
                                      const newChk = [...eq.checklistResults];
                                      newChk[cIdx] = {
                                        ...chk,
                                        afterStatus: e.target.value as ChecklistItem['afterStatus']
                                      };
                                      updated[idx] = { ...eq, checklistResults: newChk };
                                      setEditingTicket({ ...editingTicket, equipmentList: updated });
                                    }}
                                    className={`border rounded px-1.5 py-1 font-mono font-bold text-xs ${
                                      chk.afterStatus === 'OK'
                                        ? 'bg-emerald-50 border-emerald-400 text-emerald-700'
                                        : chk.afterStatus === 'NG'
                                        ? 'bg-red-50 border-red-400 text-red-600'
                                        : 'bg-slate-100 border-slate-300 text-slate-600'
                                    }`}
                                  >
                                    <option value="OK">OK</option>
                                    <option value="NG">NG</option>
                                    <option value="N/A">N/A</option>
                                  </select>
                                </div>

                                <button
                                  type="button"
                                  onClick={() => {
                                    const updated = [...editingTicket.equipmentList];
                                    updated[idx] = {
                                      ...eq,
                                      checklistResults: eq.checklistResults.filter(
                                        (_, i) => i !== cIdx
                                      )
                                    };
                                    setEditingTicket({ ...editingTicket, equipmentList: updated });
                                  }}
                                  className="text-red-600 hover:text-red-700 font-bold px-1 cursor-pointer"
                                  title="ลบหัวข้อนี้"
                                >
                                  ×
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      <datalist id="engineering-units-list">
                        <option value="MΩ" />
                        <option value="Ω" />
                        <option value="V" />
                        <option value="A" />
                        <option value="V / A" />
                        <option value="mm/s" />
                        <option value="mm" />
                        <option value="°C" />
                        <option value="bar" />
                        <option value="m³/h" />
                        <option value="rpm" />
                        <option value="%" />
                        <option value="-" />
                      </datalist>
                    </div>

                    {/* =========================================================
                        AI ENGINEERING TROUBLESHOOTER + TECH MANUAL EDITOR
                        (Pump Handbook, EASA Motor, Manufacturer Guides)
                    ========================================================= */}
                    <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-3">
                      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-2.5 border-b border-slate-100 pb-2.5">
                        <div>
                          <span className="font-bold text-slate-900 flex items-center gap-1.5 text-xs sm:text-sm">
                            <BookOpen className="w-4 h-4 text-orange-600" />
                            <span>
                              ผลการวิเคราะห์ปัญหาและแนวทางแก้ไข (AI Assistant + ช่างวิเคราะห์และเพิ่มข้อมูลเองได้)
                            </span>
                          </span>
                          <p className="text-[11px] text-slate-500 mt-0.5">
                            อ้างอิง Karassik&apos;s Pump Handbook, EASA AR100 Motor Standard และคู่มือผู้ผลิต ({eq.brand})
                          </p>
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          {/* Button 1: Run Gemini AI Troubleshooter */}
                          <button
                            type="button"
                            disabled={isAiLoading}
                            onClick={() => handleRunAiTroubleshoot(idx)}
                            className="bg-slate-900 hover:bg-slate-800 disabled:opacity-60 text-white px-3 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            {isAiLoading ? (
                              <>
                                <Loader2 className="w-3.5 h-3.5 animate-spin text-orange-400" />
                                <span>กำลังวิเคราะห์ด้วย AI...</span>
                              </>
                            ) : (
                              <>
                                <Cpu className="w-3.5 h-3.5 text-orange-400" />
                                <span>วิเคราะห์ด้วย AI (Pump Handbook / EASA)</span>
                              </>
                            )}
                          </button>

                          {/* Button 2: Save current technician analysis as reusable template */}
                          <button
                            type="button"
                            onClick={() => {
                              setSaveTemplateForIdx(saveTemplateForIdx === idx ? null : idx);
                              setNewTemplateTitle(`${eq.type} (${eq.brand}) - ${eq.symptom}`);
                            }}
                            className="bg-orange-50 hover:bg-orange-100 text-orange-700 border border-orange-200 px-2.5 py-1.5 rounded-lg font-semibold inline-flex items-center gap-1 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <BookmarkPlus className="w-3.5 h-3.5" />
                            <span>บันทึกเป็นต้นแบบของช่าง</span>
                          </button>
                        </div>
                      </div>

                      {/* Quick Select from Built-in & Custom Technician Knowledge Base */}
                      <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
                        <span className="text-[11px] font-semibold text-slate-700 shrink-0">
                          เลือกดึงข้อมูลจากคลังมาตรฐานวิศวกรรม / ต้นแบบที่ช่างบันทึกไว้:
                        </span>
                        <select
                          defaultValue=""
                          onChange={(e) => {
                            if (e.target.value) {
                              applyKnowledgeTemplate(idx, e.target.value);
                              e.target.value = '';
                            }
                          }}
                          className="w-full sm:w-auto flex-grow border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white text-xs font-medium text-slate-800"
                        >
                          <option value="">
                            -- คลิกเลือกต้นแบบการวิเคราะห์ (Pump Handbook / EASA / คลังความรู้ช่าง) --
                          </option>
                          {knowledgeTemplates.map((tpl) => (
                            <option key={tpl.id} value={tpl.id}>
                              [{tpl.category}] {tpl.title}
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* Save Custom Template Inline Box */}
                      {saveTemplateForIdx === idx && (
                        <div className="bg-orange-50 p-3 rounded-lg border border-orange-200 flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                          <input
                            type="text"
                            value={newTemplateTitle}
                            onChange={(e) => setNewTemplateTitle(e.target.value)}
                            placeholder="ตั้งชื่อต้นแบบการวิเคราะห์เพื่อใช้ซ้ำในงานถัดไป..."
                            className="flex-grow border border-orange-300 rounded-lg px-2.5 py-1.5 bg-white text-xs font-semibold"
                          />
                          <button
                            type="button"
                            onClick={() => handleSaveCustomTemplate(idx)}
                            className="bg-orange-600 hover:bg-orange-700 text-white px-3.5 py-1.5 rounded-lg font-semibold text-xs inline-flex items-center justify-center gap-1 whitespace-nowrap cursor-pointer"
                          >
                            <Check className="w-3.5 h-3.5" />
                            <span>ยืนยันบันทึกเข้าคลัง</span>
                          </button>
                        </div>
                      )}

                      {/* AI Recommended Checklist & Spare Parts Quick-Add Chips */}
                      {aiSuggestion &&
                        (aiSuggestion.recommendedChecklist.length > 0 ||
                          aiSuggestion.recommendedParts.length > 0) && (
                          <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 space-y-2">
                            {aiSuggestion.recommendedChecklist.length > 0 && (
                              <div className="space-y-1">
                                <p className="text-[11px] font-bold text-slate-700">
                                  หัวข้อตรวจเช็คเพิ่มเติมที่ AI แนะนำ (คลิกเพื่อเพิ่มลงตาราง Checklist):
                                </p>
                                <div className="flex flex-wrap gap-1.5">
                                  {aiSuggestion.recommendedChecklist.map((item, rIdx) => (
                                    <button
                                      key={rIdx}
                                      type="button"
                                      onClick={() => {
                                        const meta = getChecklistMetaForTitle(item);
                                        const updated = [...editingTicket.equipmentList];
                                        updated[idx] = {
                                          ...eq,
                                          checklistResults: [
                                            ...eq.checklistResults,
                                            {
                                              item,
                                              measuredValue: meta.measuredValue,
                                              unit: meta.unit,
                                              acceptanceCriteria: meta.acceptanceCriteria,
                                              beforeStatus: 'NG',
                                              afterStatus: 'OK'
                                            }
                                          ]
                                        };
                                        setEditingTicket({
                                          ...editingTicket,
                                          equipmentList: updated
                                        });
                                        showToast(`เพิ่มหัวข้อ "${item}" ลงตารางตรวจเช็คแล้ว`, 'success');
                                      }}
                                      className="bg-white hover:bg-orange-50 text-slate-800 border border-slate-300 hover:border-orange-400 px-2.5 py-1 rounded-md text-[11px] font-medium inline-flex items-center gap-1 cursor-pointer"
                                    >
                                      <Plus className="w-3 h-3 text-orange-600" />
                                      <span>{item}</span>
                                    </button>
                                  ))}
                                </div>
                              </div>
                            )}

                            {aiSuggestion.recommendedParts.length > 0 && (
                              <div className="space-y-1 pt-1">
                                <p className="text-[11px] font-bold text-slate-700">
                                  อะไหล่ที่ AI แนะนำให้ตรวจสอบ/เปลี่ยน (คลิกเพื่อนำชื่อไปกรอกช่องเบิกอะไหล่):
                                </p>
                                <div className="flex flex-wrap gap-1.5">
                                  {aiSuggestion.recommendedParts.map((part, pIdx) => (
                                    <button
                                      key={pIdx}
                                      type="button"
                                      onClick={() => {
                                        setPartDrafts((prev) => ({
                                          ...prev,
                                          [idx]: {
                                            name: part.name,
                                            qty: String(part.qty || 1),
                                            price: prev[idx]?.price || ''
                                          }
                                        }));
                                        showToast(
                                          `นำ "${part.name}" ไปใส่ช่องเพิ่มอะไหล่แล้ว กรุณาระบุราคาและกดเพิ่มอะไหล่`,
                                          'success'
                                        );
                                      }}
                                      className="bg-white hover:bg-amber-50 text-slate-800 border border-slate-300 hover:border-amber-400 px-2.5 py-1 rounded-md text-[11px] font-medium inline-flex items-center gap-1 cursor-pointer"
                                    >
                                      <Plus className="w-3 h-3 text-amber-600" />
                                      <span>
                                        {part.name} ({part.qty} {part.unit})
                                      </span>
                                    </button>
                                  ))}
                                </div>
                              </div>
                            )}
                          </div>
                        )}

                      {/* Editable Analysis Fields (Technician can edit or type freely) */}
                      <div>
                        <label className="font-semibold text-slate-700 mb-1 block">
                          ผลการวิเคราะห์ทางเทคนิค (ช่างสามารถพิมพ์วิเคราะห์เองหรือแก้ไขเพิ่มเติมได้):
                        </label>
                        <textarea
                          rows={2}
                          value={eq.analysis || ''}
                          onChange={(e) => {
                            const updated = [...editingTicket.equipmentList];
                            updated[idx] = { ...eq, analysis: e.target.value };
                            setEditingTicket({ ...editingTicket, equipmentList: updated });
                          }}
                          placeholder="ระบุผลการวิเคราะห์ทางเทคนิค หรือกดปุ่มวิเคราะห์ด้วย AI ด้านบน..."
                          className="w-full border border-slate-300 rounded-lg p-2 focus:outline-none focus:ring-2 focus:ring-orange-500"
                        />
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            สาเหตุรากฐาน (Root Cause):
                          </label>
                          <input
                            type="text"
                            value={eq.cause || ''}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, cause: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            placeholder="ระบุสาเหตุของปัญหา..."
                            className="w-full border border-slate-300 rounded-lg p-2"
                          />
                        </div>
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            มาตรฐานวิศวกรรม / คู่มือที่ใช้อ้างอิง:
                          </label>
                          <input
                            type="text"
                            value={eq.standardsReference || ''}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, standardsReference: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            placeholder="เช่น Karassik Pump Handbook / EASA AR100 / คู่มือ STAC"
                            className="w-full border border-slate-300 rounded-lg p-2"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            เงื่อนไขการประกัน / คิดเงิน:
                          </label>
                          <select
                            value={eq.faultOrigin}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, faultOrigin: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            className="w-full border border-slate-300 rounded-lg p-2 font-semibold text-orange-600 bg-orange-50/50"
                          >
                            <option>เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)</option>
                            <option>เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)</option>
                          </select>
                        </div>
                        <div>
                          <label className="font-semibold text-slate-700 mb-1 block">
                            แนวทางแก้ไขและซ่อมบำรุง (Corrective Action):
                          </label>
                          <input
                            type="text"
                            value={eq.correctiveAction || ''}
                            onChange={(e) => {
                              const updated = [...editingTicket.equipmentList];
                              updated[idx] = { ...eq, correctiveAction: e.target.value };
                              setEditingTicket({ ...editingTicket, equipmentList: updated });
                            }}
                            placeholder="ระบุแนวทางแก้ไขและค่ามาตรฐานหลังซ่อม..."
                            className="w-full border border-slate-300 rounded-lg p-2"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Spare Parts Section (Separating Internal Warranty Cost vs Customer Billed Amount) */}
                    <div className="bg-amber-50/80 border border-amber-200 p-3.5 rounded-xl space-y-3">
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-amber-200/80 pb-2">
                        <div>
                          <span className="font-bold text-amber-950 block">
                            รายการอะไหล่ &amp; แยกต้นทุนภายใน vs ยอดเรียกเก็บลูกค้า (Spare Parts &amp; Billing)
                          </span>
                          <p className="text-[11px] text-amber-800">
                            กรณีอยู่ในประกัน (In-Warranty / FOC) ระบบจะบันทึกเป็น <b>&ldquo;ต้นทุนภายในบริษัท&rdquo;</b> และแสดง <b>&ldquo;ยอดเรียกเก็บลูกค้า = 0 บาท&rdquo;</b> เพื่อไม่ให้ลูกค้าเข้าใจผิดว่าต้องจ่าย
                          </p>
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            const updated = [...editingTicket.equipmentList];
                            const setAllInternal = !eq.faultOrigin?.includes('อยู่ในประกัน');
                            updated[idx] = {
                              ...eq,
                              faultOrigin: setAllInternal
                                ? 'เกิดจากผลิตภัณฑ์ / อยู่ในประกัน (ฟรี / FOC)'
                                : 'เกิดจากการใช้งาน / หมดรับประกัน (คิดค่าบริการ)',
                              parts: (eq.parts || []).map((p) => ({
                                ...p,
                                billingType: setAllInternal ? 'WARRANTY_INTERNAL' : 'CUSTOMER_CHARGE'
                              }))
                            };
                            setEditingTicket({ ...editingTicket, equipmentList: updated });
                            showToast(
                              setAllInternal
                                ? 'ปรับรายการอะไหล่ทั้งหมดเป็น "เคลมในประกัน (เรียกเก็บลูกค้า 0 บาท)" แล้ว'
                                : 'ปรับรายการอะไหล่ทั้งหมดเป็น "คิดเงินลูกค้า (Out-of-Warranty)" แล้ว',
                              'success'
                            );
                          }}
                          className="bg-white hover:bg-amber-100 text-amber-900 border border-amber-300 px-2.5 py-1 rounded-lg text-[11px] font-bold cursor-pointer"
                        >
                          สลับทั้งชุดเป็น: {eq.faultOrigin?.includes('อยู่ในประกัน') ? '💳 คิดเงินลูกค้า' : '🛡️ เคลมในประกัน (ฟรี 0฿)'}
                        </button>
                      </div>

                      <div className="space-y-1.5">
                        {eq.parts && eq.parts.length > 0 ? (
                          eq.parts.map((p, pIdx) => {
                            const isInternal = isPartInternalWarranty(p, eq, editingTicket);
                            return (
                              <div
                                key={pIdx}
                                className="flex flex-col sm:flex-row justify-between sm:items-center gap-2 py-2 px-2.5 bg-white rounded-lg border border-amber-200"
                              >
                                <div className="flex flex-wrap items-center gap-2">
                                  <span className="font-semibold text-slate-900">
                                    • {p.name} ({p.qty} {p.unit})
                                  </span>
                                  <select
                                    value={isInternal ? 'WARRANTY_INTERNAL' : 'CUSTOMER_CHARGE'}
                                    onChange={(e) => {
                                      const val = e.target.value as
                                        | 'WARRANTY_INTERNAL'
                                        | 'CUSTOMER_CHARGE';
                                      const updated = [...editingTicket.equipmentList];
                                      const newParts = [...eq.parts];
                                      newParts[pIdx] = { ...p, billingType: val };
                                      updated[idx] = { ...eq, parts: newParts };
                                      setEditingTicket({ ...editingTicket, equipmentList: updated });
                                    }}
                                    className={`text-[11px] font-bold px-2 py-0.5 rounded-md border ${
                                      isInternal
                                        ? 'bg-emerald-50 border-emerald-300 text-emerald-800'
                                        : 'bg-amber-50 border-amber-300 text-amber-900'
                                    }`}
                                  >
                                    <option value="WARRANTY_INTERNAL">
                                      🛡️ ต้นทุนภายใน / เคลมในประกัน (ลูกค้าจ่าย 0 บาท)
                                    </option>
                                    <option value="CUSTOMER_CHARGE">
                                      💳 เรียกเก็บลูกค้า (คิดค่าบริการ)
                                    </option>
                                  </select>
                                </div>
                                <div className="flex items-center justify-between sm:justify-end gap-3">
                                  <div className="text-right font-mono text-[11px] tabular-nums">
                                    <span className="text-slate-500 mr-2">
                                      ต้นทุน: {p.total.toLocaleString()} THB
                                    </span>
                                    <span
                                      className={`font-bold px-2 py-0.5 rounded ${
                                        isInternal
                                          ? 'bg-emerald-100 text-emerald-800'
                                          : 'bg-orange-100 text-orange-800'
                                      }`}
                                    >
                                      เรียกเก็บลูกค้า: {isInternal ? '0 THB (ฟรี)' : `${p.total.toLocaleString()} THB`}
                                    </span>
                                  </div>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const updated = [...editingTicket.equipmentList];
                                      updated[idx] = {
                                        ...eq,
                                        parts: eq.parts.filter((_, i) => i !== pIdx)
                                      };
                                      setEditingTicket({ ...editingTicket, equipmentList: updated });
                                    }}
                                    className="text-red-600 font-bold px-1.5 cursor-pointer"
                                  >
                                    ×
                                  </button>
                                </div>
                              </div>
                            );
                          })
                        ) : (
                          <p className="text-slate-400 italic">ยังไม่มีรายการอะไหล่</p>
                        )}
                      </div>

                      {/* Add Part Form Row */}
                      <div className="grid grid-cols-1 sm:grid-cols-12 gap-2 pt-1">
                        <input
                          type="text"
                          value={draft.name}
                          onChange={(e) =>
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: { ...draft, name: e.target.value }
                            })
                          }
                          placeholder="ชื่ออะไหล่..."
                          className="sm:col-span-4 border border-amber-300 rounded-lg p-1.5 bg-white"
                        />
                        <input
                          type="number"
                          value={draft.qty}
                          onChange={(e) =>
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: { ...draft, qty: e.target.value }
                            })
                          }
                          placeholder="จน."
                          className="sm:col-span-1 border border-amber-300 rounded-lg p-1.5 bg-white font-mono text-center"
                        />
                        <input
                          type="text"
                          value={draft.unit || 'PCS'}
                          onChange={(e) =>
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: { ...draft, unit: e.target.value.toUpperCase() }
                            })
                          }
                          placeholder="หน่วย"
                          className="sm:col-span-1 border border-amber-300 rounded-lg p-1.5 bg-white font-mono text-center"
                        />
                        <input
                          type="number"
                          value={draft.price}
                          onChange={(e) =>
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: { ...draft, price: e.target.value }
                            })
                          }
                          placeholder="ราคาต้นทุน/หน่วย"
                          className="sm:col-span-2 border border-amber-300 rounded-lg p-1.5 bg-white font-mono"
                        />
                        <select
                          value={
                            draft.billingType ||
                            (eq.faultOrigin?.includes('อยู่ในประกัน') ||
                            editingTicket.warranty.includes('In-Warranty')
                              ? 'WARRANTY_INTERNAL'
                              : 'CUSTOMER_CHARGE')
                          }
                          onChange={(e) =>
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: {
                                ...draft,
                                billingType: e.target.value as
                                  | 'WARRANTY_INTERNAL'
                                  | 'CUSTOMER_CHARGE'
                              }
                            })
                          }
                          className="sm:col-span-2 border border-amber-300 rounded-lg p-1.5 bg-white font-semibold text-[11px]"
                        >
                          <option value="WARRANTY_INTERNAL">🛡️ ในประกัน (ลูกค้า 0฿)</option>
                          <option value="CUSTOMER_CHARGE">💳 เรียกเก็บลูกค้า</option>
                        </select>
                        <button
                          type="button"
                          onClick={() => {
                            if (!draft.name.trim()) {
                              showToast('กรุณากรอกชื่ออะไหล่', 'error');
                              return;
                            }
                            const qty = parseFloat(draft.qty) || 1;
                            const price = parseFloat(draft.price) || 0;
                            const chosenBillingType: 'WARRANTY_INTERNAL' | 'CUSTOMER_CHARGE' =
                              draft.billingType ||
                              (eq.faultOrigin?.includes('อยู่ในประกัน') ||
                              editingTicket.warranty.includes('In-Warranty')
                                ? 'WARRANTY_INTERNAL'
                                : 'CUSTOMER_CHARGE');
                            const updated = [...editingTicket.equipmentList];
                            updated[idx] = {
                              ...eq,
                              parts: [
                                ...(eq.parts || []),
                                {
                                  name: draft.name.trim(),
                                  qty,
                                  unit: (draft.unit || 'PCS').trim(),
                                  price,
                                  total: qty * price,
                                  billingType: chosenBillingType
                                }
                              ]
                            };
                            const updatedTicket: ServiceTicket = {
                              ...editingTicket,
                              status: 'PENDING_PARTS_APPROVAL',
                              equipmentList: updated
                            };
                            setEditingTicket(updatedTicket);
                            const updatedList = tickets.map((item) =>
                              item.id === updatedTicket.id ? updatedTicket : item
                            );
                            setTickets(updatedList);
                            syncToGoogleSheetAndLine(
                              updatedList,
                              formatTicketLineMessage(
                                updatedTicket,
                                '⏳ [ขออนุมัติเบิกอะไหล่ด่วน! - รอ Supervisor อนุมัติ]'
                              )
                            );
                            setPartDrafts({
                              ...partDrafts,
                              [idx]: { name: '', qty: '1', unit: 'PCS', price: '' }
                            });
                            playNotificationChime(soundEnabled);
                            showToast(
                              'เพิ่มอะไหล่และส่งแจ้งเตือนขออนุมัติไปยัง Supervisor แล้ว',
                              'success'
                            );
                          }}
                          className="sm:col-span-2 bg-amber-600 hover:bg-amber-700 text-white px-3 py-1.5 rounded-lg font-semibold whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          + เพิ่มอะไหล่
                        </button>
                      </div>

                      {/* Equipment Cost Breakdown Strip */}
                      {(() => {
                        const internalSub = (eq.parts || []).reduce(
                          (sum, p) =>
                            sum + (isPartInternalWarranty(p, eq, editingTicket) ? p.total : 0),
                          0
                        );
                        const customerSub = (eq.parts || []).reduce(
                          (sum, p) =>
                            sum + (!isPartInternalWarranty(p, eq, editingTicket) ? p.total : 0),
                          0
                        );
                        return (
                          <div className="bg-slate-900 text-white p-2.5 rounded-lg flex flex-wrap items-center justify-between gap-2 text-[11px]">
                            <span className="text-slate-300">
                              🛡️ ต้นทุนภายในบริษัท (เคลมในประกัน):{' '}
                              <b className="font-mono text-emerald-400">
                                {internalSub.toLocaleString()} THB
                              </b>
                            </span>
                            <span className="text-white">
                              💳 ยอดเรียกเก็บลูกค้าสุทธิ (สำหรับอุปกรณ์นี้):{' '}
                              <b className="font-mono text-orange-400 text-xs">
                                {customerSub === 0
                                  ? '0.00 THB (อยู่ในประกัน - ลูกค้าไม่ต้องจ่าย)'
                                  : `${customerSub.toLocaleString()} THB`}
                              </b>
                            </span>
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Mobile Digital Signature Capture Section inside Action Modal */}
            <div className="bg-emerald-50/70 p-4 rounded-xl border-2 border-emerald-200 space-y-3 text-xs">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-emerald-200/80 pb-2">
                <div>
                  <h4 className="font-bold text-sm text-emerald-950 flex items-center gap-1.5">
                    <PenTool className="w-4 h-4 text-emerald-700" />
                    <span>
                      ลงลายมือชื่อดิจิทัลหน้างาน (Mobile Signature Sign-Off — สำหรับมือถือ / แท็บเล็ต)
                    </span>
                  </h4>
                  <p className="text-[11px] text-emerald-800 mt-0.5">
                    ให้ลูกค้าและช่างผู้ปฏิบัติงานใช้นิ้วเซ็นชื่อตรวจรับงานซ่อมบนหน้าจอมือถือได้ทันที (ลายเซ็นจะไปแสดงในใบ Service Report &amp; PDF อัตโนมัติ)
                  </p>
                </div>
                {editingTicket.customerSignature && (
                  <span className="bg-emerald-600 text-white px-2.5 py-1 rounded-full text-[11px] font-bold">
                    ✓ มีลายเซ็นลูกค้าพร้อมออกใบรายงานแล้ว
                  </span>
                )}
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                <SignaturePadCanvas
                  label="ลายเซ็นลูกค้าผู้ตรวจรับงาน (Customer Sign-Off)"
                  sublabel="ให้ลูกค้าใช้นิ้วเซ็นชื่อเพื่อยืนยันผลการซ่อม/ตรวจเช็ค"
                  signerNameLabel="ชื่อ-นามสกุล ลูกค้าผู้ลงนามตรวจรับงาน:"
                  signerNameValue={
                    editingTicket.customerSignerName ??
                    (editingTicket.siteContact || editingTicket.contact)
                  }
                  onSignerNameChange={(val) =>
                    setEditingTicket({
                      ...editingTicket,
                      customerSignerName: val
                    })
                  }
                  initialSignature={editingTicket.customerSignature}
                  signedAt={editingTicket.customerSignedAt}
                  compact
                  onSaveSignature={(dataUrl, timestamp) =>
                    setEditingTicket((prev) =>
                      prev
                        ? {
                            ...prev,
                            customerSignature: dataUrl,
                            customerSignedAt: timestamp,
                            customerSignerName:
                              prev.customerSignerName || prev.siteContact || prev.contact
                          }
                        : prev
                    )
                  }
                  onClearSignature={() =>
                    setEditingTicket((prev) =>
                      prev
                        ? {
                            ...prev,
                            customerSignature: '',
                            customerSignedAt: ''
                          }
                        : prev
                    )
                  }
                />

                <SignaturePadCanvas
                  label="ลายเซ็นช่างผู้ปฏิบัติงาน (Technician Signature)"
                  sublabel={`ช่างผู้รับผิดชอบ: ${editingTicket.assignedTech || 'ยังไม่ระบุ'}`}
                  initialSignature={editingTicket.technicianSignature}
                  signedAt={editingTicket.technicianSignedAt}
                  compact
                  onSaveSignature={(dataUrl, timestamp) =>
                    setEditingTicket((prev) =>
                      prev
                        ? {
                            ...prev,
                            technicianSignature: dataUrl,
                            technicianSignedAt: timestamp
                          }
                        : prev
                    )
                  }
                  onClearSignature={() =>
                    setEditingTicket((prev) =>
                      prev
                        ? {
                            ...prev,
                            technicianSignature: '',
                            technicianSignedAt: ''
                          }
                        : prev
                    )
                  }
                />
              </div>
            </div>

            {/* Modal Footer Buttons */}
            <div className="flex flex-wrap justify-end gap-2.5 pt-3 border-t border-slate-200 text-xs">
              <button
                type="button"
                onClick={() => setEditingTicket(null)}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-semibold cursor-pointer"
              >
                ปิดหน้าต่าง
              </button>
              <button
                type="button"
                onClick={() => {
                  // Save any custom types/brands typed in the modal into master lists
                  const newTypes = editingTicket.equipmentList
                    .map((eq) => eq.type.trim().toUpperCase())
                    .filter((t) => t && !DEFAULT_EQUIPMENT_MASTERS.includes(t));
                  const newBr = editingTicket.equipmentList
                    .map((eq) => eq.brand.trim().toUpperCase())
                    .filter((b) => b && !DEFAULT_BRANDS.includes(b));

                  if (newTypes.length > 0) {
                    setMasterEquipmentTypes((prev) => Array.from(new Set([...prev, ...newTypes])));
                  }
                  if (newBr.length > 0) {
                    setMasterBrands((prev) => Array.from(new Set([...prev, ...newBr])));
                  }

                  const updatedList = tickets.map((item) =>
                    item.id === editingTicket.id ? editingTicket : item
                  );
                  setTickets(updatedList);
                  syncToGoogleSheetAndLine(updatedList);
                  setEditingTicket(null);
                  showToast(`บันทึกข้อมูลใบงานและลายเซ็น ${editingTicket.id} เรียบร้อยแล้ว`, 'success');
                }}
                className="bg-orange-500 hover:bg-orange-600 text-white px-5 py-2 rounded-lg font-semibold cursor-pointer"
              >
                บันทึกข้อมูล &amp; ลายเซ็น
              </button>
              {editingTicket.customerSignature && editingTicket.status !== 'CLOSE' && (
                <button
                  type="button"
                  onClick={() => {
                    const signedAndClosed: ServiceTicket = {
                      ...editingTicket,
                      status: 'CLOSE',
                      isLocked: true
                    };
                    const updatedList = tickets.map((item) =>
                      item.id === signedAndClosed.id ? signedAndClosed : item
                    );
                    setTickets(updatedList);
                    syncToGoogleSheetAndLine(
                      updatedList,
                      formatTicketLineMessage(
                        signedAndClosed,
                        '🎉 [ลูกค้าเซ็นชื่อตรวจรับงานบนมือถือและปิดจบใบงาน (CLOSE) เรียบร้อยแล้ว]'
                      )
                    );
                    setEditingTicket(null);
                    playNotificationChime(soundEnabled);
                    showToast(
                      `บันทึกลายเซ็นลูกค้าและปิดจบใบงาน ${signedAndClosed.id} เรียบร้อยแล้ว`,
                      'success'
                    );
                  }}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-lg font-semibold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  <span>ยืนยันลายเซ็นลูกค้า &amp; ปิดจบงาน</span>
                </button>
              )}
              {(editingTicket.status === 'REPAIRING' ||
                editingTicket.status === 'REJECTED' ||
                editingTicket.status === 'ASSIGNED') && (
                <button
                  type="button"
                  onClick={() => {
                    const submitted: ServiceTicket = {
                      ...editingTicket,
                      status: 'PENDING_SUPERVISOR_APPROVAL'
                    };
                    const updatedList = tickets.map((item) =>
                      item.id === submitted.id ? submitted : item
                    );
                    setTickets(updatedList);
                    syncToGoogleSheetAndLine(
                      updatedList,
                      formatTicketLineMessage(
                        submitted,
                        '✍️ [ช่างส่งรายงานซ่อมแล้ว! - รอ Supervisor อนุมัติ]'
                      )
                    );
                    setEditingTicket(null);
                    playNotificationChime(soundEnabled);
                    showToast(
                      'ส่งรายงานให้ Supervisor อนุมัติ และแจ้งเตือนผ่าน LINE เรียบร้อยแล้ว',
                      'success'
                    );
                  }}
                  className="bg-purple-600 hover:bg-purple-700 text-white px-5 py-2 rounded-lg font-semibold cursor-pointer"
                >
                  ส่ง Supervisor อนุมัติรายงาน
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* 7.5 Dedicated Customer Mobile Sign-Off Modal */}
      {customerSignOffModalTicket && (
        <div className="fixed inset-0 bg-slate-900/70 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-lg rounded-2xl border border-slate-200 p-5 sm:p-6 space-y-4 shadow-2xl">
            <div className="flex justify-between items-start border-b border-slate-200 pb-3">
              <div>
                <p className="text-[11px] font-bold text-emerald-700 uppercase">
                  MOBILE CUSTOMER SIGN-OFF · ตรวจรับงานซ่อม
                </p>
                <h3 className="font-bold text-base text-slate-900">
                  ลงลายมือชื่อตรวจรับงานซ่อม ({customerSignOffModalTicket.id})
                </h3>
                <p className="text-xs text-slate-500">
                  บริษัท: {customerSignOffModalTicket.company}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setCustomerSignOffModalTicket(null)}
                className="text-slate-400 hover:text-slate-700 text-xl font-bold cursor-pointer"
              >
                ×
              </button>
            </div>

            <SignaturePadCanvas
              label="ลายเซ็นลูกค้าผู้ตรวจรับงาน (Customer Signature)"
              sublabel="ใช้นิ้วหรือปากกาเซ็นชื่อบนหน้าจอมือถือเพื่อยืนยันการตรวจรับงานซ่อม"
              signerNameLabel="ชื่อ-นามสกุล ลูกค้าผู้ตรวจรับงาน *"
              signerNameValue={
                customerSignOffModalTicket.customerSignerName ??
                (customerSignOffModalTicket.siteContact || customerSignOffModalTicket.contact)
              }
              onSignerNameChange={(val) =>
                setCustomerSignOffModalTicket({
                  ...customerSignOffModalTicket,
                  customerSignerName: val
                })
              }
              initialSignature={customerSignOffModalTicket.customerSignature}
              signedAt={customerSignOffModalTicket.customerSignedAt}
              onSaveSignature={(dataUrl, timestamp) =>
                setCustomerSignOffModalTicket((prev) =>
                  prev
                    ? {
                        ...prev,
                        customerSignature: dataUrl,
                        customerSignedAt: timestamp,
                        customerSignerName:
                          prev.customerSignerName || prev.siteContact || prev.contact
                      }
                    : prev
                )
              }
              onClearSignature={() =>
                setCustomerSignOffModalTicket((prev) =>
                  prev
                    ? {
                        ...prev,
                        customerSignature: '',
                        customerSignedAt: ''
                      }
                    : prev
                )
              }
            />

            <div className="flex flex-wrap justify-end gap-2 pt-2 border-t border-slate-200 text-xs">
              <button
                type="button"
                onClick={() => setCustomerSignOffModalTicket(null)}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-semibold cursor-pointer"
              >
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={() => {
                  if (!customerSignOffModalTicket.customerSignature) {
                    showToast('กรุณาใช้นิ้วเซ็นชื่อในกรอบด้านบนก่อนกดยืนยัน', 'warning');
                    return;
                  }
                  updateTicketSignatureFields(
                    customerSignOffModalTicket.id,
                    {
                      customerSignature: customerSignOffModalTicket.customerSignature,
                      customerSignerName:
                        customerSignOffModalTicket.customerSignerName ||
                        customerSignOffModalTicket.siteContact ||
                        customerSignOffModalTicket.contact,
                      customerSignedAt:
                        customerSignOffModalTicket.customerSignedAt ||
                        new Date().toISOString().replace('T', ' ').substring(0, 16),
                      status: 'CLOSE',
                      isLocked: true
                    },
                    '🎉 [ลูกค้าเซ็นชื่อตรวจรับงานบนมือถือและปิดจบใบงาน (CLOSE) เรียบร้อยแล้ว]'
                  );
                  const signedId = customerSignOffModalTicket.id;
                  setCustomerSignOffModalTicket(null);
                  showToast(
                    `บันทึกลายเซ็นตรวจรับงานใบงาน ${signedId} และปิดจบงานเรียบร้อย!`,
                    'success'
                  );
                }}
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-lg font-bold inline-flex items-center gap-1.5 cursor-pointer"
              >
                <CheckCircle2 className="w-4 h-4" />
                <span>ยืนยันลายเซ็น &amp; ปิดจบใบงาน</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7.6 Dedicated Technician Mobile E-Signature Modal */}
      {technicianSignModalTicket && (
        <div className="fixed inset-0 bg-slate-900/70 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-lg rounded-2xl border border-slate-200 p-5 sm:p-6 space-y-4 shadow-2xl">
            <div className="flex justify-between items-start border-b border-slate-200 pb-3">
              <div>
                <p className="text-[11px] font-bold text-blue-700 uppercase">
                  TECHNICIAN E-SIGNATURE · ลายเซ็นช่างผู้ปฏิบัติงาน
                </p>
                <h3 className="font-bold text-base text-slate-900">
                  ลงลายมือชื่อช่างผู้ปฏิบัติงาน ({technicianSignModalTicket.id})
                </h3>
                <p className="text-xs text-slate-500">
                  บริษัท: {technicianSignModalTicket.company} · ช่าง: {technicianSignModalTicket.assignedTech || 'ช่างเทคนิค'}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setTechnicianSignModalTicket(null)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <SignaturePadCanvas
              label="ลายเซ็นช่างผู้ปฏิบัติงาน (Technician Signature)"
              sublabel="ใช้นิ้วหรือปากกาเซ็นชื่อบนหน้าจอมือถือ/แท็บเล็ตเพื่อรับรองการตรวจซ่อม"
              signerNameLabel="ชื่อ-นามสกุล ช่างผู้ปฏิบัติงาน *"
              signerNameValue={technicianSignModalTicket.assignedTech || ''}
              onSignerNameChange={(val) =>
                setTechnicianSignModalTicket({
                  ...technicianSignModalTicket,
                  assignedTech: val
                })
              }
              initialSignature={technicianSignModalTicket.technicianSignature}
              signedAt={technicianSignModalTicket.technicianSignedAt}
              onSaveSignature={(dataUrl, timestamp) =>
                setTechnicianSignModalTicket((prev) =>
                  prev
                    ? {
                        ...prev,
                        technicianSignature: dataUrl,
                        technicianSignedAt: timestamp
                      }
                    : prev
                )
              }
              onClearSignature={() =>
                setTechnicianSignModalTicket((prev) =>
                  prev
                    ? {
                        ...prev,
                        technicianSignature: '',
                        technicianSignedAt: ''
                      }
                    : prev
                )
              }
            />

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-200 text-xs">
              <button
                type="button"
                onClick={() => setTechnicianSignModalTicket(null)}
                className="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100 font-semibold cursor-pointer"
              >
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={() => {
                  if (!technicianSignModalTicket.technicianSignature) {
                    showToast('กรุณาใช้นิ้วเซ็นชื่อในกรอบด้านบนก่อนกดยืนยัน', 'warning');
                    return;
                  }
                  updateTicketSignatureFields(
                    technicianSignModalTicket.id,
                    {
                      technicianSignature: technicianSignModalTicket.technicianSignature,
                      technicianSignedAt:
                        technicianSignModalTicket.technicianSignedAt ||
                        new Date().toISOString().replace('T', ' ').substring(0, 16)
                    },
                    `✍️ [ช่าง ${technicianSignModalTicket.assignedTech || 'เทคนิค'} ลงลายเซ็นรับรองการปฏิบัติงานแล้ว]`
                  );
                  const signedId = technicianSignModalTicket.id;
                  setTechnicianSignModalTicket(null);
                  showToast(`บันทึกลายเซ็นช่างใบงาน ${signedId} เรียบร้อยแล้ว`, 'success');
                }}
                className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-5 py-2 rounded-xl shadow cursor-pointer inline-flex items-center gap-1.5"
              >
                <Check className="w-4 h-4" />
                <span>บันทึกลายเซ็นช่าง</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Google Workspace Direct Google Sheets Integration Modal */}
      <GoogleSheetsIntegrationModal
        isOpen={isGoogleSheetsModalOpen}
        onClose={() => setIsGoogleSheetsModalOpen(false)}
        tickets={tickets}
        onTicketsImported={handleTicketsImportedFromGoogleSheets}
        connectedSheetId={googleSheetId}
        onSetConnectedSheetId={handleSetConnectedSheetId}
        connectedSheetName={googleSheetName}
        autoSyncEnabled={googleAutoSync}
        onToggleAutoSync={handleToggleGoogleAutoSync}
        lastSyncedAt={googleLastSyncedAt}
        onSetLastSyncedAt={handleSetGoogleLastSyncedAt}
        showToast={showToast}
      />

      {/* 8. Supervisor-Only Google Sheets & LINE Notification Setup Modal */}
      {gasModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-2xl rounded-2xl border border-slate-200 p-6 space-y-4 max-h-[92vh] overflow-y-auto shadow-2xl text-xs sm:text-sm">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <div>
                <h3 className="font-bold text-base text-slate-900">
                  ตั้งค่าฐานข้อมูล Google Sheets (Real-Time Cloud Sync)
                </h3>
                <p className="text-[11px] text-slate-500">
                  ซิงค์ข้อมูลใบแจ้งซ่อม รายชื่อช่าง และประวัติทั้งหมดขึ้น Google Sheets อัตโนมัติแบบเรียลไทม์
                </p>
              </div>
              <button
                type="button"
                onClick={() => setGasModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-lg font-bold cursor-pointer"
              >
                ×
              </button>
            </div>

            {/* Direct Google Sheets Integration Promo Banner */}
            <div className="bg-gradient-to-r from-emerald-600 to-teal-800 text-white p-3.5 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
              <div>
                <p className="font-bold text-xs flex items-center gap-1.5 text-emerald-200">
                  ✨ แนะนำ: เชื่อมต่อ Google Sheets โดยตรงด้วย Google Account
                </p>
                <p className="text-[11px] text-emerald-100 mt-0.5">
                  ไม่ต้องสร้าง Apps Script เอง! กดเพียง 1 คลิกเพื่อสร้างตารางใหม่บน Google Drive หรือเลือกตารางที่มีอยู่
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setGasModalOpen(false);
                  setIsGoogleSheetsModalOpen(true);
                }}
                className="bg-white hover:bg-emerald-50 text-slate-950 font-bold px-3.5 py-1.5 rounded-xl text-xs whitespace-nowrap shadow-sm cursor-pointer transition-colors"
              >
                เปิดการเชื่อมต่อโดยตรง ↗
              </button>
            </div>

            <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl space-y-1 text-xs text-emerald-900">
              <p className="font-bold">📌 วิธีเชื่อมต่อให้ข้อมูลอัปเดตลง Google Sheet อัตโนมัติ 100%:</p>
              <p>1. ใน Google Sheet ไปที่เมนู <b>ส่วนขยาย (Extensions) &gt; Apps Script</b> ลบโค้ดเก่าแล้ววางโค้ดด้านล่างลงไป กดบันทึก 💾</p>
              <p>2. กดปุ่มสีน้ำเงินมุมขวาบน <b>&ldquo;การทำให้ใช้งานได้ (Deploy)&rdquo; &gt; &ldquo;การทำให้ใช้งานได้รายการใหม่ (New deployment)&rdquo;</b> เลือกประเภท <b>เว็บแอป (Web app)</b> และตั้งค่าผู้มีสิทธิ์เข้าถึงเป็น <b>&ldquo;ทุกคน (Anyone)&rdquo;</b></p>
              <p>3. คัดลอก <b>Web App URL (ที่ลงท้ายด้วย /exec)</b> มาวางในช่องขั้นตอนที่ 2 ด้านล่าง แล้วกดบันทึก</p>
            </div>

            <div className="space-y-2">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <label className="font-bold text-slate-800">
                  ขั้นตอนที่ 1: คัดลอกโค้ด Apps Script (รุ่น V12.1 แก้บั๊กข้อมูลไม่อัปเดต &amp; รูปภาพเกินขนาด) ไปวางใน Google Sheet
                </label>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE);
                    showToast('คัดลอกโค้ด Google Apps Script (V12.1) เรียบร้อย!', 'success');
                  }}
                  className="bg-slate-900 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Copy className="w-3.5 h-3.5" />
                  <span>คัดลอกโค้ด Apps Script (V12.1)</span>
                </button>
              </div>
              <textarea
                readOnly
                rows={5}
                value={GOOGLE_APPS_SCRIPT_TEMPLATE}
                className="w-full border border-slate-300 rounded-xl p-2.5 font-mono text-[11px] bg-slate-900 text-emerald-300"
              />
              <div className="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-900 space-y-1">
                <p className="font-bold">📌 สำคัญมาก (สาเหตุที่ข้อมูลไม่อัปเดตลง Google Sheet):</p>
                <p>
                  1. ใน Google Sheet ไปที่ <b>ส่วนขยาย (Extensions) &gt; Apps Script</b> ลบโค้ดเก่าออกทั้งหมด แล้ววางโค้ด V12.1 ด้านบนแทน จากนั้นกด <b>บันทึก (Save 💾)</b>
                </p>
                <p>
                  2. กดปุ่มสีน้ำเงินมมุมขวาบน <b>การทำให้ใช้งานได้ (Deploy) &gt; การทำให้ใช้งานได้รายการใหม่ (New deployment)</b> (ต้องกด New deployment เสมอเมื่อแก้โค้ด)
                </p>
                <p>
                  3. ช่อง <b>ผู้ที่มีสิทธิ์เข้าถึง (Who has access)</b> ต้องเลือกเป็น <b>ทุกคน (Anyone)</b> แล้วกด Deploy จากนั้นคัดลอกลิงก์ <b>Web App URL (ที่ลงท้ายด้วย /exec)</b> มาวางในขั้นตอนที่ 2 ด้านล่าง
                </p>
              </div>
            </div>

            <div className="space-y-2 border-t border-slate-200 pt-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <label className="block font-bold text-slate-800">
                  ขั้นตอนที่ 2: วางลิงก์ Google Apps Script Web App URL (ต้องลงท้ายด้วย /exec)
                </label>
                {sheetSyncStatus && (
                  <span
                    className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full ${
                      sheetSyncStatus.ok
                        ? 'bg-emerald-100 text-emerald-800'
                        : 'bg-rose-100 text-rose-800'
                    }`}
                  >
                    {sheetSyncStatus.ok ? '✅ ' : '⚠️ '}
                    {sheetSyncStatus.text} ({sheetSyncStatus.time})
                  </span>
                )}
              </div>
              <div className="flex gap-2 items-center">
                <input
                  type="url"
                  value={gasWebAppUrl}
                  onChange={(e) => setGasWebAppUrl(e.target.value)}
                  placeholder="https://script.google.com/macros/s/......./exec"
                  className="flex-1 border border-slate-300 rounded-xl p-2 font-mono text-xs"
                />
                {gasWebAppUrl.trim() && (
                  <button
                    type="button"
                    onClick={() => {
                      if (window.confirm('คุณต้องการลบลิงก์ Web App เดิมออก เพื่อใส่ลิงก์ใหม่ใช่หรือไม่?')) {
                        setGasWebAppUrl('');
                        localStorage.removeItem('sahakim_gas_url');
                      }
                    }}
                    title="ลบลิงก์เดิมเพื่อใส่ลิงก์ใหม่"
                    className="bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 px-3 py-2 rounded-xl text-xs font-bold cursor-pointer whitespace-nowrap inline-flex items-center gap-1"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>ล้างลิงก์เดิม</span>
                  </button>
                )}
              </div>
              {gasWebAppUrl.trim() && gasWebAppUrl.includes('docs.google.com/spreadsheets') && (
                <div className="bg-rose-50 border border-rose-300 text-rose-800 rounded-xl p-2.5 text-xs font-semibold">
                  ❌ ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet (docs.google.com/spreadsheets) ซึ่งไม่สามารถรับข้อมูลอัตโนมัติได้! กรุณาใช้ลิงก์จากเมนู Deploy &gt; New deployment ใน Apps Script ที่ขึ้นต้นด้วย <code className="font-mono bg-white px-1 rounded">https://script.google.com/macros/s/.../exec</code>
                </div>
              )}
              <div className="flex flex-wrap gap-2 pt-1">
                <button
                  type="button"
                  disabled={isSyncingSheet}
                  onClick={forceSyncToGoogleSheetNow}
                  className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-lg text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  <span>
                    {isSyncingSheet
                      ? 'กำลังส่งข้อมูลขึ้น Google Sheet...'
                      : `อัปเดตข้อมูลปัจจุบันขึ้น Google Sheet ทันที (${tickets.length} ใบงาน)`}
                  </span>
                </button>
                <button
                  type="button"
                  disabled={isSyncingSheet}
                  onClick={() => loadTicketsFromGoogleSheet(false)}
                  className="bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-white px-3.5 py-2 rounded-lg text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>ดึงข้อมูลจาก Google Sheet ลงเว็บ</span>
                </button>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-200">
              <button
                type="button"
                onClick={() => setGasModalOpen(false)}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-semibold cursor-pointer"
              >
                ปิดหน้าต่าง
              </button>
              <button
                type="button"
                disabled={isSyncingSheet}
                onClick={async () => {
                  if (gasWebAppUrl.trim()) {
                    await forceSyncToGoogleSheetNow();
                  } else {
                    showToast('บันทึกการตั้งค่าแล้ว (ยังไม่มีการเชื่อมต่อ Google Sheet)', 'warning');
                    setGasModalOpen(false);
                  }
                }}
                className="bg-orange-500 hover:bg-orange-600 disabled:opacity-50 text-white px-5 py-2 rounded-xl font-semibold cursor-pointer"
              >
                บันทึก &amp; ซิงค์ข้อมูลขึ้น Google Sheet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 9. User Manual Modal (คู่มือการใช้งาน Web App After Sales Service) */}
      {userManualOpen && (
        <div className="print-modal-root fixed inset-0 bg-slate-900/70 z-50 flex items-center justify-center p-3 sm:p-5 overflow-y-auto">
          <div
            id="user-manual-print-area"
            className="print-modal-content bg-white w-full max-w-4xl rounded-2xl border border-slate-200 p-5 sm:p-8 space-y-6 max-h-[92vh] overflow-y-auto shadow-2xl text-xs sm:text-sm"
          >
            <div className="flex flex-wrap justify-between items-start gap-4 border-b-2 border-orange-500 pb-4">
              <div>
                <p className="text-xs font-bold text-orange-600 tracking-wider">
                  SAHA KIM MOTOR CO., LTD. · ENTERPRISE SERVICE SYSTEM V12.1
                </p>
                <h2 className="text-lg sm:text-2xl font-bold text-slate-900 mt-0.5">
                  คู่มือการใช้งานระบบบริการหลังการขาย (After-Sales Service Web App Manual)
                </h2>
                <p className="text-xs text-slate-500 mt-1">
                  สำหรับลูกค้า, เจ้าหน้าที่ประสานงาน (Service Support), ช่างเทคนิค (Technician) และหัวหน้างาน (Supervisor)
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2 no-print">
                <button
                  type="button"
                  disabled={isGeneratingPdf}
                  onClick={() =>
                    downloadElementAsPdf(
                      'user-manual-print-area',
                      'Sahakim-After-Sales-Service-Manual.pdf'
                    )
                  }
                  className="bg-orange-500 hover:bg-orange-600 disabled:opacity-60 text-white px-4 py-2 rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  {isGeneratingPdf ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Download className="w-3.5 h-3.5" />
                  )}
                  <span>บันทึกคู่มือเป็น PDF (Save PDF)</span>
                </button>
                <button
                  type="button"
                  onClick={() => window.print()}
                  className="bg-slate-800 hover:bg-slate-900 text-white px-3.5 py-2 rounded-xl text-xs font-bold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Printer className="w-3.5 h-3.5" />
                  <span>สั่งพิมพ์</span>
                </button>
                <button
                  type="button"
                  onClick={() => setUserManualOpen(false)}
                  className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-3 py-2 rounded-xl text-xs font-bold cursor-pointer"
                >
                  ปิดหน้าต่าง ✕
                </button>
              </div>
            </div>

            {/* ภาพรวม Flow การทำงาน */}
            <div className="bg-slate-900 text-white p-4 sm:p-5 rounded-xl space-y-2">
              <p className="font-bold text-orange-400 text-sm">
                📌 ลำดับขั้นตอนมาตรฐานของใบงาน (Workflow Status ตั้งแต่เปิดงานจนปิดจบ)
              </p>
              <p className="text-xs sm:text-sm leading-relaxed text-slate-200">
                <b>1. NEW</b> (ลูกค้าแจ้งซ่อมใหม่) ➔ <b>2. ASSIGNED</b> (Service Support นัดวันและมอบหมายช่าง) ➔{' '}
                <b>3. IN_PROGRESS</b> (ช่างเข้าตรวจเช็คหน้างาน) ➔ <b>4. PENDING_PARTS_APPROVAL</b> (ช่างขออนุมัติเบิกอะไหล่) ➔{' '}
                <b>5. REPAIRING</b> (Supervisor อนุมัติอะไหล่แล้ว / กำลังซ่อม) ➔ <b>6. PENDING_SUPERVISOR_APPROVAL</b> (ช่างส่งรายงานซ่อมให้ตรวจ) ➔{' '}
                <b>7. CUSTOMER_ACCEPTANCE</b> (Supervisor อนุมัติและล็อกใบงาน / รอลูกค้าตรวจรับ) ➔ <b>8. CLOSE</b> (ลูกค้ากดตรวจรับงาน / ปิดงานสมบูรณ์)
              </p>
              <p className="text-[11px] text-amber-300 pt-1">
                🔗 ลิงก์ใช้งานจริงสำหรับลูกค้าและทีมงาน: https://2020skmproduct-service.github.io/sahakim-service-portal/
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* ส่วนที่ 1: ลูกค้า */}
              <div className="border border-orange-200 bg-orange-50/50 rounded-xl p-4 space-y-2.5">
                <h3 className="font-bold text-base text-orange-800 flex items-center gap-2">
                  <span className="bg-orange-500 text-white w-6 h-6 rounded-full inline-flex items-center justify-center text-xs">
                    1
                  </span>
                  <span>สำหรับลูกค้า (Customer Portal)</span>
                </h3>
                <ul className="space-y-1.5 text-slate-700 list-disc pl-5">
                  <li>
                    <b>การเข้าใช้งาน:</b> สแกน QR Code ที่ติดอยู่หน้าเครื่องจักร/ตู้คอนโทรล หรือคลิกลิงก์เว็บแอป แล้วเลือกเมนู{' '}
                    <b>&ldquo;1. ฝั่งลูกค้า (Customer Portal)&rdquo;</b>
                  </li>
                  <li>
                    <b>การเลือกบริษัท:</b> เลือกชื่อบริษัทจากรายการ หรือพิมพ์ชื่อบริษัทของท่านเพื่อดูเฉพาะใบงานของบริษัทตนเอง
                  </li>
                  <li>
                    <b>การแจ้งซ่อมใหม่:</b> กดแท็บ <b>&ldquo;+ แจ้งซ่อมใหม่&rdquo;</b> กรอกชื่อผู้ติดต่อ เบอร์โทร สถานที่หน้างาน เงื่อนไขประกัน (In-Warranty / Out-of-Warranty) และวันที่ต้องการให้นัดเข้าซ่อม
                  </li>
                  <li>
                    <b>การเพิ่มอุปกรณ์หลายรายการ:</b> เลือกประเภทอุปกรณ์ (Motor, Pump, Blower, Gen Set, Sewage, Control ฯลฯ หรือเลือก ETC. เพื่อพิมพ์เอง) ระบุอาการเสีย และแนบรูปถ่ายหน้างาน จากนั้นกด <b>&ldquo;ยืนยันการส่งใบแจ้งซ่อม&rdquo;</b> (ระบบจะเด้งแจ้งเตือนเข้า LINE บอทของบริษัททันที)
                  </li>
                  <li>
                    <b>การติดตามสถานะ &amp; ตรวจรับงาน:</b> ดูสถานะงานแบบเรียลไทม์ กดปุ่ม <b>&ldquo;ดูใบรายงาน&rdquo;</b> เพื่อดูผลตรวจเช็คมาตรฐาน และเมื่อสถานะเป็น <b>CUSTOMER_ACCEPTANCE</b> ให้กดปุ่ม <b>&ldquo;ตรวจรับงาน&rdquo;</b> เพื่อปิดใบงาน (CLOSE)
                  </li>
                </ul>
              </div>

              {/* ส่วนที่ 2: Service Support */}
              <div className="border border-blue-200 bg-blue-50/50 rounded-xl p-4 space-y-2.5">
                <h3 className="font-bold text-base text-blue-900 flex items-center gap-2">
                  <span className="bg-blue-600 text-white w-6 h-6 rounded-full inline-flex items-center justify-center text-xs">
                    2
                  </span>
                  <span>สำหรับเจ้าหน้าที่รับเรื่อง (Service Team)</span>
                </h3>
                <p className="text-xs font-bold text-blue-800 bg-blue-100 px-2.5 py-1 rounded-lg w-fit">
                  🔑 รหัสผ่านเข้าใช้งานเริ่มต้น (PIN): 1234 (Supervisor เปลี่ยนรหัสได้)
                </p>
                <ul className="space-y-1.5 text-slate-700 list-disc pl-5">
                  <li>
                    <b>การเข้าสู่ระบบ:</b> เลือกเมนู <b>&ldquo;2. เจ้าหน้าที่ภายใน&rdquo;</b> ➔ เลือกบทบาท <b>SERVICE TEAM</b> ➔ กรอกรหัส PIN
                  </li>
                  <li>
                    <b>การรับงานใหม่ (NEW) &amp; จ่ายงานให้ช่าง:</b> เมื่อมีใบงานสถานะ <b>NEW</b> ให้กดปุ่มสีน้ำเงิน{' '}
                    <b>&ldquo;นัดวัน / จ่ายงานช่าง&rdquo;</b>
                  </li>
                  <li>
                    <b>การระบุช่างและหัวหน้าช่าง:</b> เลือกวันที่ยืนยันนัดหมาย เลือกชื่อ <b>ช่างผู้ปฏิบัติงาน</b> และ <b>หัวหน้าช่างผู้ควบคุมงาน</b> จากนั้นกดยืนยัน สถานะจะเปลี่ยนเป็น <b>ASSIGNED</b> พร้อมส่งแจ้งเตือนเข้า LINE ทันที
                  </li>
                  <li>
                    <b>การติดตามงานค้างเกินกำหนด (Overdue):</b> หากมีงานเลยกำหนดนัดหมาย แถบสีแดงด้านบนจะแสดงจำนวนงานค้าง สามารถกดปุ่ม{' '}
                    <b>&ldquo;แจ้งเตือนงานค้างเข้า LINE ด่วน&rdquo;</b> เพื่อเร่งรัดงานได้ทันที
                  </li>
                </ul>
              </div>

              {/* ส่วนที่ 3: ช่างเทคนิค */}
              <div className="border border-amber-200 bg-amber-50/50 rounded-xl p-4 space-y-2.5">
                <h3 className="font-bold text-base text-amber-900 flex items-center gap-2">
                  <span className="bg-amber-600 text-white w-6 h-6 rounded-full inline-flex items-center justify-center text-xs">
                    3
                  </span>
                  <span>สำหรับช่างเทคนิค (Technician)</span>
                </h3>
                <p className="text-xs font-bold text-amber-900 bg-amber-100 px-2.5 py-1 rounded-lg w-fit">
                  🔑 รหัส PIN ช่างเริ่มต้น: 1111, 2222, 3333 (ตามรายชื่อช่างในระบบ)
                </p>
                <ul className="space-y-1.5 text-slate-700 list-disc pl-5">
                  <li>
                    <b>การเปิดบันทึกงานซ่อม:</b> เข้าสู่ระบบในฐานะ <b>TECHNICIAN</b> ➔ กดปุ่ม <b>&ldquo;จัดการ / บันทึกผล&rdquo;</b> ในใบงานที่ได้รับมอบหมาย
                  </li>
                  <li>
                    <b>การใช้ AI วิเคราะห์อาการเสีย &amp; คลังความรู้:</b> กดปุ่มสีม่วง <b>&ldquo;วิเคราะห์ด้วย AI (Pump Handbook / EASA)&rdquo;</b> หรือเลือกจากคลังต้นแบบ เพื่อให้ระบบช่วยร่างผลวิเคราะห์ทางเทคนิค สาเหตุ มาตรฐานอ้างอิง และแนวทางแก้ไข (ช่างสามารถพิมพ์แก้ไขเพิ่มเติมได้ตามหน้างานจริง)
                  </li>
                  <li>
                    <b>การบันทึก Checklist &amp; รูปถ่าย Before/After:</b> บันทึกผลตรวจเช็คมาตรฐาน (Before = NG/OK, After = OK) และอัปโหลดรูปถ่ายก่อนซ่อม-หลังซ่อม
                  </li>
                  <li>
                    <b>การขออนุมัติเบิกอะไหล่:</b> กรอกชื่ออะไหล่ จำนวน และราคา แล้วกด <b>&ldquo;เบิกอะไหล่&rdquo;</b> สถานะจะเปลี่ยนเป็น <b>PENDING_PARTS_APPROVAL</b> และแจ้งเตือน Supervisor ทาง LINE ทันที
                  </li>
                  <li>
                    <b>การส่งรายงานให้ Supervisor อนุมัติ:</b> เมื่อซ่อมเสร็จสมบูรณ์ กดปุ่มสีม่วง <b>&ldquo;ส่ง Supervisor อนุมัติรายงาน&rdquo;</b>
                  </li>
                </ul>
              </div>

              {/* ส่วนที่ 4: Supervisor */}
              <div className="border border-emerald-200 bg-emerald-50/50 rounded-xl p-4 space-y-2.5">
                <h3 className="font-bold text-base text-emerald-900 flex items-center gap-2">
                  <span className="bg-emerald-600 text-white w-6 h-6 rounded-full inline-flex items-center justify-center text-xs">
                    4
                  </span>
                  <span>สำหรับหัวหน้างาน / ผู้ดูแลระบบ (Supervisor)</span>
                </h3>
                <p className="text-xs font-bold text-emerald-900 bg-emerald-100 px-2.5 py-1 rounded-lg w-fit">
                  🔑 รหัสผ่าน Supervisor (Admin) เริ่มต้น: 9999
                </p>
                <ul className="space-y-1.5 text-slate-700 list-disc pl-5">
                  <li>
                    <b>การอนุมัติอะไหล่ &amp; อนุมัติใบรายงาน:</b> กดปุ่มสีเขียว <b>&ldquo;อนุมัติอะไหล่&rdquo;</b> หรือ <b>&ldquo;อนุมัติรายงาน&rdquo;</b> เลือกชื่อผู้อนุมัติ (Supervisor หรือหัวหน้าช่าง) เมื่ออนุมัติรายงานแล้ว ใบงานจะถูกล็อก (Locked) ป้องกันการแก้ไขโดยพลการ
                  </li>
                  <li>
                    <b>การส่งกลับแก้ไข (Reject) &amp; ปลดล็อกใบงาน:</b> หากข้อมูลไม่ครบ กดปุ่มสีแดง <b>&ldquo;ส่งกลับ&rdquo;</b> พร้อมระบุเหตุผล หรือกด <b>&ldquo;ปลดล็อก&rdquo;</b> เพื่อให้ช่างแก้ไขข้อมูล
                  </li>
                  <li>
                    <b>การจัดการช่าง/ผู้อนุมัติ &amp; รหัสผ่าน:</b> กดปุ่ม <b>&ldquo;จัดการช่าง/รหัส&rdquo;</b> เพื่อเพิ่ม-ลบรายชื่อช่าง กำหนดตำแหน่ง (ช่างเทคนิค / หัวหน้าช่าง) และตั้งรหัสผ่านใหม่
                  </li>
                  <li>
                    <b>การอัปเดตข้อมูลและส่งออก:</b> กด <b>&ldquo;อัปเดตข้อมูลเข้า LINE&rdquo;</b> เพื่อส่งสรุปทุกใบงานเข้ากลุ่ม LINE, กด <b>&ldquo;Excel (CSV)&rdquo;</b> เพื่อดึงข้อมูลลงตาราง หรือกด <b>&ldquo;ดาวน์โหลด HTML&rdquo;</b> เพื่ออัปเดตไฟล์ <code>index.html</code> ขึ้น GitHub Pages
                  </li>
                </ul>
              </div>
            </div>

            {/* ส่วนที่ 5: การตั้งค่าให้ LINE Bot ตอบกลับและแจ้งเตือนทันที (ไม่หน่วง) */}
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 space-y-2">
              <h4 className="font-bold text-amber-950 text-sm">
                ⚡ เคล็ดลับสำคัญ: ตั้งค่า LINE Official Account (@389fyhbl - Sahakim.service bot) ให้ตอบกลับและแจ้งเตือนทันทีภายใน 0.5 วินาที
              </h4>
              <ol className="list-decimal pl-5 space-y-1 text-slate-700 text-xs">
                <li>
                  เข้าไปที่แอปหรือเว็บ <b>LINE Official Account Manager (manager.line.biz)</b> เลือกบัญชี <b>Sahakim.service bot (@389fyhbl)</b>
                </li>
                <li>
                  ไปที่เมนู <b>ตั้งค่า (Settings) &gt; การตอบกลับ (Response settings)</b>:
                  <ul className="list-disc pl-5 mt-1 space-y-0.5 text-amber-950 font-medium">
                    <li>
                      โหมดการตอบกลับ (Response mode): เลือกเป็น <b>&ldquo;บอท (Bot)&rdquo;</b>
                    </li>
                    <li>
                      ข้อความตอบกลับอัตโนมัติ (Auto-response): เลือกเป็น <b>&ldquo;ปิด (Disabled)&rdquo;</b> (เพื่อไม่ให้ข้อความขออภัยของ LINE ไปขวางหรือหน่วงการตอบของระบบ)
                    </li>
                    <li>
                      Webhook: เลือกเป็น <b>&ldquo;เปิด (Enabled)&rdquo;</b>
                    </li>
                  </ul>
                </li>
                <li>
                  ในหน้าเว็บแอป เมื่อต้องการส่งแจ้งเตือนใบงานเข้ากลุ่ม LINE ให้กดปุ่มสีเขียว <b>&ldquo;ส่ง LINE&rdquo;</b> ท้ายใบงาน ระบบจะยิงข้อความแจ้งเตือนจากบอทเข้ากลุ่มและแชทโดยตรงทันที (ไม่ต้องเลือกห้องแชทเอง)
                </li>
              </ol>
            </div>
          </div>
        </div>
      )}

      {/* =====================================================================
          MASTER DATA MANAGEMENT MODAL (แก้ไขข้อมูลตั้งต้น: ยี่ห้อ, ประเภทเครื่องจักร, Checklist, โลโก้/เบอร์โทร, รายชื่อช่าง)
      ===================================================================== */}
      {masterDataModalOpen && (
        <div className="fixed inset-0 bg-slate-900/75 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-4xl rounded-2xl border-2 border-slate-300 p-5 sm:p-6 space-y-5 max-h-[92vh] overflow-y-auto shadow-2xl">
            {/* Modal Header */}
            <div className="flex flex-wrap justify-between items-start gap-3 border-b border-slate-200 pb-3.5">
              <div>
                <span className="bg-emerald-100 text-emerald-800 text-[11px] font-bold px-2.5 py-0.5 rounded-full">
                  MASTER DATA &amp; SYSTEM CONFIGURATION
                </span>
                <h3 className="font-bold text-base sm:text-lg text-slate-900 mt-1">
                  ⚙️ ตั้งค่าและแก้ไขข้อมูลตั้งต้นของระบบ (เพิ่ม / แก้ไข / ลบ ได้เองทันที)
                </h3>
                <p className="text-xs text-slate-500">
                  ปรับแก้ ยี่ห้อเครื่องจักร, ประเภทเครื่องจักร, รายการมาตรฐานใน Checklist, โลโก้/เบอร์โทรบริษัท และรายชื่อช่าง โดยระบบจะบันทึกอัตโนมัติทันที
                </p>
              </div>
              <button
                type="button"
                onClick={() => setMasterDataModalOpen(false)}
                className="bg-slate-900 hover:bg-slate-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer"
              >
                ✕ ปิดหน้าต่าง
              </button>
            </div>

            {/* 5 Category Tabs */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
              <button
                type="button"
                onClick={() => setMasterDataTab('brands')}
                className={`p-2.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                  masterDataTab === 'brands'
                    ? 'bg-orange-500 text-white border-orange-600 shadow-xs'
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                🏷️ 1. ยี่ห้อเครื่องจักร ({masterBrands.length})
              </button>
              <button
                type="button"
                onClick={() => setMasterDataTab('equipment')}
                className={`p-2.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                  masterDataTab === 'equipment'
                    ? 'bg-blue-600 text-white border-blue-700 shadow-xs'
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                ⚙️ 2. ประเภทเครื่องจักร ({masterEquipmentTypes.length})
              </button>
              <button
                type="button"
                onClick={() => setMasterDataTab('checklists')}
                className={`p-2.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                  masterDataTab === 'checklists'
                    ? 'bg-emerald-600 text-white border-emerald-700 shadow-xs'
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                📋 3. มาตรฐาน Checklist ({Object.keys(masterChecklists).length} หมวด)
              </button>
              <button
                type="button"
                onClick={() => setMasterDataTab('company')}
                className={`p-2.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                  masterDataTab === 'company'
                    ? 'bg-indigo-600 text-white border-indigo-700 shadow-xs'
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                🏢 4. โลโก้ / เบอร์โทรบริษัท
              </button>
              <button
                type="button"
                onClick={() => setMasterDataTab('staff')}
                className={`col-span-2 sm:col-span-1 p-2.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                  masterDataTab === 'staff'
                    ? 'bg-slate-900 text-white border-slate-900 shadow-xs'
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                👷 5. รายชื่อช่าง ({technicians.length})
              </button>
            </div>

            {/* TAB 1: BRANDS (ยี่ห้อเครื่องจักร) */}
            {masterDataTab === 'brands' && (
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-900">
                      🏷️ จัดการยี่ห้อเครื่องจักรตั้งต้น (DEFAULT_BRANDS)
                    </h4>
                    <p className="text-slate-500">
                      เพิ่มยี่ห้อใหม่ พิมพ์แก้ไขชื่อยี่ห้อเดิมในช่องได้โดยตรง หรือลบยี่ห้อที่ไม่ต้องการ
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const defaults = DEFAULT_BRANDS.filter((x) => !x.startsWith('ETC.'));
                      setMasterBrands(defaults);
                      showToast('รีเซ็ตรายชื่อยี่ห้อเครื่องจักรกลับเป็นค่าเริ่มต้นแล้ว', 'warning');
                    }}
                    className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-semibold cursor-pointer"
                  >
                    ↺ รีเซ็ตเป็นค่าเริ่มต้น
                  </button>
                </div>

                {/* Add new brand */}
                <div className="flex flex-col sm:flex-row gap-2 bg-white p-3 rounded-xl border border-orange-200">
                  <input
                    type="text"
                    value={newMasterBrandInput}
                    onChange={(e) => setNewMasterBrandInput(e.target.value)}
                    placeholder="พิมพ์ชื่อยี่ห้อเครื่องจักรที่ต้องการเพิ่ม (เช่น KSB, TSURUMI, MITSUBISHI, WEG)..."
                    className="flex-1 border border-slate-300 rounded-lg px-3 py-2 font-semibold text-slate-900"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const cleaned = newMasterBrandInput.trim().toUpperCase();
                      if (!cleaned) {
                        showToast('กรุณากรอกชื่อยี่ห้อเครื่องจักร', 'error');
                        return;
                      }
                      if (masterBrands.includes(cleaned)) {
                        showToast(`มียี่ห้อ "${cleaned}" อยู่ในระบบแล้ว`, 'warning');
                        return;
                      }
                      setMasterBrands((prev) => [...prev, cleaned]);
                      setNewMasterBrandInput('');
                      showToast(`เพิ่มยี่ห้อ "${cleaned}" เข้าระบบเรียบร้อย!`, 'success');
                    }}
                    className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-lg font-bold shrink-0 cursor-pointer"
                  >
                    + เพิ่มยี่ห้อเครื่องจักร
                  </button>
                </div>

                {/* Editable list of brands */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 max-h-80 overflow-y-auto pr-1">
                  {masterBrands.map((brand, bIdx) => (
                    <div
                      key={bIdx}
                      className="flex items-center gap-1.5 bg-white p-2 rounded-lg border border-slate-200"
                    >
                      <span className="text-[11px] font-mono text-slate-400 w-6 text-center shrink-0">
                        {bIdx + 1}.
                      </span>
                      <input
                        type="text"
                        value={brand}
                        onChange={(e) => {
                          const val = e.target.value.toUpperCase();
                          setMasterBrands((prev) =>
                            prev.map((item, i) => (i === bIdx ? val : item))
                          );
                        }}
                        className="flex-1 border border-slate-200 focus:border-orange-500 rounded px-2 py-1 font-bold text-slate-900 min-w-0"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          setMasterBrands((prev) => prev.filter((_, i) => i !== bIdx));
                          showToast(`ลบยี่ห้อ "${brand}" แล้ว`, 'warning');
                        }}
                        className="text-red-500 hover:text-red-700 px-2 py-1 rounded hover:bg-red-50 font-bold shrink-0 cursor-pointer"
                        title="ลบยี่ห้อนี้"
                      >
                        ลบ
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* TAB 2: EQUIPMENT TYPES (ประเภทเครื่องจักร) */}
            {masterDataTab === 'equipment' && (
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-900">
                      ⚙️ จัดการประเภทเครื่องจักรตั้งต้น (DEFAULT_EQUIPMENT_MASTERS)
                    </h4>
                    <p className="text-slate-500">
                      เพิ่มประเภทเครื่องจักรใหม่ พิมพ์แก้ไขชื่อประเภทเดิมได้ทันที หรือลบรายการที่ไม่ใช้
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const defaults = DEFAULT_EQUIPMENT_MASTERS.filter((x) => !x.startsWith('ETC.'));
                      setMasterEquipmentTypes(defaults);
                      showToast('รีเซ็ตประเภทเครื่องจักรกลับเป็นค่าเริ่มต้นแล้ว', 'warning');
                    }}
                    className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-semibold cursor-pointer"
                  >
                    ↺ รีเซ็ตเป็นค่าเริ่มต้น
                  </button>
                </div>

                {/* Add new equipment type */}
                <div className="flex flex-col sm:flex-row gap-2 bg-white p-3 rounded-xl border border-blue-200">
                  <input
                    type="text"
                    value={newMasterEquipInput}
                    onChange={(e) => setNewMasterEquipInput(e.target.value)}
                    placeholder="พิมพ์ประเภทเครื่องจักรที่ต้องการเพิ่ม (เช่น SPLIT CASE PUMP, SUBMERSIBLE PUMP, VFD PANEL)..."
                    className="flex-1 border border-slate-300 rounded-lg px-3 py-2 font-semibold text-slate-900"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const cleaned = newMasterEquipInput.trim().toUpperCase();
                      if (!cleaned) {
                        showToast('กรุณากรอกชื่อประเภทเครื่องจักร', 'error');
                        return;
                      }
                      if (masterEquipmentTypes.includes(cleaned)) {
                        showToast(`มีประเภทเครื่องจักร "${cleaned}" อยู่ในระบบแล้ว`, 'warning');
                        return;
                      }
                      setMasterEquipmentTypes((prev) => [...prev, cleaned]);
                      setNewMasterEquipInput('');
                      showToast(`เพิ่มประเภทเครื่องจักร "${cleaned}" เรียบร้อย!`, 'success');
                    }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg font-bold shrink-0 cursor-pointer"
                  >
                    + เพิ่มประเภทเครื่องจักร
                  </button>
                </div>

                {/* Editable list of equipment types */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 max-h-80 overflow-y-auto pr-1">
                  {masterEquipmentTypes.map((eqType, eIdx) => (
                    <div
                      key={eIdx}
                      className="flex items-center gap-1.5 bg-white p-2 rounded-lg border border-slate-200"
                    >
                      <span className="text-[11px] font-mono text-slate-400 w-6 text-center shrink-0">
                        {eIdx + 1}.
                      </span>
                      <input
                        type="text"
                        value={eqType}
                        onChange={(e) => {
                          const val = e.target.value.toUpperCase();
                          setMasterEquipmentTypes((prev) =>
                            prev.map((item, i) => (i === eIdx ? val : item))
                          );
                        }}
                        className="flex-1 border border-slate-200 focus:border-blue-500 rounded px-2 py-1 font-bold text-slate-900 min-w-0"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          setMasterEquipmentTypes((prev) => prev.filter((_, i) => i !== eIdx));
                          showToast(`ลบประเภทเครื่องจักร "${eqType}" แล้ว`, 'warning');
                        }}
                        className="text-red-500 hover:text-red-700 px-2 py-1 rounded hover:bg-red-50 font-bold shrink-0 cursor-pointer"
                      >
                        ลบ
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* TAB 3: CHECKLIST TEMPLATES (แก้ไขรายการมาตรฐานใน Checklist) */}
            {masterDataTab === 'checklists' && (
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-900">
                      📋 จัดการรายการมาตรฐานใน Checklist แยกตามประเภทเครื่องจักร
                    </h4>
                    <p className="text-slate-500">
                      เลือกหมวดเครื่องจักรเพื่อเพิ่ม/แก้ไข/ลบหัวข้อตรวจเช็คมาตรฐาน (ใช้เป็นค่าตั้งต้นเวลาเปิดใบงานใหม่หรือกดโหลดมาตรฐาน)
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setMasterChecklists({ ...CHECKLIST_TEMPLATES });
                      showToast('รีเซ็ตตาราง Checklist ทั้งหมดกลับเป็นค่ามาตรฐานแล้ว', 'warning');
                    }}
                    className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-semibold cursor-pointer"
                  >
                    ↺ รีเซ็ต Checklist เป็นค่าเริ่มต้น
                  </button>
                </div>

                {/* Category selector + Add new category */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 bg-white p-3 rounded-xl border border-emerald-200">
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      เลือกหมวดเครื่องจักรที่ต้องการแก้ไข Checklist:
                    </label>
                    <select
                      value={selectedChecklistCategory}
                      onChange={(e) => setSelectedChecklistCategory(e.target.value)}
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 font-bold text-emerald-900 bg-emerald-50/50"
                    >
                      {Object.keys(masterChecklists).map((cat) => (
                        <option key={cat} value={cat}>
                          หมวด: {cat} ({(masterChecklists[cat] || []).length} หัวข้อ)
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      หรือสร้างหมวด Checklist ใหม่ (ตรงตามชื่อประเภทเครื่องจักร):
                    </label>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        value={newChecklistCategoryInput}
                        onChange={(e) => setNewChecklistCategoryInput(e.target.value)}
                        placeholder="เช่น BOOSTER PUMP, CHILLER..."
                        className="flex-1 border border-slate-300 rounded-lg px-3 py-2 font-semibold"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          const cat = newChecklistCategoryInput.trim().toUpperCase();
                          if (!cat) return;
                          if (!masterChecklists[cat]) {
                            setMasterChecklists((prev) => ({
                              ...prev,
                              [cat]: [...(prev.PUMP || CHECKLIST_TEMPLATES.PUMP)]
                            }));
                          }
                          setSelectedChecklistCategory(cat);
                          setNewChecklistCategoryInput('');
                          showToast(`เปิดหมวด Checklist "${cat}" เรียบร้อย`, 'success');
                        }}
                        className="bg-slate-900 hover:bg-slate-800 text-white px-3.5 py-2 rounded-lg font-bold shrink-0 cursor-pointer"
                      >
                        + เพิ่มหมวด
                      </button>
                    </div>
                  </div>
                </div>

                {/* Add new checklist item to selected category */}
                <div className="flex flex-col sm:flex-row gap-2 bg-white p-3 rounded-xl border border-slate-200">
                  <input
                    type="text"
                    value={newChecklistItemInput}
                    onChange={(e) => setNewChecklistItemInput(e.target.value)}
                    placeholder={`พิมพ์หัวข้อตรวจเช็คใหม่สำหรับหมวด ${selectedChecklistCategory}...`}
                    className="flex-1 border border-slate-300 rounded-lg px-3 py-2 font-semibold text-slate-900"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const text = newChecklistItemInput.trim();
                      if (!text) {
                        showToast('กรุณาพิมพ์หัวข้อตรวจเช็คที่ต้องการเพิ่ม', 'error');
                        return;
                      }
                      setMasterChecklists((prev) => ({
                        ...prev,
                        [selectedChecklistCategory]: [
                          ...(prev[selectedChecklistCategory] || []),
                          text
                        ]
                      }));
                      setNewChecklistItemInput('');
                      showToast(`เพิ่มหัวข้อตรวจเช็คในหมวด ${selectedChecklistCategory} แล้ว`, 'success');
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-bold shrink-0 cursor-pointer"
                  >
                    + เพิ่มหัวข้อ Checklist
                  </button>
                </div>

                {/* Editable checklist rows */}
                <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                  {(masterChecklists[selectedChecklistCategory] || []).map((itemText, cIdx) => (
                    <div
                      key={cIdx}
                      className="flex items-center gap-2 bg-white p-2.5 rounded-lg border border-slate-200"
                    >
                      <span className="font-mono font-bold text-emerald-700 w-7 text-center shrink-0">
                        #{cIdx + 1}
                      </span>
                      <input
                        type="text"
                        value={itemText}
                        onChange={(e) => {
                          const val = e.target.value;
                          setMasterChecklists((prev) => {
                            const list = [...(prev[selectedChecklistCategory] || [])];
                            list[cIdx] = val;
                            return { ...prev, [selectedChecklistCategory]: list };
                          });
                        }}
                        className="flex-1 border border-slate-200 focus:border-emerald-500 rounded px-2.5 py-1.5 font-medium text-slate-900"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          setMasterChecklists((prev) => {
                            const list = (prev[selectedChecklistCategory] || []).filter(
                              (_, i) => i !== cIdx
                            );
                            return { ...prev, [selectedChecklistCategory]: list };
                          });
                          showToast('ลบหัวข้อตรวจเช็คแล้ว', 'warning');
                        }}
                        className="text-red-500 hover:text-red-700 px-2.5 py-1 rounded hover:bg-red-50 font-bold shrink-0 cursor-pointer"
                      >
                        ลบ
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* TAB 4: COMPANY PROFILE (เปลี่ยนโลโก้ / ชื่อบริษัท / เบอร์โทร / รูปเจ้าหน้าที่ภายใน) */}
            {masterDataTab === 'company' && (
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-900">
                      🏢 แก้ไขข้อมูลบริษัท โลโก้ เบอร์โทรศัพท์ และรูปประจำตัวเจ้าหน้าที่ภายใน
                    </h4>
                    <p className="text-slate-500">
                      ข้อมูลที่แก้ไขจะไปแสดงที่แถบหัวเว็บ (Header) และหัวกระดาษใบรายงาน (Service Report) ทันที
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setCompanyProfile({
                        nameTh: 'บริษัท สหคิมมอเตอร์ จำกัด',
                        nameEn: 'SAHA KIM MOTOR CO., LTD.',
                        phone: '02-XXX-XXXX',
                        tagline:
                          'บริการซ่อม บำรุงรักษา และจำหน่ายปั๊มน้ำ มอเตอร์ Blower, Gen Set, Sewage, Aerator และระบบควบคุมอุตสาหกรรม',
                        logoUrl: '',
                        staffMascotUrl: SKM_SERVICE_MASCOT_LOGO
                      });
                      showToast('รีเซ็ตข้อมูลบริษัทและโลโก้กลับเป็นค่าเริ่มต้นแล้ว', 'warning');
                    }}
                    className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-semibold cursor-pointer"
                  >
                    ↺ รีเซ็ตข้อมูลบริษัท
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-white p-4 rounded-xl border border-slate-200">
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      ชื่อบริษัท (ภาษาไทย) *
                    </label>
                    <input
                      type="text"
                      value={companyProfile.nameTh}
                      onChange={(e) =>
                        setCompanyProfile((prev) => ({ ...prev, nameTh: e.target.value }))
                      }
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 font-bold text-slate-900"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      ชื่อบริษัท (ภาษาอังกฤษ) *
                    </label>
                    <input
                      type="text"
                      value={companyProfile.nameEn}
                      onChange={(e) =>
                        setCompanyProfile((prev) => ({ ...prev, nameEn: e.target.value }))
                      }
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 font-bold text-orange-600"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      เบอร์โทรศัพท์ / ช่องทางติดต่อ (แสดงบนหัวใบ Service Report)
                    </label>
                    <input
                      type="text"
                      value={companyProfile.phone}
                      onChange={(e) =>
                        setCompanyProfile((prev) => ({ ...prev, phone: e.target.value }))
                      }
                      placeholder="เช่น 02-123-4567, 081-234-5678"
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 font-semibold text-slate-900"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      รายละเอียดบริการ / ที่อยู่บริษัท (ใต้ชื่อบริษัทในใบรายงาน)
                    </label>
                    <input
                      type="text"
                      value={companyProfile.tagline}
                      onChange={(e) =>
                        setCompanyProfile((prev) => ({ ...prev, tagline: e.target.value }))
                      }
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 text-slate-800"
                    />
                  </div>
                </div>

                {/* Logo Upload & Internal Staff Mascot Upload */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {/* Company Logo */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-2.5">
                    <p className="font-bold text-slate-900">
                      1. โลโก้บริษัท (แสดงมุมซ้ายบนเว็บ &amp; หัวใบ Service Report)
                    </p>
                    <div className="flex items-center gap-3">
                      {companyProfile.logoUrl ? (
                        <img
                          src={companyProfile.logoUrl}
                          alt="Company Logo"
                          className="w-16 h-16 rounded-xl object-contain border border-slate-300 bg-slate-50 p-1 shrink-0"
                        />
                      ) : (
                        <div className="w-16 h-16 rounded-xl border border-dashed border-slate-300 bg-slate-50 flex items-center justify-center text-[10px] text-slate-400 text-center p-1 shrink-0">
                          ยังไม่ใส่รูปโลโก้
                        </div>
                      )}
                      <div className="space-y-1.5 flex-1">
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                              const reader = new FileReader();
                              reader.onload = () => {
                                if (typeof reader.result === 'string') {
                                  setCompanyProfile((prev) => ({
                                    ...prev,
                                    logoUrl: reader.result as string
                                  }));
                                  showToast('อัปโหลดโลโก้บริษัทเรียบร้อย!', 'success');
                                }
                              };
                              reader.readAsDataURL(e.target.files[0]);
                            }
                          }}
                          className="w-full text-[11px] border border-slate-300 rounded-lg p-1 bg-slate-50 cursor-pointer"
                        />
                        <div className="flex gap-1.5">
                          <input
                            type="text"
                            value={companyProfile.logoUrl}
                            onChange={(e) =>
                              setCompanyProfile((prev) => ({ ...prev, logoUrl: e.target.value }))
                            }
                            placeholder="หรือวางลิงก์ URL รูปโลโก้..."
                            className="flex-1 border border-slate-300 rounded-lg px-2 py-1 text-[11px]"
                          />
                          {companyProfile.logoUrl && (
                            <button
                              type="button"
                              onClick={() =>
                                setCompanyProfile((prev) => ({ ...prev, logoUrl: '' }))
                              }
                              className="bg-red-100 hover:bg-red-200 text-red-700 px-2.5 py-1 rounded-lg font-bold text-[11px] cursor-pointer"
                            >
                              นำออก
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Internal Staff Mascot Image */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-2.5">
                    <p className="font-bold text-slate-900">
                      2. รูปเจ้าหน้าที่ภายใน (การ์ดที่ 2 หน้าหลัก &amp; หน้าล็อกอิน)
                    </p>
                    <div className="flex items-center gap-3">
                      <img
                        src={companyProfile.staffMascotUrl || SKM_SERVICE_MASCOT_LOGO}
                        alt="Staff Mascot"
                        className="w-16 h-16 rounded-xl object-cover border-2 border-orange-500 bg-slate-900 shrink-0"
                      />
                      <div className="space-y-1.5 flex-1">
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                              const reader = new FileReader();
                              reader.onload = () => {
                                if (typeof reader.result === 'string') {
                                  setCompanyProfile((prev) => ({
                                    ...prev,
                                    staffMascotUrl: reader.result as string
                                  }));
                                  showToast('เปลี่ยนรูปเจ้าหน้าที่ภายในเรียบร้อย!', 'success');
                                }
                              };
                              reader.readAsDataURL(e.target.files[0]);
                            }
                          }}
                          className="w-full text-[11px] border border-slate-300 rounded-lg p-1 bg-slate-50 cursor-pointer"
                        />
                        <div className="flex gap-1.5">
                          <input
                            type="text"
                            value={companyProfile.staffMascotUrl}
                            onChange={(e) =>
                              setCompanyProfile((prev) => ({
                                ...prev,
                                staffMascotUrl: e.target.value
                              }))
                            }
                            placeholder="หรือวางลิงก์ URL รูปเจ้าหน้าที่..."
                            className="flex-1 border border-slate-300 rounded-lg px-2 py-1 text-[11px]"
                          />
                          <button
                            type="button"
                            onClick={() =>
                              setCompanyProfile((prev) => ({
                                ...prev,
                                staffMascotUrl: SKM_SERVICE_MASCOT_LOGO
                              }))
                            }
                            className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-2.5 py-1 rounded-lg font-bold text-[11px] shrink-0 cursor-pointer"
                          >
                            ค่าเริ่มต้น
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* TAB 5: STAFF & TECHNICIANS (เพิ่ม/แก้ไขรายชื่อช่าง, ตำแหน่ง, รหัส PIN, ผู้อนุมัติ) */}
            {masterDataTab === 'staff' && (
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-900">
                      👷 จัดการรายชื่อช่างเทคนิค, หัวหน้าช่าง, ผู้อนุมัติ และรหัสผ่าน (PIN)
                    </h4>
                    <p className="text-slate-500">
                      เพิ่มรายชื่อช่างใหม่ แก้ไขชื่อช่าง ตำแหน่ง ความเชี่ยวชาญ หรือเปลี่ยนรหัส PIN ได้ทันที
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setMasterDataModalOpen(false);
                      setTechManagerOpen(true);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg font-bold cursor-pointer"
                  >
                    🔑 ตั้งค่ารหัสผ่าน Admin / Supervisor เพิ่มเติม
                  </button>
                </div>

                {/* Add new technician */}
                <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-2.5">
                  <p className="font-bold text-slate-900">+ เพิ่มรายชื่อช่างเทคนิค / หัวหน้าช่างใหม่:</p>
                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-2">
                    <input
                      type="text"
                      value={newTechName}
                      onChange={(e) => setNewTechName(e.target.value)}
                      placeholder="ชื่อ-นามสกุลช่าง *"
                      className="border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold"
                    />
                    <select
                      value={newTechPosition}
                      onChange={(e) => setNewTechPosition(e.target.value)}
                      className="border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white font-semibold text-indigo-700"
                    >
                      {TECH_POSITIONS.map((pos) => (
                        <option key={pos} value={pos}>
                          {pos}
                        </option>
                      ))}
                    </select>
                    <input
                      type="text"
                      value={newTechSpec}
                      onChange={(e) => setNewTechSpec(e.target.value)}
                      placeholder="ความเชี่ยวชาญ (เช่น PUMP & MOTOR)"
                      className="border border-slate-300 rounded-lg px-2.5 py-1.5"
                    />
                    <div className="flex gap-1.5">
                      <input
                        type="text"
                        value={newTechPin}
                        onChange={(e) => setNewTechPin(e.target.value)}
                        placeholder="PIN (เช่น 4444)"
                        className="w-24 border border-slate-300 rounded-lg px-2 py-1.5 font-mono font-bold text-center"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          if (!newTechName.trim()) {
                            showToast('กรุณากรอกชื่อ-นามสกุลช่าง', 'error');
                            return;
                          }
                          let maxNum = 0;
                          technicians.forEach((t) => {
                            const match = t.id.match(/\d+/);
                            if (match) {
                              const num = parseInt(match[0], 10);
                              if (num > maxNum) maxNum = num;
                            }
                          });
                          const nextNum = maxNum + 1;
                          const nextId = `TECH-${nextNum < 10 ? '0' + nextNum : nextNum}`;
                          const newTechItem: Technician = {
                            id: nextId,
                            name: newTechName.trim(),
                            position: newTechPosition,
                            specialization: newTechSpec.trim() || 'PUMP & MOTOR SPECIALIST',
                            status: 'Available',
                            pin: newTechPin.trim() || '1111'
                          };
                          const updated = [...technicians, newTechItem];
                          setNewTechName('');
                          setNewTechSpec('');
                          setNewTechPin('');
                          handleSaveTechniciansList(
                            updated,
                            `เพิ่มรายชื่อช่าง "${newTechItem.name}" เรียบร้อย!`
                          );
                        }}
                        className="bg-orange-500 hover:bg-orange-600 text-white px-3 py-1.5 rounded-lg font-bold shrink-0 cursor-pointer"
                      >
                        + เพิ่มช่าง
                      </button>
                    </div>
                  </div>
                </div>

                {/* Editable Technician Rows */}
                <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
                  {technicians.map((tc, idx) => (
                    <div
                      key={tc.id}
                      className="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center bg-white p-2.5 rounded-xl border border-slate-200"
                    >
                      <div className="sm:col-span-4 flex items-center gap-1.5">
                        <span className="font-mono text-[11px] text-slate-400 shrink-0">
                          {tc.id}
                        </span>
                        <input
                          type="text"
                          value={tc.name}
                          onChange={(e) => {
                            const val = e.target.value;
                            setTechnicians((prev) =>
                              prev.map((item, i) => (i === idx ? { ...item, name: val } : item))
                            );
                          }}
                          className="w-full border border-slate-200 rounded px-2 py-1 font-bold text-slate-900"
                        />
                      </div>
                      <div className="sm:col-span-3">
                        <select
                          value={tc.position || 'ช่างเทคนิค (Technician)'}
                          onChange={(e) => {
                            const val = e.target.value;
                            setTechnicians((prev) =>
                              prev.map((item, i) => (i === idx ? { ...item, position: val } : item))
                            );
                          }}
                          className="w-full border border-slate-200 rounded px-2 py-1 font-semibold text-indigo-700 bg-slate-50"
                        >
                          {TECH_POSITIONS.map((pos) => (
                            <option key={pos} value={pos}>
                              {pos}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="sm:col-span-3">
                        <input
                          type="text"
                          value={tc.specialization}
                          onChange={(e) => {
                            const val = e.target.value;
                            setTechnicians((prev) =>
                              prev.map((item, i) =>
                                i === idx ? { ...item, specialization: val } : item
                              )
                            );
                          }}
                          className="w-full border border-slate-200 rounded px-2 py-1 text-slate-700"
                        />
                      </div>
                      <div className="sm:col-span-2 flex items-center justify-end gap-1.5">
                        <input
                          type="text"
                          value={tc.pin}
                          onChange={(e) => {
                            const val = e.target.value;
                            setTechnicians((prev) =>
                              prev.map((item, i) => (i === idx ? { ...item, pin: val } : item))
                            );
                          }}
                          className="w-16 border border-slate-300 rounded px-1.5 py-1 font-mono font-bold text-center text-orange-600"
                          title="รหัส PIN ช่าง"
                        />
                        <button
                          type="button"
                          onClick={() => {
                            if (
                              window.confirm(
                                `ยืนยันการลบรายชื่อช่าง "${tc.name}" (${tc.id}) ออกจากระบบหรือไม่?`
                              )
                            ) {
                              const updated = technicians.filter((_, i) => i !== idx);
                              handleSaveTechniciansList(
                                updated,
                                `ลบรายชื่อช่าง ${tc.name} เรียบร้อยแล้ว`
                              );
                            }
                          }}
                          className="text-red-500 hover:text-red-700 px-2 py-1 rounded hover:bg-red-50 font-bold cursor-pointer"
                        >
                          ลบ
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Modal Footer */}
            <div className="flex flex-wrap items-center justify-between gap-2 pt-3 border-t border-slate-200 text-xs">
              <span className="text-emerald-700 font-semibold">
                ✓ ข้อมูลทุกหมวดบันทึกอัตโนมัติและซิงค์ออนไลน์ให้ทุกเครื่องทันที
              </span>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => setResetConfirmModalOpen(true)}
                  className="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-300 px-3.5 py-2 rounded-xl font-bold cursor-pointer inline-flex items-center gap-1.5"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>ล้างใบงานตัวอย่าง & ตารางช่างว่าง 🟢</span>
                </button>
                <button
                  type="button"
                  onClick={handleOpenGithubExportModal}
                  className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2 rounded-xl font-bold cursor-pointer"
                >
                  ⬇️ ดาวน์โหลด index.html (สำหรับ GitHub)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    localStorage.setItem('sahakim_technicians', JSON.stringify(technicians));
                    localStorage.setItem('sahakim_supervisors', JSON.stringify(supervisors));
                    localStorage.setItem('sahakim_master_brands_v1', JSON.stringify(masterBrands));
                    localStorage.setItem('sahakim_master_equipment_types_v1', JSON.stringify(masterEquipmentTypes));
                    localStorage.setItem('sahakim_master_checklists_v1', JSON.stringify(masterChecklists));
                    localStorage.setItem('sahakim_company_profile_v1', JSON.stringify(companyProfile));
                    pushOnlineSharedState(tickets, 'MASTER_DATA_SAVED', technicians, supervisors);
                    setMasterDataModalOpen(false);
                    showToast('บันทึกและอัปเดตข้อมูลตั้งต้นออนไลน์ไปยังทุกเครื่องเรียบร้อยแล้ว!', 'success');
                  }}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-xl font-bold cursor-pointer"
                >
                  ✓ บันทึกและอัปเดตออนไลน์
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* =====================================================================
          RESET TO CLEAN PRODUCTION MODAL (ล้างใบงานตัวอย่าง & ตารางช่างว่างเขียว 🟢)
      ===================================================================== */}
      {resetConfirmModalOpen && (
        <div className="fixed inset-0 bg-slate-900/80 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-lg rounded-2xl border-2 border-orange-500 p-5 sm:p-6 space-y-4 max-h-[92vh] overflow-y-auto shadow-2xl text-xs sm:text-sm">
            <div className="flex items-start justify-between gap-3 border-b border-slate-200 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 rounded-xl bg-orange-100 text-orange-600 flex items-center justify-center font-bold shrink-0">
                  <RotateCcw className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base sm:text-lg text-slate-900">
                    รีเซ็ตระบบเพื่อเริ่มต้นใช้งานจริง
                  </h3>
                  <p className="text-xs text-slate-500">
                    ล้างใบงานตัวอย่างออกทั้งหมด เพื่อให้เริ่มกรอกข้อมูลงานจริงทันที
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setResetConfirmModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold p-1 cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="bg-amber-50 border border-amber-300 rounded-xl p-3.5 space-y-2 text-slate-800 text-xs">
              <p className="font-bold text-amber-900 flex items-center gap-1.5">
                <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                <span>คำแนะนำสำหรับการเริ่มใช้งานจริง:</span>
              </p>
              <ul className="list-disc pl-5 space-y-1 text-slate-700">
                <li>
                  เมื่อกด <b>&ldquo;ล้างใบงานตัวอย่าง & เริ่มใช้งานจริง&rdquo;</b> ใบงานตัวอย่างจะถูกลบออกทั้งหมด (ใบงาน = 0 รายการ)
                </li>
                <li>
                  <b>ตารางงานช่างจะกลับมาว่างทั้งหมด (สีเขียว 🟢 ทุกวัน)</b> พร้อมรับงานใหม่ทันที
                </li>
                <li>
                  เมื่อคุณเปิดใบแจ้งซ่อมจริงและกำหนดวันรับงาน ตารางจะอัปเดตสถานะคิวงานตามจริงทันที
                </li>
                <li>
                  ข้อมูลช่าง, หัวหน้าช่าง, ประเภทเครื่องจักร, และแบรนด์สินค้าจะยังคงอยู่ครบถ้วน ไม่สูญหาย
                </li>
              </ul>
            </div>

            <div className="space-y-2 pt-1">
              <button
                type="button"
                onClick={handleResetToProductionClean}
                className="w-full bg-orange-600 hover:bg-orange-700 text-white font-bold py-3 px-4 rounded-xl shadow-md flex items-center justify-center gap-2 transition cursor-pointer"
              >
                <RotateCcw className="w-4 h-4" />
                <span>ยืนยัน: ล้างใบงานตัวอย่าง & ตารางช่างว่างเขียว 🟢 เริ่มต้นทำงานจริง</span>
              </button>

              <button
                type="button"
                onClick={handleRestoreDemoTickets}
                className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl flex items-center justify-center gap-2 transition cursor-pointer"
              >
                <span>🔄 กู้คืนข้อมูลออนไลน์ล่าสุด (รวมใบงาน SKM-69-003 บิ๊กซี วงศ์สว่าง)</span>
              </button>

              <button
                type="button"
                onClick={() => setResetConfirmModalOpen(false)}
                className="w-full text-slate-500 hover:text-slate-700 py-2 text-xs font-semibold text-center cursor-pointer"
              >
                ยกเลิก (ปิดหน้าต่าง)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* =====================================================================
          GITHUB INDEX.HTML DOWNLOAD & 1-CLICK COPY ASSISTANT MODAL
      ===================================================================== */}
      {githubExportModalOpen && (
        <div className="fixed inset-0 bg-slate-900/80 z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto no-print">
          <div className="bg-white w-full max-w-3xl rounded-2xl border-2 border-orange-500 p-5 sm:p-6 space-y-5 max-h-[92vh] overflow-y-auto shadow-2xl text-xs sm:text-sm">
            {/* Header */}
            <div className="flex flex-wrap items-start justify-between gap-3 border-b border-slate-200 pb-3.5">
              <div>
                <span className="bg-emerald-100 text-emerald-800 text-[11px] font-extrabold px-2.5 py-0.5 rounded-full">
                  ✅ ระบบสั่งดาวน์โหลดไฟล์ index.html ให้แล้ว!
                </span>
                <h3 className="text-base sm:text-xl font-extrabold text-slate-900 mt-1">
                  📦 หาไฟล์ที่ดาวน์โหลดไม่เจอ? หรือต้องการอัปเดต GitHub แบบง่ายที่สุด (ไม่ต้องหาไฟล์)
                </h3>
                <p className="text-xs text-slate-600 mt-0.5">
                  เลือกใช้วิธีที่สะดวกที่สุดด้านล่างได้เลยค่ะ (แนะนำ <b>วิธีที่ 1: กดคัดลอกโค้ดไปวางใน GitHub โดยตรง</b> ไม่ต้องเปิดหาไฟล์ในเครื่องเลย!)
                </p>
              </div>
              <button
                type="button"
                onClick={() => setGithubExportModalOpen(false)}
                className="bg-slate-900 hover:bg-slate-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold cursor-pointer"
              >
                ✕ ปิดหน้าต่าง
              </button>
            </div>

            {/* METHOD 1: 1-Click Copy Code (No file hunting needed!) */}
            <div className="bg-emerald-50 border-2 border-emerald-500 rounded-2xl p-4 sm:p-5 space-y-3.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <span className="bg-emerald-600 text-white text-[11px] font-bold px-2.5 py-0.5 rounded-full">
                    ⭐ แนะนำ! วิธีที่ 1 (ง่ายที่สุด — ไม่ต้องหาไฟล์ที่ดาวน์โหลดในคอมพิวเตอร์)
                  </span>
                  <h4 className="font-extrabold text-sm sm:text-base text-emerald-950 mt-1">
                    กดปุ่ม &ldquo;คัดลอกโค้ด index.html ทั้งหมด&rdquo; แล้วนำไปวาง (Paste) ใน GitHub ได้เลย
                  </h4>
                </div>
              </div>

              <button
                type="button"
                onClick={handleCopyStandaloneHtmlToClipboard}
                className={`w-full py-3.5 px-4 rounded-xl font-extrabold text-sm sm:text-base flex items-center justify-center gap-2 shadow-md transition-all cursor-pointer ${
                  copiedHtmlSuccess
                    ? 'bg-emerald-700 text-white ring-4 ring-emerald-300'
                    : 'bg-emerald-600 hover:bg-emerald-700 text-white'
                }`}
              >
                <Copy className="w-5 h-5" />
                <span>
                  {copiedHtmlSuccess
                    ? '✅ คัดลอกโค้ด index.html สำเร็จแล้ว! (นำไปกด Ctrl + V วางใน GitHub ได้เลย)'
                    : '📋 คลิกที่นี่เพื่อคัดลอกโค้ด index.html ทั้งหมด (ไม่ต้องหาไฟล์)'}
                </span>
              </button>

              <div className="bg-white rounded-xl p-3.5 border border-emerald-200 space-y-1.5 text-xs text-slate-800">
                <p className="font-bold text-emerald-900">
                  ขั้นตอนการนำโค้ดที่คัดลอกไปอัปเดตใน GitHub (ใช้เวลา 30 วินาที):
                </p>
                <ol className="list-decimal pl-5 space-y-1">
                  <li>
                    กดปุ่มสีเขียว <b>&ldquo;📋 คลิกที่นี่เพื่อคัดลอกโค้ด index.html ทั้งหมด&rdquo;</b> ด้านบน
                  </li>
                  <li>
                    เปิดหน้า GitHub ของคุณ (<b>2020skmproduct-service/sahakim-service-portal</b>) ➔ คลิกที่ชื่อไฟล์{' '}
                    <code className="bg-slate-100 px-1.5 py-0.5 rounded font-mono font-bold text-orange-600">
                      index.html
                    </code>
                  </li>
                  <li>
                    คลิกปุ่ม <b>รูปดินสอ ✏️ (Edit this file)</b> ที่อยู่มุมขวาบนของกรอบโค้ด
                  </li>
                  <li>
                    คลิกในพื้นที่โค้ด แล้วกดปุ่ม <b><kbd className="bg-slate-200 px-1.5 py-0.5 rounded">Ctrl</kbd> + <kbd className="bg-slate-200 px-1.5 py-0.5 rounded">A</kbd></b> (เลือกโค้ดเก่าทั้งหมด) ➔ กด <b><kbd className="bg-slate-200 px-1.5 py-0.5 rounded">Delete</kbd></b> ลบของเก่าออก ➔ กด <b><kbd className="bg-slate-200 px-1.5 py-0.5 rounded">Ctrl</kbd> + <kbd className="bg-slate-200 px-1.5 py-0.5 rounded">V</kbd></b> (วางโค้ดใหม่ลงไป)
                  </li>
                  <li>
                    กดปุ่มสีเขียว <b>&ldquo;Commit changes...&rdquo;</b> มุมขวาบน ยืนยันอีก 1 ครั้ง ➔ รอ 1 นาที เว็บจริงจะอัปเดตทันทีค่ะ!
                  </li>
                </ol>
              </div>
            </div>

            {/* METHOD 2: Where is the downloaded file on the computer + Save As */}
            <div className="bg-orange-50/80 border border-orange-300 rounded-2xl p-4 sm:p-5 space-y-3">
              <h4 className="font-extrabold text-sm sm:text-base text-orange-950">
                📁 วิธีที่ 2: หากต้องการใช้ไฟล์ <code className="font-mono bg-white px-1.5 py-0.5 rounded">index.html</code> ที่ดาวน์โหลดลงเครื่อง (ไฟล์อยู่ที่ไหน?)
              </h4>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                <div className="bg-white p-3.5 rounded-xl border border-orange-200 space-y-1.5">
                  <p className="font-bold text-orange-800">
                    🔍 1. วิธีเปิดดูไฟล์ที่เพิ่งดาวน์โหลดทันที (เร็วที่สุด):
                  </p>
                  <p className="text-slate-700 leading-relaxed">
                    บนคีย์บอร์ดคอมพิวเตอร์ ให้กดปุ่ม{' '}
                    <b className="bg-slate-900 text-white px-2 py-0.5 rounded font-mono">Ctrl + J</b>{' '}
                    พร้อมกัน (สำหรับ Chrome หรือ Edge) หน้าต่างประวัติการดาวน์โหลดจะเด้งขึ้นมา จะเห็นไฟล์{' '}
                    <b>index.html</b> อยู่บนสุด ➔ กดคำว่า <b>&ldquo;แสดงในโฟลเดอร์&rdquo; (Show in folder / รูปแฟ้ม 📁)</b> ได้เลยค่ะ
                  </p>
                </div>

                <div className="bg-white p-3.5 rounded-xl border border-orange-200 space-y-1.5">
                  <p className="font-bold text-orange-800">
                    💻 2. โฟลเดอร์ที่เก็บไฟล์ในเครื่องคอมพิวเตอร์:
                  </p>
                  <p className="text-slate-700 leading-relaxed">
                    เปิด <b>This PC / File Explorer (รูปแฟ้มสีเหลือง)</b> ➔ คลิกเมนู{' '}
                    <b className="text-blue-700">Downloads (ดาวน์โหลด)</b> ทางซ้ายมือ จะพบไฟล์ชื่อ{' '}
                    <b>index.html</b> อยู่ในนั้นค่ะ{' '}
                    <span className="text-rose-600 font-semibold">
                      (ระวัง: หากกดดาวน์โหลดซ้ำหลายครั้ง ชื่ออาจมีวงเล็บ เช่น index (1).html ตอนอัปโหลดให้เปลี่ยนชื่อกลับเป็น index.html)
                    </span>
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2.5 pt-1">
                <button
                  type="button"
                  onClick={handleSaveHtmlWithFilePicker}
                  className="bg-orange-500 hover:bg-orange-600 text-white px-4 py-2.5 rounded-xl font-bold text-xs inline-flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <Download className="w-4 h-4" />
                  <span>💾 เลือกโฟลเดอร์เก็บไฟล์เอง (เช่น เลือกวางไว้หน้า Desktop)</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    downloadStandaloneHtml(
                      tickets,
                      technicians,
                      knowledgeTemplates,
                      masterEquipmentTypes,
                      masterBrands,
                      supPassword,
                      servicePin,
                      gasWebAppUrl,
                      lineChannelToken,
                      lineTargetId,
                      supervisors
                    );
                    showToast('ดาวน์โหลดไฟล์ index.html ซ้ำอีกครั้งแล้ว! (กด Ctrl + J เพื่อดูไฟล์)', 'success');
                  }}
                  className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2.5 rounded-xl font-bold text-xs inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>⬇️ กดดาวน์โหลดไฟล์ index.html ซ้ำอีกครั้ง</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

