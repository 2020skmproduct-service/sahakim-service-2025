import { ServiceTicket } from '../types';

export interface SpreadsheetFileInfo {
  id: string;
  name: string;
  modifiedTime?: string;
  webViewLink?: string;
}

export const SHEET_HEADERS = [
  'ID',
  'เลขที่เอกสาร (Doc No)',
  'ชื่อบริษัท / ลูกค้า (Company)',
  'เลขประจำตัวผู้เสียภาษี (Tax ID)',
  'ผู้ติดต่อ (Contact)',
  'สถานที่หน้างาน (Site Location)',
  'ประเภทงานบริการ (Service Type)',
  'ระดับความเร่งด่วน (Urgency)',
  'สถานะรับประกัน (Warranty)',
  'ประเภทอุปกรณ์ (Equipment Type)',
  'ยี่ห้อ (Brand)',
  'รุ่น (Model)',
  'อาการเสียที่แจ้ง (Symptom)',
  'ช่างผู้รับผิดชอบ (Assigned Tech)',
  'หัวหน้าช่าง (Supervisor)',
  'สถานะใบงาน (Status)',
  'จำนวนรายการอะไหล่ (Parts Count)',
  'มูลค่าอะไหล่รวม THB (Total Amount)',
  'การเซ็นรับมอบงาน (Customer Signed)',
  'วันที่นัดหมาย (Desired Date)',
  'วันที่สร้างใบงาน (Created Date)'
];

export function extractSpreadsheetId(input: string): string {
  if (!input) return '';
  const trimmed = input.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return trimmed;
}

export function ticketToRowValues(t: ServiceTicket): (string | number)[] {
  const primaryEquip = t.equipmentList && t.equipmentList.length > 0 ? t.equipmentList[0] : null;
  const allParts = t.equipmentList
    ? t.equipmentList.flatMap((eq) => (Array.isArray(eq.parts) ? eq.parts : []))
    : [];
  const totalAmount = allParts.reduce((sum, p) => sum + (Number(p.total) || (Number(p.qty) || 0) * (Number(p.price) || 0)), 0);

  const equipTypes = t.equipmentList?.map((e) => e.type).filter(Boolean).join(', ') || primaryEquip?.type || '-';
  const brands = t.equipmentList?.map((e) => e.brand).filter(Boolean).join(', ') || primaryEquip?.brand || '-';
  const models = t.equipmentList?.map((e) => e.model).filter(Boolean).join(', ') || primaryEquip?.model || '-';
  const symptoms = t.equipmentList?.map((e) => e.symptom).filter(Boolean).join('; ') || primaryEquip?.symptom || '-';

  return [
    t.id || '',
    t.docNo || t.id || '',
    t.company || '',
    t.taxId || '',
    t.contact || '',
    t.location || '',
    t.serviceType || 'ซ่อมบำรุง / Overhaul',
    t.urgency || 'ปกติ',
    t.warranty || 'IN_WARRANTY',
    equipTypes,
    brands,
    models,
    symptoms,
    t.assignedTech || 'ยังไม่ระบุ',
    t.approvedBySupervisor || '-',
    t.status || 'NEW',
    allParts.length,
    totalAmount,
    t.customerSignature ? `เซ็นแล้ว (${t.customerSignerName || 'ลูกค้า'})` : 'ยังไม่เซ็น',
    t.desiredDate || t.confirmedDate || '-',
    t.createdAt || new Date().toISOString()
  ];
}

/**
 * Creates a brand-new Google Spreadsheet in the user's Google Drive
 * with pre-formatted Tickets and Summary tabs.
 */
export async function createSahakimSpreadsheet(
  accessToken: string,
  tickets: ServiceTicket[]
): Promise<{ spreadsheetId: string; spreadsheetUrl: string; title: string }> {
  const title = `SAHA KIM MOTOR - Service & Repair Database (${new Date().toLocaleDateString('th-TH')})`;

  const createBody = {
    properties: {
      title
    },
    sheets: [
      {
        properties: {
          title: 'Tickets',
          gridProperties: {
            frozenRowCount: 1,
            columnCount: 26
          }
        }
      },
      {
        properties: {
          title: 'Summary_Dashboard',
          gridProperties: {
            frozenRowCount: 1,
            columnCount: 10
          }
        }
      }
    ]
  };

  const createRes = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(createBody)
  });

  if (!createRes.ok) {
    const errData = await createRes.json().catch(() => null);
    throw new Error(errData?.error?.message || `Failed to create spreadsheet: HTTP ${createRes.status}`);
  }

  const createdData = await createRes.json();
  const spreadsheetId = createdData.spreadsheetId;
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
  const ticketsSheetId = createdData.sheets?.[0]?.properties?.sheetId ?? 0;

  // Format header row and write values
  await syncTicketsToSpreadsheet(accessToken, spreadsheetId, tickets, ticketsSheetId);

  return {
    spreadsheetId,
    spreadsheetUrl,
    title
  };
}

/**
 * Syncs (exports/overwrites) ticket data to the Google Spreadsheet
 */
export async function syncTicketsToSpreadsheet(
  accessToken: string,
  spreadsheetId: string,
  tickets: ServiceTicket[],
  ticketsSheetId?: number
): Promise<{ success: boolean; count: number }> {
  const cleanId = extractSpreadsheetId(spreadsheetId);
  if (!cleanId) {
    throw new Error('Spreadsheet ID ไม่ถูกต้อง');
  }

  // 1. Fetch metadata to ensure Tickets sheet exists and find its sheetId
  const metaRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${cleanId}?fields=sheets.properties`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!metaRes.ok) {
    const errData = await metaRes.json().catch(() => null);
    throw new Error(errData?.error?.message || `ไม่สามารถเข้าถึง Google Sheet ได้ (HTTP ${metaRes.status})`);
  }

  const metaData = await metaRes.json();
  const sheets: any[] = metaData.sheets || [];
  let targetSheet = sheets.find((s) => s.properties?.title === 'Tickets');
  let currentSheetId = ticketsSheetId ?? targetSheet?.properties?.sheetId ?? 0;

  // If "Tickets" sheet doesn't exist, create it
  if (!targetSheet) {
    const addSheetRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${cleanId}:batchUpdate`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [
          {
            addSheet: {
              properties: {
                title: 'Tickets',
                gridProperties: { frozenRowCount: 1, columnCount: 26 }
              }
            }
          }
        ]
      })
    });
    if (addSheetRes.ok) {
      const addedData = await addSheetRes.json();
      currentSheetId = addedData.replies?.[0]?.addSheet?.properties?.sheetId ?? 0;
    }
  }

  // 2. Prepare row values
  const rows = [
    SHEET_HEADERS,
    ...tickets.map((t) => ticketToRowValues(t))
  ];

  // 3. Non-destructive update: write current rows without clearing the sheet.
  // Existing rows beyond the new dataset are intentionally left untouched;
  // cleanup must be a separate, explicitly authorized operation.

  // 4. Write data to Tickets sheet
  const updateRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${cleanId}/values/Tickets!A1?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        range: 'Tickets!A1',
        majorDimension: 'ROWS',
        values: rows
      })
    }
  );

  if (!updateRes.ok) {
    const errData = await updateRes.json().catch(() => null);
    throw new Error(errData?.error?.message || `เกิดข้อผิดพลาดในการบันทึกข้อมูลลงตาราง (HTTP ${updateRes.status})`);
  }

  // 5. Apply professional styling to header row (Navy #0f172a, white text, bold)
  if (currentSheetId !== undefined) {
    try {
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${cleanId}:batchUpdate`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          requests: [
            {
              repeatCell: {
                range: {
                  sheetId: currentSheetId,
                  startRowIndex: 0,
                  endRowIndex: 1,
                  startColumnIndex: 0,
                  endColumnIndex: SHEET_HEADERS.length
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: { red: 15 / 255, green: 23 / 255, blue: 42 / 255 },
                    textFormat: {
                      foregroundColor: { red: 1, green: 1, blue: 1 },
                      bold: true,
                      fontSize: 11
                    },
                    horizontalAlignment: 'CENTER'
                  }
                },
                fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
              }
            },
            {
              autoResizeDimensions: {
                dimensions: {
                  sheetId: currentSheetId,
                  dimension: 'COLUMNS',
                  startIndex: 0,
                  endIndex: SHEET_HEADERS.length
                }
              }
            }
          ]
        })
      });
    } catch {
      // Non-fatal styling error
    }
  }

  // 6. Write Summary_Dashboard tab
  try {
    const totalTickets = tickets.length;
    const closedCount = tickets.filter((t) => t.status === 'CLOSE').length;
    const inProgressCount = tickets.filter(
      (t) => ['INSPECTION', 'REPAIRING', 'PENDING_PARTS_APPROVAL', 'PENDING_SUPERVISOR_APPROVAL'].includes(t.status)
    ).length;
    const newCount = tickets.filter((t) => t.status === 'NEW' || t.status === 'ASSIGNED').length;

    let totalRevenue = 0;
    tickets.forEach((t) => {
      t.equipmentList?.forEach((eq) => {
        eq.parts?.forEach((p) => {
          totalRevenue += Number(p.total) || (Number(p.qty) || 0) * (Number(p.price) || 0);
        });
      });
    });

    const summaryRows = [
      ['หัวข้อสรุปภาพรวม (Metric)', 'ค่าสถิติ (Value)', 'หน่วย (Unit)'],
      ['จำนวนใบงานทั้งหมด (Total Tickets)', totalTickets, 'ใบงาน'],
      ['ใบงานใหม่ / รอจ่ายงาน (New / Assigned)', newCount, 'ใบงาน'],
      ['กำลังดำเนินงาน / ซ่อมแซม (In Progress)', inProgressCount, 'ใบงาน'],
      ['ปิดงานสำเร็จเรียบร้อย (Closed)', closedCount, 'ใบงาน'],
      ['มูลค่าอะไหล่และงานซ่อมรวม (Total Cost/Revenue)', totalRevenue, 'บาท (THB)'],
      ['อัปเดตล่าสุดเมื่อ (Last Synced At)', new Date().toLocaleString('th-TH'), '']
    ];

    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${cleanId}/values/Summary_Dashboard!A1?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          range: 'Summary_Dashboard!A1',
          majorDimension: 'ROWS',
          values: summaryRows
        })
      }
    );
  } catch {
    // Non-fatal summary error
  }

  return { success: true, count: tickets.length };
}

/**
 * Loads and reconstructs tickets from the Google Spreadsheet
 */
export async function loadTicketsFromSpreadsheet(
  accessToken: string,
  spreadsheetId: string
): Promise<ServiceTicket[]> {
  const cleanId = extractSpreadsheetId(spreadsheetId);
  if (!cleanId) {
    throw new Error('Spreadsheet ID ไม่ถูกต้อง');
  }

  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${cleanId}/values/Tickets!A2:U5000`,
    {
      headers: { Authorization: `Bearer ${accessToken}` }
    }
  );

  if (!res.ok) {
    const errData = await res.json().catch(() => null);
    throw new Error(errData?.error?.message || `ไม่สามารถดึงข้อมูลจากตารางได้ (HTTP ${res.status})`);
  }

  const data = await res.json();
  const rows: any[][] = data.values || [];

  const parsedTickets: ServiceTicket[] = rows
    .filter((row) => row && row[0])
    .map((row, idx) => {
      const id = String(row[0] || `SKM-${Date.now()}-${idx}`).trim();
      const docNo = String(row[1] || id).trim();
      const company = String(row[2] || '').trim();
      const taxId = String(row[3] || '').trim();
      const contact = String(row[4] || '').trim();
      const location = String(row[5] || '').trim();
      const serviceType = String(row[6] || 'ซ่อมบำรุง / Overhaul').trim();
      const urgency = String(row[7] || 'ปกติ').trim();
      const warranty = String(row[8] || 'IN_WARRANTY').trim();
      const equipType = String(row[9] || 'MOTOR').trim();
      const brand = String(row[10] || 'BROOK CROMPTON').trim();
      const model = String(row[11] || '-').trim();
      const symptom = String(row[12] || '-').trim();
      const assignedTech = String(row[13] || '').trim();
      const supervisor = String(row[14] || '').trim();
      const statusRaw = String(row[15] || 'NEW').trim().toUpperCase();
      const totalAmount = Number(row[17]) || 0;
      const customerSigned = String(row[18] || '').includes('เซ็นแล้ว');
      const desiredDate = String(row[19] || '').trim();
      const createdAt = String(row[20] || new Date().toISOString()).trim();

      const validStatuses = [
        'NEW',
        'ASSIGNED',
        'INSPECTION',
        'REPAIRING',
        'PENDING_PARTS_APPROVAL',
        'PENDING_SUPERVISOR_APPROVAL',
        'REJECTED',
        'CUSTOMER_ACCEPTANCE',
        'CLOSE'
      ];
      const status = validStatuses.includes(statusRaw) ? (statusRaw as any) : 'NEW';

      return {
        id,
        docNo,
        company,
        taxId,
        contact,
        siteContact: contact,
        location,
        serviceType,
        urgency,
        warranty,
        desiredDate,
        confirmedDate: desiredDate,
        status,
        isLocked: status === 'CLOSE',
        rejectReason: '',
        createdAt,
        assignedTech,
        approvedBySupervisor: supervisor,
        customerSignature: customerSigned ? 'SIGNED_IN_SHEET' : undefined,
        equipmentList: [
          {
            type: equipType,
            brand,
            model,
            symptom,
            beforePhotos: [],
            afterPhotos: [],
            analysis: '',
            cause: '',
            faultOrigin: '',
            correctiveAction: '',
            parts: totalAmount > 0 ? [{ name: 'รายการอะไหล่ซ่อมบำรุง', qty: 1, unit: 'ชุด', price: totalAmount, total: totalAmount }] : [],
            partsApproved: true,
            checklistResults: []
          }
        ]
      };
    });

  return parsedTickets;
}

/**
 * Read-only import of the company's authoritative technician master tab.
 * This function never writes to the source spreadsheet.
 */
export async function loadTechniciansFromSpreadsheet(
  accessToken: string,
  spreadsheetId: string
): Promise<import('../types').Technician[]> {
  const cleanId = extractSpreadsheetId(spreadsheetId);
  if (!cleanId) throw new Error('Spreadsheet ID ไม่ถูกต้อง');

  const range = encodeURIComponent("'รายชื่อช่าง_Technicians'!A2:F5000");
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${cleanId}/values/${range}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!res.ok) {
    const errData = await res.json().catch(() => null);
    throw new Error(errData?.error?.message || `อ่านทะเบียนช่างไม่สำเร็จ (HTTP ${res.status})`);
  }

  const data = await res.json();
  const rows: any[][] = Array.isArray(data.values) ? data.values : [];
  const seen = new Set<string>();
  return rows
    .filter((row) => row && row[0] != null && String(row[0]).trim() && row[1] && String(row[1]).trim())
    .map((row) => {
      const id = String(row[0]).trim();
      const name = String(row[1]).trim();
      const key = id.toLowerCase();
      if (seen.has(key)) return null;
      seen.add(key);
      const rawPin = row[5] == null ? '' : String(row[5]).trim();
      return {
        id,
        name,
        position: String(row[2] || 'ช่างเทคนิค').trim(),
        specialization: String(row[3] || '').trim(),
        employmentStatus: String(row[4] || '').trim(),
        status: 'Available' as const,
        pin: rawPin || id
      };
    })
    .filter((tech) => tech !== null) as import('../types').Technician[];
}

/**
 * Lists user spreadsheets from Google Drive API
 */
export async function listUserSpreadsheets(accessToken: string): Promise<SpreadsheetFileInfo[]> {
  const query = encodeURIComponent("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
  const url = `https://www.googleapis.com/drive/v3/files?q=${query}&orderBy=modifiedTime desc&pageSize=15&fields=files(id,name,modifiedTime,webViewLink)`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => null);
    throw new Error(errData?.error?.message || `Failed to fetch Google Drive files: HTTP ${res.status}`);
  }

  const data = await res.json();
  return (data.files || []).map((f: any) => ({
    id: f.id,
    name: f.name,
    modifiedTime: f.modifiedTime,
    webViewLink: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}`
  }));
}
