import { ServiceTicket } from '../types';

export function getStatusColor(status: string): number {
  switch (status) {
    case 'NEW':
      return 0xf97316; // Orange
    case 'ASSIGNED':
      return 0x2563eb; // Blue
    case 'INSPECTION':
      return 0x0284c7; // Sky blue
    case 'REPAIRING':
      return 0xd97706; // Amber
    case 'PENDING_PARTS_APPROVAL':
    case 'PENDING_SUPERVISOR_APPROVAL':
      return 0x9333ea; // Purple
    case 'REJECTED':
      return 0xdc2626; // Red
    case 'CUSTOMER_ACCEPTANCE':
      return 0x059669; // Emerald
    case 'CLOSE':
      return 0x10b981; // Green
    default:
      return 0x475569; // Slate
  }
}

export function buildDiscordTicketPayload(ticket: ServiceTicket, eventHeading?: string) {
  const portalUrl = 'https://2020skmproduct-service.github.io/sahakim-service-portal/';
  const eqSummary = (ticket.equipmentList || [])
    .map((eq) => `• ${eq.type} (${eq.brand} ${eq.model || '-'}) - อาการ: ${eq.symptom || '-'}`)
    .join('\n');

  let totalPartsPrice = 0;
  (ticket.equipmentList || []).forEach((eq) => {
    (eq.parts || []).forEach((p) => {
      totalPartsPrice += Number(p.total) || (Number(p.qty) || 0) * (Number(p.price) || 0);
    });
  });

  const fields = [
    { name: '📌 สถานะใบงาน', value: `\`${ticket.status}\``, inline: true },
    { name: '🚨 ความเร่งด่วน', value: ticket.urgency || 'ปกติ', inline: true },
    { name: '🛡️ การรับประกัน', value: ticket.warranty || '-', inline: true },
    { name: '🏢 บริษัท / ลูกค้า', value: ticket.company || '-', inline: true },
    { name: '📍 สถานที่หน้างาน', value: ticket.location || '-', inline: true },
    { name: '📞 ผู้ติดต่อ', value: ticket.siteContact || ticket.contact || '-', inline: true },
    { name: '👷 ช่างผู้ปฏิบัติงาน', value: ticket.assignedTech || 'ยังไม่ระบุช่าง', inline: true },
    { name: '📅 วันนัดหมาย', value: ticket.confirmedDate || ticket.desiredDate || '-', inline: true },
    { name: '💰 ยอดอะไหล่รวม', value: `${totalPartsPrice.toLocaleString()} THB`, inline: true }
  ];

  if (ticket.leadTech) {
    fields.push({ name: '👨‍🔧 หัวหน้าช่าง', value: ticket.leadTech, inline: true });
  }
  if (ticket.approvedBySupervisor) {
    fields.push({ name: '✅ ผู้อนุมัติ', value: ticket.approvedBySupervisor, inline: true });
  }
  if (ticket.rejectReason) {
    fields.push({ name: '❌ เหตุผลส่งกลับ', value: ticket.rejectReason, inline: false });
  }

  return {
    content: eventHeading ? `🔔 **${eventHeading}**` : '🔔 **[แจ้งเตือนระบบงานซ่อม - บริษัท สหคิมมอเตอร์ จำกัด]**',
    embeds: [
      {
        title: `📋 ใบงาน: ${ticket.id} (${ticket.docNo || ticket.id})`,
        description: eqSummary ? `🔧 **รายการอุปกรณ์:**\n${eqSummary}` : undefined,
        url: portalUrl,
        color: getStatusColor(ticket.status),
        fields,
        footer: {
          text: 'SAHA KIM MOTOR CO., LTD. · Service Management System V12.1'
        },
        timestamp: new Date().toISOString()
      }
    ]
  };
}

export async function sendDiscordNotification(
  webhookUrl: string,
  payload: { message?: string; ticket?: ServiceTicket; eventHeading?: string }
): Promise<{ ok: boolean; message?: string }> {
  if (!webhookUrl || !webhookUrl.trim().startsWith('https://discord.com/api/webhooks/')) {
    return { ok: false, message: 'URL ของ Discord Webhook ไม่ถูกต้อง (ต้องขึ้นต้นด้วย https://discord.com/api/webhooks/...)' };
  }

  let bodyData: any = {};
  if (payload.ticket) {
    bodyData = buildDiscordTicketPayload(payload.ticket, payload.eventHeading);
  } else if (payload.message) {
    bodyData = {
      content: payload.message
    };
  }

  // 1. Try sending through server proxy first to avoid CORS or network blocks
  try {
    const res = await fetch('/api/discord/notify', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        webhookUrl: webhookUrl.trim(),
        body: bodyData
      })
    });
    if (res.ok) {
      return { ok: true };
    }
  } catch {
    // fallback to direct browser fetch below
  }

  // 2. Direct browser fetch fallback (Discord Webhooks natively support CORS)
  try {
    const directRes = await fetch(webhookUrl.trim(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyData)
    });
    if (directRes.ok) {
      return { ok: true };
    }
    const errText = await directRes.text();
    return { ok: false, message: errText || `HTTP ${directRes.status}` };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'ไม่สามารถเชื่อมต่อ Discord ได้';
    return { ok: false, message: msg };
  }
}
