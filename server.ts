import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Default credentials for Sahakim.service bot (@389fyhbl)
const DEFAULT_LINE_CHANNEL_ID = process.env.LINE_CHANNEL_ID || '2012276060';
const DEFAULT_LINE_CHANNEL_SECRET =
  process.env.LINE_CHANNEL_SECRET || 'db6f7a7f17771925cc3dd2ff0aa38ab5';
const LINE_CONFIG_FILE = '/tmp/sahakim-line-config.json';
const LEGACY_LINE_CONFIG_FILE = path.join(__dirname, '.line-config.json');

interface LineStoredConfig {
  targetId: string;
  groupIds: string[];
  userIds: string[];
  gasWebAppUrl?: string;
  cachedAccessToken?: string;
  tokenExpiresAt?: number;
  updatedAt?: string;
}

let memoryLineConfig: LineStoredConfig | null = null;

function sanitizeTicketsForSheet(tickets: any[]): any[] {
  if (!Array.isArray(tickets)) return [];
  return tickets.map((t) => {
    const cleanEquip = Array.isArray(t.equipmentList)
      ? t.equipmentList.map((eq: any) => ({
          ...eq,
          beforePhotos: Array.isArray(eq.beforePhotos)
            ? eq.beforePhotos.map((p: string) =>
                typeof p === 'string' &&
                (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800)
                  ? 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=600'
                  : p
              )
            : [],
          afterPhotos: Array.isArray(eq.afterPhotos)
            ? eq.afterPhotos.map((p: string) =>
                typeof p === 'string' &&
                (p.startsWith('data:') || p.startsWith('blob:') || p.length > 800)
                  ? 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=600'
                  : p
              )
            : []
        }))
      : [];
    return {
      ...t,
      equipmentList: cleanEquip
    };
  });
}

function loadLineConfig(): LineStoredConfig {
  if (memoryLineConfig) {
    return memoryLineConfig;
  }
  try {
    if (fs.existsSync(LINE_CONFIG_FILE)) {
      const raw = fs.readFileSync(LINE_CONFIG_FILE, 'utf-8');
      const parsed = JSON.parse(raw);
      memoryLineConfig = {
        targetId: parsed.targetId || '',
        groupIds: Array.isArray(parsed.groupIds) ? parsed.groupIds : [],
        userIds: Array.isArray(parsed.userIds) ? parsed.userIds : [],
        gasWebAppUrl: parsed.gasWebAppUrl || '',
        cachedAccessToken: parsed.cachedAccessToken,
        tokenExpiresAt: parsed.tokenExpiresAt,
        updatedAt: parsed.updatedAt
      };
      return memoryLineConfig;
    }
    if (fs.existsSync(LEGACY_LINE_CONFIG_FILE)) {
      const raw = fs.readFileSync(LEGACY_LINE_CONFIG_FILE, 'utf-8');
      const parsed = JSON.parse(raw);
      fs.writeFileSync(LINE_CONFIG_FILE, JSON.stringify(parsed, null, 2), 'utf-8');
      memoryLineConfig = parsed;
      return parsed;
    }
  } catch (e) {
    console.error('Failed to read LINE config:', e);
  }
  memoryLineConfig = { targetId: '', groupIds: [], userIds: [] };
  return memoryLineConfig;
}

function saveLineConfig(cfg: LineStoredConfig) {
  memoryLineConfig = cfg;
  try {
    fs.writeFileSync(LINE_CONFIG_FILE, JSON.stringify(cfg, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to write LINE config:', e);
  }
}

async function getOAuthAccessToken(forceRefresh = false): Promise<string> {
  const cfg = loadLineConfig();
  const now = Date.now();
  if (
    !forceRefresh &&
    cfg.cachedAccessToken &&
    cfg.tokenExpiresAt &&
    cfg.tokenExpiresAt > now + 60_000
  ) {
    return cfg.cachedAccessToken;
  }

  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', DEFAULT_LINE_CHANNEL_ID);
  params.append('client_secret', DEFAULT_LINE_CHANNEL_SECRET);

  const response = await fetch('https://api.line.me/v2/oauth/accessToken', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString()
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`LINE OAuth error (${response.status}): ${errText}`);
  }

  const data = (await response.json()) as { access_token: string; expires_in: number };
  cfg.cachedAccessToken = data.access_token;
  cfg.tokenExpiresAt = now + (data.expires_in || 2592000) * 1000;
  saveLineConfig(cfg);
  return data.access_token;
}

async function getValidLineAccessToken(_customToken?: string): Promise<string> {
  // Use in-memory cached OAuth token directly to eliminate extra /v2/bot/info HTTP roundtrips
  return getOAuthAccessToken(false);
}

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    if (req.method === 'OPTIONS') {
      return res.sendStatus(204);
    }
    next();
  });

  app.use(express.json({ limit: '10mb' }));

  // Health check
  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok' });
  });

  const PUBLIC_WEBHOOK_RELAY_URL =
    process.env.LINE_WEBHOOK_RELAY_URL || 'https://smee.io/sahakim-line-2012276060';

  async function processLineWebhookEvents(events: any[]) {
    if (!Array.isArray(events) || events.length === 0) return;
    const cfg = loadLineConfig();
    let configChanged = false;

    for (const ev of events) {
      const groupId = ev?.source?.groupId;
      const roomId = ev?.source?.roomId;
      const userId = ev?.source?.userId;
      const capturedTarget = groupId || roomId || userId;

      // Ignore synthetic test IDs
      if (capturedTarget === 'C_TEST_123') continue;

      const isNewGroup = Boolean(groupId && !cfg.groupIds.includes(groupId));

      if (groupId && !cfg.groupIds.includes(groupId)) {
        cfg.groupIds.unshift(groupId);
        configChanged = true;
      }
      if (userId && !cfg.userIds.includes(userId)) {
        cfg.userIds.unshift(userId);
        configChanged = true;
      }
      // Always prioritize groupId as primary targetId when a group event arrives
      if (groupId && cfg.targetId !== groupId) {
        cfg.targetId = groupId;
        configChanged = true;
      } else if (!cfg.targetId && capturedTarget) {
        cfg.targetId = capturedTarget;
        configChanged = true;
      }

      const rawMsgText = (ev?.message?.text || '').trim();
      const msgText = rawMsgText.toLowerCase();
      const isTicketShareMessage =
        rawMsgText.includes('เลขที่ใบงาน:') ||
        rawMsgText.includes('[แจ้งเตือนใบงาน') ||
        rawMsgText.includes('[มีใบแจ้งซ่อมใหม่');

      // Always reply immediately in 1-on-1 chat (!groupId), or when invited/triggered in group chat
      const shouldReply =
        !groupId ||
        isNewGroup ||
        isTicketShareMessage ||
        ev?.type === 'join' ||
        ev?.type === 'memberJoined' ||
        ev?.type === 'follow' ||
        msgText === '/id' ||
        msgText === 'id' ||
        msgText.includes('สวัสดี') ||
        msgText.includes('ทดสอบ') ||
        msgText.includes('แจ้งเตือน') ||
        msgText.includes('คู่มือ') ||
        msgText.includes('สถานะ');

      if (ev?.replyToken && shouldReply) {
        try {
          const accessToken = await getValidLineAccessToken();
          const replyText = isTicketShareMessage
            ? `✅ บอทได้รับข้อมูลใบงานและกระจายแจ้งเตือนไปยังทีมงาน/กลุ่ม LINE ทันทีเรียบร้อยแล้วครับ`
            : groupId
              ? `✅ เชื่อมต่อระบบแจ้งเตือน บริษัท สหคิมมอเตอร์ จำกัด เข้ากับกลุ่มนี้สำเร็จแล้ว!\n📌 Group ID: ${groupId}\n🔗 เปิดระบบงานซ่อม: https://2020skmproduct-service.github.io/sahakim-service-portal/`
              : `⚡ ระบบอัตโนมัติ บริษัท สหคิมมอเตอร์ จำกัด (@389fyhbl)\n📌 สถานะบอท: ออนไลน์พร้อมแจ้งเตือนทันที\n📌 User ID: ${userId || '-'}${cfg.groupIds[0] ? `\n📌 Group ID ที่เชื่อมแล้ว: ${cfg.groupIds[0]}` : '\n💡 หากต้องการให้บอทแจ้งเตือนในกลุ่ม: เชิญบอทเข้ากลุ่ม LINE แล้วพิมพ์ "สวัสดี" ในกลุ่ม'}\n🔗 เปิดเว็บแอปแจ้งซ่อม/ติดตามงาน:\nhttps://2020skmproduct-service.github.io/sahakim-service-portal/`;

          await fetch('https://api.line.me/v2/bot/message/reply', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${accessToken}`
            },
            body: JSON.stringify({
              replyToken: ev.replyToken,
              messages: [{ type: 'text', text: replyText }]
            })
          });
        } catch (replyErr) {
          console.error('LINE auto-reply error:', replyErr);
        }
      }

      if (isTicketShareMessage && !groupId) {
        dispatchLineNotification(rawMsgText).catch((fwdErr) => {
          console.error('Failed to forward shared ticket message:', fwdErr);
        });
      }
    }

    if (configChanged) {
      cfg.updatedAt = new Date().toISOString();
      saveLineConfig(cfg);
      console.log('Updated LINE config with captured targetId:', cfg.targetId);
    }
  }

  // Subscribe to public webhook relay so LINE Developers Webhook works without Google auth walls
  function startWebhookRelayListener() {
    let reconnectTimer: NodeJS.Timeout | null = null;
    const connect = async () => {
      try {
        const res = await fetch(PUBLIC_WEBHOOK_RELAY_URL, {
          headers: { Accept: 'text/event-stream' }
        });
        if (!res.ok || !res.body) {
          throw new Error(`Relay HTTP ${res.status}`);
        }
        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed.startsWith('data:')) continue;
            const jsonStr = trimmed.slice(5).trim();
            if (!jsonStr || jsonStr === '{}') continue;
            try {
              const payload = JSON.parse(jsonStr);
              const events = payload?.body?.events || payload?.events;
              if (Array.isArray(events) && events.length > 0) {
                processLineWebhookEvents(events).catch((err) =>
                  console.error('Webhook events error:', err)
                );
              }
              const notifyBody = payload?.body?.sahakimNotify ? payload.body : payload?.sahakimNotify ? payload : null;
              if (notifyBody && notifyBody.message) {
                dispatchLineNotification(
                  String(notifyBody.message),
                  notifyBody.targetId,
                  notifyBody.channelToken
                ).catch((err) => console.warn('Relay notify warning:', err?.message || err));
              }
            } catch {
              // Ignore malformed SSE lines
            }
          }
        }
      } catch (err) {
        console.error('Webhook relay connection closed, reconnecting in 2s:', err);
      }
      if (!reconnectTimer) {
        reconnectTimer = setTimeout(() => {
          reconnectTimer = null;
          connect();
        }, 2000);
      }
    };
    connect();
  }

  async function ensureLineWebhookEndpoint() {
    try {
      const token = await getValidLineAccessToken();
      const getRes = await fetch('https://api.line.me/v2/bot/channel/webhook/endpoint', {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (getRes.ok) {
        const current = (await getRes.json()) as { endpoint?: string; active?: boolean };
        if (current.endpoint === PUBLIC_WEBHOOK_RELAY_URL && current.active) {
          return;
        }
      }
      await fetch('https://api.line.me/v2/bot/channel/webhook/endpoint', {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${token}`
        },
        body: JSON.stringify({ endpoint: PUBLIC_WEBHOOK_RELAY_URL })
      });
      console.log('Successfully configured LINE Webhook Endpoint to:', PUBLIC_WEBHOOK_RELAY_URL);
    } catch (err) {
      console.error('Failed to auto-configure LINE webhook endpoint:', err);
    }
  }

  let lineMonthlyQuotaExceededUntil = 0;

  async function dispatchLineNotification(
    message: string,
    targetId?: string,
    channelToken?: string
  ) {
    const cfg = loadLineConfig();
    const cleanProvidedTarget =
      targetId && String(targetId).trim() !== 'BROADCAST' ? String(targetId).trim() : '';

    if (cleanProvidedTarget) {
      if (cleanProvidedTarget.startsWith('C') && !cfg.groupIds.includes(cleanProvidedTarget)) {
        cfg.groupIds.unshift(cleanProvidedTarget);
      }
      if (cfg.targetId !== cleanProvidedTarget) {
        cfg.targetId = cleanProvidedTarget;
        cfg.updatedAt = new Date().toISOString();
        saveLineConfig(cfg);
      }
    }

    // Check if LINE Monthly Limit was already hit recently
    if (Date.now() < lineMonthlyQuotaExceededUntil) {
      console.warn('[LINE API] Monthly quota limit active (429). Notification skipped gracefully.');
      return {
        status: 'quota_exceeded',
        quotaExceeded: true,
        message: 'โควต้าส่งข้อความ LINE รายเดือนเต็มแล้ว (You have reached your monthly limit - 429)',
        targetId: cleanProvidedTarget || 'NONE',
        pushedTo: []
      };
    }

    const accessToken = await getValidLineAccessToken(channelToken);

    const pushTargets = Array.from(
      new Set(
        [...cfg.groupIds, cleanProvidedTarget, cfg.targetId].filter((id) =>
          Boolean(id && id !== 'BROADCAST')
        )
      )
    );

    const pushedTo: string[] = [];
    let quotaHit = false;
    let lastError = '';

    // Only send targeted push if targets exist; only broadcast if explicitly requested or no specific target
    if (pushTargets.length > 0) {
      await Promise.allSettled(
        pushTargets.map(async (toId) => {
          try {
            const pushRes = await fetch('https://api.line.me/v2/bot/message/push', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${accessToken}`
              },
              body: JSON.stringify({
                to: toId,
                messages: [{ type: 'text', text: String(message) }]
              })
            });
            if (pushRes.ok) {
              pushedTo.push(toId);
            } else {
              const body = await pushRes.text();
              if (pushRes.status === 429 || body.includes('monthly limit')) {
                quotaHit = true;
              }
              lastError = body;
            }
          } catch (err: unknown) {
            lastError = err instanceof Error ? err.message : String(err);
          }
        })
      );
    } else if (targetId === 'BROADCAST' || !cleanProvidedTarget) {
      try {
        const bcRes = await fetch('https://api.line.me/v2/bot/message/broadcast', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${accessToken}`
          },
          body: JSON.stringify({
            messages: [{ type: 'text', text: String(message) }]
          })
        });
        if (bcRes.ok) {
          pushedTo.push('BROADCAST');
        } else {
          const body = await bcRes.text();
          if (bcRes.status === 429 || body.includes('monthly limit')) {
            quotaHit = true;
          }
          lastError = body;
        }
      } catch (err: unknown) {
        lastError = err instanceof Error ? err.message : String(err);
      }
    }

    if (quotaHit) {
      lineMonthlyQuotaExceededUntil = Date.now() + 15 * 60 * 1000;
      console.warn('[LINE API] Monthly limit reached (429): You have reached your monthly limit.');
      return {
        status: 'quota_exceeded',
        quotaExceeded: true,
        message: 'โควต้าส่งข้อความ LINE รายเดือนเต็มแล้ว (You have reached your monthly limit)',
        targetId: cleanProvidedTarget || 'NONE',
        pushedTo: []
      };
    }

    if (pushedTo.length === 0 && lastError) {
      console.warn('[LINE API] Send push warning:', lastError);
      return {
        status: 'error',
        error: lastError,
        targetId: cleanProvidedTarget || 'NONE',
        pushedTo: []
      };
    }

    const pushedGroup = pushedTo.find((id) => id.startsWith('C') || id.startsWith('R'));
    return {
      status: 'ok',
      mode: pushedGroup ? 'group' : pushedTo.length > 0 ? 'push' : 'none',
      targetId: pushedGroup || pushedTo[0] || 'NONE',
      pushedTo
    };
  }

  // LINE Webhook verification & event receiver (captures Group ID / User ID automatically)
  app.get('/api/line/webhook', (_req, res) => {
    res.status(200).json({ status: 'ok', bot: '@389fyhbl', relayUrl: PUBLIC_WEBHOOK_RELAY_URL });
  });

  app.post('/api/line/webhook', async (req, res) => {
    try {
      const events = Array.isArray(req.body?.events) ? req.body.events : [];
      await processLineWebhookEvents(events);
      return res.status(200).json({ status: 'ok' });
    } catch (err) {
      console.error('LINE Webhook Error:', err);
      return res.status(200).json({ status: 'ok' });
    }
  });

  // =========================================================================
  // REAL-TIME ONLINE SHARED DATA STORE (Syncs Tickets, Schedule, Techs & Master Data across all devices)
  // =========================================================================
  const ONLINE_STATE_FILE = path.join(__dirname, '.sahakim-online-state.json');
  const TMP_ONLINE_STATE_FILE = '/tmp/sahakim-online-state.json';

  interface OnlineSharedState {
    version: number;
    updatedAt: string;
    updatedBy: string;
    tickets: any[];
    technicians?: any[];
    supervisors?: string[];
    masterBrands?: string[];
    masterEquipmentTypes?: string[];
    masterChecklists?: Record<string, string[]>;
    companyProfile?: any;
  }

  let memoryOnlineState: OnlineSharedState | null = null;
  const onlineSseClients = new Set<express.Response>();

  function loadOnlineState(): OnlineSharedState {
    if (memoryOnlineState) return memoryOnlineState;
    try {
      const targetFile = fs.existsSync(ONLINE_STATE_FILE)
        ? ONLINE_STATE_FILE
        : fs.existsSync(TMP_ONLINE_STATE_FILE)
        ? TMP_ONLINE_STATE_FILE
        : '';
      if (targetFile) {
        const raw = fs.readFileSync(targetFile, 'utf-8');
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.tickets)) {
          memoryOnlineState = {
            version: Number(parsed.version) || 1,
            updatedAt: parsed.updatedAt || new Date().toISOString(),
            updatedBy: parsed.updatedBy || 'SYSTEM',
            tickets: parsed.tickets,
            technicians: parsed.technicians,
            supervisors: parsed.supervisors,
            masterBrands: parsed.masterBrands,
            masterEquipmentTypes: parsed.masterEquipmentTypes,
            masterChecklists: parsed.masterChecklists,
            companyProfile: parsed.companyProfile
          };
          return memoryOnlineState;
        }
      }
    } catch (err) {
      console.error('Failed to read online state:', err);
    }
    memoryOnlineState = {
      version: 0,
      updatedAt: new Date().toISOString(),
      updatedBy: 'INIT',
      tickets: []
    };
    return memoryOnlineState;
  }

  function saveOnlineState(nextState: OnlineSharedState) {
    memoryOnlineState = nextState;
    const jsonStr = JSON.stringify(nextState, null, 2);
    try {
      fs.writeFileSync(ONLINE_STATE_FILE, jsonStr, 'utf-8');
    } catch {
      // Fallback to /tmp if workspace dir is read-only
    }
    try {
      fs.writeFileSync(TMP_ONLINE_STATE_FILE, jsonStr, 'utf-8');
    } catch (err) {
      console.error('Failed to write online state to /tmp:', err);
    }
  }

  function broadcastOnlineState(state: OnlineSharedState, clientId?: string) {
    const payload = JSON.stringify({
      type: 'state_updated',
      version: state.version,
      updatedAt: state.updatedAt,
      updatedBy: state.updatedBy,
      clientId: clientId || '',
      ticketCount: Array.isArray(state.tickets) ? state.tickets.length : 0
    });
    onlineSseClients.forEach((clientRes) => {
      try {
        clientRes.write(`data: ${payload}\n\n`);
      } catch {
        onlineSseClients.delete(clientRes);
      }
    });
  }

  app.get('/api/online/events', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders?.();

    const current = loadOnlineState();
    res.write(
      `data: ${JSON.stringify({
        type: 'connected',
        version: current.version,
        updatedAt: current.updatedAt,
        ticketCount: current.tickets.length
      })}\n\n`
    );

    onlineSseClients.add(res);
    const keepAlive = setInterval(() => {
      try {
        res.write(': keep-alive\n\n');
      } catch {
        clearInterval(keepAlive);
        onlineSseClients.delete(res);
      }
    }, 20000);

    req.on('close', () => {
      clearInterval(keepAlive);
      onlineSseClients.delete(res);
    });
  });

  app.get('/api/online/state', (_req, res) => {
    const state = loadOnlineState();
    res.json({
      status: 'ok',
      version: state.version,
      updatedAt: state.updatedAt,
      updatedBy: state.updatedBy,
      tickets: state.tickets,
      data: state,
      ...state,
      onlineClients: onlineSseClients.size
    });
  });

  app.post('/api/online/state', (req, res) => {
    try {
      const current = loadOnlineState();
      const body = req.body || {};
      const nextVersion = (current.version || 0) + 1;
      const nowIso = new Date().toISOString();

      const nextState: OnlineSharedState = {
        version: nextVersion,
        updatedAt: nowIso,
        updatedBy: body.updatedBy || 'ONLINE_USER',
        tickets: Array.isArray(body.tickets) ? body.tickets : current.tickets,
        technicians: Array.isArray(body.technicians) ? body.technicians : current.technicians,
        supervisors: Array.isArray(body.supervisors) ? body.supervisors : current.supervisors,
        masterBrands: Array.isArray(body.masterBrands) ? body.masterBrands : current.masterBrands,
        masterEquipmentTypes: Array.isArray(body.masterEquipmentTypes)
          ? body.masterEquipmentTypes
          : current.masterEquipmentTypes,
        masterChecklists:
          body.masterChecklists && typeof body.masterChecklists === 'object'
            ? body.masterChecklists
            : current.masterChecklists,
        companyProfile:
          body.companyProfile && typeof body.companyProfile === 'object'
            ? body.companyProfile
            : current.companyProfile
      };

      saveOnlineState(nextState);
      broadcastOnlineState(nextState, body.clientId);

      res.json({
        status: 'ok',
        version: nextState.version,
        updatedAt: nextState.updatedAt,
        ticketCount: nextState.tickets.length,
        onlineClients: onlineSseClients.size
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown online state save error';
      res.status(500).json({ status: 'error', message: msg });
    }
  });

  // Direct download endpoint for standalone index.html
  app.get('/download/index.html', async (_req, res) => {
    try {
      const { generateStandaloneHtmlString } = await import('./src/standaloneGenerator.js');
      const state = loadOnlineState();
      const html = generateStandaloneHtmlString(
        state.tickets || [],
        state.technicians || [],
        [],
        [],
        [],
        'admin123',
        '1234',
        '',
        '',
        '',
        state.supervisors || []
      );
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.setHeader('Content-Disposition', 'attachment; filename="index.html"');
      res.send(html);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : 'Export error';
      res.status(500).send(`Error generating index.html: ${msg}`);
    }
  });

  // Get current LINE bot configuration & auto-detected Group ID + Google Sheet URL
  app.get('/api/line/config', async (_req, res) => {
    try {
      const cfg = loadLineConfig();
      const token = await getValidLineAccessToken();
      res.json({
        status: 'ok',
        channelId: DEFAULT_LINE_CHANNEL_ID,
        botBasicId: '@389fyhbl',
        webhookRelayUrl: PUBLIC_WEBHOOK_RELAY_URL,
        hasValidToken: Boolean(token),
        targetId: cfg.groupIds[0] || cfg.targetId || cfg.userIds[0] || '',
        groupIds: cfg.groupIds,
        userIds: cfg.userIds,
        gasWebAppUrl: cfg.gasWebAppUrl || '',
        updatedAt: cfg.updatedAt || null
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown error';
      res.status(500).json({ status: 'error', message: msg });
    }
  });

  // Server-side proxy to sync tickets to Google Apps Script Web App (avoids browser CORS & 50k cell overflow)
  app.post('/api/gsheet/sync', async (req, res) => {
    try {
      const { gasWebAppUrl, tickets, lineNotifyText, lineChannelToken, lineTargetId } =
        req.body || {};
      const cfg = loadLineConfig();
      const trimmedUrl = typeof gasWebAppUrl === 'string' ? gasWebAppUrl.trim() : '';

      if (trimmedUrl && trimmedUrl !== cfg.gasWebAppUrl) {
        cfg.gasWebAppUrl = trimmedUrl;
        cfg.updatedAt = new Date().toISOString();
        saveLineConfig(cfg);
      }

      const activeGasUrl = trimmedUrl || cfg.gasWebAppUrl || '';
      if (!activeGasUrl) {
        return res.status(400).json({
          status: 'no_url',
          message:
            'ยังไม่ได้ระบุลิงก์ Google Apps Script Web App URL (ที่ลงท้ายด้วย /exec) ในหน้าตั้งค่า'
        });
      }

      if (
        activeGasUrl.includes('docs.google.com/spreadsheets') &&
        !activeGasUrl.includes('script.google.com')
      ) {
        return res.status(400).json({
          status: 'invalid_url',
          message:
            'ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet (docs.google.com) กรุณาใช้ลิงก์ Web App URL จากเมนู ส่วนขยาย > Apps Script > การทำให้ใช้งานได้ (Deploy) ที่ลงท้ายด้วย /exec'
        });
      }

      const cleanTickets = sanitizeTicketsForSheet(Array.isArray(tickets) ? tickets : []);

      const gasRes = await fetch(activeGasUrl, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          tickets: cleanTickets,
          lineNotifyText: lineNotifyText || '',
          lineChannelToken: lineChannelToken || '',
          lineTargetId: lineTargetId || cfg.groupIds[0] || cfg.targetId || ''
        })
      });

      const rawText = await gasRes.text();
      if (
        rawText.includes('<!DOCTYPE html>') ||
        rawText.includes('accounts.google.com') ||
        rawText.includes('ServiceLogin')
      ) {
        return res.status(403).json({
          status: 'auth_error',
          message:
            'Google Apps Script ปฏิเสธสิทธิ์เข้าถึง หรือยังไม่ได้อัปเดตการ Deploy — กรุณาตั้งค่า "ผู้ที่มีสิทธิ์เข้าถึง (Who has access)" เป็น "ทุกคน (Anyone)" แล้วกด Deploy เวอร์ชันใหม่'
        });
      }

      let parsedGas: any = null;
      try {
        parsedGas = JSON.parse(rawText);
      } catch {
        // non-JSON response
      }

      if (parsedGas && parsedGas.status === 'error') {
        return res.status(500).json({
          status: 'gas_error',
          message: `Google Apps Script แจ้งข้อผิดพลาด: ${parsedGas.message || 'Unknown error'}`
        });
      }

      return res.json({
        status: 'ok',
        syncedCount: cleanTickets.length,
        gasWebAppUrl: activeGasUrl,
        gasResponse: parsedGas || rawText.slice(0, 200)
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown Google Sheet sync error';
      console.error('Google Sheet Sync Error:', msg);
      return res.status(500).json({ status: 'error', message: msg });
    }
  });

  // Server-side proxy to load tickets from Google Apps Script Web App
  app.get('/api/gsheet/load', async (req, res) => {
    try {
      const cfg = loadLineConfig();
      const queryUrl = typeof req.query.url === 'string' ? req.query.url.trim() : '';
      const activeGasUrl = queryUrl || cfg.gasWebAppUrl || '';
      if (!activeGasUrl) {
        return res.status(400).json({
          status: 'no_url',
          message: 'ยังไม่ได้ระบุลิงก์ Google Apps Script Web App URL (/exec)'
        });
      }
      if (
        activeGasUrl.includes('docs.google.com/spreadsheets') &&
        !activeGasUrl.includes('script.google.com')
      ) {
        return res.status(400).json({
          status: 'invalid_url',
          message:
            'ลิงก์ที่วางเป็นลิงก์หน้าตาราง Google Sheet กรุณาใช้ลิงก์ Web App URL จาก Apps Script ที่ลงท้ายด้วย /exec'
        });
      }

      const gasRes = await fetch(activeGasUrl, {
        method: 'GET',
        redirect: 'follow'
      });
      const rawText = await gasRes.text();
      if (rawText.includes('<!DOCTYPE html>') || rawText.includes('accounts.google.com')) {
        return res.status(403).json({
          status: 'auth_error',
          message:
            'ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบว่าตั้งค่า Deploy Web App เป็น "ทุกคน (Anyone)" แล้ว'
        });
      }
      const data = JSON.parse(rawText);
      return res.json(data);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown error loading Google Sheet';
      return res.status(500).json({ status: 'error', message: msg });
    }
  });

  // Send LINE Notification directly from backend (Push to Group/User AND Broadcast to all bot friends)
  app.post('/api/line/notify', async (req, res) => {
    try {
      const { message, targetId, channelToken } = req.body || {};
      if (!message || !String(message).trim()) {
        return res.status(400).json({ error: 'Message is required' });
      }
      const result = await dispatchLineNotification(String(message), targetId, channelToken);
      return res.json(result);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown LINE notification error';
      console.warn('LINE Notify Warning:', msg);
      return res.json({
        status: 'error',
        error: msg,
        quotaExceeded: msg.includes('429') || msg.includes('monthly limit')
      });
    }
  });

  // AI Troubleshooting Endpoint referencing Pump Handbook, EASA Motor Standards & Manufacturer Guides
  app.post('/api/ai/troubleshoot', async (req, res) => {
    try {
      const { equipmentType, brand, model, symptom, detail, checklistResults } = req.body;

      if (!process.env.GEMINI_API_KEY) {
        return res.status(503).json({
          error: 'GEMINI_API_KEY is not configured on the server.'
        });
      }

      const ai = new GoogleGenAI({
        apiKey: process.env.GEMINI_API_KEY,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build'
          }
        }
      });

      const checklistSummary = Array.isArray(checklistResults)
        ? checklistResults.map((c: { item: string; beforeStatus: string }) => `${c.item}: ${c.beforeStatus}`).join(', ')
        : '';

      const prompt = `คุณคือหัวหน้าวิศวกรเครื่องกลและไฟฟ้าอาวุโสของ บริษัท สหคิมมอเตอร์ จำกัด (Saha Kim Motor Co., Ltd.)
จงวิเคราะห์ปัญหาอาการเสียของอุปกรณ์ต่อไปนี้โดยอ้างอิงหลักวิศวกรรมจาก:
1. Karassik's Pump Handbook (Hydraulic Institute Standards, Cavitation, NPSH, Mechanical Seal, Bearing & Shaft Alignment)
2. มาตรฐาน EASA AR100 / IEEE 43 (สำหรับมอเตอร์ไฟฟ้า, Insulation Resistance Megger Test, Winding Burnout, Voltage Unbalance, Bearing Fluting)
3. คู่มือ Troubleshooting ของแบรนด์ผู้ผลิตชั้นนำ (${brand || 'Industrial Standard'} เช่น BULL POWER, BROOK CROMPTON, STAC, UNOMACH, VARISCO, FYBROC, EBARA, GRUNDFOS)

ข้อมูลอุปกรณ์ที่ตรวจพบ:
- ประเภทอุปกรณ์ (Equipment Type): ${equipmentType || 'PUMP / MOTOR'}
- แบรนด์ (Brand): ${brand || '-'}
- รุ่น (Model): ${model || '-'}
- อาการเสียที่แจ้ง (Symptom): ${symptom || 'ทำงานผิดปกติ'}
- รายละเอียดหน้างานเพิ่มเติม (Detail): ${detail || '-'}
- ผลตรวจเช็คเบื้องต้น (Before Checklist): ${checklistSummary || '-'}

กรุณาวิเคราะห์และส่งคำตอบกลับเป็น JSON ตามโครงสร้างที่กำหนด (ภาษาไทยกระชับ ชัดเจน นำไปใส่ใบ Service Report ได้ทันที และช่างสามารถแก้ไขต่อยอดได้):`;

      const response = await ai.models.generateContent({
        model: 'gemini-3.8-flash',
        contents: prompt,
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              analysis: {
                type: Type.STRING,
                description:
                  'ผลการวิเคราะห์ทางเทคนิคอย่างละเอียด (อ้างอิงหลักการจาก Pump Handbook หรือมาตรฐาน EASA AR100)'
              },
              cause: {
                type: Type.STRING,
                description:
                  'สาเหตุรากฐานของปัญหา (Root Cause) ทั้งด้าน Mechanical / Electrical / Hydraulic'
              },
              correctiveAction: {
                type: Type.STRING,
                description:
                  'แนวทางแก้ไขและขั้นตอนการซ่อมบำรุงตามมาตรฐาน (รวมค่ามาตรฐานที่ควรวัด เช่น ค่า Megger, Alignment Tolerance, Clearances)'
              },
              standardsReference: {
                type: Type.STRING,
                description:
                  'ข้อกำหนดมาตรฐานที่ใช้อ้างอิง (เช่น EASA AR100, IEEE 43, Karassik Pump Handbook, ISO 10816 Vibration, หรือคู่มือแบรนด์ผู้ผลิต)'
              },
              recommendedChecklist: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description:
                  'รายการหัวข้อตรวจเช็คหน้างานเพิ่มเติม 2-4 ข้อที่แนะนำให้ช่างตรวจสอบเพิ่ม'
              },
              recommendedParts: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING },
                    qty: { type: Type.NUMBER },
                    unit: { type: Type.STRING }
                  },
                  required: ['name', 'qty', 'unit']
                },
                description: 'รายการอะไหล่ที่แนะนำให้ตรวจสอบหรือเปลี่ยนตามอาการเสียนี้'
              }
            },
            required: [
              'analysis',
              'cause',
              'correctiveAction',
              'standardsReference',
              'recommendedChecklist',
              'recommendedParts'
            ]
          }
        }
      });

      const rawText = response.text || '{}';
      const parsed = JSON.parse(rawText);
      return res.json(parsed);
    } catch (error: unknown) {
      const errMessage = error instanceof Error ? error.message : 'Unknown AI error';
      console.error('AI Troubleshoot Error:', errMessage);
      return res.status(500).json({ error: errMessage });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
    setTimeout(() => {
      startWebhookRelayListener();
      ensureLineWebhookEndpoint().catch((err) => {
        console.error('Background LINE endpoint init error:', err);
      });
    }, 500);
  });
}

startServer();
