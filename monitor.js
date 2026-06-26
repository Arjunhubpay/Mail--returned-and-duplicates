/**
 * PayOps Email Monitor
 * Schedule: Every 1 hour, 9:00 AM – 6:30 PM UAE (UTC+4), Monday–Friday
 *
 * Env vars required:
 *   GRAPH_TENANT_ID        – Azure AD tenant ID
 *   GRAPH_CLIENT_ID        – App registration client ID
 *   GRAPH_CLIENT_SECRET    – App registration client secret
 *   SLACK_BOT_TOKEN        – Slack bot OAuth token (xoxb-...)
 *   PAYOPS_MAILBOX         – Shared mailbox to monitor (default: payments@hubpay.ae)
 */

'use strict';

const PAYOPS_MAILBOX = process.env.PAYOPS_MAILBOX || 'payments@hubpay.ae';
const SLACK_CHANNEL  = 'C0AUYL9CTD3'; // #automation-payops

// ── Corpay senders ────────────────────────────────────────────────────────────
const CORPAY_SENDER_DOMAINS = ['@corpay.com', '@cambridgefx.com'];

// ── Return/reversal keywords (subject OR body) ────────────────────────────────
const CORPAY_KEYWORDS = [
  'funds returned', 'funds have been returned',
  'funds returned', // subject pattern "Hubpay Ltd - Deal: XXXXX-1 (Funds Returned)"
  'rejected and returned', 'rejected by the beneficiary', 'rejected by the beneficiary bank',
  'placed into your holding balance', 'holding balance',
  'recall could not take place', 'funds rejected',
  'returned to your', 'please advise on how you would like us to proceed',
  'request to reverse', 'reverse unidentified funds',
  'reversal', 'payment reversal',
  'return of funds', 'return to originating', 'return to sender',
  'funds have been declared null and void',
  'proof of return',
  'unidentified funds',
];

const ZAND_KEYWORDS = [
  'possible duplicate', 'duplicate queue',
  'confirm if we can release', 'release the below',
  'request to reverse', 'reverse unidentified funds',
  'return of funds', 'return to originating', 'return to sender',
  'reversal', 'payment reversal',
  'unidentified funds',
  'funds returned', 'rejected',
];

// ── Hubpay repliers ───────────────────────────────────────────────────────────
const HUBPAY_REPLIERS = [
  'arjun.krishnan@hubpay.ae',
  'mohsin.mohammed@hubpay.ae',
  'kiran.shahzadi@hubpay.ae',
  'payments@hubpay.ae',
  'zeba@hubpay.ae',
];

// ─────────────────────────────────────────────────────────────────────────────

function isWeekdayUAE() {
  const now = new Date();
  // UAE is UTC+4
  const uaeOffset = 4 * 60;
  const uaeMs = now.getTime() + (uaeOffset - now.getTimezoneOffset()) * 60000;
  const uaeDate = new Date(uaeMs);
  const day = uaeDate.getDay(); // 0=Sun, 6=Sat
  return day >= 1 && day <= 5;
}

function oneHourAgo() {
  return new Date(Date.now() - 60 * 60 * 1000).toISOString();
}

function matchesKeywords(email, keywords) {
  const haystack = `${email.subject || ''} ${email.bodyPreview || ''}`.toLowerCase();
  return keywords.some(kw => haystack.includes(kw.toLowerCase()));
}

// ── Microsoft Graph helpers ───────────────────────────────────────────────────

async function getGraphToken() {
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET } = process.env;
  if (!GRAPH_TENANT_ID || !GRAPH_CLIENT_ID || !GRAPH_CLIENT_SECRET) {
    throw new Error('Missing GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET');
  }
  const url = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res = await fetch(url, { method: 'POST', body });
  const json = await res.json();
  if (!json.access_token) throw new Error(`Token error: ${JSON.stringify(json)}`);
  return json.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

async function searchMessages(token, mailbox, senderDomainOrAddress, since) {
  const filter = [
    `receivedDateTime ge ${since}`,
    `contains(from/emailAddress/address,'${senderDomainOrAddress}')`,
  ].join(' and ');

  const path =
    `/users/${encodeURIComponent(mailbox)}/messages` +
    `?$filter=${encodeURIComponent(filter)}` +
    `&$select=id,subject,bodyPreview,receivedDateTime,from,conversationId` +
    `&$top=50&$orderby=receivedDateTime desc`;

  const data = await graphGet(token, path);
  return data.value || [];
}

async function threadHasHubpayReply(token, mailbox, conversationId) {
  const path =
    `/users/${encodeURIComponent(mailbox)}/messages` +
    `?$filter=${encodeURIComponent(`conversationId eq '${conversationId}'`)}` +
    `&$select=from,sentDateTime&$top=50`;
  const data = await graphGet(token, path);
  const messages = data.value || [];
  return messages.some(m => {
    const addr = (m.from?.emailAddress?.address || '').toLowerCase();
    return HUBPAY_REPLIERS.includes(addr);
  });
}

// ── Slack helper ──────────────────────────────────────────────────────────────

async function postToSlack(rows) {
  const token = process.env.SLACK_BOT_TOKEN;
  if (!token) throw new Error('Missing SLACK_BOT_TOKEN');

  const now = new Date();
  const uaeTime = now.toLocaleString('en-AE', {
    timeZone: 'Asia/Dubai',
    dateStyle: 'medium',
    timeStyle: 'short',
  });

  const tableRows = rows
    .map((r, i) => `| ${i + 1} | ${r.from} | ${r.subject} | ${r.received} |`)
    .join('\n');

  const text =
    `:rotating_light: *PayOps — Unattended Emails Requiring Action*\n` +
    `_${uaeTime} | ${rows.length} unattended email(s) awaiting response_\n\n` +
    `| # | From | Subject | Received |\n` +
    `|---|------|---------|----------|\n` +
    `${tableRows}\n\n` +
    `:warning: Kindly check and respond to the above.\n` +
    `_PayOps Monitor • Every 1hr | 9:00 AM – 6:30 PM UAE | Mon–Fri_`;

  const res = await fetch('https://slack.com/api/chat.postMessage', {
    method:  'POST',
    headers: {
      Authorization:  `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ channel: SLACK_CHANNEL, text, mrkdwn: true }),
  });
  const json = await res.json();
  if (!json.ok) throw new Error(`Slack error: ${json.error}`);
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  if (!isWeekdayUAE()) {
    console.log('Weekend — skipping.');
    return;
  }

  const token = await getGraphToken();
  const since = oneHourAgo();
  const unattended = [];

  // CHECK 1 — Corpay
  for (const domain of CORPAY_SENDER_DOMAINS) {
    const msgs = await searchMessages(token, PAYOPS_MAILBOX, domain, since);
    for (const msg of msgs) {
      if (!matchesKeywords(msg, CORPAY_KEYWORDS)) continue;
      const replied = await threadHasHubpayReply(token, PAYOPS_MAILBOX, msg.conversationId);
      if (!replied) {
        unattended.push({
          from:     msg.from?.emailAddress?.address || domain,
          subject:  msg.subject,
          received: new Date(msg.receivedDateTime).toLocaleString('en-AE', {
            timeZone:  'Asia/Dubai',
            dateStyle: 'medium',
            timeStyle: 'short',
          }),
        });
      }
    }
  }

  // CHECK 2 — Zand
  const zandMsgs = await searchMessages(token, PAYOPS_MAILBOX, 'customerservice@zand.ae', since);
  for (const msg of zandMsgs) {
    if (!matchesKeywords(msg, ZAND_KEYWORDS)) continue;
    const replied = await threadHasHubpayReply(token, PAYOPS_MAILBOX, msg.conversationId);
    if (!replied) {
      unattended.push({
        from:     'Zand',
        subject:  msg.subject,
        received: new Date(msg.receivedDateTime).toLocaleString('en-AE', {
          timeZone:  'Asia/Dubai',
          dateStyle: 'medium',
          timeStyle: 'short',
        }),
      });
    }
  }

  if (unattended.length === 0) {
    console.log('No unattended emails — done.');
    return;
  }

  await postToSlack(unattended);
  console.log(`Posted ${unattended.length} unattended email(s) to Slack.`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
