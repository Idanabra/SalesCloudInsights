/**
 * SAP Sales Cloud Insights — Outlook Taskpane v1.5
 */

'use strict';

const _VERSION = '2.6';
console.log('[SAP Insights] taskpane.js version', _VERSION);

/* ─────────────────────────────────────────────────────────────────────────
   Settings helpers (localStorage)
───────────────────────────────────────────────────────────────────────── */

const STORAGE_KEY = 'sap_sc_insights_v2';

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (s.url && s.user && s.pass) return s;  // salesCycleId is optional
  } catch (_) { /* ignore */ }
  return null;
}

function saveSettingsToStorage(s) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
}

function basicAuthHeader(user, pass) {
  return 'Basic ' + btoa(unescape(encodeURIComponent(user + ':' + pass)));
}

/* ─────────────────────────────────────────────────────────────────────────
   DOM helpers
───────────────────────────────────────────────────────────────────────── */

function el(id) { return document.getElementById(id); }

function showEl(id)  { el(id).classList.remove('hidden'); }
function hideEl(id)  { el(id).classList.add('hidden'); }

function escHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ─────────────────────────────────────────────────────────────────────────
   Status banner
───────────────────────────────────────────────────────────────────────── */

function showStatus(type, title, detail, withSpinner) {
  const banner = el('status-banner');
  const iconMap = { success: '✔', error: '✖', info: 'ℹ', saving: '' };
  const spinnerHtml = withSpinner
    ? '<span class="status-spinner"></span>'
    : `<span class="status-icon">${iconMap[type] ?? ''}</span>`;

  banner.className = `${type}`;
  banner.innerHTML = `
    ${spinnerHtml}
    <div class="status-text">
      <strong>${escHtml(title)}</strong>
      ${detail ? `<span>${escHtml(detail)}</span>` : ''}
    </div>`;
  showEl('status-banner');
}

function hideStatus() { hideEl('status-banner'); }

/* ─────────────────────────────────────────────────────────────────────────
   SAP API helpers
───────────────────────────────────────────────────────────────────────── */

async function sapFetch(settings, path, options = {}) {
  const url = `${settings.url.replace(/\/$/, '')}${path}`;
  const headers = {
    'Authorization': basicAuthHeader(settings.user, settings.pass),
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    ...options.headers,
  };

  const response = await fetch(url, { ...options, headers });

  if (!response.ok) {
    let msg = `HTTP ${response.status}`;
    try {
      const body = await response.json();
      msg = body?.error?.message?.value ?? body?.message ?? msg;
    } catch (_) { /* ignore */ }
    throw new Error(msg);
  }

  // 204 No Content or non-JSON responses — return null rather than throwing
  if (response.status === 204) return null;
  const ct = response.headers.get('content-type') ?? '';
  if (!ct.includes('json')) return null;
  return response.json();
}

async function fetchOpportunities(settings, searchQuery = '') {
  let path = '/sap/c4c/api/v1/opportunity-service/opportunities';
  const params = new URLSearchParams();

  params.set('$top', '100');
  params.set('$select', 'id,displayId,name,OwnerName,ownerName,owner,LifeCycleStatusCode,salesCycleCode');

  // Filter by Sales Cycle if configured
  if (settings.salesCycleId?.trim()) {
    params.set('$filter', `salesCycleCode eq '${settings.salesCycleId.trim()}'`);
  }

  if (searchQuery.trim()) {
    params.set('$search', searchQuery.trim());
  }

  path += '?' + params.toString();

  const json = await sapFetch(settings, path);
  const all = json?.value ?? json?.data?.value ?? [];

  // Client-side search filter (fallback if $search isn't supported)
  let results = all;
  if (searchQuery.trim()) {
    const q = searchQuery.trim().toLowerCase();
    results = all.filter(o => {
      const name   = (o.name  ?? o.Name  ?? '').toLowerCase();
      const dispId = String(o.displayId ?? o.DisplayID ?? o.id ?? '').toLowerCase();
      return name.includes(q) || dispId.includes(q);
    });
  }

  // Sort A–Z by name
  results.sort((a, b) =>
    oppName(a).localeCompare(oppName(b), undefined, { sensitivity: 'base' })
  );

  return results;
}

/* ─────────────────────────────────────────────────────────────────────────
   Opportunity field extraction
───────────────────────────────────────────────────────────────────────── */

function oppField(o, ...keys) {
  for (const k of keys) {
    if (o[k] != null && o[k] !== '') return o[k];
  }
  return null;
}

function oppName(o) {
  return oppField(o, 'name', 'Name', 'subject', 'Subject') ?? '(Unnamed Opportunity)';
}

function oppDisplayId(o) {
  return String(oppField(o, 'displayId', 'DisplayID', 'ExternalID', 'id') ?? o.id ?? '');
}

function oppUUID(o) {
  return String(o.id ?? o.ObjectID ?? o.opportunityId ?? '');
}

function oppOwner(o) {
  const raw = oppField(o, 'OwnerName', 'ownerName', 'ResponsibleName', 'SalesRepresentativeName', 'owner');
  if (raw == null) return null;
  // SAP sometimes returns owner as a complex object — extract the display string
  if (typeof raw === 'object') {
    return raw.content ?? raw.name ?? raw.displayName ?? raw.fullName ?? null;
  }
  return String(raw) || null;
}

/* ─────────────────────────────────────────────────────────────────────────
   Plain-text cleanup — strip markdown/URI noise injected by Office.js
───────────────────────────────────────────────────────────────────────── */

function cleanPlainText(text) {
  return text
    // [label](url) → label
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    // label <url> or just <url> where url starts with http/https/mailto/tel
    .replace(/[ \t]*<(?:https?|mailto|tel)[^>]*>/g, '')
    // bare <url> leftovers
    .replace(/<(?:https?|mailto|tel)[^>]*>/g, '')
    // normalise line endings
    .replace(/\r\n/g, '\n')
    // blank out whitespace-only lines
    .replace(/^[ \t]+$/gm, '')
    // collapse 3+ blank lines → 2
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/* ─────────────────────────────────────────────────────────────────────────
   Office.js helpers — wrapped in Promises
───────────────────────────────────────────────────────────────────────── */

function getMailBodyTextAsync() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: null },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
          resolve('');
          return;
        }
        resolve(cleanPlainText(result.value));
      }
    );
  });
}

/**
 * Returns only the inner HTML content of the <body> element to avoid
 * sending the full Word HTML document wrapper to SAP.
 */
function getMailBodyHtmlAsync() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: null },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
          resolve(null);
          return;
        }
        const html = result.value;
        const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
        resolve(bodyMatch ? bodyMatch[1] : html);
      }
    );
  });
}

/**
 * Resolve cid: inline image references to base64 data URIs.
 * Requires Mailbox 1.8+ (getAttachmentContentAsync). Gracefully skips if unavailable.
 * Without this, images in email signatures/bodies show as broken in SAP.
 */
async function resolveCidImages(html) {
  if (!html || !html.includes('cid:')) return html;

  const item = Office.context.mailbox.item;
  if (typeof item.getAttachmentContentAsync !== 'function') {
    console.warn('[SAP Insights] getAttachmentContentAsync not available — cid: images will be broken');
    return html;
  }

  const inlineAtts = (item.attachments ?? []).filter(a => a.isInline);
  if (!inlineAtts.length) return html;

  const fetched = await Promise.all(
    inlineAtts.map(att => new Promise(resolve => {
      item.getAttachmentContentAsync(att.id, r => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded
          ? { att, value: r.value }
          : null);
      });
    }))
  );

  let out = html;
  for (const r of fetched) {
    if (!r?.value?.content) continue;
    if (r.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) continue;

    const dataUri = `data:${r.att.contentType};base64,${r.value.content}`;
    // cid: references look like: cid:image001.png@01AB...  — match by attachment name
    const safeName = r.att.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    out = out.replace(
      new RegExp(`src=["']cid:[^"']*${safeName}[^"']*["']`, 'gi'),
      () => `src="${dataUri}"`   // arrow fn avoids $ special meaning in replace
    );
  }

  return out;
}

/** Minimal HTML wrapper for plain text — used when HTML fetch fails. */
function plainToHtml(plain) {
  const paragraphs = plain.split('\n')
    .map(l => l.trim())
    .map(l => l
      ? `<p>${l.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>`
      : '<br>')
    .join('\n');
  return `<!DOCTYPE html><html><body dir="auto">${paragraphs}</body></html>`;
}

function emailList(arr) {
  if (!Array.isArray(arr)) return [];
  return arr.map(r => r?.emailAddress ?? '').filter(Boolean);
}

async function readOutlookEmail() {
  const item = Office.context.mailbox.item;

  const [plainContent, rawHtml] = await Promise.all([
    getMailBodyTextAsync(),
    getMailBodyHtmlAsync(),
  ]);

  // Replace cid: inline image references with base64 data URIs so SAP can render them
  const richTextBody = await resolveCidImages(rawHtml);

  const sentOn = item.dateTimeCreated
    ? new Date(item.dateTimeCreated).toISOString()
    : new Date().toISOString();

  const fromEmail    = item.from?.emailAddress ?? '';
  const mailboxEmail = Office.context.mailbox.userProfile?.emailAddress ?? '';
  const direction    = fromEmail.toLowerCase() === mailboxEmail.toLowerCase()
    ? 'OUTBOUND' : 'INBOUND';

  return {
    subject:      item.subject ?? '',
    messageId:    item.internetMessageId ?? null,
    from:         fromEmail,
    to:           emailList(item.to),
    cc:           emailList(item.cc),
    bcc:          [],
    plainContent,
    richTextBody,
    sentOn,
    direction,
  };
}

/* ─────────────────────────────────────────────────────────────────────────
   Build SAP email payload
───────────────────────────────────────────────────────────────────────── */

function buildEmailPayload(email, opp) {
  const payload = {
    subject:            email.subject,
    messageId:          email.messageId,
    transmissionStatus: 'CREATE',
    direction:          email.direction,
    dataOrigin:         'MANUAL',
    isDraft:            false,
    isAutoReply:        false,
    isBounce:           false,
    sentOn:             email.sentOn,
    from:               email.from,
    toRecipients:       email.to,
    ccRecipients:       email.cc,
    bccRecipients:      email.bcc,
    plainContent:       email.plainContent,
    accounts:             [],
    contacts:             [],
    individualCustomers:  [],
    employees:            [],
    attachments:          [],
    relatedData: {
      interactionNumber:         null,
      interactionOutboundNumber: null,
    },
    relatedObjects: [
      {
        objectId:  oppUUID(opp),
        displayId: oppDisplayId(opp),
        type:      '72',
        role:      'PREDECESSOR',
      },
    ],
  };

  return payload;
}

/** Ensure the HTML being uploaded to SAP is a complete document, not a fragment. */
function wrapHtmlDocument(html) {
  if (!html) return null;
  if (/^\s*<!DOCTYPE/i.test(html) || /^\s*<html/i.test(html)) return html;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
}

/**
 * Correct flow derived from SAP KBA 3567599 + document-service API docs:
 *
 *  Step A: POST /document-service/documents
 *          { fileName: "Email.html", category: "DOCUMENT", type: "10001",
 *            hostObjectId: emailId, hostObjectType: "Email" }
 *          Response: { id: docId, uploadUrl: "<presigned S3 PUT URL>" }
 *
 *  Step B: PUT HTML to uploadUrl — NO Authorization header (pre-signed S3 embeds it)
 *
 *  Step C: PATCH /email-service/emails/{emailId}
 *          { richTextDocumentId: docId }
 *          SAP sets richTextPreSignedURL automatically once the doc is linked.
 *
 * Returns { docId, ok: true } on success or throws.
 */
async function attachRichTextToEmail(settings, emailId, htmlContent) {
  const base = settings.url.replace(/\/$/, '');
  const auth = basicAuthHeader(settings.user, settings.pass);

  // ── Step A: create document record linked to the email ──────────────────
  const docBody = {
    fileName:       '__OriginalContent.html',
    category:       'DOCUMENT',
    type:           '11005',
    hostObjectId:   emailId,
    hostObjectType: '39',
  };

  const docResp = await fetch(`${base}/sap/c4c/api/v1/document-service/documents`, {
    method:  'POST',
    headers: { 'Authorization': auth, 'Content-Type': 'application/json', 'Accept': 'application/json' },
    body:    JSON.stringify(docBody),
  });
  const docText = await docResp.text();
  console.log('[SAP Insights] document-service POST →', docResp.status, docText.slice(0, 300));

  if (!docResp.ok) {
    if (docResp.status === 401 || docResp.status === 403) throw new Error(`document-service auth error HTTP ${docResp.status}`);
    throw new Error(`document-service: HTTP ${docResp.status} — ${docText.slice(0, 200)}`);
  }

  const meta = JSON.parse(docText);

  const docId     = meta?.id ?? meta?.value?.id ?? meta?.value?.[0]?.id ?? meta?.data?.id;
  const uploadUrl = meta?.uploadUrl ?? meta?.value?.uploadUrl ?? meta?.value?.[0]?.uploadUrl
                  ?? meta?.presignedUrl ?? meta?.signedUrl ?? meta?.data?.uploadUrl;

  if (!docId || !uploadUrl) {
    throw new Error(`document-service: missing id/uploadUrl — ${docText.slice(0, 200)}`);
  }

  // ── Step B: upload HTML to pre-signed S3 URL (NO auth header) ───────────
  const s3Resp = await fetch(uploadUrl, {
    method:  'PUT',
    headers: { 'Content-Type': 'text/html; charset=utf-8' },
    body:    htmlContent,
  });

  console.log('[SAP Insights] S3 PUT →', s3Resp.status);
  if (!s3Resp.ok) throw new Error(`S3 upload HTTP ${s3Resp.status}`);

  // ── Step C: PATCH email to set richTextDocumentId ────────────────────────
  const patchResp = await fetch(
    `${base}/sap/c4c/api/v1/email-service/emails/${emailId}`,
    {
      method:  'PATCH',
      headers: { 'Authorization': auth, 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body:    JSON.stringify({ richTextDocumentId: docId }),
    }
  );

  console.log('[SAP Insights] PATCH email richTextDocumentId →', patchResp.status);
  // Non-fatal if PATCH fails — the document is already linked via hostObjectId

  return docId;
}

/* ─────────────────────────────────────────────────────────────────────────
   Main application controller
───────────────────────────────────────────────────────────────────────── */

const app = (() => {

  let _settings     = null;
  let _allOpps      = [];
  let _filteredOpps = [];
  let _selectedOpp  = null;
  let _searchTimer  = null;
  let _loading      = false;

  /* ── View management ─────────────────────────────────────── */

  function showMainView() {
    hideEl('settings-view');
    showEl('main-view');
    el('refresh-btn').style.display = '';
  }

  function showSettingsView() {
    hideEl('main-view');
    showEl('settings-view');
    el('refresh-btn').style.display = 'none';

    if (_settings) {
      el('cfg-url').value         = _settings.url          ?? '';
      el('cfg-user').value        = _settings.user         ?? '';
      el('cfg-pass').value        = _settings.pass         ?? '';
      el('cfg-sales-cycle').value = _settings.salesCycleId ?? '';
    }

    el('cancel-settings-btn').style.display = _settings ? '' : 'none';
  }

  /* ── Opportunity rendering ───────────────────────────────── */

  function renderOpportunities(list) {
    const container = el('opp-list');
    container.innerHTML = '';

    const label = list.length === 1 ? '1 Opportunity' : `${list.length} Opportunities`;
    el('opp-count-label').textContent = label;

    list.forEach(opp => {
      const item = document.createElement('div');
      item.className = 'opp-item';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', 'false');
      item.dataset.id = oppUUID(opp);

      const name   = escHtml(oppName(opp));
      const dispId = escHtml(oppDisplayId(opp));
      const owner  = oppOwner(opp);

      item.innerHTML = `
        <div class="opp-name" title="${name}">${name}</div>
        <div class="opp-meta">
          ${dispId ? `<span class="opp-tag id">ID: ${dispId}</span>` : ''}
          ${owner ? `<span class="opp-tag owner">${escHtml(owner)}</span>` : ''}
        </div>`;

      item.addEventListener('click', () => selectOpportunity(opp, item));
      container.appendChild(item);
    });
  }

  function selectOpportunity(opp, itemEl) {
    el('opp-list').querySelectorAll('.opp-item.selected').forEach(e => {
      e.classList.remove('selected');
      e.setAttribute('aria-selected', 'false');
    });

    _selectedOpp = opp;
    itemEl.classList.add('selected');
    itemEl.setAttribute('aria-selected', 'true');

    el('save-btn').disabled = false;
    hideStatus();
  }

  /* ── Load / search opportunities ────────────────────────── */

  async function loadOpportunities(searchQuery = '') {
    if (_loading) return;
    if (!_settings) { showSettingsView(); return; }

    _loading = true;
    _selectedOpp = null;
    el('save-btn').disabled = true;
    hideStatus();

    hideEl('opp-section');
    hideEl('empty-state');
    hideEl('error-state');
    showEl('loading-state');

    try {
      const opps = await fetchOpportunities(_settings, searchQuery);
      _allOpps      = opps;
      _filteredOpps = opps;

      hideEl('loading-state');

      if (opps.length === 0) {
        el('empty-msg').textContent = searchQuery
          ? `No opportunities match "${searchQuery}".`
          : 'No open opportunities found.';
        showEl('empty-state');
      } else {
        renderOpportunities(opps);
        showEl('opp-section');
      }
    } catch (err) {
      hideEl('loading-state');
      el('error-msg').textContent = err.message ?? 'Failed to load opportunities.';
      showEl('error-state');

      if (err.message.includes('401') || err.message.includes('403') || err.message.includes('Failed to fetch')) {
        showStatus('error', 'Connection failed', err.message);
      }
    } finally {
      _loading = false;
    }
  }

  /* ── Search debounce ─────────────────────────────────────── */

  function onSearchInput(value) {
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(() => loadOpportunities(value), 350);
  }

  /* ── Settings ────────────────────────────────────────────── */

  function showSettings() {
    showSettingsView();
  }

  function cancelSettings() {
    showMainView();
  }

  function saveSettings() {
    const url  = el('cfg-url').value.trim().replace(/\/$/, '');
    const user = el('cfg-user').value.trim();
    const pass = el('cfg-pass').value;

    if (!url || !user || !pass) {
      el('cfg-url') .style.borderColor = url  ? '' : 'var(--error)';
      el('cfg-user').style.borderColor = user ? '' : 'var(--error)';
      el('cfg-pass').style.borderColor = pass ? '' : 'var(--error)';
      return;
    }

    ['cfg-url', 'cfg-user', 'cfg-pass'].forEach(id => {
      el(id).style.borderColor = '';
    });

    _settings = { url, user, pass, salesCycleId: el('cfg-sales-cycle').value.trim() };
    saveSettingsToStorage(_settings);
    showMainView();
    loadOpportunities();
  }

  /* ── Save Email to SAP ───────────────────────────────────── */

  async function saveEmailToSAP() {
    if (!_selectedOpp) return;
    if (!_settings) { showSettingsView(); return; }

    el('save-btn').disabled = true;
    showStatus('saving', 'שומר ב-SAP...', `מקשר ל: ${oppName(_selectedOpp)}`, true);

    try {
      // Step 1: Read Outlook email
      let emailData;
      try {
        emailData = await readOutlookEmail();
      } catch (err) {
        throw new Error('Could not read email data: ' + err.message);
      }

      // Prepare full HTML document from email body
      const htmlDoc = wrapHtmlDocument(
        emailData.richTextBody || plainToHtml(emailData.plainContent)
      );

      // Step 2: POST email (plain metadata — HTML attached separately)
      const payload = buildEmailPayload(emailData, _selectedOpp);
      const result  = await sapFetch(_settings, '/sap/c4c/api/v1/email-service/emails', {
        method: 'POST',
        body:   JSON.stringify(payload),
      });

      // SAP wraps the created entity in { value: { id, ... } } (object, not array)
      const emailId = result?.id ?? result?.value?.id ?? result?.value?.[0]?.id ?? result?.data?.id;
      console.log('[SAP Insights] POST email → id:', emailId, 'response:', JSON.stringify(result).slice(0, 300));

      // Step 3: Attach HTML via document-service (requires emailId from step 2)
      //   POST /document-service/documents { hostObjectId: emailId, hostObjectType: "Email" }
      //   PUT  HTML to S3 upload URL (no auth)
      //   PATCH email with richTextDocumentId
      let htmlStatus = 'טקסט בלבד (אין מזהה מייל)';

      if (emailId) {
        try {
          const docId = await attachRichTextToEmail(_settings, emailId, htmlDoc);
          htmlStatus = `HTML ✔ docId: ${docId.slice(0, 8)}...`;
        } catch (docErr) {
          console.warn('[SAP Insights] HTML attach failed:', docErr.message);
          htmlStatus = `HTML נכשל: ${docErr.message.slice(0, 100)}`;
        }
      }

      showStatus(
        'success',
        'המייל נשמר ב-SAP',
        `מזהה: ${emailId || '—'} | ${htmlStatus} | הזדמנות: ${oppName(_selectedOpp)}`
      );

      _selectedOpp = null;
      el('opp-list').querySelectorAll('.opp-item.selected').forEach(e => {
        e.classList.remove('selected');
        e.setAttribute('aria-selected', 'false');
      });

    } catch (err) {
      showStatus('error', 'Save failed', err.message);
      el('save-btn').disabled = false;
    }
  }

  /* ── Initialisation ──────────────────────────────────────── */

  function init() {
    _settings = loadSettings();

    if (!_settings) {
      showSettingsView();
    } else {
      showMainView();
      loadOpportunities();
    }
  }

  return { init, loadOpportunities, onSearchInput, showSettings, cancelSettings, saveSettings, saveEmailToSAP };

})();

/* ─────────────────────────────────────────────────────────────────────────
   Office.onReady entry point
───────────────────────────────────────────────────────────────────────── */

Office.onReady((info) => {
  if (info.host !== null && info.host !== Office.HostType.Outlook) {
    document.getElementById('app').innerHTML =
      '<div style="padding:16px;color:#a80000">This add-in is designed for Microsoft Outlook.</div>';
    return;
  }

  app.init();
});
