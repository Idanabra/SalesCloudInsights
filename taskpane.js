/**
 * SAP Sales Cloud Insights — Outlook Taskpane v1.5
 */

'use strict';

const _VERSION = '1.5';
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
    if (s.url && s.user && s.pass) return s;
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

  return response.json();
}

async function fetchOpportunities(settings, searchQuery = '') {
  let path = '/sap/c4c/api/v1/opportunity-service/opportunities';
  const params = new URLSearchParams();

  params.set('$top', '100');
  params.set('$select', 'id,displayId,name,OwnerName,ownerName,owner,LifeCycleStatusCode');

  if (searchQuery.trim()) {
    params.set('$search', searchQuery.trim());
  }

  path += '?' + params.toString();

  const json = await sapFetch(settings, path);
  const all = json?.value ?? json?.data?.value ?? [];

  if (searchQuery.trim()) {
    const q = searchQuery.trim().toLowerCase();
    return all.filter(o => {
      const name   = (o.name  ?? o.Name  ?? '').toLowerCase();
      const dispId = String(o.displayId ?? o.DisplayID ?? o.id ?? '').toLowerCase();
      return name.includes(q) || dispId.includes(q);
    });
  }

  return all;
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
  return oppField(o, 'OwnerName', 'ownerName', 'owner', 'ResponsibleName', 'SalesRepresentativeName') ?? '—';
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

  const [plainContent, richTextBody] = await Promise.all([
    getMailBodyTextAsync(),
    getMailBodyHtmlAsync(),
  ]);

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

  // Send HTML body under all known SAP field names; SAP ignores unknown fields.
  const htmlContent = email.richTextBody || plainToHtml(email.plainContent);
  payload.richTextBody = htmlContent;
  payload.richText     = htmlContent;
  payload.htmlBody     = htmlContent;
  payload.htmlContent  = htmlContent;

  return payload;
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
      el('cfg-url').value  = _settings.url  ?? '';
      el('cfg-user').value = _settings.user ?? '';
      el('cfg-pass').value = _settings.pass ?? '';
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
      const owner  = escHtml(oppOwner(opp));

      item.innerHTML = `
        <div class="opp-name" title="${name}">${name}</div>
        <div class="opp-meta">
          ${dispId ? `<span class="opp-tag id">ID: ${dispId}</span>` : ''}
          ${owner !== '—' ? `<span class="opp-tag owner">👤 ${owner}</span>` : ''}
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

    _settings = { url, user, pass };
    saveSettingsToStorage(_settings);
    showMainView();
    loadOpportunities();
  }

  /* ── Save Email to SAP ───────────────────────────────────── */

  async function saveEmailToSAP() {
    if (!_selectedOpp) return;
    if (!_settings) { showSettingsView(); return; }

    el('save-btn').disabled = true;
    showStatus('saving', 'Saving to SAP…', `Linking to: ${oppName(_selectedOpp)}`, true);

    try {
      let emailData;
      try {
        emailData = await readOutlookEmail();
      } catch (err) {
        throw new Error('Could not read email data: ' + err.message);
      }

      const payload = buildEmailPayload(emailData, _selectedOpp);

      const path = '/sap/c4c/api/v1/email-service/emails';
      await sapFetch(_settings, path, {
        method: 'POST',
        body: JSON.stringify(payload),
      });

      showStatus(
        'success',
        'Email saved to SAP',
        `Linked to opportunity: ${oppName(_selectedOpp)} (${oppDisplayId(_selectedOpp)})`
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
