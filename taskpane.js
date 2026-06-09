/**
 * SAP Sales Cloud Insights — Outlook Taskpane
 * Connects Outlook Desktop to SAP Sales Cloud V2 via Basic Auth REST APIs.
 *
 * Flow:
 *  1. Office.onReady() initialises the add-in.
 *  2. If settings are missing → show Settings view.
 *  3. Otherwise → fetch opportunities and show Main view.
 *  4. User selects an opportunity → "Save Email to SAP" button activates.
 *  5. On save: reads Outlook item metadata, POSTs to SAP email-service.
 */

'use strict';

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

/**
 * Show a status banner above the opportunity list.
 * @param {'success'|'error'|'info'|'saving'} type
 * @param {string} title
 * @param {string} [detail]
 * @param {boolean} [withSpinner]
 */
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

/**
 * Fetch opportunities from SAP Sales Cloud V2.
 * Uses $search when a query is provided; falls back to $top=50 otherwise.
 * Client-side filtering is applied as a second pass.
 *
 * @param {object} settings
 * @param {string} [searchQuery]
 * @returns {Promise<Array>}
 */
async function fetchOpportunities(settings, searchQuery = '') {
  let path = '/sap/c4c/api/v1/opportunity-service/opportunities';
  const params = new URLSearchParams();

  params.set('$top', '100');
  params.set('$select', 'id,displayId,name,OwnerName,ownerName,owner,LifeCycleStatusCode');

  if (searchQuery.trim()) {
    // $search is supported in SAP Sales Cloud V2 OData services
    params.set('$search', searchQuery.trim());
  }

  path += '?' + params.toString();

  const json = await sapFetch(settings, path);
  // OData v4 wraps results in .value
  const all = json?.value ?? json?.data?.value ?? [];

  // Client-side filter as fallback (in case $search isn't supported on the tenant)
  if (searchQuery.trim()) {
    const q = searchQuery.trim().toLowerCase();
    return all.filter(o => {
      const name  = (o.name  ?? o.Name  ?? '').toLowerCase();
      const dispId = String(o.displayId ?? o.DisplayID ?? o.id ?? '').toLowerCase();
      return name.includes(q) || dispId.includes(q);
    });
  }

  return all;
}

/* ─────────────────────────────────────────────────────────────────────────
   Opportunity field extraction
   SAP V2 field names can vary by tenant config; we try common variants.
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
  // The UUID used for relatedObjects linkage
  return String(o.id ?? o.ObjectID ?? o.opportunityId ?? '');
}

function oppOwner(o) {
  return oppField(o, 'OwnerName', 'ownerName', 'owner', 'ResponsibleName', 'SalesRepresentativeName') ?? '—';
}

/* ─────────────────────────────────────────────────────────────────────────
   Office.js helpers — wrapped in Promises
───────────────────────────────────────────────────────────────────────── */

/**
 * Get the plain-text body of the current mail item.
 * @returns {Promise<string>}
 */
function getMailBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: null },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value ?? '');
        } else {
          reject(new Error(result.error?.message ?? 'Failed to read email body'));
        }
      }
    );
  });
}

/**
 * Extract a simple array of email strings from an EmailAddressDetails array.
 * @param {Array} arr
 * @returns {string[]}
 */
function emailList(arr) {
  if (!Array.isArray(arr)) return [];
  return arr.map(r => r?.emailAddress ?? '').filter(Boolean);
}

/**
 * Collect all email metadata from the current Outlook item.
 * NOTE: item.bcc is NOT available in read mode (Outlook API restriction).
 *
 * @returns {Promise<object>}
 */
async function readOutlookEmail() {
  const item = Office.context.mailbox.item;

  const [body] = await Promise.all([getMailBodyAsync()]);

  return {
    subject:    item.subject ?? '',
    from:       item.from?.emailAddress ?? '',
    to:         emailList(item.to),
    cc:         emailList(item.cc),
    bcc:        [], // Not available in read mode — Outlook API restriction
    body,
  };
}

/* ─────────────────────────────────────────────────────────────────────────
   Build SAP email payload
───────────────────────────────────────────────────────────────────────── */

/**
 * @param {object} email   – from readOutlookEmail()
 * @param {object} opp     – selected SAP opportunity record
 * @returns {object}       – request body for POST /email-service/emails
 */
function buildEmailPayload(email, opp) {
  return {
    subject:       email.subject,
    from:          email.from,
    toRecipients:  email.to,
    ccRecipients:  email.cc,
    bccRecipients: email.bcc,
    body:          email.body,
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
        type:      '72',       // SAP type code for Opportunity
        role:      'PREDECESSOR',
      },
    ],
  };
}

/* ─────────────────────────────────────────────────────────────────────────
   Main application controller
───────────────────────────────────────────────────────────────────────── */

const app = (() => {

  let _settings     = null;   // Current saved settings
  let _allOpps      = [];     // Full opportunity list from SAP
  let _filteredOpps = [];     // After search filter
  let _selectedOpp  = null;   // Currently selected opportunity object
  let _searchTimer  = null;   // Debounce timer ID
  let _loading      = false;  // Guard against parallel loads

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

    // Pre-fill form from saved settings
    if (_settings) {
      el('cfg-url').value  = _settings.url  ?? '';
      el('cfg-user').value = _settings.user ?? '';
      el('cfg-pass').value = _settings.pass ?? '';
    }

    // Hide Cancel if no settings yet (first-run)
    if (!_settings) {
      el('cancel-settings-btn').style.display = 'none';
    } else {
      el('cancel-settings-btn').style.display = '';
    }
  }

  /* ── Opportunity rendering ───────────────────────────────── */

  function renderOpportunities(list) {
    const container = el('opp-list');
    container.innerHTML = '';

    // Update count label
    const label = list.length === 1
      ? '1 Opportunity'
      : `${list.length} Opportunities`;
    el('opp-count-label').textContent = label;

    list.forEach(opp => {
      const item = document.createElement('div');
      item.className = 'opp-item';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', 'false');
      item.dataset.id = oppUUID(opp);

      const name    = escHtml(oppName(opp));
      const dispId  = escHtml(oppDisplayId(opp));
      const owner   = escHtml(oppOwner(opp));

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
    // Deselect any previous selection
    container('opp-list').querySelectorAll('.opp-item.selected').forEach(el => {
      el.classList.remove('selected');
      el.setAttribute('aria-selected', 'false');
    });

    _selectedOpp = opp;
    itemEl.classList.add('selected');
    itemEl.setAttribute('aria-selected', 'true');

    el('save-btn').disabled = false;
    hideStatus();
  }

  function container(id) { return document.getElementById(id); }

  /* ── Load / search opportunities ────────────────────────── */

  async function loadOpportunities(searchQuery = '') {
    if (_loading) return;
    if (!_settings) { showSettingsView(); return; }

    _loading = true;
    _selectedOpp = null;
    el('save-btn').disabled = true;
    hideStatus();

    // Show loading state, hide others
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

      // If it looks like an auth or config error, nudge user to settings
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

    // Reset error borders
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
      // Step 1: Read the current Outlook email
      let emailData;
      try {
        emailData = await readOutlookEmail();
      } catch (err) {
        throw new Error('Could not read email data: ' + err.message);
      }

      // Step 2: Build the POST payload
      const payload = buildEmailPayload(emailData, _selectedOpp);

      // Step 3: POST to SAP email-service
      const path = '/sap/c4c/api/v1/email-service/emails';
      await sapFetch(_settings, path, {
        method: 'POST',
        body: JSON.stringify(payload),
      });

      // Step 4: Show success
      showStatus(
        'success',
        'Email saved to SAP',
        `Linked to opportunity: ${oppName(_selectedOpp)} (${oppDisplayId(_selectedOpp)})`
      );

      // Clear selection after save
      _selectedOpp = null;
      container('opp-list').querySelectorAll('.opp-item.selected').forEach(e => {
        e.classList.remove('selected');
        e.setAttribute('aria-selected', 'false');
      });

    } catch (err) {
      showStatus('error', 'Save failed', err.message);
      el('save-btn').disabled = false; // Re-enable so user can retry
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
   All Office.js calls must happen after this resolves.
───────────────────────────────────────────────────────────────────────── */

Office.onReady((info) => {
  // Only activate for Outlook; guard against running in a plain browser tab
  if (info.host !== null && info.host !== Office.HostType.Outlook) {
    document.getElementById('app').innerHTML =
      '<div style="padding:16px;color:#a80000">This add-in is designed for Microsoft Outlook.</div>';
    return;
  }

  app.init();
});
