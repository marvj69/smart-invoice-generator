const CLOUD_SAVE_DEBOUNCE_MS = 900;
let cloudAuthUser = null;
let cloudSaveTimer = null;
let cloudSyncPaused = false;
let cloudLastSavedAt = null;
let authUiMode = 'login';

function setCloudSyncStatus(message, type = 'info') {
    const statusEl = document.getElementById('cloudSyncStatus');
    if (statusEl) {
        statusEl.textContent = message;
        statusEl.dataset.status = type;
    }

    const button = document.getElementById('accountStatusButton');
    if (button) {
        button.dataset.status = type;
        button.title = cloudAuthUser ? `Signed in as ${cloudAuthUser.email}` : 'Account';
    }
}

function apiErrorMessage(error, fallback) {
    if (error && error.message) return error.message;
    return fallback || 'Request failed.';
}

async function apiRequest(path, options = {}) {
    const headers = {
        Accept: 'application/json',
        ...(options.headers || {})
    };

    const request = {
        ...options,
        credentials: 'include',
        headers
    };

    if (options.body && typeof options.body === 'object' && !(options.body instanceof FormData)) {
        request.body = JSON.stringify(options.body);
        request.headers = {
            ...headers,
            'Content-Type': 'application/json'
        };
    }

    const response = await fetch(path, request);
    let payload = {};
    try {
        payload = await response.json();
    } catch (_error) {
        payload = {};
    }

    if (!response.ok) {
        const error = new Error(payload.error || `Request failed with status ${response.status}`);
        error.status = response.status;
        error.code = payload.code || '';
        throw error;
    }

    return payload;
}

function normalizeTemplateForCloud(template, index) {
    const source = template && typeof template === 'object' ? template : {};
    const numericId = Number(source.id);
    const rawName = String(source.name || 'Untitled Template').trim();
    const normalizedName = typeof normalizeTemplateNameValue === 'function'
        ? normalizeTemplateNameValue(rawName)
        : rawName;
    return {
        id: Number.isFinite(numericId) && numericId > 0 ? numericId : Date.now() + index,
        name: (normalizedName || 'Untitled Template').slice(0, 120) || 'Untitled Template',
        date: String(source.date || new Date().toLocaleDateString()).trim(),
        data: normalizeInvoiceData(source.data || {})
    };
}

function normalizeTemplatesForCloud(templates) {
    return (Array.isArray(templates) ? templates : [])
        .map(normalizeTemplateForCloud)
        .filter(template => template.name);
}

function numberValue(value) {
    const parsed = parseFloat(value);
    return Number.isFinite(parsed) ? parsed : 0;
}

function hasMeaningfulCurrentInvoice(data) {
    if (!data || typeof data !== 'object') return false;
    if (data.companyName || data.companyDetails || data.clientName || data.clientDetails || data.logo || data.notes) return true;
    if (numberValue(data.taxRate) > 0 || numberValue(data.discountValue) > 0) return true;
    if (!Array.isArray(data.items)) return false;

    return data.items.some(item => {
        if (!item || typeof item !== 'object') return false;
        return Boolean(
            String(item.address || '').trim() ||
            String(item.work || '').trim() ||
            String(item.description || '').trim() ||
            numberValue(item.rate) > 0
        );
    });
}

function hasMeaningfulCloudData(data) {
    if (!data || typeof data !== 'object') return false;
    const defaultCompany = normalizeDefaultCompanyProfile(data.defaultCompany || {});
    return Boolean(
        hasMeaningfulCurrentInvoice(data.currentInvoice) ||
        normalizeTemplatesForCloud(data.templates).length ||
        defaultCompany.companyName ||
        defaultCompany.companyDetails
    );
}

function writeDefaultCompanyLocal(profile) {
    defaultCompanyProfile = normalizeDefaultCompanyProfile(profile || {});

    if (defaultCompanyProfile.companyName || defaultCompanyProfile.companyDetails) {
        localStorage.setItem(DEFAULT_COMPANY_STORAGE_KEY, JSON.stringify(defaultCompanyProfile));
    } else {
        localStorage.removeItem(DEFAULT_COMPANY_STORAGE_KEY);
    }
    localStorage.removeItem(LEGACY_DEFAULT_BILLING_STORAGE_KEY);
    loadDefaultCompanySettings();
}

function writeTemplatesLocal(templates) {
    savedTemplates = normalizeTemplatesForCloud(templates);
    localStorage.setItem('invoiceTemplates', JSON.stringify(savedTemplates));
    const modal = document.getElementById('templateModal');
    if (modal && !modal.classList.contains('hidden')) {
        renderTemplatesList();
    }
}

function getLocalPersistenceSnapshot() {
    const wasPaused = cloudSyncPaused;
    cloudSyncPaused = true;
    try {
        if (typeof updateInvoice === 'function') {
            updateInvoice();
        }

        return {
            currentInvoice: normalizeInvoiceData(invoiceData || {}),
            templates: normalizeTemplatesForCloud(savedTemplates),
            defaultCompany: normalizeDefaultCompanyProfile(defaultCompanyProfile || {})
        };
    } finally {
        cloudSyncPaused = wasPaused;
    }
}

function applyCloudDataToLocal(data) {
    const wasPaused = cloudSyncPaused;
    cloudSyncPaused = true;
    try {
        const cloudData = data && typeof data === 'object' ? data : {};
        writeDefaultCompanyLocal(cloudData.defaultCompany || {});
        writeTemplatesLocal(cloudData.templates || []);

        if (hasMeaningfulCurrentInvoice(cloudData.currentInvoice)) {
            applyInvoiceDataToForm(cloudData.currentInvoice);
        }
    } finally {
        cloudSyncPaused = wasPaused;
    }
}

function formatCloudSaveTime(date) {
    if (!date) return '';
    return date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
}

function updateAccountUi() {
    const signedIn = Boolean(cloudAuthUser);
    const signedOutPanel = document.getElementById('authSignedOutPanel');
    const signedInPanel = document.getElementById('authSignedInPanel');
    const emailEl = document.getElementById('signedInEmail');
    const icon = document.getElementById('accountStatusIcon');

    if (signedOutPanel) signedOutPanel.classList.toggle('hidden', signedIn);
    if (signedInPanel) signedInPanel.classList.toggle('hidden', !signedIn);
    if (emailEl) emailEl.textContent = signedIn ? cloudAuthUser.email : '';
    if (icon) icon.className = signedIn ? 'fas fa-user-check' : 'fas fa-user';

    document.body.classList.toggle('cloud-authenticated', signedIn);

    if (!signedIn) {
        setCloudSyncStatus('Local only', 'info');
    } else if (cloudLastSavedAt) {
        setCloudSyncStatus(`Saved at ${formatCloudSaveTime(cloudLastSavedAt)}`, 'success');
    } else {
        setCloudSyncStatus('Signed in', 'success');
    }
}

function setAuthMode(mode) {
    authUiMode = mode === 'signup' ? 'signup' : 'login';
    const isSignup = authUiMode === 'signup';
    const title = document.getElementById('authModeTitle');
    const submit = document.getElementById('authSubmitButton');
    const password = document.getElementById('authPassword');
    const message = document.getElementById('authMessage');

    document.querySelectorAll('[data-auth-mode]').forEach(button => {
        const active = button.dataset.authMode === authUiMode;
        button.classList.toggle('bg-gray-900', active);
        button.classList.toggle('text-white', active);
        button.classList.toggle('bg-white', !active);
        button.classList.toggle('text-gray-700', !active);
    });

    if (title) title.textContent = isSignup ? 'Create Account' : 'Sign In';
    if (submit) submit.textContent = isSignup ? 'Create Account' : 'Sign In';
    if (password) password.autocomplete = isSignup ? 'new-password' : 'current-password';
    if (message) message.textContent = '';
}

function openAccountModal() {
    const modal = document.getElementById('accountModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    closeSettingsMenu();
    closeChatTemplateBubble();
    closeMobileActionSheet();

    requestAnimationFrame(() => {
        const email = document.getElementById('authEmail');
        if (!cloudAuthUser && email) email.focus();
    });
}

function closeAccountModal() {
    const modal = document.getElementById('accountModal');
    if (modal) modal.classList.add('hidden');
}

async function saveCloudData(options = {}) {
    if (!cloudAuthUser) return;
    if (cloudSaveTimer) {
        clearTimeout(cloudSaveTimer);
        cloudSaveTimer = null;
    }

    const payload = getLocalPersistenceSnapshot();
    setCloudSyncStatus('Saving...', 'saving');

    try {
        await apiRequest('/api/user-data', {
            method: 'PUT',
            body: payload
        });
        cloudLastSavedAt = new Date();
        updateAccountUi();
        if (options.announce) {
            showToast('Account data synced');
        }
    } catch (error) {
        setCloudSyncStatus(apiErrorMessage(error, 'Sync failed'), 'error');
        if (!options.silent) {
            showToast(apiErrorMessage(error, 'Could not sync account data'), 'error');
        }
        throw error;
    }
}

function scheduleCloudSave() {
    if (cloudSyncPaused || !cloudAuthUser) return;
    if (cloudSaveTimer) clearTimeout(cloudSaveTimer);
    cloudSaveTimer = setTimeout(() => {
        saveCloudData({ silent: true }).catch(() => {});
    }, CLOUD_SAVE_DEBOUNCE_MS);
}

async function loadCloudDataAfterAuth() {
    const localSnapshot = getLocalPersistenceSnapshot();
    setCloudSyncStatus('Loading...', 'saving');

    const payload = await apiRequest('/api/user-data');
    const remoteData = payload.data || {};

    if (hasMeaningfulCloudData(remoteData)) {
        applyCloudDataToLocal(remoteData);
        cloudLastSavedAt = remoteData.updatedAt ? new Date(remoteData.updatedAt) : null;
        updateAccountUi();
        showToast('Account data loaded');
        return;
    }

    applyCloudDataToLocal(localSnapshot);
    await saveCloudData({ silent: true });
    showToast('Local data saved to your account');
}

async function handleAuthSubmit(event) {
    event.preventDefault();
    const emailInput = document.getElementById('authEmail');
    const passwordInput = document.getElementById('authPassword');
    const message = document.getElementById('authMessage');
    const submit = document.getElementById('authSubmitButton');
    const email = emailInput ? emailInput.value.trim() : '';
    const password = passwordInput ? passwordInput.value : '';
    const endpoint = authUiMode === 'signup' ? '/api/auth/signup' : '/api/auth/login';

    if (message) message.textContent = '';
    if (submit) submit.disabled = true;
    setCloudSyncStatus(authUiMode === 'signup' ? 'Creating account...' : 'Signing in...', 'saving');

    try {
        const payload = await apiRequest(endpoint, {
            method: 'POST',
            body: { email, password }
        });
        cloudAuthUser = payload.user;
        if (passwordInput) passwordInput.value = '';
        updateAccountUi();
        await loadCloudDataAfterAuth();
    } catch (error) {
        const text = apiErrorMessage(error, authUiMode === 'signup' ? 'Could not create account.' : 'Could not sign in.');
        if (passwordInput) passwordInput.value = '';
        if (message) message.textContent = text;
        setCloudSyncStatus(text, 'error');
    } finally {
        if (submit) submit.disabled = false;
    }
}

async function logoutUser() {
    try {
        if (cloudAuthUser) {
            await saveCloudData({ silent: true }).catch(() => {});
        }
        await apiRequest('/api/auth/logout', { method: 'POST' });
        cloudAuthUser = null;
        cloudLastSavedAt = null;
        updateAccountUi();
        showToast('Signed out', 'info');
    } catch (error) {
        showToast(apiErrorMessage(error, 'Could not sign out'), 'error');
    }
}

async function syncNow() {
    if (!cloudAuthUser) {
        openAccountModal();
        return;
    }
    await saveCloudData({ announce: true }).catch(() => {});
}

async function loadCurrentAccount() {
    updateAccountUi();
    try {
        const payload = await apiRequest('/api/auth/me');
        cloudAuthUser = payload.user || null;
        updateAccountUi();
        if (cloudAuthUser) {
            await loadCloudDataAfterAuth();
        }
    } catch (error) {
        cloudAuthUser = null;
        updateAccountUi();
        if (error.code === 'DATABASE_NOT_CONFIGURED') {
            setCloudSyncStatus('Connect a Vercel database', 'error');
        } else {
            setCloudSyncStatus('Cloud sync unavailable', 'error');
        }
    }
}

function initAccountSync() {
    const form = document.getElementById('authForm');
    const modal = document.getElementById('accountModal');

    if (form) {
        form.addEventListener('submit', handleAuthSubmit);
    }

    if (modal) {
        modal.addEventListener('click', (event) => {
            if (event.target === modal) closeAccountModal();
        });
    }

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            closeAccountModal();
        }
    });

    setAuthMode('login');
    loadCurrentAccount();
}

document.addEventListener('DOMContentLoaded', initAccountSync);
