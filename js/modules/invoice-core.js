// State Management
        const PDF_WORKER_SRC = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        const GEMINI_API_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta/models';
        const GEMINI_DEFAULT_MODEL = 'gemini-3-flash-preview';
        const GEMINI_FALLBACK_MODELS = Object.freeze(['gemini-3-flash-preview-02-05', 'gemini-2.5-flash']);
        const GEMINI_REQUEST_TIMEOUT_MS = 90000;
        const GEMINI_THINKING_LEVEL_MINIMAL = 'minimal';
        const GEMINI_FALLBACK_THINKING_BUDGET = 1024;
        const GEMINI_API_KEY_STORAGE_KEY = 'invoice_get_gemini_api_key';
        const GEMINI_MODEL_STORAGE_KEY = 'invoice_get_gemini_model';
        const DEFAULT_COMPANY_STORAGE_KEY = 'invoice_get_default_company';
        const LEGACY_DEFAULT_BILLING_STORAGE_KEY = 'invoice_get_default_billing';
        const CHAT_TO_TEMPLATE_MAX_CHARS = 12000;
        const IMPORT_DEBUG_MAX_ENTRIES = 180;
        const MOBILE_PREVIEW_BREAKPOINT = 1024;
        const MM_TO_PX = 96 / 25.4;
        const PAPER_WIDTH_MM = 215.9;
        const PAPER_HEIGHT_MM = 279.4;
        const MOBILE_PREVIEW_HORIZONTAL_PADDING = 16;
        const MOBILE_PREVIEW_VERTICAL_PADDING = 44;
        const GEMINI_INVOICE_SCHEMA = Object.freeze({
            type: 'object',
            additionalProperties: false,
            required: [
                'documentType',
                'companyName',
                'companyDetails',
                'invoiceDate',
                'clientName',
                'clientDetails',
                'items',
                'taxRate',
                'discountType',
                'discountValue',
                'notes'
            ],
            properties: {
                documentType: {
                    type: 'string',
                    enum: ['Invoice', 'Bid']
                },
                companyName: {
                    type: 'string'
                },
                companyDetails: {
                    type: 'string',
                    description: 'Address block format: line 1 street number + street name, line 2 city + state + ZIP, line 3 phone (if present)'
                },
                invoiceDate: {
                    type: 'string',
                    description: 'Use YYYY-MM-DD if known, otherwise empty string'
                },
                clientName: {
                    type: 'string'
                },
                clientDetails: {
                    type: 'string',
                    description: 'Address block format: line 1 street number + street name, line 2 city + state + ZIP, line 3 phone (if present)'
                },
                items: {
                    type: 'array',
                    items: {
                        type: 'object',
                        additionalProperties: false,
                        required: ['description', 'address', 'work', 'quantity', 'rate'],
                        properties: {
                            description: { type: 'string' },
                            address: {
                                type: 'string',
                                description: 'Property address block: line 1 street number + street name, line 2 city + state + ZIP'
                            },
                            work: { type: 'string' },
                            quantity: { type: 'number' },
                            rate: { type: 'number' }
                        }
                    }
                },
                taxRate: {
                    type: 'number'
                },
                discountType: {
                    type: 'string',
                    enum: ['fixed', 'percentage']
                },
                discountValue: {
                    type: 'number'
                },
                notes: {
                    type: 'string'
                }
            }
        });
        const PDF_PAYLOAD_MARKERS = Object.freeze({
            start: 'INVGET_PAYLOAD_BEGIN',
            end: 'INVGET_PAYLOAD_END',
            chunkPrefix: 'INVGET_PAYLOAD_CHUNK_'
        });
        let toastTimerId = null;
        let importDebugEntries = [];
        if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
        }

        function normalizeDefaultCompanyProfile(data) {
            const source = data && typeof data === 'object' ? data : {};
            return {
                companyName: String(source.companyName || source.clientName || source.name || '').trim(),
                companyDetails: normalizeAddressBlock(source.companyDetails || source.clientDetails || source.details || source.address || '', { includePhone: true })
            };
        }

        function loadDefaultCompanyProfileFromStorage() {
            const raw = localStorage.getItem(DEFAULT_COMPANY_STORAGE_KEY) || localStorage.getItem(LEGACY_DEFAULT_BILLING_STORAGE_KEY);
            if (!raw) return normalizeDefaultCompanyProfile({});

            try {
                return normalizeDefaultCompanyProfile(JSON.parse(raw));
            } catch (error) {
                console.warn('Failed to parse default company profile', error);
                return normalizeDefaultCompanyProfile({});
            }
        }

        function getDefaultCompanyProfile() {
            defaultCompanyProfile = normalizeDefaultCompanyProfile(defaultCompanyProfile);
            return defaultCompanyProfile;
        }

        function applyDefaultCompanyFallback(data, sourceLabel = '') {
            const normalized = normalizeInvoiceData(data);
            const defaults = getDefaultCompanyProfile();
            const hasDefaultName = Boolean(normalizeSpace(defaults.companyName));
            const hasDefaultDetails = Boolean(normalizeSpace(defaults.companyDetails));

            if (!hasDefaultName && !hasDefaultDetails) {
                return normalized;
            }

            const nameMissing = !normalizeSpace(normalized.companyName);
            const detailsMissing = !normalizeSpace(normalized.companyDetails);
            let applied = false;

            if (nameMissing && hasDefaultName) {
                normalized.companyName = defaults.companyName;
                applied = true;
            }

            if (detailsMissing && hasDefaultDetails) {
                normalized.companyDetails = defaults.companyDetails;
                applied = true;
            }

            if (applied) {
                appendImportDebug('Applied default company fallback', {
                    source: sourceLabel || 'unknown',
                    filledName: nameMissing && hasDefaultName,
                    filledDetails: detailsMissing && hasDefaultDetails
                });
            }

            return normalized;
        }

        function loadDefaultCompanySettings() {
            defaultCompanyProfile = loadDefaultCompanyProfileFromStorage();
            const nameInput = document.getElementById('defaultCompanyName');
            const detailsInput = document.getElementById('defaultCompanyDetails');
            if (nameInput) nameInput.value = defaultCompanyProfile.companyName;
            if (detailsInput) detailsInput.value = defaultCompanyProfile.companyDetails;
        }

        function saveDefaultCompanySettings(showConfirmation = false) {
            const nameInput = document.getElementById('defaultCompanyName');
            const detailsInput = document.getElementById('defaultCompanyDetails');
            const normalizedProfile = normalizeDefaultCompanyProfile({
                companyName: nameInput ? nameInput.value : '',
                companyDetails: detailsInput ? detailsInput.value : ''
            });

            defaultCompanyProfile = normalizedProfile;

            if (normalizedProfile.companyName || normalizedProfile.companyDetails) {
                localStorage.setItem(DEFAULT_COMPANY_STORAGE_KEY, JSON.stringify(normalizedProfile));
            } else {
                localStorage.removeItem(DEFAULT_COMPANY_STORAGE_KEY);
            }
            localStorage.removeItem(LEGACY_DEFAULT_BILLING_STORAGE_KEY);

            if (nameInput) nameInput.value = normalizedProfile.companyName;
            if (detailsInput) detailsInput.value = normalizedProfile.companyDetails;

            if (showConfirmation) {
                showToast('Default company saved');
            }
        }

        function clearDefaultCompanySettings() {
            defaultCompanyProfile = normalizeDefaultCompanyProfile({});
            localStorage.removeItem(DEFAULT_COMPANY_STORAGE_KEY);
            localStorage.removeItem(LEGACY_DEFAULT_BILLING_STORAGE_KEY);
            const nameInput = document.getElementById('defaultCompanyName');
            const detailsInput = document.getElementById('defaultCompanyDetails');
            if (nameInput) nameInput.value = '';
            if (detailsInput) detailsInput.value = '';
            showToast('Default company cleared', 'info');
        }

        let defaultCompanyProfile = loadDefaultCompanyProfileFromStorage();

        function getDefaultLineItem() {
            return { description: '', quantity: 1, rate: 0, address: '', work: '' };
        }

        function getDefaultInvoiceData() {
            const defaultCompany = getDefaultCompanyProfile();
            return {
                documentType: 'Invoice',
                companyName: defaultCompany.companyName,
                companyDetails: defaultCompany.companyDetails,
                logo: null,
                invoiceDate: new Date().toISOString().split('T')[0],
                clientName: '',
                clientDetails: '',
                items: [getDefaultLineItem()],
                taxRate: 0,
                discountType: 'fixed',
                discountValue: 0,
                notes: ''
            };
        }

        let invoiceData = getDefaultInvoiceData();

        function migrateV1TemplatesIfNeeded() {
            const raw = localStorage.getItem('invoiceTemplates');
            if (!raw) {
                return [];
            }

            let parsed;
            try {
                parsed = JSON.parse(raw);
            } catch (error) {
                console.warn('Failed to parse stored templates', error);
                return [];
            }

            if (Array.isArray(parsed)) {
                return parsed;
            }

            if (!parsed || typeof parsed !== 'object') {
                return [];
            }

            const entries = Object.entries(parsed);
            if (entries.length === 0) {
                return [];
            }

            const parser = new DOMParser();
            const text = (el) => (el && el.textContent ? el.textContent.trim() : '');
            const value = (el) => {
                if (!el) return '';
                const direct = (el.value || '').trim();
                if (direct) return direct;
                const attr = el.getAttribute('value');
                return attr ? attr.trim() : '';
            };
            const number = (input) => {
                const cleaned = String(input || '').replace(/[^0-9.\-]/g, '');
                const parsedNumber = parseFloat(cleaned);
                return Number.isFinite(parsedNumber) ? parsedNumber : 0;
            };

            const baseId = Date.now();
            const migrated = entries.map(([name, pageHTML], index) => {
                const doc = parser.parseFromString(String(pageHTML || ''), 'text/html');
                const scope = doc.body || doc;

                const companyNodes = scope.querySelectorAll('.company-details [contenteditable="true"]');
                const clientNodes = scope.querySelectorAll('.billing-details [contenteditable="true"]');

                const companyName = text(companyNodes[0]);
                const companyDetails = Array.from(companyNodes)
                    .slice(1)
                    .map(text)
                    .filter(Boolean)
                    .join('\n');

                const clientName = text(clientNodes[0]);
                const clientDetails = Array.from(clientNodes)
                    .slice(1)
                    .map(text)
                    .filter(Boolean)
                    .join('\n');

                const items = Array.from(scope.querySelectorAll('.items-table tbody tr'))
                    .map(row => {
                        const description = text(row.querySelector('.description'));
                        const quantity = number(text(row.querySelector('.quantity')));
                        const rate = number(text(row.querySelector('.price')));
                        return { description, quantity: quantity || 0, rate: rate || 0 };
                    })
                    .filter(item => item.description || item.quantity || item.rate);

                const invoiceDate = value(scope.querySelector('#invoice-date'));
                const docTypeNode = scope.querySelector('#invoice-type');
                const docTypeRaw = docTypeNode ? (docTypeNode.value || docTypeNode.textContent || '').trim() : '';
                const documentType = docTypeRaw.toLowerCase() === 'bid' ? 'Bid' : 'Invoice';
                const taxRate = number(text(scope.querySelector('#tax-rate')));
                const notes = text(scope.querySelector('.invoice-footer [contenteditable="true"]'));

                return {
                    id: baseId + index,
                    name,
                    date: new Date().toLocaleDateString(),
                    data: {
                        companyName,
                        companyDetails,
                        logo: null,
                        documentType,
                        invoiceDate,
                        clientName,
                        clientDetails,
                        items: items.length ? items : [{ description: '', quantity: 1, rate: 0 }],
                        taxRate,
                        discountType: 'fixed',
                        discountValue: 0,
                        notes
                    }
                };
            });

            try {
                localStorage.setItem('invoiceTemplates_v1_backup', raw);
                localStorage.setItem('invoiceTemplates', JSON.stringify(migrated));
            } catch (error) {
                console.warn('Failed to persist migrated templates', error);
            }

            return migrated;
        }

        let savedTemplates = migrateV1TemplatesIfNeeded();
        let updateScheduled = false;
        let mobilePreviewOpen = false;

        function scheduleUpdate() {
            if (updateScheduled) return;
            updateScheduled = true;
            requestAnimationFrame(() => {
                updateScheduled = false;
                updateInvoice();
            });
        }

        function stringifyDebugValue(value) {
            if (value === undefined || value === null) return '';
            if (typeof value === 'string') return value;
            try {
                return JSON.stringify(value);
            } catch (error) {
                return String(value);
            }
        }

        function renderImportDebugLog() {
            const logEl = document.getElementById('importDebugLog');
            if (!logEl) return;
            logEl.value = importDebugEntries.join('\n');
            logEl.scrollTop = logEl.scrollHeight;
        }

        function appendImportDebug(message, details = '') {
            const timestamp = new Date().toLocaleTimeString('en-US', { hour12: false });
            const base = normalizeSpace(message) || 'Debug event';
            const detailText = normalizeSpace(stringifyDebugValue(details));
            const entry = detailText ? `[${timestamp}] ${base} | ${detailText}` : `[${timestamp}] ${base}`;
            importDebugEntries.push(entry);
            if (importDebugEntries.length > IMPORT_DEBUG_MAX_ENTRIES) {
                importDebugEntries = importDebugEntries.slice(-IMPORT_DEBUG_MAX_ENTRIES);
            }
            renderImportDebugLog();
            console.info('[ImportDebug]', entry);
        }

        function clearImportDebugLog() {
            importDebugEntries = [];
            renderImportDebugLog();
            showToast('Import debug log cleared', 'info');
        }

        function getConfiguredGeminiApiKey() {
            const input = document.getElementById('geminiApiKey');
            const valueFromInput = normalizeSpace(input ? input.value : '');
            if (valueFromInput) return valueFromInput;
            return normalizeSpace(localStorage.getItem(GEMINI_API_KEY_STORAGE_KEY) || '');
        }

        function getConfiguredGeminiModel() {
            const input = document.getElementById('geminiModel');
            const valueFromInput = normalizeSpace(input ? input.value : '');
            if (valueFromInput) return valueFromInput;
            return normalizeSpace(localStorage.getItem(GEMINI_MODEL_STORAGE_KEY) || GEMINI_DEFAULT_MODEL) || GEMINI_DEFAULT_MODEL;
        }

        function loadGeminiSettings() {
            const savedKey = normalizeSpace(localStorage.getItem(GEMINI_API_KEY_STORAGE_KEY) || '');
            const savedModelRaw = normalizeSpace(localStorage.getItem(GEMINI_MODEL_STORAGE_KEY) || '');
            let savedModel = savedModelRaw || GEMINI_DEFAULT_MODEL;
            if (savedModel === 'gemini-3-flash-preview-02-05' || savedModel === 'gemini-2.5-flash-lite') {
                savedModel = GEMINI_DEFAULT_MODEL;
                localStorage.setItem(GEMINI_MODEL_STORAGE_KEY, savedModel);
            }
            const keyInput = document.getElementById('geminiApiKey');
            const modelInput = document.getElementById('geminiModel');

            if (keyInput) keyInput.value = savedKey;
            if (modelInput) modelInput.value = savedModel;
            appendImportDebug('Gemini settings loaded', {
                model: savedModel,
                hasApiKey: Boolean(savedKey)
            });
        }

        function saveGeminiSettings(showConfirmation = false) {
            const apiKey = getConfiguredGeminiApiKey();
            const model = getConfiguredGeminiModel();
            localStorage.setItem(GEMINI_API_KEY_STORAGE_KEY, apiKey);
            localStorage.setItem(GEMINI_MODEL_STORAGE_KEY, model);
            const keyInput = document.getElementById('geminiApiKey');
            const modelInput = document.getElementById('geminiModel');
            if (keyInput) keyInput.value = apiKey;
            if (modelInput) modelInput.value = model;
            appendImportDebug('Gemini settings saved', { model, hasApiKey: Boolean(apiKey) });
            if (showConfirmation) {
                showToast('Gemini settings saved');
            }
        }

        function clearGeminiSettings() {
            localStorage.removeItem(GEMINI_API_KEY_STORAGE_KEY);
            localStorage.removeItem(GEMINI_MODEL_STORAGE_KEY);
            const keyInput = document.getElementById('geminiApiKey');
            const modelInput = document.getElementById('geminiModel');
            if (keyInput) keyInput.value = '';
            if (modelInput) modelInput.value = GEMINI_DEFAULT_MODEL;
            appendImportDebug('Gemini settings cleared');
            showToast('Gemini settings cleared', 'info');
        }

        function closeSettingsMenu() {
            const menu = document.getElementById('settingsMenu');
            const button = document.getElementById('settingsMenuButton');
            if (menu) menu.classList.add('hidden');
            if (button) button.setAttribute('aria-expanded', 'false');
        }

        function toggleSettingsMenu(event) {
            if (event) event.stopPropagation();
            const menu = document.getElementById('settingsMenu');
            const button = document.getElementById('settingsMenuButton');
            if (!menu) return;
            const willOpen = menu.classList.contains('hidden');
            menu.classList.toggle('hidden', !willOpen);
            if (button) button.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
            if (willOpen) {
                closeChatTemplateBubble();
            }
        }

        function isMobilePreviewViewport() {
            return window.innerWidth < MOBILE_PREVIEW_BREAKPOINT;
        }

        function isStandalonePwa() {
            return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
        }

        function updateAppViewportHeight() {
            const viewportHeight = window.visualViewport ? window.visualViewport.height : window.innerHeight;
            if (!viewportHeight) return;
            document.documentElement.style.setProperty('--app-height', `${Math.round(viewportHeight)}px`);
        }

        function preventNativePinchZoom(event) {
            if (event.touches && event.touches.length > 1) {
                event.preventDefault();
            }
        }

        function setupMobilePwaExperience() {
            updateAppViewportHeight();
            document.body.classList.toggle('pwa-standalone', isStandalonePwa());

            if (window.visualViewport && typeof window.visualViewport.addEventListener === 'function') {
                window.visualViewport.addEventListener('resize', updateAppViewportHeight);
            }
            window.addEventListener('resize', updateAppViewportHeight, { passive: true });
            window.addEventListener('orientationchange', updateAppViewportHeight, { passive: true });

            document.addEventListener('gesturestart', (event) => {
                event.preventDefault();
            }, { passive: false });

            document.addEventListener('touchmove', preventNativePinchZoom, { passive: false });
        }

        function resetMobilePreviewScale() {
            const preview = document.getElementById('invoice-preview');
            const wrap = document.getElementById('mobilePreviewScaleWrap');
            if (preview) {
                preview.style.transform = '';
                preview.style.transformOrigin = '';
            }
            if (wrap) {
                wrap.style.width = '';
                wrap.style.height = '';
                wrap.style.minHeight = '';
            }
        }

        function applyMobilePreviewScale() {
            const preview = document.getElementById('invoice-preview');
            const wrap = document.getElementById('mobilePreviewScaleWrap');
            if (!preview || !wrap) return;

            if (!mobilePreviewOpen || !isMobilePreviewViewport()) {
                resetMobilePreviewScale();
                return;
            }

            const baseWidthPx = PAPER_WIDTH_MM * MM_TO_PX;
            const baseHeightPx = PAPER_HEIGHT_MM * MM_TO_PX;
            const availableWidth = Math.max(240, window.innerWidth - MOBILE_PREVIEW_HORIZONTAL_PADDING);
            const availableHeight = Math.max(280, window.innerHeight - MOBILE_PREVIEW_VERTICAL_PADDING);
            const scale = Math.min(1, availableWidth / baseWidthPx, availableHeight / baseHeightPx);
            const scaledWidth = baseWidthPx * scale;
            const scaledHeight = baseHeightPx * scale;

            preview.style.transform = `scale(${scale})`;
            preview.style.transformOrigin = 'top left';
            wrap.style.width = `${scaledWidth}px`;
            wrap.style.height = `${scaledHeight}px`;
            wrap.style.minHeight = `${scaledHeight}px`;
        }

        function openMobilePreview() {
            if (!isMobilePreviewViewport()) return;
            updateInvoice();
            closeSettingsMenu();
            closeChatTemplateBubble();
            mobilePreviewOpen = true;
            document.body.classList.add('mobile-preview-open');
            requestAnimationFrame(applyMobilePreviewScale);
        }

        function closeMobilePreview() {
            mobilePreviewOpen = false;
            document.body.classList.remove('mobile-preview-open');
            resetMobilePreviewScale();
        }

        // Initialize
        document.addEventListener('DOMContentLoaded', () => {
            setupMobilePwaExperience();
            loadGeminiSettings();
            loadDefaultCompanySettings();
            applyInvoiceDataToForm(invoiceData);
        });

        // Update Preview
        function updateInvoice() {
            // Get values from inputs
            invoiceData.companyName = document.getElementById('companyName').value;
            invoiceData.companyDetails = document.getElementById('companyDetails').value;
            invoiceData.documentType = document.getElementById('documentType').value;
            invoiceData.invoiceDate = document.getElementById('invoiceDate').value;
            invoiceData.clientName = document.getElementById('clientName').value;
            invoiceData.clientDetails = document.getElementById('clientDetails').value;
            invoiceData.taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
            invoiceData.discountType = document.getElementById('discountType').value;
            invoiceData.discountValue = parseFloat(document.getElementById('discountValue').value) || 0;
            invoiceData.notes = document.getElementById('notes').value;

            // Update items from DOM
            const itemRows = document.querySelectorAll('.line-item');
            invoiceData.items = Array.from(itemRows).map(row => {
                const addressInput = row.querySelector('.item-address');
                const workInput = row.querySelector('.item-work');
                const descInput = row.querySelector('.item-desc');
                const address = addressInput ? addressInput.value : '';
                const work = workInput ? workInput.value : '';
                const fallbackDescription = descInput ? descInput.value : '';
                const descriptionParts = [];

                if (address.trim()) descriptionParts.push(address.trim());
                if (work.trim()) descriptionParts.push(work.trim());
                if (!descriptionParts.length && fallbackDescription.trim()) {
                    descriptionParts.push(fallbackDescription.trim());
                }

                return {
                    description: descriptionParts.join('\n'),
                    address,
                    work,
                    quantity: parseFloat(row.querySelector('.item-qty').value) || 0,
                    rate: parseFloat(row.querySelector('.item-rate').value) || 0
                };
            });

            // Render Preview
            renderPreview();
            calculateTotals();
        }

        function renderPreview() {
            // Basic Info
            document.getElementById('previewCompanyName').textContent = invoiceData.companyName || 'Your Company';
            document.getElementById('previewCompanyDetails').textContent = invoiceData.companyDetails;
            document.getElementById('previewDocumentType').textContent = (invoiceData.documentType || 'Invoice').toUpperCase();
            document.getElementById('previewInvoiceDate').textContent = formatDate(invoiceData.invoiceDate);
            document.getElementById('previewClientName').textContent = invoiceData.clientName || 'Client Name';
            document.getElementById('previewClientDetails').textContent = invoiceData.clientDetails;

            // Logo
            const logoContainer = document.getElementById('previewLogo');
            if (invoiceData.logo) {
                logoContainer.innerHTML = `<img src="${invoiceData.logo}" class="max-w-[200px] max-h-[100px] object-contain">`;
                logoContainer.classList.remove('hidden');
            } else {
                logoContainer.classList.add('hidden');
            }

            // Items
            const tbody = document.getElementById('previewItemsList');
            tbody.innerHTML = '';
            
            invoiceData.items.forEach((item, index) => {
                if (item.description || item.quantity || item.rate) {
                    const amount = item.quantity * item.rate;
                    const row = document.createElement('tr');
                    row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                    const descriptionHtml = formatDescriptionText(item.description);
                    row.innerHTML = `
                        <td class="py-3 pr-4 text-gray-900">${descriptionHtml}</td>
                        <td class="py-3 px-4 text-right text-gray-600">${item.quantity}</td>
                        <td class="py-3 px-4 text-right text-gray-600">$${formatMoney(item.rate)}</td>
                        <td class="py-3 pl-4 text-right font-medium">$${formatMoney(amount)}</td>
                    `;
                    tbody.appendChild(row);
                }
            });

            // Notes
            const notesSection = document.getElementById('previewNotesSection');
            if (invoiceData.notes) {
                document.getElementById('previewNotes').textContent = invoiceData.notes;
                notesSection.classList.remove('hidden');
            } else {
                notesSection.classList.add('hidden');
            }
        }

        function calculateTotals() {
            let subtotal = invoiceData.items.reduce((sum, item) => {
                return sum + (item.quantity * item.rate);
            }, 0);

            let taxAmount = subtotal * (invoiceData.taxRate / 100);
            
            let discountAmount = 0;
            if (invoiceData.discountType === 'percentage') {
                discountAmount = subtotal * (invoiceData.discountValue / 100);
            } else {
                discountAmount = invoiceData.discountValue;
            }

            let total = subtotal + taxAmount - discountAmount;

            // Update Preview
            document.getElementById('previewSubtotal').textContent = '$' + formatMoney(subtotal);
            document.getElementById('previewTaxRate').textContent = invoiceData.taxRate;
            document.getElementById('previewTaxAmount').textContent = '$' + formatMoney(taxAmount);
            document.getElementById('previewTotal').textContent = '$' + formatMoney(total);
            
            const discountRow = document.getElementById('previewDiscountRow');
            if (discountAmount > 0) {
                document.getElementById('previewDiscountAmount').textContent = '-$' + formatMoney(discountAmount);
                discountRow.style.display = 'flex';
            } else {
                discountRow.style.display = 'none';
            }
        }

        // Line Items Management
        function addLineItem(itemData = null, shouldRefresh = true) {
            const container = document.getElementById('lineItemsContainer');
            const normalizedItem = normalizeLineItem(itemData || getDefaultLineItem());

            const div = document.createElement('div');
            div.className = 'line-item bg-gray-50 p-3 rounded-lg border border-gray-200 space-y-2';
            div.innerHTML = `
                <div>
                    <label class="block text-xs font-medium text-gray-600 mb-1">Property Address</label>
                    <input type="text" placeholder="e.g., 885 County Rd CKL, Champion, MI 49814" 
                        class="item-address w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500"
                        oninput="scheduleUpdate()">
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-600 mb-1">Work Done</label>
                    <textarea rows="2" placeholder="e.g., Install handrail for front steps" 
                        class="item-work w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 resize-none"
                        oninput="scheduleUpdate()"></textarea>
                </div>
                <div class="flex gap-2">
                    <input type="number" placeholder="Qty" min="0" step="1" value="1"
                        class="item-qty w-20 px-2 py-1 border border-gray-300 rounded text-sm text-right"
                        oninput="scheduleUpdate()">
                    <input type="number" placeholder="Rate" min="0" step="0.01" value="0"
                        class="item-rate flex-1 px-2 py-1 border border-gray-300 rounded text-sm text-right"
                        oninput="scheduleUpdate()">
                    <button onclick="removeLineItem(this)" class="px-2 text-red-500 hover:text-red-700">
                        <i class="fas fa-trash-alt"></i>
                    </button>
                </div>
            `;

            div.querySelector('.item-address').value = normalizedItem.address;
            div.querySelector('.item-work').value = normalizedItem.work;
            div.querySelector('.item-qty').value = normalizedItem.quantity;
            div.querySelector('.item-rate').value = normalizedItem.rate;

            container.appendChild(div);
            if (shouldRefresh) {
                updateInvoice();
            }
        }

        function removeLineItem(btn) {
            const items = document.querySelectorAll('.line-item');
            if (items.length > 1) {
                btn.closest('.line-item').remove();
                updateInvoice();
            } else {
                showToast('At least one item is required', 'error');
            }
        }

        // Logo Handling
        function handleLogoUpload(input) {
            if (input.files && input.files[0]) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    invoiceData.logo = e.target.result;
                    updateInvoice();
                    showToast('Logo uploaded successfully');
                };
                reader.readAsDataURL(input.files[0]);
            }
        }

        // PDF Generation
        function formatFilenameDate(dateString) {
            const raw = String(dateString || '').trim();
            if (!raw) return '';

            const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (isoMatch) {
                return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
            }

            const parsed = new Date(raw);
            if (Number.isNaN(parsed.getTime())) {
                return '';
            }

            const year = parsed.getFullYear();
            const month = String(parsed.getMonth() + 1).padStart(2, '0');
            const day = String(parsed.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        function normalizeFilenamePart(value, maxLength = 48) {
            let cleaned = String(value || '')
                .normalize('NFKD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/[\r\n]+/g, ' ')
                .replace(/[\\/:*?"<>|]+/g, ' ')
                .replace(/[^\w\s&'.,\-()]+/g, ' ')
                .replace(/\s+/g, ' ')
                .trim()
                .replace(/^[.\-_\s]+|[.\-_\s]+$/g, '');

            if (!cleaned) return '';
            if (cleaned.length <= maxLength) return cleaned;

            const slice = cleaned.slice(0, maxLength).trim();
            const lastSpace = slice.lastIndexOf(' ');
            if (lastSpace > 12) {
                cleaned = slice.slice(0, lastSpace);
            } else {
                cleaned = slice;
            }

            return cleaned.replace(/[.\-_\s]+$/g, '');
        }

        function extractCityFromText(text) {
            const raw = String(text || '').replace(/\s+/g, ' ').trim();
            if (!raw) return '';

            const commaFormat = raw.match(/,\s*([^,]+?)\s*,\s*[A-Z]{2}\b/);
            if (commaFormat) {
                return commaFormat[1].trim();
            }

            const stateZipFormat = raw.match(/(?:^|,\s*)([A-Za-z][A-Za-z .'-]*?)\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/);
            if (stateZipFormat) {
                return stateZipFormat[1].trim();
            }

            return '';
        }

        function getCityForFilename(data) {
            if (Array.isArray(data.items)) {
                for (const item of data.items) {
                    if (!item || typeof item !== 'object') continue;

                    const address = String(item.address || '').trim();
                    const parsed = parseDescriptionFields(item.description || '');
                    const addressFallback = String(parsed.address || '').trim();
                    const city = extractCityFromText(address || addressFallback);
                    if (city) return city;
                }
            }

            const clientDetailsLines = String(data.clientDetails || '')
                .split(/\r?\n/)
                .map(line => line.trim())
                .filter(Boolean);

            for (const line of clientDetailsLines) {
                const city = extractCityFromText(line);
                if (city) return city;
            }

            return '';
        }

        function getDescriptorForFilename(data) {
            const candidates = [];

            if (Array.isArray(data.items)) {
                for (const item of data.items) {
                    if (!item || typeof item !== 'object') continue;
                    const work = String(item.work || '').trim();
                    if (work) candidates.push(work);

                    if (!work) {
                        const parsed = parseDescriptionFields(item.description || '');
                        if (parsed.work) candidates.push(parsed.work);
                    }
                }
            }

            candidates.push(String(data.documentType || '').trim());

            const stopWords = new Set(['the', 'and', 'for', 'with', 'from', 'to', 'of', 'a', 'an', 'at', 'on', 'in']);

            for (const candidate of candidates) {
                const words = normalizeFilenamePart(candidate, 80)
                    .split(/\s+/)
                    .map(word => word.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, ''))
                    .filter(Boolean)
                    .filter(word => !stopWords.has(word.toLowerCase()));

                if (words.length >= 2) {
                    return `${words[0]} ${words[1]}`;
                }

                if (words.length === 1) {
                    const docWord = normalizeFilenamePart(data.documentType || 'Work', 20).split(/\s+/).find(Boolean) || 'Work';
                    if (words[0].toLowerCase() !== docWord.toLowerCase()) {
                        return `${words[0]} ${docWord}`;
                    }
                    return `${words[0]} Work`;
                }
            }

            return 'Service Work';
        }

        function buildPdfFileName(data) {
            const descriptorRaw = normalizeFilenamePart(getDescriptorForFilename(data), 36) || 'Service Work';
            const descriptorWords = descriptorRaw.split(/\s+/).filter(Boolean);
            const descriptor = descriptorWords.length >= 2
                ? `${descriptorWords[0]} ${descriptorWords[1]}`
                : `${descriptorWords[0] || 'Service'} Work`;

            const city = normalizeFilenamePart(getCityForFilename(data), 28) || 'Unknown City';
            const datePart = formatFilenameDate(data.invoiceDate || new Date().toISOString().split('T')[0]);
            const safeDate = datePart || formatFilenameDate(new Date().toISOString()) || String(Date.now());
            let baseName = `${descriptor} - ${city} - ${safeDate}`.replace(/\s{2,}/g, ' ').trim();

            return `${baseName}.pdf`;
        }

        function encodeBase64Url(value) {
            const raw = String(value || '');
            if (!raw) return '';

            try {
                const bytes = new TextEncoder().encode(raw);
                let binary = '';
                bytes.forEach(byte => {
                    binary += String.fromCharCode(byte);
                });

                return btoa(binary)
                    .replace(/\+/g, '-')
                    .replace(/\//g, '_')
                    .replace(/=+$/g, '');
            } catch (error) {
                console.warn('Failed to encode embedded invoice payload', error);
                return '';
            }
        }

        function decodeBase64Url(value) {
            const cleaned = String(value || '').replace(/[^A-Za-z0-9_-]/g, '');
            if (!cleaned) return '';

            try {
                let padded = cleaned
                    .replace(/-/g, '+')
                    .replace(/_/g, '/');
                while (padded.length % 4 !== 0) {
                    padded += '=';
                }

                const binary = atob(padded);
                const bytes = new Uint8Array(binary.length);
                for (let index = 0; index < binary.length; index += 1) {
                    bytes[index] = binary.charCodeAt(index);
                }

                return new TextDecoder().decode(bytes);
            } catch (error) {
                console.warn('Failed to decode embedded invoice payload', error);
                return '';
            }
        }

        function createEmbeddedInvoicePayloadText(data) {
            const normalizedData = normalizeInvoiceData(data);
            const encodedPayload = encodeBase64Url(JSON.stringify(normalizedData));
            if (!encodedPayload) return '';

            const chunks = encodedPayload.match(/.{1,180}/g) || [];
            const lines = [PDF_PAYLOAD_MARKERS.start];
            chunks.forEach((chunk, index) => {
                lines.push(`${PDF_PAYLOAD_MARKERS.chunkPrefix}${String(index + 1).padStart(3, '0')}:${chunk}`);
            });
            lines.push(PDF_PAYLOAD_MARKERS.end);
            return lines.join('\n');
        }

        function extractEmbeddedInvoicePayloadFromText(rawText) {
            const text = String(rawText || '');
            if (!text.includes(PDF_PAYLOAD_MARKERS.start)) return null;
            if (!text.includes(PDF_PAYLOAD_MARKERS.end)) return null;

            const startIndex = text.indexOf(PDF_PAYLOAD_MARKERS.start);
            const endIndex = text.indexOf(PDF_PAYLOAD_MARKERS.end, startIndex);
            if (endIndex < 0 || endIndex <= startIndex) return null;

            const payloadBlock = text.slice(startIndex, endIndex);
            const chunkPattern = new RegExp(`^${PDF_PAYLOAD_MARKERS.chunkPrefix}(\\d{3})\\s*:\\s*([A-Za-z0-9_\\-\\s]+)$`);
            const chunks = [];

            const lines = payloadBlock
                .split(/\r?\n/)
                .map(normalizeSpace)
                .filter(Boolean);

            lines.forEach(line => {
                const match = line.match(chunkPattern);
                if (!match) return;
                const order = Number(match[1]);
                const chunk = String(match[2] || '').replace(/[^A-Za-z0-9_-]/g, '');
                if (!chunk) return;
                chunks.push({ order, chunk });
            });

            if (!chunks.length) return null;

            chunks.sort((a, b) => a.order - b.order);
            const decoded = decodeBase64Url(chunks.map(entry => entry.chunk).join(''));
            if (!decoded) return null;

            try {
                const parsed = JSON.parse(decoded);
                const invoiceLikeData = extractInvoiceDataFromJson(parsed);
                if (!invoiceLikeData) return null;
                return normalizeInvoiceData(invoiceLikeData);
            } catch (error) {
                console.warn('Could not parse embedded payload from PDF', error);
                return null;
            }
        }

        function embedPayloadTextInPdf(pdf, payloadText) {
            if (!pdf || typeof pdf.text !== 'function') return;

            const lines = String(payloadText || '')
                .split('\n')
                .map(line => line.trim())
                .filter(Boolean);
            if (!lines.length) return;

            const lastPage = typeof pdf.getNumberOfPages === 'function'
                ? Math.max(1, pdf.getNumberOfPages())
                : 1;
            if (typeof pdf.setPage === 'function') {
                pdf.setPage(lastPage);
            }

            const pageHeight = pdf.internal && pdf.internal.pageSize && typeof pdf.internal.pageSize.getHeight === 'function'
                ? pdf.internal.pageSize.getHeight()
                : 297;
            const startY = Math.max(2, pageHeight - Math.max(5, lines.length * 0.8));
            const previousSize = typeof pdf.getFontSize === 'function' ? pdf.getFontSize() : null;

            try {
                if (typeof pdf.setFontSize === 'function') {
                    pdf.setFontSize(1);
                }
                if (typeof pdf.setTextColor === 'function') {
                    pdf.setTextColor(255, 255, 255);
                }
                pdf.text(lines, 1, startY, { lineHeightFactor: 0.8 });
            } catch (error) {
                console.warn('Failed to embed import payload into PDF', error);
            } finally {
                if (Number.isFinite(previousSize) && typeof pdf.setFontSize === 'function') {
                    pdf.setFontSize(previousSize);
                }
                if (typeof pdf.setTextColor === 'function') {
                    pdf.setTextColor(0, 0, 0);
                }
            }
        }

        function generatePDF() {
            updateInvoice();
            const element = document.getElementById('invoice-preview');
            const fileName = buildPdfFileName(invoiceData);
            const embeddedPayloadText = createEmbeddedInvoicePayloadText(invoiceData);
            showToast('Generating PDF...', 'info');

            const exportNode = element.cloneNode(true);
            exportNode.style.minHeight = 'auto';
            exportNode.style.height = 'auto';
            exportNode.style.boxShadow = 'none';
            exportNode.style.margin = '0';

            const wrapper = document.createElement('div');
            wrapper.style.position = 'fixed';
            wrapper.style.left = '-10000px';
            wrapper.style.top = '0';
            wrapper.style.background = '#ffffff';
            wrapper.style.padding = '0';
            wrapper.appendChild(exportNode);
            document.body.appendChild(wrapper);

            const PdfClass = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
            const hasHtml2Pdf = typeof window.html2pdf === 'function';
            const hasManualDeps = !!(PdfClass && window.html2canvas);

            if (!hasHtml2Pdf && !hasManualDeps) {
                wrapper.remove();
                showToast('PDF generator not available', 'error');
                return;
            }

            if (hasHtml2Pdf) {
                const worker = window.html2pdf().set({
                    filename: fileName,
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: {
                        scale: 2,
                        useCORS: true,
                        backgroundColor: '#ffffff'
                    },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                    pagebreak: { mode: ['css', 'legacy'] }
                }).from(exportNode).toPdf();

                worker.get('pdf').then(pdf => {
                    embedPayloadTextInPdf(pdf, embeddedPayloadText);
                });

                worker.save().then(() => {
                    showToast('PDF downloaded successfully');
                }).catch(err => {
                    console.error(err);
                    showToast('Error generating PDF', 'error');
                }).finally(() => {
                    wrapper.remove();
                });
                return;
            }

            window.html2canvas(exportNode, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            }).then(canvas => {
                const pdf = new PdfClass('p', 'mm', 'a4');
                const pageWidth = pdf.internal.pageSize.getWidth();
                const pageHeight = pdf.internal.pageSize.getHeight();
                const imgData = canvas.toDataURL('image/jpeg', 0.98);
                const imgWidth = pageWidth;
                const imgHeight = canvas.height * imgWidth / canvas.width;

                if (imgHeight <= pageHeight) {
                    pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);
                } else if (imgHeight <= pageHeight + 2) {
                    const scale = pageHeight / imgHeight;
                    const scaledWidth = imgWidth * scale;
                    const scaledHeight = imgHeight * scale;
                    const xOffset = (pageWidth - scaledWidth) / 2;
                    pdf.addImage(imgData, 'JPEG', xOffset, 0, scaledWidth, scaledHeight);
                } else {
                    let heightLeft = imgHeight;
                    let position = 0;

                    pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;

                    while (heightLeft > 1) {
                        position -= pageHeight;
                        pdf.addPage();
                        pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
                        heightLeft -= pageHeight;
                    }
                }

                embedPayloadTextInPdf(pdf, embeddedPayloadText);
                pdf.save(fileName);
                showToast('PDF downloaded successfully');
            }).catch(err => {
                console.error(err);
                showToast('Error generating PDF', 'error');
            }).finally(() => {
                wrapper.remove();
            });
        }

        // Template Management
        function saveTemplate() {
            const name = document.getElementById('templateName').value.trim();
            if (!name) {
                showToast('Please enter a template name', 'error');
                return;
            }

            const template = {
                id: Date.now(),
                name: name,
                date: new Date().toLocaleDateString(),
                data: JSON.parse(JSON.stringify(invoiceData))
            };

            // Check for duplicates
            const existingIndex = savedTemplates.findIndex(t => t.name === name);
            if (existingIndex >= 0) {
                if (!confirm('A template with this name exists. Overwrite?')) return;
                savedTemplates[existingIndex] = template;
            } else {
                savedTemplates.push(template);
            }

            localStorage.setItem('invoiceTemplates', JSON.stringify(savedTemplates));
            document.getElementById('templateName').value = '';
            showToast('Template saved successfully');
        }

        function loadTemplate(id) {
            const template = savedTemplates.find(t => t.id === id);
            if (!template) return;

            applyInvoiceDataToForm(template.data);
            closeTemplateManager();
            showToast('Template loaded');
        }

        function deleteTemplate(id, event) {
            event.stopPropagation();
            if (confirm('Are you sure you want to delete this template?')) {
                savedTemplates = savedTemplates.filter(t => t.id !== id);
                localStorage.setItem('invoiceTemplates', JSON.stringify(savedTemplates));
                renderTemplatesList();
                showToast('Template deleted');
            }
        }

        function openTemplateManager() {
            document.getElementById('templateModal').classList.remove('hidden');
            renderTemplatesList();
        }

        function closeTemplateManager() {
            document.getElementById('templateModal').classList.add('hidden');
        }

        function renderTemplatesList() {
            const list = document.getElementById('templatesList');
            const empty = document.getElementById('noTemplates');
            
            if (savedTemplates.length === 0) {
                list.innerHTML = '';
                empty.classList.remove('hidden');
                return;
            }

            empty.classList.add('hidden');
            list.innerHTML = savedTemplates.map(template => `
                <div onclick="loadTemplate(${template.id})" class="template-card bg-white border border-gray-200 rounded-lg p-4 cursor-pointer hover:border-indigo-300 flex justify-between items-center group">
                    <div>
                        <h3 class="font-semibold text-gray-900">${template.name}</h3>
                        <p class="text-xs text-gray-500">Saved on ${template.date}</p>
                        <p class="text-xs text-gray-400 mt-1">${template.data.items.length} items • ${template.data.companyName || 'No company'}</p>
                    </div>
                    <button onclick="deleteTemplate(${template.id}, event)" class="opacity-0 group-hover:opacity-100 p-2 text-red-500 hover:bg-red-50 rounded transition-all">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            `).join('');
        }

        // Utilities
