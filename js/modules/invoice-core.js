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
            if (typeof scheduleCloudSave === 'function') {
                scheduleCloudSave();
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
            if (typeof scheduleCloudSave === 'function') {
                scheduleCloudSave();
            }
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
        let draggedLineItem = null;
        let pointerDragState = null;
        let addressTemplateAutoSaveTimer = null;
        let activeTemplateAutoSaveTimer = null;
        let activeTemplateId = null;
        const LINE_ITEM_DRAG_START_THRESHOLD = 5;
        const LINE_ITEM_DRAG_SCROLL_ZONE = 88;
        const LINE_ITEM_DRAG_MAX_SCROLL_STEP = 16;
        const ADDRESS_TEMPLATE_AUTO_SAVE_DELAY_MS = 1000;
        const ACTIVE_TEMPLATE_AUTO_SAVE_DELAY_MS = 700;

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
            document.body.classList.remove('settings-modal-open');
        }

        function toggleSettingsMenu(eventOrForceOpen) {
            if (eventOrForceOpen && typeof eventOrForceOpen.stopPropagation === 'function') {
                eventOrForceOpen.stopPropagation();
            }
            const menu = document.getElementById('settingsMenu');
            const button = document.getElementById('settingsMenuButton');
            if (!menu) return;
            const wasOpen = !menu.classList.contains('hidden');
            const willOpen = typeof eventOrForceOpen === 'boolean'
                ? eventOrForceOpen
                : !wasOpen;
            menu.classList.toggle('hidden', !willOpen);
            if (button) button.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
            if (willOpen) {
                closeChatTemplateBubble();
                closeMobileActionSheet();
                document.body.classList.add('settings-modal-open');
                requestAnimationFrame(() => {
                    const keyInput = document.getElementById('geminiApiKey');
                    if (keyInput) keyInput.focus();
                });
            } else if (wasOpen) {
                document.body.classList.remove('settings-modal-open');
                if (window.innerWidth < 1024) {
                    closeMobileActionSheet();
                }
            }
        }


        function updateActionSheetIcon(isOpen) {
            const trigger = document.getElementById('mobileActionSheetTrigger');
            if (!trigger) return;
            trigger.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
            const icon = trigger.querySelector('i');
            if (icon) {
                icon.className = isOpen ? 'fas fa-xmark' : 'fas fa-ellipsis';
            }
        }

        function closeMobileActionSheet() {
            document.body.classList.remove('mobile-action-sheet-open');
            updateActionSheetIcon(false);
        }

        function toggleMobileActionSheet(forceOpen) {
            if (typeof forceOpen === 'boolean') {
                document.body.classList.toggle('mobile-action-sheet-open', forceOpen);
            } else {
                document.body.classList.toggle('mobile-action-sheet-open');
            }
            const isOpen = document.body.classList.contains('mobile-action-sheet-open');
            updateActionSheetIcon(isOpen);
            if (isOpen) {
                closeMobilePreview();
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

        function updateHeaderHeight() {
            const header = document.querySelector('header');
            if (header) {
                document.documentElement.style.setProperty('--header-height', `${header.offsetHeight}px`);
            }
        }

        function preventNativePinchZoom(event) {
            if (event.touches && event.touches.length > 1) {
                event.preventDefault();
            }
        }

        function setupMobilePwaExperience() {
            document.body.classList.toggle('pwa-standalone', isStandalonePwa());
            const refreshViewportMetrics = () => {
                updateAppViewportHeight();
                updateHeaderHeight();
            };
            refreshViewportMetrics();

            if (window.visualViewport && typeof window.visualViewport.addEventListener === 'function') {
                window.visualViewport.addEventListener('resize', refreshViewportMetrics);
            }
            window.addEventListener('resize', refreshViewportMetrics, { passive: true });
            window.addEventListener('orientationchange', refreshViewportMetrics, { passive: true });
            window.addEventListener('resize', () => {
                if (window.innerWidth >= 1024) {
                    closeMobileActionSheet();
                }
            }, { passive: true });

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
            closeMobileActionSheet();
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
            setupLineItemReordering();
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
                const address = normalizeLineItemAddressValue(addressInput ? addressInput.value : '');
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
            if (typeof scheduleCloudSave === 'function') {
                scheduleCloudSave();
            }
            scheduleActiveTemplateAutoSave();
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
                    row.className = `invoice-line-item ${index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}`;
                    const descriptionHtml = formatDescriptionText(item.description);
                    row.innerHTML = `
                        <td class="py-3 pr-4 text-gray-900"><div class="invoice-item-description">${descriptionHtml}</div></td>
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
        function getLineItemCount() {
            return document.querySelectorAll('#lineItemsContainer .line-item').length;
        }

        function renumberLineItems() {
            const items = document.querySelectorAll('#lineItemsContainer .line-item');
            items.forEach((item, index) => {
                const badge = item.querySelector('.line-item-badge');
                if (badge) badge.textContent = `Item ${index + 1}`;
                const handle = item.querySelector('.line-item-drag-handle');
                if (handle) handle.setAttribute('aria-label', `Reorder item ${index + 1}`);
                item.dataset.itemIndex = String(index);
            });
        }

        function getLineItemAfterDragPosition(container, y) {
            const draggableItems = [...container.querySelectorAll('.line-item:not(.dragging)')];

            return draggableItems.reduce((closest, child) => {
                const box = child.getBoundingClientRect();
                const offset = y - box.top - box.height / 2;

                if (offset < 0 && offset > closest.offset) {
                    return { offset, element: child };
                }

                return closest;
            }, { offset: Number.NEGATIVE_INFINITY, element: null }).element;
        }

        function getLineItemScrollParent(element) {
            let node = element ? element.parentElement : null;
            while (node && node !== document.body) {
                const style = window.getComputedStyle(node);
                if (/(auto|scroll)/.test(style.overflowY) && node.scrollHeight > node.clientHeight) {
                    return node;
                }
                node = node.parentElement;
            }

            return document.scrollingElement || document.documentElement;
        }

        function resetDraggedLineItemStyles(item) {
            if (!item) return;

            item.classList.remove('dragging');
            item.removeAttribute('aria-grabbed');
            item.style.removeProperty('height');
            item.style.removeProperty('left');
            item.style.removeProperty('margin');
            item.style.removeProperty('pointer-events');
            item.style.removeProperty('position');
            item.style.removeProperty('top');
            item.style.removeProperty('transform');
            item.style.removeProperty('width');
            item.style.removeProperty('z-index');
        }

        function positionDraggedLineItem() {
            if (!pointerDragState || !pointerDragState.isDragging) return;

            const x = pointerDragState.latestX - pointerDragState.offsetX;
            const y = pointerDragState.latestY - pointerDragState.offsetY;
            pointerDragState.item.style.transform = `translate3d(${Math.round(x)}px, ${Math.round(y)}px, 0) scale(1.01)`;
        }

        function moveLineItemPlaceholder() {
            if (!pointerDragState || !pointerDragState.isDragging) return;

            const dragCenterY = pointerDragState.latestY - pointerDragState.offsetY + pointerDragState.itemRect.height / 2;
            const afterElement = getLineItemAfterDragPosition(pointerDragState.container, dragCenterY);

            if (!afterElement) {
                pointerDragState.container.appendChild(pointerDragState.placeholder);
            } else if (afterElement !== pointerDragState.placeholder.nextElementSibling) {
                pointerDragState.container.insertBefore(pointerDragState.placeholder, afterElement);
            }
        }

        function updateLineItemAutoScroll() {
            if (!pointerDragState || !pointerDragState.isDragging) return;

            const scroller = pointerDragState.scrollParent;
            const viewportRect = scroller === document.scrollingElement || scroller === document.documentElement
                ? { top: 0, bottom: window.innerHeight }
                : scroller.getBoundingClientRect();

            let scrollStep = 0;
            const distanceFromTop = pointerDragState.latestY - viewportRect.top;
            const distanceFromBottom = viewportRect.bottom - pointerDragState.latestY;

            if (distanceFromTop < LINE_ITEM_DRAG_SCROLL_ZONE) {
                scrollStep = -Math.ceil((1 - Math.max(distanceFromTop, 0) / LINE_ITEM_DRAG_SCROLL_ZONE) * LINE_ITEM_DRAG_MAX_SCROLL_STEP);
            } else if (distanceFromBottom < LINE_ITEM_DRAG_SCROLL_ZONE) {
                scrollStep = Math.ceil((1 - Math.max(distanceFromBottom, 0) / LINE_ITEM_DRAG_SCROLL_ZONE) * LINE_ITEM_DRAG_MAX_SCROLL_STEP);
            }

            if (scrollStep) {
                scroller.scrollTop += scrollStep;
                moveLineItemPlaceholder();
                pointerDragState.autoScrollFrame = requestAnimationFrame(updateLineItemAutoScroll);
            } else {
                pointerDragState.autoScrollFrame = null;
            }
        }

        function scheduleLineItemAutoScroll() {
            if (!pointerDragState || !pointerDragState.isDragging || pointerDragState.autoScrollFrame) return;
            pointerDragState.autoScrollFrame = requestAnimationFrame(updateLineItemAutoScroll);
        }

        function removeLineItemMouseListeners() {
            document.removeEventListener('mousemove', handleLineItemMouseMove);
            document.removeEventListener('mouseup', handleLineItemMouseUp);
        }

        function beginLineItemDrag(event) {
            if (!pointerDragState || pointerDragState.isDragging) return;

            const item = pointerDragState.item;
            const rect = item.getBoundingClientRect();
            const placeholder = document.createElement('div');
            placeholder.className = 'line-item-placeholder';
            placeholder.style.height = `${rect.height}px`;
            placeholder.setAttribute('aria-hidden', 'true');

            item.parentNode.insertBefore(placeholder, item);

            pointerDragState.placeholder = placeholder;
            pointerDragState.itemRect = rect;
            pointerDragState.offsetX = event.clientX - rect.left;
            pointerDragState.offsetY = event.clientY - rect.top;
            pointerDragState.latestX = event.clientX;
            pointerDragState.latestY = event.clientY;
            pointerDragState.isDragging = true;
            pointerDragState.scrollParent = getLineItemScrollParent(item);

            draggedLineItem = item;
            item.classList.add('dragging');
            item.setAttribute('aria-grabbed', 'true');
            item.style.position = 'fixed';
            item.style.top = '0';
            item.style.left = '0';
            item.style.width = `${rect.width}px`;
            item.style.height = `${rect.height}px`;
            item.style.margin = '0';
            item.style.pointerEvents = 'none';
            item.style.zIndex = '70';

            pointerDragState.container.classList.add('line-items-reordering');
            document.body.classList.add('line-item-drag-active');
            positionDraggedLineItem();
            moveLineItemPlaceholder();
        }

        function finishLineItemDrag() {
            if (!pointerDragState && !draggedLineItem) return;

            const state = pointerDragState;
            const item = state ? state.item : draggedLineItem;
            removeLineItemMouseListeners();

            if (state && state.autoScrollFrame) {
                cancelAnimationFrame(state.autoScrollFrame);
            }

            if (state && state.isDragging && state.placeholder) {
                state.placeholder.replaceWith(item);
            } else if (state && state.placeholder) {
                state.placeholder.remove();
            }

            resetDraggedLineItemStyles(item);
            if (state && state.container) {
                state.container.classList.remove('line-items-reordering');
            }

            draggedLineItem = null;
            pointerDragState = null;
            document.body.classList.remove('line-item-drag-active');
            renumberLineItems();
            updateInvoice();
        }

        function cancelLineItemDrag() {
            if (!pointerDragState) return;

            const state = pointerDragState;
            removeLineItemMouseListeners();
            if (state.autoScrollFrame) {
                cancelAnimationFrame(state.autoScrollFrame);
            }
            if (state.placeholder) {
                state.placeholder.remove();
            }

            resetDraggedLineItemStyles(state.item);
            state.container.classList.remove('line-items-reordering');
            draggedLineItem = null;
            pointerDragState = null;
            document.body.classList.remove('line-item-drag-active');
        }

        function handleLineItemDragStart(event) {
            event.preventDefault();
        }

        function handleLineItemDragOver(event) {
            event.preventDefault();
        }

        function startLineItemDragTracking(event, pointerId, shouldCapture = true) {
            const item = event.target.closest('.line-item');
            const container = document.getElementById('lineItemsContainer');
            if (!item || !container) return;

            pointerDragState = {
                pointerId,
                handle: event.currentTarget,
                item,
                container,
                startX: event.clientX,
                startY: event.clientY,
                latestX: event.clientX,
                latestY: event.clientY,
                isDragging: false,
                placeholder: null,
                autoScrollFrame: null
            };

            if (shouldCapture && event.currentTarget.setPointerCapture) {
                event.currentTarget.setPointerCapture(pointerId);
            }
            event.preventDefault();
        }

        function handleLineItemPointerDown(event) {
            if (event.button !== undefined && event.button !== 0) return;
            startLineItemDragTracking(event, event.pointerId);
        }

        function moveTrackedLineItem(clientX, clientY, sourceEvent) {
            if (!pointerDragState) return;

            pointerDragState.latestX = clientX;
            pointerDragState.latestY = clientY;

            if (!pointerDragState.isDragging) {
                const deltaX = clientX - pointerDragState.startX;
                const deltaY = clientY - pointerDragState.startY;
                if (Math.hypot(deltaX, deltaY) < LINE_ITEM_DRAG_START_THRESHOLD) return;
                beginLineItemDrag({
                    clientX,
                    clientY
                });
            }

            sourceEvent.preventDefault();
            positionDraggedLineItem();
            moveLineItemPlaceholder();
            scheduleLineItemAutoScroll();
        }

        function handleLineItemPointerMove(event) {
            if (!pointerDragState || pointerDragState.pointerId !== event.pointerId) return;
            moveTrackedLineItem(event.clientX, event.clientY, event);
        }

        function handleLineItemPointerUp(event) {
            if (!pointerDragState || pointerDragState.pointerId !== event.pointerId) return;

            if (pointerDragState.handle && pointerDragState.handle.hasPointerCapture && pointerDragState.handle.hasPointerCapture(event.pointerId)) {
                pointerDragState.handle.releasePointerCapture(event.pointerId);
            }
            finishLineItemDrag();
        }

        function handleLineItemMouseDown(event) {
            if (pointerDragState || event.button !== 0) return;

            startLineItemDragTracking(event, 'mouse', false);
            document.addEventListener('mousemove', handleLineItemMouseMove);
            document.addEventListener('mouseup', handleLineItemMouseUp);
        }

        function handleLineItemMouseMove(event) {
            if (!pointerDragState || pointerDragState.pointerId !== 'mouse') return;
            moveTrackedLineItem(event.clientX, event.clientY, event);
        }

        function handleLineItemMouseUp() {
            if (!pointerDragState || pointerDragState.pointerId !== 'mouse') return;
            finishLineItemDrag();
        }

        function handleLineItemDragKeydown(event) {
            if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;

            const item = event.target.closest('.line-item');
            const container = document.getElementById('lineItemsContainer');
            if (!item || !container) return;

            const sibling = event.key === 'ArrowUp' ? item.previousElementSibling : item.nextElementSibling;
            if (!sibling || !sibling.classList.contains('line-item')) return;

            event.preventDefault();
            if (event.key === 'ArrowUp') {
                container.insertBefore(item, sibling);
            } else {
                container.insertBefore(sibling, item);
            }

            item.classList.remove('line-item-keyboard-moved');
            requestAnimationFrame(() => item.classList.add('line-item-keyboard-moved'));
            renumberLineItems();
            updateInvoice();
            event.currentTarget.focus({ preventScroll: true });
        }

        function handleLineItemDragEscape(event) {
            if (event.key !== 'Escape' || !pointerDragState || !pointerDragState.isDragging) return;
            event.preventDefault();
            cancelLineItemDrag();
        }

        function setupLineItemReordering() {
            const container = document.getElementById('lineItemsContainer');
            if (!container) return;

            document.addEventListener('keydown', handleLineItemDragEscape);
        }

        function addLineItem(itemData = null, shouldRefresh = true) {
            const container = document.getElementById('lineItemsContainer');
            const normalizedItem = normalizeLineItem(itemData || getDefaultLineItem());
            const itemNumber = getLineItemCount() + 1;

            const div = document.createElement('div');
            div.className = 'line-item bg-gray-50 p-3 rounded-lg border border-gray-200 space-y-2';
            div.innerHTML = `
                <div class="line-item-header">
                    <div class="line-item-title">
                        <button type="button" class="line-item-drag-handle"
                            ondragstart="handleLineItemDragStart(event)"
                            onpointerdown="handleLineItemPointerDown(event)" onpointermove="handleLineItemPointerMove(event)" onpointerup="handleLineItemPointerUp(event)" onpointercancel="handleLineItemPointerUp(event)"
                            onmousedown="handleLineItemMouseDown(event)"
                            onkeydown="handleLineItemDragKeydown(event)"
                            title="Drag or use arrow keys to reorder item" aria-label="Reorder item ${itemNumber}">
                            <i class="fas fa-grip-vertical" aria-hidden="true"></i>
                        </button>
                        <span class="line-item-badge">Item ${itemNumber}</span>
                    </div>
                    <button type="button" onclick="removeLineItem(this)" class="px-2 text-red-500 hover:text-red-700" title="Remove item">
                        <i class="fas fa-trash-alt text-xs"></i>
                    </button>
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-600 mb-1">Property Address</label>
                    <input type="text" placeholder="e.g., 885 County Rd CKL, Champion, MI 49814"
                        class="item-address w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500"
                        oninput="scheduleUpdate(); scheduleAddressTemplateAutoSave(this)" onchange="handleLineItemAddressSubmitted(this)">
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-600 mb-1">Work Done</label>
                    <textarea rows="2" placeholder="e.g., Install handrail for front steps"
                        class="item-work w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 resize-none"
                        oninput="scheduleUpdate()"></textarea>
                </div>
                <div class="line-item-numbers">
                    <div class="line-item-number-field">
                        <label>Qty</label>
                        <input type="number" placeholder="Qty" min="0" step="1" value="1"
                            class="item-qty w-full px-2 py-1 border border-gray-300 rounded text-sm text-right"
                            oninput="scheduleUpdate()">
                    </div>
                    <div class="line-item-number-field">
                        <label>Rate</label>
                        <input type="number" placeholder="Rate" min="0" step="0.01" value="0"
                            class="item-rate w-full px-2 py-1 border border-gray-300 rounded text-sm text-right"
                            oninput="scheduleUpdate()">
                    </div>
                </div>
            `;

            div.querySelector('.item-address').value = normalizeLineItemAddressValue(normalizedItem.address);
            div.querySelector('.item-work').value = normalizedItem.work;
            div.querySelector('.item-qty').value = normalizedItem.quantity;
            div.querySelector('.item-rate').value = normalizedItem.rate;

            container.appendChild(div);
            renumberLineItems();
            if (shouldRefresh) {
                updateInvoice();
            }
        }

        function removeLineItem(btn) {
            const items = document.querySelectorAll('.line-item');
            if (items.length > 1) {
                btn.closest('.line-item').remove();
                renumberLineItems();
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

        function clearInvoice() {
            if (!confirm('Clear this invoice and start over?')) return;

            clearActiveTemplate();
            applyInvoiceDataToForm(getDefaultInvoiceData());
            clearAddressTemplateAutoSaveMarkers();

            const logoInput = document.getElementById('logoInput');
            const importInput = document.getElementById('invoiceImportInput');
            const templateNameInput = document.getElementById('templateName');
            if (logoInput) logoInput.value = '';
            if (importInput) importInput.value = '';
            if (templateNameInput) templateNameInput.value = '';

            closeSettingsMenu();
            closeChatTemplateBubble();
            closeMobileActionSheet();
            showToast('Invoice cleared', 'info');
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

        function cleanCityCandidateForFilename(value) {
            return normalizeSpace(value)
                .replace(/^[,.\-_\s]+|[,.\-_\s]+$/g, '');
        }

        function stripStreetPrefixFromCityCandidate(value) {
            let candidate = cleanCityCandidateForFilename(value);
            if (!candidate) return '';

            const commaParts = candidate.split(/\s*,\s*/).map(cleanCityCandidateForFilename).filter(Boolean);
            if (commaParts.length > 1) {
                candidate = commaParts[commaParts.length - 1];
            }

            const streetSuffixPattern = /\b(?:county\s+(?:road|rd)|co\.?\s*rd|cr|street|st|road|rd|avenue|ave|boulevard|blvd|drive|dr|lane|ln|court|ct|place|pl|terrace|ter|parkway|pkwy|circle|cir|trail|trl|way|highway|hwy)\.?\b/i;
            const streetSuffixMatch = candidate.match(streetSuffixPattern);
            if (!streetSuffixMatch) {
                return /^\d/.test(candidate) ? '' : candidate;
            }

            const streetSuffix = streetSuffixMatch[0];
            candidate = cleanCityCandidateForFilename(candidate.slice(streetSuffixMatch.index + streetSuffix.length));
            if (!candidate) return '';

            if (/^(?:county\s+(?:road|rd)|co\.?\s*rd|cr|highway|hwy)$/i.test(streetSuffix)) {
                candidate = cleanCityCandidateForFilename(candidate.replace(/^(?:[A-Z]{1,5}\d*|\d+[A-Z]?)\s+(?=[A-Za-z])/, ''));
            }

            return candidate;
        }

        function extractCityFromText(text) {
            const rawValue = String(text || '').trim();
            if (!rawValue) return '';

            const rawLines = rawValue
                .replace(/\r/g, '\n')
                .split(/\n+/)
                .map(line => normalizeCitySpacingInLine(line))
                .filter(Boolean);

            const normalizedCandidates = rawLines.length
                ? [...rawLines, normalizeCitySpacingInLine(rawLines.join(' '))]
                : [normalizeCitySpacingInLine(rawValue)];

            for (const candidate of normalizedCandidates) {
                const commaStateZip = candidate.match(/^(.+?)\s*,\s*[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b/);
                if (commaStateZip) {
                    const city = stripStreetPrefixFromCityCandidate(commaStateZip[1]);
                    if (city) return city;
                }

                const commaState = candidate.match(/^(.+?)\s*,\s*[A-Za-z]{2}\b/);
                if (commaState) {
                    const city = stripStreetPrefixFromCityCandidate(commaState[1]);
                    if (city) return city;
                }

                const stateZip = candidate.match(/^(.+?)\s+[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b/);
                if (stateZip) {
                    const city = stripStreetPrefixFromCityCandidate(stateZip[1]);
                    if (city) return city;
                }
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
                    const cityCandidates = [address, addressFallback].filter(Boolean);
                    for (const candidate of cityCandidates) {
                        const city = extractCityFromText(candidate);
                        if (city) return city;
                    }
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

        function getDocumentTypeForFilename(data) {
            return /\bbid\b/i.test(String(data.documentType || '')) ? 'Bid' : 'Invoice';
        }

        function isLikelyAddressOnlyFilenameCandidate(value) {
            const text = normalizeSpace(value);
            if (!text) return false;

            return /^\d+\s+/.test(text)
                || /,\s*[A-Za-z]{2}\b/.test(text)
                || /\b[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b/.test(text)
                || /\b(?:county\s+(?:road|rd)|co\.?\s*rd|cr|street|st|road|rd|avenue|ave|boulevard|blvd|drive|dr|lane|ln|court|ct|place|pl|terrace|ter|parkway|pkwy|circle|cir|trail|trl|way|highway|hwy)\.?\b/i.test(text);
        }

        function normalizeWorkCandidateForFilename(value, documentType) {
            const documentWord = getDocumentTypeForFilename({ documentType }).toLowerCase();
            const documentPrefixPattern = new RegExp(`^${documentWord}\\b\\s*(?:for\\b\\s*)?`, 'i');
            return normalizeFilenamePart(value, 80)
                .replace(documentPrefixPattern, '')
                .replace(/^(?:for|to)\b\s*/i, '')
                .trim();
        }

        function getDescriptorForFilename(data) {
            const candidates = [];
            const documentType = getDocumentTypeForFilename(data);

            if (Array.isArray(data.items)) {
                for (const item of data.items) {
                    if (!item || typeof item !== 'object') continue;
                    const work = String(item.work || '').trim();
                    if (work) candidates.push(work);

                    const description = String(item.description || '').trim();
                    const parsed = parseDescriptionFields(description);
                    if (parsed.work) candidates.push(parsed.work);

                    if (!work && !parsed.work && description) {
                        const address = normalizeSpace(item.address || '');
                        const normalizedDescription = normalizeSpace(description);
                        if (
                            normalizedDescription !== address &&
                            !isLikelyAddressOnlyFilenameCandidate(normalizedDescription)
                        ) {
                            candidates.push(normalizedDescription);
                        }
                    }
                }
            }

            for (const candidate of candidates) {
                const words = normalizeWorkCandidateForFilename(candidate, documentType)
                    .split(/\s+/)
                    .map(word => word.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, ''))
                    .filter(Boolean);

                if (words.length >= 2) {
                    return words.slice(0, 5).join(' ');
                }

                if (words.length === 1) {
                    return `${words[0]} Work`;
                }
            }

            return 'Service Work';
        }

        function sanitizeFileName(baseName, extension = 'pdf') {
            const reservedWindowsNames = new Set([
                'con', 'prn', 'aux', 'nul',
                'com1', 'com2', 'com3', 'com4', 'com5', 'com6', 'com7', 'com8', 'com9',
                'lpt1', 'lpt2', 'lpt3', 'lpt4', 'lpt5', 'lpt6', 'lpt7', 'lpt8', 'lpt9'
            ]);

            let safeBase = String(baseName || '')
                .replace(/[\x00-\x1f\x80-\x9f]+/g, ' ')
                .replace(/[\\/:*?"<>|]+/g, ' ')
                .replace(/\s+/g, ' ')
                .trim()
                .replace(/[.\s]+$/g, '');

            if (!safeBase) {
                safeBase = 'Invoice';
            }

            if (reservedWindowsNames.has(safeBase.toLowerCase())) {
                safeBase = `Invoice ${safeBase}`;
            }

            let safeExtension = String(extension || 'pdf')
                .replace(/[^\w]+/g, '')
                .toLowerCase();
            if (!safeExtension) safeExtension = 'pdf';

            return `${safeBase}.${safeExtension}`;
        }

        function buildPdfFileName(data) {
            const sourceData = data && typeof data === 'object' ? data : {};
            const documentType = normalizeFilenamePart(getDocumentTypeForFilename(sourceData), 12) || 'Invoice';
            const descriptorRaw = normalizeFilenamePart(getDescriptorForFilename(sourceData), 48) || 'Service Work';
            const descriptorWords = descriptorRaw.split(/\s+/).filter(Boolean);
            const descriptor = descriptorWords.length >= 2
                ? descriptorRaw
                : `${descriptorWords[0] || 'Service'} Work`;

            const city = normalizeFilenamePart(getCityForFilename(sourceData), 28) || 'Unknown City';
            const datePart = formatFilenameDate(sourceData.invoiceDate || new Date().toISOString().split('T')[0]);
            const safeDate = datePart || formatFilenameDate(new Date().toISOString()) || String(Date.now());
            let baseName = `${documentType} - ${descriptor} - ${city} - ${safeDate}`.replace(/\s{2,}/g, ' ').trim();

            return sanitizeFileName(baseName, 'pdf');
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

        function decodeEmbeddedInvoicePayloadChunks(chunks) {
            const uniqueChunks = new Map();
            (Array.isArray(chunks) ? chunks : []).forEach(entry => {
                const order = Number(entry && entry.order);
                const chunk = String(entry && entry.chunk ? entry.chunk : '').replace(/[^A-Za-z0-9_-]/g, '');
                if (!Number.isFinite(order) || order < 1 || !chunk || uniqueChunks.has(order)) return;
                uniqueChunks.set(order, chunk);
            });

            if (!uniqueChunks.size) return null;

            const decoded = decodeBase64Url(Array.from(uniqueChunks.entries())
                .sort((a, b) => a[0] - b[0])
                .map(entry => entry[1])
                .join(''));
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

        function extractEmbeddedInvoicePayloadFromText(rawText) {
            const text = String(rawText || '');
            if (!text.includes(PDF_PAYLOAD_MARKERS.chunkPrefix)) return null;

            const startIndex = text.includes(PDF_PAYLOAD_MARKERS.start)
                ? text.indexOf(PDF_PAYLOAD_MARKERS.start)
                : 0;
            const endIndex = text.includes(PDF_PAYLOAD_MARKERS.end)
                ? text.indexOf(PDF_PAYLOAD_MARKERS.end, startIndex)
                : text.length;
            if (endIndex < 0 || endIndex <= startIndex) return null;

            const payloadBlock = text.slice(startIndex, endIndex);
            const chunkPattern = new RegExp(`${PDF_PAYLOAD_MARKERS.chunkPrefix}(\\d{3})\\s*:\\s*([A-Za-z0-9_\\-\\s]+?)(?=\\s*(?:,?\\s*${PDF_PAYLOAD_MARKERS.chunkPrefix}\\d{3}\\s*:|${PDF_PAYLOAD_MARKERS.end}|$))`, 'g');
            const chunks = [];
            let match;

            while ((match = chunkPattern.exec(payloadBlock)) !== null) {
                const order = Number(match[1]);
                const chunk = String(match[2] || '').replace(/[^A-Za-z0-9_-]/g, '');
                if (!chunk) continue;
                chunks.push({ order, chunk });
            }

            return decodeEmbeddedInvoicePayloadChunks(chunks);
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

        function getSelectablePdfTextNodes(root) {
            if (!root || typeof document.createTreeWalker !== 'function') return [];

            const nodes = [];
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
                acceptNode(node) {
                    const text = normalizeSpace(node && node.nodeValue);
                    if (!text) return NodeFilter.FILTER_REJECT;

                    const parent = node.parentElement;
                    if (!parent) return NodeFilter.FILTER_REJECT;

                    const style = window.getComputedStyle(parent);
                    if (
                        style.display === 'none' ||
                        style.visibility === 'hidden' ||
                        style.opacity === '0' ||
                        style.fontSize === '0px'
                    ) {
                        return NodeFilter.FILTER_REJECT;
                    }

                    return NodeFilter.FILTER_ACCEPT;
                }
            });

            while (walker.nextNode()) {
                nodes.push(walker.currentNode);
            }

            return nodes;
        }

        function addSelectableTextLayerToPdf(pdf, sourceNode) {
            if (!pdf || !sourceNode || typeof pdf.text !== 'function') return;
            if (!sourceNode.getBoundingClientRect || typeof document.createRange !== 'function') return;

            const rootRect = sourceNode.getBoundingClientRect();
            if (!rootRect.width || !rootRect.height) return;

            const pageHeight = pdf.internal && pdf.internal.pageSize && typeof pdf.internal.pageSize.getHeight === 'function'
                ? pdf.internal.pageSize.getHeight()
                : 297;
            const xScale = PAPER_WIDTH_MM / rootRect.width;
            const pageCount = typeof pdf.getNumberOfPages === 'function'
                ? Math.max(1, pdf.getNumberOfPages())
                : 1;
            const originalPage = typeof pdf.getCurrentPageInfo === 'function'
                ? pdf.getCurrentPageInfo().pageNumber
                : pageCount;
            const previousSize = typeof pdf.getFontSize === 'function' ? pdf.getFontSize() : null;

            try {
                if (typeof pdf.setTextColor === 'function') {
                    pdf.setTextColor(255, 255, 255);
                }

                getSelectablePdfTextNodes(sourceNode).forEach(node => {
                    const text = normalizeSpace(node.nodeValue);
                    if (!text) return;

                    const range = document.createRange();
                    range.selectNodeContents(node);
                    const rects = Array.from(range.getClientRects())
                        .filter(rect => rect.width > 0 && rect.height > 0);
                    range.detach();

                    if (!rects.length) return;

                    const rect = rects[0];
                    const parentStyle = window.getComputedStyle(node.parentElement);
                    const fontSizePx = parseFloat(parentStyle.fontSize) || rect.height || 12;
                    const x = Math.max(0, (rect.left - rootRect.left) * xScale);
                    const absoluteY = Math.max(0, (rect.top - rootRect.top + fontSizePx * 0.82) * xScale);
                    const pageNumber = Math.min(pageCount, Math.max(1, Math.floor(absoluteY / pageHeight) + 1));
                    const y = absoluteY - ((pageNumber - 1) * pageHeight);
                    const fontSizeMm = Math.max(1, fontSizePx * xScale);

                    if (typeof pdf.setPage === 'function') {
                        pdf.setPage(pageNumber);
                    }
                    if (typeof pdf.setFontSize === 'function') {
                        pdf.setFontSize(fontSizeMm);
                    }

                    pdf.text(text, x, y, {
                        baseline: 'alphabetic',
                        maxWidth: Math.max(1, rect.width * xScale),
                        renderingMode: 'invisible'
                    });
                });
            } catch (error) {
                console.warn('Failed to add selectable PDF text layer', error);
            } finally {
                if (Number.isFinite(previousSize) && typeof pdf.setFontSize === 'function') {
                    pdf.setFontSize(previousSize);
                }
                if (typeof pdf.setTextColor === 'function') {
                    pdf.setTextColor(0, 0, 0);
                }
                if (typeof pdf.setPage === 'function') {
                    pdf.setPage(originalPage);
                }
            }
        }

        function generatePDF() {
            updateInvoice();
            const exportInvoiceData = normalizeInvoiceData(invoiceData);
            const element = document.getElementById('invoice-preview');
            const fileName = buildPdfFileName(exportInvoiceData);
            const embeddedPayloadText = createEmbeddedInvoicePayloadText(exportInvoiceData);
            showToast('Generating PDF...', 'info');

            const exportNode = element.cloneNode(true);
            exportNode.style.transform = 'none';
            exportNode.style.transformOrigin = 'top left';
            exportNode.style.width = `${PAPER_WIDTH_MM}mm`;
            exportNode.style.minHeight = 'auto';
            exportNode.style.height = 'auto';
            exportNode.style.aspectRatio = 'auto';
            exportNode.style.boxShadow = 'none';
            exportNode.style.margin = '0';
            exportNode.style.overflow = 'visible';

            // Normalize layout for export so long invoices can flow across pages cleanly.
            const exportContentColumn = exportNode.querySelector('.h-full.flex.flex-col');
            if (exportContentColumn) {
                exportContentColumn.style.height = 'auto';
                exportContentColumn.style.minHeight = '0';
                exportContentColumn.style.display = 'block';
            }

            const exportItemsSection = exportNode.querySelector('.flex-1.mb-12');
            if (exportItemsSection) {
                exportItemsSection.style.flex = '0 0 auto';
            }

            const exportNotesSection = exportNode.querySelector('#previewNotesSection');
            if (exportNotesSection) {
                exportNotesSection.style.marginTop = '0';
            }

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
                    pagebreak: {
                        mode: ['css', 'legacy'],
                        avoid: ['thead', 'tr', '.invoice-line-item', '#previewNotesSection']
                    }
                }).from(exportNode).toPdf();

                const preparePdf = worker.get('pdf').then(pdf => {
                    addSelectableTextLayerToPdf(pdf, exportNode);
                    embedPayloadTextInPdf(pdf, embeddedPayloadText);
                });

                preparePdf.then(() => worker.save()).then(() => {
                    const savedTemplate = saveDownloadedPdfTemplate(exportInvoiceData);
                    showToast(savedTemplate ? 'PDF downloaded and template saved' : 'PDF downloaded successfully');
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

                addSelectableTextLayerToPdf(pdf, exportNode);
                embedPayloadTextInPdf(pdf, embeddedPayloadText);
                pdf.save(fileName);
                const savedTemplate = saveDownloadedPdfTemplate(exportInvoiceData);
                showToast(savedTemplate ? 'PDF downloaded and template saved' : 'PDF downloaded successfully');
            }).catch(err => {
                console.error(err);
                showToast('Error generating PDF', 'error');
            }).finally(() => {
                wrapper.remove();
            });
        }

        // Template Management
        function normalizeLineItemAddressValue(value) {
            return normalizeAddressForSingleLine(value, { includePhone: false });
        }

        function normalizeTemplateNameValue(value) {
            const raw = normalizeSpace(value);
            if (!raw) return '';
            return normalizeAddressForSingleLine(raw, { includePhone: false }) || raw;
        }

        function persistSavedTemplates() {
            localStorage.setItem('invoiceTemplates', JSON.stringify(savedTemplates));
            const modal = document.getElementById('templateModal');
            if (modal && !modal.classList.contains('hidden')) {
                renderTemplatesList();
            }
            if (typeof scheduleCloudSave === 'function') {
                scheduleCloudSave();
            }
        }

        function createTemplateSnapshot(name, existingTemplate = null, data = invoiceData) {
            const existingId = Number(existingTemplate && existingTemplate.id);
            return {
                id: Number.isFinite(existingId) && existingId > 0 ? existingId : Date.now(),
                name: name,
                date: new Date().toLocaleDateString(),
                data: cloneInvoiceDataSnapshot(data)
            };
        }

        function activateTemplate(template) {
            const id = Number(template && template.id);
            activeTemplateId = Number.isFinite(id) && id > 0 ? id : null;
        }

        function clearActiveTemplate() {
            activeTemplateId = null;
            if (activeTemplateAutoSaveTimer) {
                clearTimeout(activeTemplateAutoSaveTimer);
                activeTemplateAutoSaveTimer = null;
            }
            if (addressTemplateAutoSaveTimer) {
                clearTimeout(addressTemplateAutoSaveTimer);
                addressTemplateAutoSaveTimer = null;
            }
        }

        function clearAddressTemplateAutoSaveMarkers() {
            document.querySelectorAll('.item-address').forEach(input => {
                if (!input || !input.dataset) return;
                delete input.dataset.autosavedTemplateId;
                delete input.dataset.autosavedTemplateName;
            });
        }

        function cloneInvoiceDataSnapshot(data = invoiceData) {
            return JSON.parse(JSON.stringify(normalizeInvoiceData(data)));
        }

        function stringifyTemplateData(data) {
            try {
                return JSON.stringify(data || {});
            } catch (_error) {
                return '';
            }
        }

        function saveActiveTemplateSnapshot(options = {}) {
            const activeIndex = findTemplateIndexById(activeTemplateId);
            if (activeIndex < 0) {
                activeTemplateId = null;
                return null;
            }

            const existingTemplate = savedTemplates[activeIndex];
            const data = cloneInvoiceDataSnapshot();
            if (stringifyTemplateData(existingTemplate && existingTemplate.data) === stringifyTemplateData(data)) {
                return existingTemplate;
            }

            const template = {
                ...existingTemplate,
                date: new Date().toLocaleDateString(),
                data
            };
            savedTemplates[activeIndex] = template;
            persistSavedTemplates();
            if (options.message) {
                showToast(options.message);
            }
            return template;
        }

        function scheduleActiveTemplateAutoSave() {
            if (!activeTemplateId) return;
            if (activeTemplateAutoSaveTimer) clearTimeout(activeTemplateAutoSaveTimer);
            activeTemplateAutoSaveTimer = setTimeout(() => {
                activeTemplateAutoSaveTimer = null;
                saveActiveTemplateSnapshot();
            }, ACTIVE_TEMPLATE_AUTO_SAVE_DELAY_MS);
        }

        function findTemplateIndexById(id) {
            const numericId = Number(id);
            if (!Number.isFinite(numericId) || numericId <= 0) return -1;
            return savedTemplates.findIndex(template => Number(template && template.id) === numericId);
        }

        function upsertTemplate(name, options = {}) {
            const normalizedName = normalizeTemplateNameValue(name);
            if (!normalizedName) return null;

            const idIndex = findTemplateIndexById(options.templateId);
            const nameIndex = savedTemplates.findIndex(template => normalizeTemplateNameValue(template && template.name) === normalizedName);
            let existingIndex = idIndex >= 0 ? idIndex : nameIndex;

            if (idIndex >= 0 && nameIndex >= 0 && idIndex !== nameIndex) {
                savedTemplates.splice(idIndex, 1);
                existingIndex = nameIndex > idIndex ? nameIndex - 1 : nameIndex;
            }

            const existingTemplate = existingIndex >= 0 ? savedTemplates[existingIndex] : null;
            const template = createTemplateSnapshot(normalizedName, existingTemplate, options.data || invoiceData);

            if (existingIndex >= 0) {
                if (options.confirmOverwrite && !confirm('A template with this name exists. Overwrite?')) {
                    return null;
                }
                savedTemplates[existingIndex] = template;
            } else {
                savedTemplates.push(template);
            }

            persistSavedTemplates();
            if (options.message) {
                showToast(options.message);
            }
            activateTemplate(template);
            return template;
        }

        function commitAddressTemplate(input, options = {}) {
            const address = normalizeLineItemAddressValue(input && input.value);
            if (!address) return null;
            if (input && input.value !== address) {
                input.value = address;
            }

            if (!(input && input.dataset && input.dataset.autosavedTemplateId) && activeTemplateId) {
                updateInvoice();
                return saveActiveTemplateSnapshot();
            }

            const previousAddress = input && input.dataset ? normalizeTemplateNameValue(input.dataset.autosavedTemplateName) : '';
            if (previousAddress === address) return null;

            updateInvoice();
            const template = upsertTemplate(address, {
                templateId: input && input.dataset ? input.dataset.autosavedTemplateId : '',
                message: options.message
            });

            if (template && input && input.dataset) {
                input.dataset.autosavedTemplateId = String(template.id);
                input.dataset.autosavedTemplateName = template.name;
            }

            return template;
        }

        function getPrimaryTemplateAddress(data) {
            if (!data || typeof data !== 'object' || !Array.isArray(data.items)) return '';

            for (const item of data.items) {
                if (!item || typeof item !== 'object') continue;

                const address = normalizeLineItemAddressValue(item.address || '');
                if (address) return address;

                const parsed = parseDescriptionFields(item.description || '');
                const parsedAddress = normalizeLineItemAddressValue(parsed.address || '');
                if (parsedAddress) return parsedAddress;
            }

            return '';
        }

        function getImportedPdfTemplateName(file) {
            const address = getPrimaryTemplateAddress(invoiceData);
            if (address) return address;

            const fileName = String(file && file.name ? file.name : '').trim();
            const baseName = normalizeSpace(fileName.replace(/\.[^.]+$/, ''));
            return baseName || 'Imported PDF';
        }

        function buildDownloadedPdfTemplateName(data) {
            const sourceData = data && typeof data === 'object' ? data : {};
            const address = normalizeFilenamePart(getPrimaryTemplateAddress(sourceData), 56);
            const documentType = normalizeFilenamePart(getDocumentTypeForFilename(sourceData), 12) || 'Invoice';
            const descriptorRaw = normalizeFilenamePart(getDescriptorForFilename(sourceData), 32) || 'Service Work';
            const descriptorWords = descriptorRaw.split(/\s+/).filter(Boolean);
            const descriptor = descriptorWords.length >= 2
                ? descriptorRaw
                : `${descriptorWords[0] || 'Service'} Work`;
            const datePart = formatFilenameDate(sourceData.invoiceDate || new Date().toISOString().split('T')[0]);
            const safeDate = datePart || formatFilenameDate(new Date().toISOString()) || String(Date.now());

            if (address) {
                return [address, documentType, descriptor, safeDate].filter(Boolean).join(' - ');
            }

            const city = normalizeFilenamePart(getCityForFilename(sourceData), 28);
            return [documentType, descriptor, city, safeDate].filter(Boolean).join(' - ');
        }

        function saveDownloadedPdfTemplate(data) {
            const template = upsertTemplate(buildDownloadedPdfTemplateName(data), { data });
            if (template) {
                clearAddressTemplateAutoSaveMarkers();
            }
            return template;
        }

        function saveImportedPdfTemplate(file) {
            updateInvoice();
            const template = upsertTemplate(getImportedPdfTemplateName(file));
            clearAddressTemplateAutoSaveMarkers();
            return template;
        }

        function scheduleAddressTemplateAutoSave(input) {
            if (addressTemplateAutoSaveTimer) clearTimeout(addressTemplateAutoSaveTimer);
            addressTemplateAutoSaveTimer = setTimeout(() => {
                addressTemplateAutoSaveTimer = null;
                commitAddressTemplate(input, { message: 'Template saved from address' });
            }, ADDRESS_TEMPLATE_AUTO_SAVE_DELAY_MS);
        }

        function handleLineItemAddressSubmitted(input) {
            if (addressTemplateAutoSaveTimer) {
                clearTimeout(addressTemplateAutoSaveTimer);
                addressTemplateAutoSaveTimer = null;
            }
            commitAddressTemplate(input, { message: 'Template saved from address' });
        }

        function saveTemplate() {
            const name = document.getElementById('templateName').value.trim();
            if (!name) {
                showToast('Please enter a template name', 'error');
                return;
            }

            if (upsertTemplate(name, { confirmOverwrite: true, message: 'Template saved successfully' })) {
                clearAddressTemplateAutoSaveMarkers();
                document.getElementById('templateName').value = '';
            }
        }

        function loadTemplate(id) {
            const template = savedTemplates.find(t => t.id === id);
            if (!template) return;

            clearActiveTemplate();
            const normalizedName = normalizeTemplateNameValue(template.name);
            if (normalizedName && normalizedName !== template.name) {
                template.name = normalizedName;
                persistSavedTemplates();
            }
            applyInvoiceDataToForm(template.data);
            activateTemplate(template);
            clearAddressTemplateAutoSaveMarkers();
            closeTemplateManager();
            showToast('Template loaded');
        }

        function deleteTemplate(id, event) {
            event.stopPropagation();
            if (confirm('Are you sure you want to delete this template?')) {
                savedTemplates = savedTemplates.filter(t => t.id !== id);
                if (Number(activeTemplateId) === Number(id)) {
                    clearActiveTemplate();
                }
                persistSavedTemplates();
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
            list.innerHTML = savedTemplates.map(template => {
                const templateId = Number(template.id);
                const safeId = Number.isFinite(templateId) ? templateId : 0;
                const itemCount = Array.isArray(template.data && template.data.items) ? template.data.items.length : 0;
                const companyName = template.data && template.data.companyName ? template.data.companyName : 'No company';
                const templateName = normalizeTemplateNameValue(template.name) || 'Untitled Template';
                return `
                <div onclick="loadTemplate(${safeId})" class="template-card bg-white border border-gray-200 rounded-lg p-4 cursor-pointer hover:border-indigo-300 flex justify-between items-center group">
                    <div>
                        <h3 class="font-semibold text-gray-900">${escapeHtml(templateName)}</h3>
                        <p class="text-xs text-gray-500">Saved on ${escapeHtml(template.date)}</p>
                        <p class="text-xs text-gray-400 mt-1">${itemCount} items • ${escapeHtml(companyName)}</p>
                    </div>
                    <button onclick="deleteTemplate(${safeId}, event)" class="opacity-0 group-hover:opacity-100 p-2 text-red-500 hover:bg-red-50 rounded transition-all">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            `;
            }).join('');
        }

        // Utilities
