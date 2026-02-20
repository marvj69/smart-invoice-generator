        async function readFileAsText(file) {
            if (file && typeof file.text === 'function') {
                return file.text();
            }

            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ''));
                reader.onerror = () => reject(new Error('Unable to read file'));
                reader.readAsText(file);
            });
        }

        async function readFileAsDataUrl(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ''));
                reader.onerror = () => reject(new Error('Unable to encode file'));
                reader.readAsDataURL(file);
            });
        }

        function getGeminiApiErrorMessage(payload, fallbackMessage) {
            const fallback = normalizeSpace(fallbackMessage) || 'Gemini request failed';
            if (!payload || typeof payload !== 'object') return fallback;
            if (!payload.error || typeof payload.error !== 'object') return fallback;

            const message = normalizeSpace(payload.error.message || '');
            if (message) return message;

            const status = normalizeSpace(payload.error.status || '');
            const code = payload.error.code;
            if (status && code !== undefined) return `${status} (${code})`;
            if (status) return status;
            if (code !== undefined) return `Error ${code}`;
            return fallback;
        }

        function extractGeminiTextResponse(payload) {
            if (!payload || typeof payload !== 'object') return '';
            const candidates = Array.isArray(payload.candidates) ? payload.candidates : [];

            for (const candidate of candidates) {
                const content = candidate && typeof candidate === 'object' ? candidate.content : null;
                const parts = content && Array.isArray(content.parts) ? content.parts : [];
                const text = parts
                    .map(part => (part && typeof part.text === 'string' ? part.text : ''))
                    .join('\n')
                    .trim();
                if (text) return text;
            }

            return '';
        }

        function parseJsonFromGeminiText(rawText) {
            const trimmed = String(rawText || '').trim();
            if (!trimmed) return null;

            const attempts = [trimmed];
            const fencedMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
            if (fencedMatch && fencedMatch[1]) {
                attempts.push(fencedMatch[1].trim());
            }

            const objectStart = trimmed.indexOf('{');
            const objectEnd = trimmed.lastIndexOf('}');
            if (objectStart >= 0 && objectEnd > objectStart) {
                attempts.push(trimmed.slice(objectStart, objectEnd + 1));
            }

            for (const attempt of attempts) {
                try {
                    return JSON.parse(attempt);
                } catch (error) {
                    continue;
                }
            }

            return null;
        }

        function buildGeminiInvoicePrompt() {
            return [
                'Extract invoice fields from this PDF and return JSON only.',
                'Use the response schema exactly.',
                'Rules:',
                '- Do not invent details that are not present.',
                '- For unknown text, use an empty string.',
                '- For unknown numeric values, use 0.',
                '- invoiceDate must be YYYY-MM-DD when possible, otherwise empty string.',
                '- discountType must be "fixed" or "percentage".',
                '- Keep every line item, and the details of that line item, as it appears in the PDF.',
                '- companyDetails and clientDetails must be address blocks in this order:',
                '  line 1: street number + street name',
                '  line 2: city, state ZIP',
                '  line 3: phone number only if present',
                '- each item.address must be two lines when possible:',
                '  line 1: street number + street name',
                '  line 2: city, state ZIP'
            ].join('\n');
        }

        function buildGeminiChatInvoicePrompt(userInput) {
            const todayIso = new Date().toISOString().split('T')[0];
            return [
                'Convert this user request into invoice JSON and return JSON only.',
                'Use the response schema exactly.',
                'Goal: infer as many fields as possible while staying grounded in provided details.',
                'Rules:',
                '- Prioritize explicit details from the user.',
                '- Infer documentType as "Bid" when quote/estimate/proposal language is used; otherwise "Invoice".',
                '- Infer missing quantity/rate/amount only when two of those values are provided.',
                `- Resolve relative dates (today/tomorrow/next Friday) using ${todayIso} as today.`,
                '- invoiceDate must be YYYY-MM-DD when possible, otherwise empty string.',
                '- Normalize companyDetails, clientDetails, and item.address into multiline address blocks when possible.',
                '- Do not invent specific names, street numbers, prices, or dates when they are not implied.',
                '- For unknown text use empty string. For unknown numeric values use 0.',
                'User request:',
                String(userInput || '').trim()
            ].join('\n');
        }

        function buildGeminiThinkingConfig(model) {
            const normalizedModel = normalizeSpace(model).toLowerCase();
            if (normalizedModel.startsWith('gemini-3')) {
                return { thinkingLevel: GEMINI_THINKING_LEVEL_MINIMAL };
            }
            return { thinkingBudget: GEMINI_FALLBACK_THINKING_BUDGET };
        }

        function buildGeminiInvoiceRequestBody(model, base64Payload, useSchema = true) {
            const requestBody = {
                contents: [
                    {
                        role: 'user',
                        parts: [
                            { text: buildGeminiInvoicePrompt() },
                            {
                                inlineData: {
                                    mimeType: 'application/pdf',
                                    data: base64Payload
                                }
                            }
                        ]
                    }
                ],
                generationConfig: {
                    temperature: 0,
                    responseMimeType: 'application/json',
                    thinkingConfig: buildGeminiThinkingConfig(model)
                }
                
            };

            if (useSchema) {
                requestBody.generationConfig.responseJsonSchema = GEMINI_INVOICE_SCHEMA;
            }

            return requestBody;
        }

        function buildGeminiChatInvoiceRequestBody(model, userPrompt, useSchema = true) {
            const requestBody = {
                contents: [
                    {
                        role: 'user',
                        parts: [
                            { text: buildGeminiChatInvoicePrompt(userPrompt) }
                        ]
                    }
                ],
                generationConfig: {
                    temperature: 0,
                    responseMimeType: 'application/json',
                    thinkingConfig: buildGeminiThinkingConfig(model)
                }
            };

            if (useSchema) {
                requestBody.generationConfig.responseJsonSchema = GEMINI_INVOICE_SCHEMA;
            }

            return requestBody;
        }

        function normalizeErrorMessage(error) {
            if (!error) return '';
            if (typeof error === 'string') return normalizeSpace(error);
            return normalizeSpace(error.message || String(error));
        }

        function shouldRetryWithoutSchema(message) {
            const text = normalizeSpace(message).toLowerCase();
            if (!text) return false;
            return text.includes('responsejsonschema')
                || text.includes('response schema')
                || text.includes('invalid argument')
                || text.includes('unsupported');
        }

        function isModelUnavailableMessage(message) {
            const text = normalizeSpace(message).toLowerCase();
            if (!text) return false;
            return (text.includes('model') && text.includes('not found'))
                || text.includes('model not found')
                || text.includes('unknown model')
                || text.includes('permission denied');
        }

        function getGeminiModelCandidates(preferredModel) {
            const first = normalizeSpace(preferredModel) || GEMINI_DEFAULT_MODEL;
            const candidates = [first, GEMINI_DEFAULT_MODEL, ...GEMINI_FALLBACK_MODELS];
            const unique = [];
            candidates.forEach(model => {
                const normalized = normalizeSpace(model);
                if (!normalized) return;
                if (!unique.includes(normalized)) {
                    unique.push(normalized);
                }
            });
            return unique;
        }

        async function callGeminiGenerateContent({ apiKey, model, requestBody }) {
            const endpoint = `${GEMINI_API_BASE_URL}/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`;
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), GEMINI_REQUEST_TIMEOUT_MS);
            const hasSchema = Boolean(
                requestBody
                && requestBody.generationConfig
                && requestBody.generationConfig.responseJsonSchema
            );
            const firstContent = requestBody && Array.isArray(requestBody.contents) ? requestBody.contents[0] : null;
            const parts = firstContent && Array.isArray(firstContent.parts) ? firstContent.parts : [];
            const inlineDataPart = parts.find(part => part && part.inlineData && typeof part.inlineData === 'object');
            const inlineData = inlineDataPart ? inlineDataPart.inlineData : null;
            appendImportDebug('Gemini request started', {
                model,
                hasSchema,
                hasInlinePdf: Boolean(inlineData && normalizeSpace(inlineData.data)),
                inlineMimeType: inlineData && inlineData.mimeType ? inlineData.mimeType : '',
                base64Length: inlineData && typeof inlineData.data === 'string' ? inlineData.data.length : 0,
                thinkingConfig: requestBody && requestBody.generationConfig ? requestBody.generationConfig.thinkingConfig : null
            });

            try {
                const response = await fetch(endpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody),
                    signal: controller.signal
                });

                let payload = null;
                try {
                    payload = await response.json();
                } catch (error) {
                    payload = null;
                }
                appendImportDebug('Gemini response received', { model, status: response.status, ok: response.ok });

                if (!response.ok) {
                    const errorMessage = getGeminiApiErrorMessage(payload, `Gemini request failed (${response.status})`);
                    appendImportDebug('Gemini response error', { model, error: errorMessage });
                    throw new Error(errorMessage);
                }

                return payload;
            } catch (error) {
                const message = normalizeErrorMessage(error);
                appendImportDebug('Gemini request failed', { model, error: message || 'Unknown error' });
                if (message.toLowerCase().includes('abort')) {
                    throw new Error('Gemini request timed out while processing invoice data');
                }
                if (message === 'failed to fetch' || message.includes('networkerror')) {
                    throw new Error('Could not reach Gemini API. Check your internet connection and API key.');
                }
                throw error;
            } finally {
                clearTimeout(timeoutId);
            }
        }

        function parseGeminiInvoicePayload(payload, sourceLabel = 'request') {
            const contextLabel = normalizeSpace(sourceLabel) || 'request';
            const responseText = extractGeminiTextResponse(payload);
            if (!responseText) {
                const blockReason = normalizeSpace(payload && payload.promptFeedback ? payload.promptFeedback.blockReason : '');
                if (blockReason) {
                    appendImportDebug('Gemini response blocked', blockReason);
                    throw new Error(`Gemini blocked this ${contextLabel} (${blockReason})`);
                }
                appendImportDebug('Gemini response had no text content');
                throw new Error(`Gemini returned an empty response for this ${contextLabel}`);
            }

            const parsedJson = parseJsonFromGeminiText(responseText);
            if (!parsedJson || typeof parsedJson !== 'object') {
                appendImportDebug('Gemini response JSON parse failed', responseText.slice(0, 220));
                throw new Error('Gemini response was not valid JSON');
            }

            const invoiceLikeData = extractInvoiceDataFromJson(parsedJson) || parsedJson;
            const normalized = normalizeInvoiceData(invoiceLikeData);
            if (!hasMeaningfulInvoiceData(normalized)) {
                appendImportDebug('Gemini response missing required invoice fields');
                throw new Error('Gemini did not return enough invoice data to populate the form');
            }
            appendImportDebug('Gemini response parsed successfully', {
                companyName: normalized.companyName || '',
                clientName: normalized.clientName || '',
                items: Array.isArray(normalized.items) ? normalized.items.length : 0
            });

            return normalized;
        }

        async function extractInvoiceDataFromPdfWithGemini(file) {
            const apiKey = getConfiguredGeminiApiKey();
            const model = getConfiguredGeminiModel();
            appendImportDebug('Gemini PDF extraction requested', {
                fileName: file && file.name ? file.name : '',
                fileSize: file && Number.isFinite(file.size) ? file.size : 0,
                configuredModel: model
            });

            if (!apiKey) {
                appendImportDebug('Gemini extraction stopped: missing API key');
                throw new Error('Enter your Gemini API key in Settings first');
            }

            saveGeminiSettings(false);

            const dataUrl = await readFileAsDataUrl(file);
            const base64Payload = dataUrl.includes(',') ? dataUrl.split(',')[1] : '';
            if (!base64Payload) {
                appendImportDebug('Gemini extraction failed: could not encode PDF');
                throw new Error('Could not encode PDF for Gemini request');
            }
            appendImportDebug('PDF encoded for Gemini', { base64Length: base64Payload.length });

            const modelCandidates = getGeminiModelCandidates(model);
            appendImportDebug('Gemini model candidates', modelCandidates.join(', '));
            let lastErrorMessage = '';

            for (const candidateModel of modelCandidates) {
                let tryNextModel = false;
                for (let attempt = 0; attempt < 2; attempt += 1) {
                    const useSchema = attempt === 0;
                    appendImportDebug('Gemini attempt', {
                        model: candidateModel,
                        attempt: attempt + 1,
                        useSchema
                    });
                    try {
                        const requestBody = buildGeminiInvoiceRequestBody(candidateModel, base64Payload, useSchema);
                        const firstContent = Array.isArray(requestBody.contents) ? requestBody.contents[0] : null;
                        const parts = firstContent && Array.isArray(firstContent.parts) ? firstContent.parts : [];
                        const inlineDataPart = parts.find(part => part && part.inlineData && typeof part.inlineData === 'object');
                        const inlineData = inlineDataPart ? inlineDataPart.inlineData : null;
                        if (!inlineData || inlineData.mimeType !== 'application/pdf' || !normalizeSpace(inlineData.data)) {
                            appendImportDebug('Gemini request body invalid: missing PDF payload', { model: candidateModel });
                            throw new Error('Internal error: PDF payload missing from Gemini request');
                        }
                        const payload = await callGeminiGenerateContent({
                            apiKey,
                            model: candidateModel,
                            requestBody
                        });
                        appendImportDebug('Gemini attempt succeeded', { model: candidateModel, attempt: attempt + 1 });
                        return parseGeminiInvoicePayload(payload, 'PDF');
                    } catch (error) {
                        const message = normalizeErrorMessage(error);
                        lastErrorMessage = message || 'Gemini import failed';
                        appendImportDebug('Gemini attempt failed', {
                            model: candidateModel,
                            attempt: attempt + 1,
                            error: lastErrorMessage
                        });

                        if (useSchema && shouldRetryWithoutSchema(message)) {
                            appendImportDebug('Retrying same model without schema', candidateModel);
                            continue;
                        }

                        if (isModelUnavailableMessage(message) && candidateModel !== modelCandidates[modelCandidates.length - 1]) {
                            tryNextModel = true;
                            appendImportDebug('Switching to next model candidate', candidateModel);
                            break;
                        }

                        break;
                    }
                }

                if (!tryNextModel) {
                    break;
                }
            }

            showToast('Gemini failed. Trying local PDF text parser...', 'info', 4500);
            appendImportDebug('Starting local PDF parser fallback');
            try {
                const pdfText = await extractTextFromPdf(file);
                if (normalizeSpace(pdfText)) {
                    appendImportDebug('Local PDF parser extracted text', { textLength: pdfText.length });
                    const parsedFromText = parseInvoiceTextToData(pdfText, file && file.name ? file.name : 'invoice.pdf');
                    if (hasMeaningfulInvoiceData(parsedFromText)) {
                        appendImportDebug('Local PDF parser produced usable invoice data');
                        return parsedFromText;
                    }
                    appendImportDebug('Local PDF parser text did not map to invoice fields');
                }
            } catch (fallbackError) {
                console.warn('Local PDF parser fallback failed', fallbackError);
                appendImportDebug('Local PDF parser fallback failed', normalizeErrorMessage(fallbackError));
            }

            appendImportDebug('Import failed after Gemini and fallback', lastErrorMessage || 'No detailed error');
            throw new Error(lastErrorMessage || 'Gemini import failed and local fallback could not parse the PDF');
        }

        const CHAT_TEMPLATE_PANEL_MOBILE_BREAKPOINT = 1023;
        const CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING = 8;
        const CHAT_TEMPLATE_PANEL_MAX_WIDTH = 380;
        const CHAT_TEMPLATE_PANEL_MIN_VISIBLE_HEIGHT = 140;

        function resetChatTemplatePanelLayout(panel) {
            if (!panel) return;
            panel.style.position = '';
            panel.style.left = '';
            panel.style.right = '';
            panel.style.top = '';
            panel.style.bottom = '';
            panel.style.zIndex = '';
            panel.style.width = '';
            panel.style.maxWidth = '';
            panel.style.maxHeight = '';
            panel.style.transform = '';
        }

        function updateChatTemplatePanelLayout() {
            const panel = document.getElementById('chatTemplatePanel');
            const button = document.getElementById('chatTemplateToggle');
            if (!panel || panel.classList.contains('hidden')) return;

            panel.style.transform = '';

            const viewportWidth = Math.max(window.innerWidth || 0, document.documentElement.clientWidth || 0);
            const viewportHeight = Math.max(window.innerHeight || 0, document.documentElement.clientHeight || 0);
            if (!viewportWidth || !viewportHeight) return;

            if (viewportWidth > CHAT_TEMPLATE_PANEL_MOBILE_BREAKPOINT) {
                resetChatTemplatePanelLayout(panel);
                return;
            }

            const widthCap = Math.max(240, viewportWidth - (CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING * 2));
            const panelWidth = Math.min(CHAT_TEMPLATE_PANEL_MAX_WIDTH, widthCap);
            panel.style.position = 'fixed';
            panel.style.left = 'auto';
            panel.style.right = `${CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING}px`;
            panel.style.bottom = 'auto';
            panel.style.zIndex = '50';
            panel.style.width = `${panelWidth}px`;
            panel.style.maxWidth = `${panelWidth}px`;

            const buttonRect = button ? button.getBoundingClientRect() : null;
            const preferredTop = Math.round((buttonRect ? buttonRect.bottom : CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING) + CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING);
            const maxTop = Math.max(
                CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING,
                viewportHeight - CHAT_TEMPLATE_PANEL_MIN_VISIBLE_HEIGHT - CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING
            );
            const topEdge = Math.min(Math.max(preferredTop, CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING), maxTop);
            panel.style.top = `${topEdge}px`;
            panel.style.maxHeight = `${Math.max(
                CHAT_TEMPLATE_PANEL_MIN_VISIBLE_HEIGHT,
                viewportHeight - topEdge - CHAT_TEMPLATE_PANEL_VIEWPORT_PADDING
            )}px`;
        }

        function handleChatTemplateViewportResize() {
            const panel = document.getElementById('chatTemplatePanel');
            if (!panel || panel.classList.contains('hidden')) return;
            updateChatTemplatePanelLayout();
        }

        window.addEventListener('resize', handleChatTemplateViewportResize);
        if (window.visualViewport && typeof window.visualViewport.addEventListener === 'function') {
            window.visualViewport.addEventListener('resize', handleChatTemplateViewportResize);
        }

        function toggleChatTemplateBubble(eventOrForceOpen) {
            if (eventOrForceOpen && typeof eventOrForceOpen.stopPropagation === 'function') {
                eventOrForceOpen.stopPropagation();
            }
            const panel = document.getElementById('chatTemplatePanel');
            const button = document.getElementById('chatTemplateToggle');
            if (!panel || !button) return;

            const shouldOpen = typeof eventOrForceOpen === 'boolean'
                ? eventOrForceOpen
                : panel.classList.contains('hidden');

            panel.classList.toggle('hidden', !shouldOpen);
            button.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');

            if (shouldOpen) {
                closeSettingsMenu();
                requestAnimationFrame(() => {
                    updateChatTemplatePanelLayout();
                    const input = document.getElementById('chatToTemplateInput');
                    if (input) {
                        input.focus();
                    }
                });
            } else {
                resetChatTemplatePanelLayout(panel);
            }
        }

        function closeChatTemplateBubble() {
            toggleChatTemplateBubble(false);
        }

        function setChatTemplateLoading(isLoading) {
            const button = document.getElementById('chatToTemplateButton');
            const label = document.getElementById('chatToTemplateButtonText');
            const busy = Boolean(isLoading);

            if (button) {
                button.disabled = busy;
                button.classList.toggle('opacity-60', busy);
                button.classList.toggle('cursor-not-allowed', busy);
            }
            if (label) {
                label.textContent = busy ? 'Populating...' : 'Populate Invoice';
            }
        }

        function clearChatTemplateInput() {
            const input = document.getElementById('chatToTemplateInput');
            if (!input) return;
            input.value = '';
            input.focus();
        }

        function handleChatTemplateKeydown(event) {
            if (!event) return;
            if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
                event.preventDefault();
                generateTemplateFromChat();
            }
        }

        async function extractInvoiceDataFromChatWithGemini(userPrompt) {
            const normalizedPrompt = String(userPrompt || '').trim();
            if (!normalizedPrompt) {
                throw new Error('Describe the invoice you want to generate first');
            }
            if (normalizedPrompt.length > CHAT_TO_TEMPLATE_MAX_CHARS) {
                throw new Error(`Prompt is too long. Keep it under ${CHAT_TO_TEMPLATE_MAX_CHARS} characters.`);
            }

            const apiKey = getConfiguredGeminiApiKey();
            const model = getConfiguredGeminiModel();
            appendImportDebug('Gemini chat extraction requested', {
                promptLength: normalizedPrompt.length,
                configuredModel: model
            });

            if (!apiKey) {
                appendImportDebug('Gemini chat extraction stopped: missing API key');
                throw new Error('Enter your Gemini API key in Settings first');
            }

            saveGeminiSettings(false);

            const modelCandidates = getGeminiModelCandidates(model);
            appendImportDebug('Gemini model candidates', modelCandidates.join(', '));
            let lastErrorMessage = '';

            for (const candidateModel of modelCandidates) {
                let tryNextModel = false;
                for (let attempt = 0; attempt < 2; attempt += 1) {
                    const useSchema = attempt === 0;
                    appendImportDebug('Gemini chat attempt', {
                        model: candidateModel,
                        attempt: attempt + 1,
                        useSchema
                    });
                    try {
                        const requestBody = buildGeminiChatInvoiceRequestBody(candidateModel, normalizedPrompt, useSchema);
                        const payload = await callGeminiGenerateContent({
                            apiKey,
                            model: candidateModel,
                            requestBody
                        });
                        appendImportDebug('Gemini chat attempt succeeded', { model: candidateModel, attempt: attempt + 1 });
                        return parseGeminiInvoicePayload(payload, 'chat request');
                    } catch (error) {
                        const message = normalizeErrorMessage(error);
                        lastErrorMessage = message || 'Gemini chat-to-template failed';
                        appendImportDebug('Gemini chat attempt failed', {
                            model: candidateModel,
                            attempt: attempt + 1,
                            error: lastErrorMessage
                        });

                        if (useSchema && shouldRetryWithoutSchema(message)) {
                            appendImportDebug('Retrying chat model without schema', candidateModel);
                            continue;
                        }

                        if (isModelUnavailableMessage(message) && candidateModel !== modelCandidates[modelCandidates.length - 1]) {
                            tryNextModel = true;
                            appendImportDebug('Switching chat request to next model candidate', candidateModel);
                            break;
                        }

                        break;
                    }
                }

                if (!tryNextModel) {
                    break;
                }
            }

            appendImportDebug('Gemini chat extraction failed after all attempts', lastErrorMessage || 'No detailed error');
            const parsedFromText = parseInvoiceTextToData(normalizedPrompt, 'chat-request.txt');
            if (hasMeaningfulInvoiceData(parsedFromText)) {
                appendImportDebug('Local text parser fallback produced usable invoice data for chat input');
                return parsedFromText;
            }
            throw new Error(lastErrorMessage || 'Gemini could not parse enough invoice data from that request');
        }

        async function generateTemplateFromChat() {
            const input = document.getElementById('chatToTemplateInput');
            const button = document.getElementById('chatToTemplateButton');
            if (!input) return;
            if (button && button.disabled) return;

            const prompt = String(input.value || '').trim();
            if (!prompt) {
                showToast('Describe the invoice you want first', 'error');
                input.focus();
                return;
            }

            appendImportDebug('----- New chat-to-template attempt -----');
            appendImportDebug('Chat prompt submitted', { promptLength: prompt.length });
            setChatTemplateLoading(true);
            showToast('Asking Gemini to populate invoice...', 'info');

            try {
                const parsedData = await extractInvoiceDataFromChatWithGemini(prompt);
                const dataWithCompanyDefaults = applyDefaultCompanyFallback(parsedData, 'chat template');
                applyInvoiceDataToForm(dataWithCompanyDefaults);
                appendImportDebug('Chat-to-template completed and form populated');
                showToast('Invoice populated from chat');
            } catch (error) {
                console.error('Chat-to-template failed', error);
                appendImportDebug('Chat-to-template failed', normalizeErrorMessage(error));
                showToast(error && error.message ? error.message : 'Could not populate template from chat', 'error', 12000);
            } finally {
                setChatTemplateLoading(false);
            }
        }

        function htmlToPlainText(html) {
            return String(html || '')
                .replace(/<\s*br\s*\/?>/gi, '\n')
                .replace(/<\/\s*(p|div|li|tr|h1|h2|h3|h4|h5|h6|section|article)\s*>/gi, '\n')
                .replace(/<[^>]+>/g, ' ')
                .replace(/\r/g, '');
        }

        async function extractTextFromPdf(file) {
            if (!window.pdfjsLib || typeof window.pdfjsLib.getDocument !== 'function') {
                throw new Error('PDF import is unavailable in this browser');
            }

            const buffer = await file.arrayBuffer();
            const loadingTask = window.pdfjsLib.getDocument({ data: buffer });
            const pdf = await loadingTask.promise;
            const pages = [];
            appendImportDebug('Local PDF text extraction started', { pages: pdf.numPages });

            for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
                const page = await pdf.getPage(pageNumber);
                const textContent = await page.getTextContent({ disableCombineTextItems: false });
                const linesByY = new Map();
                const inlineItems = [];

                textContent.items.forEach(item => {
                    const text = normalizeSpace(item && item.str ? item.str : '');
                    if (!text) return;

                    inlineItems.push(text);

                    const transform = Array.isArray(item.transform) ? item.transform : [];
                    const y = Math.round(Number(transform[5]) || 0);
                    const x = Number(transform[4]) || 0;
                    if (!linesByY.has(y)) linesByY.set(y, []);
                    linesByY.get(y).push({ x, text });
                });

                const pageLines = Array.from(linesByY.entries())
                    .sort((a, b) => b[0] - a[0])
                    .map(([, segments]) => segments.sort((a, b) => a.x - b.x).map(segment => segment.text).join(' '))
                    .map(normalizeSpace)
                    .filter(Boolean);

                if (!pageLines.length && inlineItems.length) {
                    pageLines.push(normalizeSpace(inlineItems.join(' ')));
                }

                pages.push(pageLines.join('\n'));
            }

            const extractedText = pages.join('\n').trim();
            if (normalizeSpace(extractedText)) {
                appendImportDebug('Local PDF text layer extracted', { textLength: extractedText.length });
                return extractedText;
            }
            appendImportDebug('No selectable text found in PDF');
            return '';
        }

        function parseInvoiceTextToData(rawText, sourceName = '') {
            const normalizedText = String(rawText || '').replace(/\r/g, '\n');
            const lines = normalizedText
                .split('\n')
                .map(normalizeSpace)
                .filter(Boolean);

            const parsed = getDefaultInvoiceData();
            parsed.documentType = normalizeDocumentType(`${sourceName}\n${lines.join('\n')}`);

            const date = findInvoiceDate(lines);
            if (date) parsed.invoiceDate = date;

            const company = extractCompanySection(lines);
            parsed.companyName = company.name;
            parsed.companyDetails = company.details;

            const client = extractClientSection(lines);
            parsed.clientName = client.name;
            parsed.clientDetails = client.details;

            const parsedItems = parseItemsFromLines(lines);
            if (parsedItems.length) {
                parsed.items = parsedItems;
            }

            const totals = parseTotalsFromLines(lines);
            parsed.taxRate = totals.taxRate;
            parsed.discountType = totals.discountType;
            parsed.discountValue = totals.discountValue;
            parsed.notes = extractNotes(lines);

            return normalizeInvoiceData(parsed);
        }

        async function parseUploadedInvoiceFile(file) {
            const name = String(file && file.name ? file.name : '').trim();
            const extension = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
            const type = String(file && file.type ? file.type : '').toLowerCase();
            appendImportDebug('Parsing uploaded file', { name, extension, type });

            if (extension === 'json' || type.includes('json')) {
                appendImportDebug('Import branch selected: JSON');
                const raw = await readFileAsText(file);
                let payload;
                try {
                    payload = JSON.parse(raw);
                } catch (error) {
                    throw new Error('The JSON file is not valid');
                }
                const invoiceLikeData = extractInvoiceDataFromJson(payload);
                if (!invoiceLikeData) {
                    throw new Error('JSON did not contain invoice or bid data');
                }

                const normalized = normalizeInvoiceData(invoiceLikeData);
                if (!hasMeaningfulInvoiceData(normalized)) {
                    throw new Error('JSON file was read, but no invoice fields were recognized');
                }
                return normalized;
            }

            if (extension === 'pdf' || type.includes('pdf')) {
                showToast('Sending PDF to Gemini for extraction...', 'info');
                appendImportDebug('Import branch selected: PDF -> Gemini');
                return extractInvoiceDataFromPdfWithGemini(file);
            }

            appendImportDebug('Import branch selected: text/html');
            let rawText = await readFileAsText(file);
            if (!normalizeSpace(rawText)) {
                throw new Error('The selected file is empty');
            }

            if (extension === 'html' || extension === 'htm' || /<[^>]+>/.test(rawText)) {
                rawText = htmlToPlainText(rawText);
            }

            const parsedFromText = parseInvoiceTextToData(rawText, name);
            if (!hasMeaningfulInvoiceData(parsedFromText)) {
                throw new Error('Could not recognize invoice fields in this file');
            }

            return parsedFromText;
        }

        async function handleInvoiceUpload(input) {
            const file = input && input.files ? input.files[0] : null;
            if (!file) return;
            appendImportDebug('----- New import attempt -----');
            appendImportDebug('File chosen', {
                name: file.name || '',
                size: Number.isFinite(file.size) ? file.size : 0,
                type: file.type || ''
            });

            showToast('Parsing uploaded file...', 'info');

            try {
                const parsedData = await parseUploadedInvoiceFile(file);
                applyInvoiceDataToForm(parsedData);
                appendImportDebug('Import completed and form populated');
                showToast(`Imported ${file.name}`);
            } catch (error) {
                console.error('Import failed', error);
                const fallbackMessage = 'Could not import this file. For PDF import, provide a Gemini API key.';
                appendImportDebug('Import failed', normalizeErrorMessage(error));
                showToast(error && error.message ? error.message : fallbackMessage, 'error', 12000);
            } finally {
                input.value = '';
            }
        }
