        function normalizeSpace(value) {
            return String(value || '')
                .replace(/\u00A0/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
        }

        function normalizeCitySpacingInLine(value) {
            let line = normalizeSpace(value);
            if (!line) return '';

            line = line.replace(/\s*,\s*/g, ', ');

            // OCR can merge county-road route letters with an uppercase city token:
            // "COUNTY ROAD CKLCHAMPION, MI 49814" -> "COUNTY ROAD CKL CHAMPION, MI 49814"
            line = line.replace(
                /(\b(?:county\s+road|county\s+rd|co\.?\s*rd|cr)\s+)([B-DF-HJ-NP-TV-Z]{3})([B-DF-HJ-NP-TV-Z][A-Z]{3,})(,\s*[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b)/gi,
                '$1$2 $3$4'
            );

            // OCR/LLM output can also merge street suffix + city in uppercase/title case:
            // "286 HEMLOCK STREPUBLIC, MI 49879" -> "286 HEMLOCK ST REPUBLIC, MI 49879"
            line = line.replace(
                /(\b(?:st|street|rd|road|ave|avenue|blvd|boulevard|dr|drive|ln|lane|ct|court|pl|place|ter|terrace|pkwy|parkway|cir|circle|trl|trail|way|hwy|highway)\.?)(?=[A-Z][A-Za-z.'-]{2,}(?:,\s*[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b))/gi,
                '$1 '
            );

            // PDFs/LLM output can merge street suffix + city (for example: "StSpringfield, IL 62704" or "123Main").
            // We enforce a space between a lowercase letter/digit/punct and an Uppercase letter in Title Case,
            // regardless of whether the state/ZIP anchor is present.
            line = line.replace(
                /([a-z0-9#.])([A-Z][a-z])/g,
                (match, p1, p2, offset, str) => {
                    // Avoid inappropriately splitting known prefixes (McAllen, DeKalb, etc.)
                    const preceding = str.substring(0, offset + 1);
                    if (/(?:^|\s)(Mc|Mac|De|Di|La|Le|El|O')$/i.test(preceding)) {
                        return match;
                    }
                    return p1 + ' ' + p2;
                }
            );

            return normalizeSpace(line.replace(/\s*,\s*/g, ', '));
        }

        function extractPhoneFromText(value) {
            const text = String(value || '');
            const match = text.match(/(?:\+?1[\s.\-]*)?(?:\(\d{3}\)|\d{3})[\s.\-]*\d{3}[\s.\-]*\d{4}(?:\s*(?:x|ext\.?)\s*\d+)?/i);
            return match ? normalizeSpace(match[0]) : '';
        }

        function formatPhoneNumber(value) {
            const raw = normalizeSpace(value);
            if (!raw) return '';

            const extMatch = raw.match(/(?:x|ext\.?)\s*(\d+)$/i);
            const extension = extMatch ? extMatch[1] : '';
            let digits = raw.replace(/\D/g, '');

            if (digits.length === 11 && digits.startsWith('1')) {
                digits = digits.slice(1);
            }

            if (digits.length !== 10) {
                return raw;
            }

            const formatted = `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
            return extension ? `${formatted} x${extension}` : formatted;
        }

        function normalizeLocalityLine(value) {
            const raw = normalizeCitySpacingInLine(String(value || '').replace(/^city\s*[:\-]\s*/i, ''));
            if (!raw) return '';

            const cityStateZip = raw.match(/^(.+?),?\s+([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/);
            if (cityStateZip) {
                const city = normalizeSpace(cityStateZip[1].replace(/,\s*$/, ''));
                const state = cityStateZip[2].toUpperCase();
                const zip = cityStateZip[3];
                return `${city}, ${state} ${zip}`;
            }

            const cityZipState = raw.match(/^(.+?),?\s+(\d{5}(?:-\d{4})?)\s+([A-Za-z]{2})$/);
            if (cityZipState) {
                const city = normalizeSpace(cityZipState[1].replace(/,\s*$/, ''));
                const zip = cityZipState[2];
                const state = cityZipState[3].toUpperCase();
                return `${city}, ${state} ${zip}`;
            }

            return raw;
        }

        function normalizeAddressBlock(value, options = {}) {
            const includePhone = options.includePhone !== false;
            const raw = String(value || '').replace(/\r/g, '\n').trim();
            if (!raw) return '';

            const extractedPhone = includePhone ? extractPhoneFromText(raw) : '';
            const formattedPhone = includePhone ? formatPhoneNumber(extractedPhone) : '';

            let textWithoutPhone = raw;
            if (extractedPhone) {
                textWithoutPhone = textWithoutPhone.replace(extractedPhone, ' ');
            }

            textWithoutPhone = textWithoutPhone
                .replace(/\b(?:phone|tel|telephone|mobile|cell)\b\s*[:\-]*/ig, ' ')
                .replace(/[ \t]+\n/g, '\n')
                .replace(/\n[ \t]+/g, '\n');

            const rawLines = textWithoutPhone
                .split(/\n+/)
                .map(line => normalizeCitySpacingInLine(line))
                .filter(Boolean);

            const commaParts = normalizeSpace(textWithoutPhone.replace(/\n/g, ', '))
                .split(/\s*,\s*/)
                .map(part => normalizeSpace(part))
                .filter(Boolean);

            let streetLine = rawLines.find(line => /^\d+\s+/.test(line)) || '';
            let localityLine = rawLines.find(line => {
                return /\d{5}(?:-\d{4})?/.test(line)
                    || /,\s*[A-Za-z]{2}\b/.test(line)
                    || /\b[A-Za-z]{2}\s+\d{5}(?:-\d{4})?\b/.test(line);
            }) || '';

            if (!streetLine && commaParts.length) {
                streetLine = commaParts.find(part => /^\d+\s+/.test(part)) || commaParts[0] || '';
            }

            if (!localityLine && commaParts.length >= 2) {
                localityLine = commaParts.slice(1).join(', ');
            }

            if ((!streetLine || !localityLine) && rawLines.length === 1) {
                const singleLine = rawLines[0];
                const oneLineMatch = singleLine.match(/^(\d+\s+[^,]+),?\s+(.+)$/);
                if (oneLineMatch) {
                    if (!streetLine) streetLine = normalizeSpace(oneLineMatch[1]);
                    if (!localityLine) localityLine = normalizeSpace(oneLineMatch[2]);
                }
            }

            if (!streetLine) {
                streetLine = rawLines[0] || '';
            }

            streetLine = normalizeCitySpacingInLine(streetLine);
            localityLine = normalizeCitySpacingInLine(localityLine);

            const suiteLine = rawLines.find(line => {
                if (line === streetLine || line === localityLine) return false;
                return /\b(?:apt|apartment|suite|ste|unit|#)\b/i.test(line);
            });

            if (suiteLine && streetLine && !streetLine.includes(suiteLine)) {
                streetLine = `${streetLine}, ${suiteLine}`;
            }

            if (!localityLine) {
                localityLine = rawLines.find(line => line !== streetLine && line !== suiteLine) || '';
            }

            const normalizedLocality = normalizeLocalityLine(localityLine);
            const extraLines = rawLines.filter(line => {
                if (line === streetLine) return false;
                if (line === localityLine) return false;
                if (line === suiteLine) return false;
                return true;
            });

            const outputLines = [];
            if (streetLine) outputLines.push(streetLine);
            if (normalizedLocality && normalizedLocality !== streetLine) outputLines.push(normalizedLocality);
            extraLines.forEach(line => outputLines.push(line));
            if (formattedPhone) outputLines.push(formattedPhone);

            return outputLines.join('\n').trim();
        }

        function toNumber(value, fallback = 0) {
            if (typeof value === 'number') {
                return Number.isFinite(value) ? value : fallback;
            }

            const cleaned = String(value || '')
                .replace(/\u2212/g, '-')
                .replace(/[^0-9.\-]/g, '');
            const parsed = parseFloat(cleaned);
            return Number.isFinite(parsed) ? parsed : fallback;
        }

        function toCurrencyNumber(value, fallback = 0) {
            const raw = String(value || '').trim();
            if (!raw) return fallback;

            const isWrappedNegative = /^\(.*\)$/.test(raw);
            const parsed = toNumber(raw, fallback);
            return isWrappedNegative ? -Math.abs(parsed) : parsed;
        }

        function normalizeDocumentType(value) {
            return /\bbid\b/i.test(String(value || '')) ? 'Bid' : 'Invoice';
        }

        function toISODateString(value) {
            const raw = normalizeSpace(String(value || '').replace(/(\d+)(st|nd|rd|th)\b/gi, '$1'));
            if (!raw) return '';

            const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
            if (isoMatch) {
                const year = Number(isoMatch[1]);
                const month = Number(isoMatch[2]);
                const day = Number(isoMatch[3]);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                }
            }

            const slashMatch = raw.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
            if (slashMatch) {
                const month = Number(slashMatch[1]);
                const day = Number(slashMatch[2]);
                const yearValue = Number(slashMatch[3]);
                const year = yearValue < 100 ? 2000 + yearValue : yearValue;
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                }
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

        function findDateToken(value) {
            const text = String(value || '');
            const patterns = [
                /\b\d{4}-\d{1,2}-\d{1,2}\b/,
                /\b\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\b/,
                /\b(?:Jan|January|Feb|February|Mar|March|Apr|April|May|Jun|June|Jul|July|Aug|August|Sep|Sept|September|Oct|October|Nov|November|Dec|December)\s+\d{1,2},?\s+\d{4}\b/i
            ];

            for (const pattern of patterns) {
                const match = text.match(pattern);
                if (match) {
                    const iso = toISODateString(match[0]);
                    if (iso) return iso;
                }
            }

            return '';
        }

        function findItemHeaderIndex(lines) {
            return lines.findIndex(line => /\bdescription\b/i.test(line) && /\bqty\b/i.test(line));
        }

        function isTotalsLine(line) {
            return /^(subtotal|tax\b|discount\b|total\b)/i.test(line);
        }

        function findInvoiceDate(lines) {
            for (let index = 0; index < lines.length; index += 1) {
                const line = lines[index];
                if (!/^date\b/i.test(line)) continue;

                const direct = line.replace(/^date\s*[:\-]?\s*/i, '');
                const directDate = findDateToken(direct) || toISODateString(direct);
                if (directDate) return directDate;

                const nextLine = lines[index + 1];
                const nextDate = findDateToken(nextLine) || toISODateString(nextLine);
                if (nextDate) return nextDate;
            }

            for (const line of lines) {
                const guessed = findDateToken(line);
                if (guessed) return guessed;
            }

            return '';
        }

        function extractCompanySection(lines) {
            const billToIndex = lines.findIndex(line => /^bill\s*to\b/i.test(line) || /^client\b/i.test(line));
            const dateIndex = lines.findIndex(line => /^date\b/i.test(line));
            const itemHeaderIndex = findItemHeaderIndex(lines);
            const docTypeIndex = lines.findIndex(line => /^(invoice|bid)\b/i.test(line));

            const endCandidates = [billToIndex, dateIndex, itemHeaderIndex, docTypeIndex].filter(index => index > 0);
            const sectionEnd = endCandidates.length ? Math.min(...endCandidates) : Math.min(lines.length, 4);

            const linesInSection = lines.slice(0, sectionEnd).filter(line => {
                return !/^date\b/i.test(line) && !/^(invoice|bid)\b/i.test(line);
            });

            if (!linesInSection.length) {
                return { name: '', details: '' };
            }

            return {
                name: linesInSection[0],
                details: linesInSection.slice(1).join('\n')
            };
        }

        function extractClientSection(lines) {
            const billToIndex = lines.findIndex(line => /^bill\s*to\b/i.test(line) || /^client\b/i.test(line));
            if (billToIndex < 0) {
                return { name: '', details: '' };
            }

            const itemHeaderIndex = findItemHeaderIndex(lines);
            const totalsIndex = lines.findIndex((line, idx) => idx > billToIndex && isTotalsLine(line));
            const notesIndex = lines.findIndex((line, idx) => idx > billToIndex && /^notes?\b/i.test(line));
            const endCandidates = [itemHeaderIndex, totalsIndex, notesIndex].filter(index => index > billToIndex);
            const sectionEnd = endCandidates.length ? Math.min(...endCandidates) : Math.min(lines.length, billToIndex + 6);
            const clientLines = lines
                .slice(billToIndex + 1, sectionEnd)
                .filter(line => !/^date\b/i.test(line) && !/^(description|qty|rate|amount)\b/i.test(line));

            if (!clientLines.length) {
                return { name: '', details: '' };
            }

            return {
                name: clientLines[0],
                details: clientLines.slice(1).join('\n')
            };
        }

        function parseItemsFromLines(lines) {
            const headerIndex = findItemHeaderIndex(lines);
            const workingLines = headerIndex >= 0 ? lines.slice(headerIndex + 1) : lines.slice();
            const items = [];
            let descriptionBuffer = [];

            const parseDescriptionBuffer = (fallbackText = '') => {
                const raw = [...descriptionBuffer, fallbackText].join('\n').trim();
                descriptionBuffer = [];
                if (!raw) return null;

                const split = parseDescriptionFields(raw);
                return normalizeLineItem({
                    description: raw,
                    address: split.address,
                    work: split.work,
                    quantity: 1,
                    rate: 0
                });
            };

            for (const line of workingLines) {
                if (!line) continue;
                if (isTotalsLine(line) || /^notes?\b/i.test(line)) break;
                if (/^(description|qty|rate|amount)\b/i.test(line)) continue;

                const normalizedLine = normalizeSpace(line);
                const fullPattern = normalizedLine.match(/^(.*\S)\s+(-?\d+(?:\.\d+)?)\s+\$?(-?[\d,]+(?:\.\d{1,2})?)\s+\$?(-?[\d,]+(?:\.\d{1,2})?)$/);
                if (fullPattern) {
                    const parsedDescription = parseDescriptionBuffer(fullPattern[1]) || normalizeLineItem({
                        description: fullPattern[1],
                        quantity: toNumber(fullPattern[2], 0),
                        rate: toCurrencyNumber(fullPattern[3], 0)
                    });

                    parsedDescription.quantity = Math.max(0, toNumber(fullPattern[2], 0));
                    parsedDescription.rate = Math.max(0, toCurrencyNumber(fullPattern[3], 0));
                    items.push(parsedDescription);
                    continue;
                }

                const compactPattern = normalizedLine.match(/^(.*\S)\s+(-?\d+(?:\.\d+)?)\s+\$?(-?[\d,]+(?:\.\d{1,2})?)$/);
                if (compactPattern && descriptionBuffer.length) {
                    const parsedDescription = parseDescriptionBuffer(compactPattern[1]) || normalizeLineItem({
                        description: compactPattern[1],
                        quantity: toNumber(compactPattern[2], 0),
                        rate: toCurrencyNumber(compactPattern[3], 0)
                    });

                    parsedDescription.quantity = Math.max(0, toNumber(compactPattern[2], 0));
                    parsedDescription.rate = Math.max(0, toCurrencyNumber(compactPattern[3], 0));
                    items.push(parsedDescription);
                    continue;
                }

                descriptionBuffer.push(normalizedLine);
            }

            if (descriptionBuffer.length) {
                const bufferedItem = parseDescriptionBuffer();
                if (bufferedItem) items.push(bufferedItem);
            }

            return items.filter(item => {
                return item.address || item.work || item.description || item.quantity || item.rate;
            });
        }

        function parseTotalsFromLines(lines) {
            let taxRate = 0;
            let discountType = 'fixed';
            let discountValue = 0;

            for (const line of lines) {
                if (/^tax\b/i.test(line)) {
                    const taxMatch = line.match(/(-?\d+(?:\.\d+)?)\s*%/);
                    if (taxMatch) {
                        taxRate = Math.max(0, toNumber(taxMatch[1], 0));
                    }
                }

                if (/^discount\b/i.test(line)) {
                    const discountPercent = line.match(/(-?\d+(?:\.\d+)?)\s*%/);
                    if (discountPercent) {
                        discountType = 'percentage';
                        discountValue = Math.max(0, toNumber(discountPercent[1], 0));
                        continue;
                    }

                    const amountCandidates = line.match(/-?\$?\(?-?[\d,]+(?:\.\d{1,2})?\)?/g);
                    if (amountCandidates && amountCandidates.length) {
                        const parsedAmount = toCurrencyNumber(amountCandidates[amountCandidates.length - 1], 0);
                        discountType = 'fixed';
                        discountValue = Math.max(0, Math.abs(parsedAmount));
                    }
                }
            }

            return { taxRate, discountType, discountValue };
        }

        function extractNotes(lines) {
            const notesIndex = lines.findIndex(line => /^notes?\b/i.test(line) || /^additional notes\b/i.test(line));
            if (notesIndex < 0) return '';
            return lines.slice(notesIndex + 1).join('\n').trim();
        }

        function normalizeLineItem(item) {
            const source = item && typeof item === 'object' ? item : {};
            let address = normalizeAddressBlock(source.address || source.propertyAddress || '', { includePhone: false });
            let work = String(source.work || source.workDone || '').trim();
            const rawDescription = String(source.description || source.item || '').trim();

            if ((!address || !work) && rawDescription) {
                const parsed = parseDescriptionFields(rawDescription);
                if (!address) address = normalizeAddressBlock(parsed.address || '', { includePhone: false });
                if (!work) work = parsed.work;
            }

            const description = [address, work].filter(Boolean).join('\n') || rawDescription;
            const quantitySource = source.quantity ?? source.qty ?? source.hours ?? source.units ?? 1;
            const rateSource = source.rate ?? source.price ?? source.unitPrice ?? source.amount ?? 0;

            return {
                description,
                address,
                work,
                quantity: Math.max(0, toNumber(quantitySource, 1)),
                rate: Math.max(0, toCurrencyNumber(rateSource, 0))
            };
        }

        function normalizeInvoiceData(data) {
            const source = data && typeof data === 'object' ? data : {};
            const normalized = getDefaultInvoiceData();

            normalized.documentType = normalizeDocumentType(source.documentType || source.type || source.docType || 'Invoice');
            normalized.companyName = String(source.companyName || source.fromName || source.businessName || source.vendorName || '').trim();
            normalized.companyDetails = normalizeAddressBlock(source.companyDetails || source.fromDetails || source.businessDetails || source.vendorDetails || '', { includePhone: true });
            normalized.clientName = String(source.clientName || source.customerName || source.billToName || '').trim();
            normalized.clientDetails = normalizeAddressBlock(source.clientDetails || source.customerDetails || source.billToDetails || '', { includePhone: true });
            normalized.logo = typeof source.logo === 'string' && source.logo ? source.logo : null;
            normalized.notes = String(source.notes || source.note || source.terms || '').trim();

            const parsedDate = toISODateString(source.invoiceDate || source.date || source.issueDate || source.createdAt || '');
            normalized.invoiceDate = parsedDate || normalized.invoiceDate;

            normalized.taxRate = Math.max(0, toNumber(source.taxRate ?? source.tax ?? source.vatRate ?? 0, 0));

            const sourceDiscountType = String(source.discountType || '').toLowerCase();
            normalized.discountType = ['percentage', 'percent', '%'].includes(sourceDiscountType) ? 'percentage' : 'fixed';
            normalized.discountValue = Math.max(0, toCurrencyNumber(source.discountValue ?? source.discount ?? 0, 0));

            const itemsSource = Array.isArray(source.items) ? source.items
                : Array.isArray(source.lineItems) ? source.lineItems
                : Array.isArray(source.services) ? source.services
                : [];

            const normalizedItems = itemsSource.map(normalizeLineItem).filter(item => {
                return item.address || item.work || item.description || item.quantity || item.rate;
            });

            normalized.items = normalizedItems.length ? normalizedItems : [getDefaultLineItem()];
            return normalized;
        }

        function hasMeaningfulInvoiceData(data) {
            if (!data || typeof data !== 'object') return false;
            if (data.companyName || data.companyDetails || data.clientName || data.clientDetails || data.notes) return true;
            if ((data.taxRate || 0) > 0 || (data.discountValue || 0) > 0) return true;
            if (!Array.isArray(data.items)) return false;

            return data.items.some(item => {
                if (!item || typeof item !== 'object') return false;
                return Boolean(
                    String(item.address || '').trim() ||
                    String(item.work || '').trim() ||
                    String(item.description || '').trim() ||
                    toNumber(item.quantity, 0) > 0 ||
                    toCurrencyNumber(item.rate, 0) > 0
                );
            });
        }

        function applyInvoiceDataToForm(data) {
            invoiceData = normalizeInvoiceData(data);

            document.getElementById('companyName').value = invoiceData.companyName;
            document.getElementById('companyDetails').value = invoiceData.companyDetails;
            document.getElementById('documentType').value = invoiceData.documentType;
            document.getElementById('invoiceDate').value = invoiceData.invoiceDate;
            document.getElementById('clientName').value = invoiceData.clientName;
            document.getElementById('clientDetails').value = invoiceData.clientDetails;
            document.getElementById('taxRate').value = invoiceData.taxRate;
            document.getElementById('discountType').value = invoiceData.discountType;
            document.getElementById('discountValue').value = invoiceData.discountValue;
            document.getElementById('notes').value = invoiceData.notes;

            const container = document.getElementById('lineItemsContainer');
            container.innerHTML = '';
            const itemsToRender = Array.isArray(invoiceData.items) && invoiceData.items.length
                ? invoiceData.items
                : [getDefaultLineItem()];

            itemsToRender.forEach(item => addLineItem(item, false));
            updateInvoice();
        }

        function isLikelyInvoicePayload(candidate) {
            if (!candidate || typeof candidate !== 'object') return false;

            const hasItemCollection = Array.isArray(candidate.items) || Array.isArray(candidate.lineItems) || Array.isArray(candidate.services);
            const hintKeys = ['companyName', 'clientName', 'invoiceDate', 'documentType', 'taxRate', 'discountValue', 'notes'];
            const hasHintKey = hintKeys.some(key => Object.prototype.hasOwnProperty.call(candidate, key));
            return hasItemCollection || hasHintKey;
        }

        function extractInvoiceDataFromJson(payload, depth = 0) {
            if (!payload || depth > 3) return null;

            if (isLikelyInvoicePayload(payload)) {
                return payload;
            }

            if (Array.isArray(payload)) {
                for (const entry of payload) {
                    const found = extractInvoiceDataFromJson(entry, depth + 1);
                    if (found) return found;
                }
                return null;
            }

            const preferredKeys = ['data', 'invoice', 'bid', 'document', 'template', 'templates', 'records', 'payload'];
            for (const key of preferredKeys) {
                if (payload[key] === undefined) continue;
                const found = extractInvoiceDataFromJson(payload[key], depth + 1);
                if (found) return found;
            }

            for (const value of Object.values(payload)) {
                if (!value || typeof value !== 'object') continue;
                const found = extractInvoiceDataFromJson(value, depth + 1);
                if (found) return found;
            }

            return null;
        }

