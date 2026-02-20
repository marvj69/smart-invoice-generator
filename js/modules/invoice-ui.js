        function parseDescriptionFields(description) {
            const raw = String(description || '').trim();
            if (!raw) {
                return { address: '', work: '' };
            }

            const lines = raw.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
            if (lines.length > 1) {
                return { address: lines[0], work: lines.slice(1).join('\n') };
            }

            const separators = [' | ', ' - ', ' — ', ' – '];
            for (const separator of separators) {
                const index = raw.indexOf(separator);
                if (index > 0) {
                    return {
                        address: raw.slice(0, index).trim(),
                        work: raw.slice(index + separator.length).trim()
                    };
                }
            }

            const zipRegex = /\d{5}(?:-\d{4})?/g;
            let match;
            let lastMatch = null;
            while ((match = zipRegex.exec(raw)) !== null) {
                lastMatch = match;
            }
            if (lastMatch) {
                const splitIndex = lastMatch.index + lastMatch[0].length;
                if (splitIndex < raw.length) {
                    return {
                        address: raw.slice(0, splitIndex).trim(),
                        work: raw.slice(splitIndex).trim()
                    };
                }
            }

            return { address: raw, work: '' };
        }

        function escapeHtml(value) {
            return String(value || '').replace(/[&<>"']/g, (char) => {
                switch (char) {
                    case '&':
                        return '&amp;';
                    case '<':
                        return '&lt;';
                    case '>':
                        return '&gt;';
                    case '"':
                        return '&quot;';
                    case "'":
                        return '&#39;';
                    default:
                        return char;
                }
            });
        }

        function formatDescriptionText(text) {
            const normalized = String(text || '').replace(/([0-9]{5})([A-Za-z])/g, '$1\n$2');
            return escapeHtml(normalized).replace(/\r?\n/g, '<br>');
        }

        function formatMoney(amount) {
            return amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
        }

        function formatDate(dateString) {
            if (!dateString) return '';
            const isoMatch = String(dateString).match(/^(\d{4})-(\d{2})-(\d{2})$/);
            let date;

            if (isoMatch) {
                const year = Number(isoMatch[1]);
                const month = Number(isoMatch[2]) - 1;
                const day = Number(isoMatch[3]);
                date = new Date(year, month, day);
            } else {
                date = new Date(dateString);
            }

            if (Number.isNaN(date.getTime())) {
                return String(dateString);
            }

            return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
        }

        function showToast(message, type = 'success', durationMs = 0) {
            const toast = document.getElementById('toast');
            const messageEl = document.getElementById('toastMessage');
            const icon = toast.querySelector('i');
            
            messageEl.textContent = message;
            
            if (type === 'error') {
                icon.className = 'fas fa-exclamation-circle text-red-400';
            } else if (type === 'info') {
                icon.className = 'fas fa-info-circle text-blue-400';
            } else {
                icon.className = 'fas fa-check-circle text-green-400';
            }
            
            toast.classList.remove('translate-y-20', 'opacity-0');
            if (toastTimerId) {
                clearTimeout(toastTimerId);
                toastTimerId = null;
            }
            const resolvedDuration = Number.isFinite(durationMs) && durationMs > 0
                ? durationMs
                : (type === 'error' ? 9000 : type === 'info' ? 4500 : 3000);
            
            toastTimerId = setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
                toastTimerId = null;
            }, resolvedDuration);
        }

        document.addEventListener('click', function(e) {
            const wrapper = document.getElementById('settingsMenuWrapper');
            const menu = document.getElementById('settingsMenu');
            if (wrapper && menu && !menu.classList.contains('hidden') && !wrapper.contains(e.target)) {
                closeSettingsMenu();
            }

            const chatWidget = document.getElementById('chatTemplateMenuWrapper');
            const chatPanel = document.getElementById('chatTemplatePanel');
            if (chatWidget && chatPanel && !chatPanel.classList.contains('hidden') && !chatWidget.contains(e.target)) {
                closeChatTemplateBubble();
            }
        });

        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                closeSettingsMenu();
                closeMobilePreview();
                closeChatTemplateBubble();
            }
        });

        // Close modal on outside click
        document.getElementById('templateModal').addEventListener('click', function(e) {
            if (e.target === this) closeTemplateManager();
        });

        document.getElementById('previewPane').addEventListener('click', function(e) {
            if (!mobilePreviewOpen) return;
            const wrap = document.getElementById('mobilePreviewScaleWrap');
            if (e.target === this || e.target === wrap) {
                closeMobilePreview();
            }
        });

        window.addEventListener('resize', function() {
            if (!mobilePreviewOpen) return;
            if (!isMobilePreviewViewport()) {
                closeMobilePreview();
                return;
            }
            applyMobilePreviewScale();
        });
