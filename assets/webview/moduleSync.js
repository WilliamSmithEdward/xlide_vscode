        const vscode = acquireVsCodeApi();
        let plan = {{planJson}};
        let selected = new Set();
        let activeId = plan.items[0]?.id;
        let applying = false;
        let applied = false;
        let showHeaders = false;
        let settingsSaving = false;
        let settingsSaveTimer;
        let tooltipTimer;
        let tooltipTarget;
        const tooltipDelayMs = 140;
        const listNodes = new Map();

        // Virtualized diff state. Only the rows inside the viewport (plus a
        // small overscan) exist in the DOM; heights start as estimates and are
        // replaced with real measurements as rows scroll into view.
        const DIFF_OVERSCAN = 12;
        let diffLines = [];
        let diffHeights = [];
        let diffOffsets = [];
        let diffCanvas = null;
        let diffNodes = new Map();
        let diffCharsPerCol = 100;
        let diffLineHeight = 18;
        let diffRowPadding = 5;
        let diffRenderedKey = null;
        let diffResizeFrame = 0;

        const el = id => document.getElementById(id);

        function clamp(value, min, max) {
            return Math.max(min, Math.min(max, value));
        }

        function setListWidth(width) {
            const layout = el('layout');
            const bounds = layout.getBoundingClientRect();
            const max = Math.max(320, bounds.width - 360);
            layout.style.setProperty('--list-width', `${clamp(width, 280, max)}px`);
        }

        function installSplitter() {
            const splitter = el('splitter');
            const layout = el('layout');
            splitter.addEventListener('pointerdown', event => {
                event.preventDefault();
                splitter.setPointerCapture(event.pointerId);
                splitter.classList.add('dragging');
                document.body.classList.add('resizing');
            });
            splitter.addEventListener('pointermove', event => {
                if (!splitter.hasPointerCapture(event.pointerId)) return;
                const bounds = layout.getBoundingClientRect();
                setListWidth(event.clientX - bounds.left);
            });
            function stopDrag(event) {
                if (splitter.hasPointerCapture(event.pointerId)) {
                    splitter.releasePointerCapture(event.pointerId);
                }
                splitter.classList.remove('dragging');
                document.body.classList.remove('resizing');
            }
            splitter.addEventListener('pointerup', stopDrag);
            splitter.addEventListener('pointercancel', stopDrag);
        }

        function selectedFromPlan() {
            return new Set(plan.items.filter(item => item.checked && item.selectable).map(item => item.id));
        }

        function isRelevantItem(item) {
            return item.selectable && item.status !== 'unchanged' && !item.status.startsWith('skipping');
        }

        function statusTone(item) {
            if (item.status === 'will-create') return 'create';
            if (item.status === 'will-remove') return 'remove';
            if (item.status === 'will-write' || item.status === 'will-update') return 'write';
            if (item.status === 'unchanged') return 'same';
            if (item.status.startsWith('skipping')) return 'skip';
            return 'error';
        }

        function copyTooltip(item, side, hasCode) {
            if (hasCode) {
                return side === 'left' ? 'Copy left code to clipboard.' : 'Copy right code to clipboard.';
            }
            const title = side === 'left' ? item.leftTitle : item.rightTitle;
            if (title.startsWith('Repo:')) {
                return item.existsInRepo ? 'No repo file code to copy.' : 'Repo file does not exist yet.';
            }
            if (title.startsWith('Workbook:')) {
                return item.existsInWorkbook ? 'No workbook module code to copy.' : 'Workbook module does not exist yet.';
            }
            return side === 'left' ? 'No left-side code to copy.' : 'No right-side code to copy.';
        }

        function shouldShowWarnings() {
            return plan.warnings.some(warning => !warning.includes('skipping import unless the module already exists in the workbook'));
        }

        function currentSettings() {
            return {
                folderPath: plan.folderPath,
                exportMode: plan.direction === 'export' ? el('syncMode').value : undefined,
                importMode: plan.direction === 'import' ? el('syncMode').value : undefined,
            };
        }

        function option(value, label, description) {
            const item = document.createElement('option');
            item.value = value;
            item.textContent = label;
            if (description) {
                item.title = description;
            }
            return item;
        }

        function setTooltip(targetOrId, text) {
            const target = typeof targetOrId === 'string' ? el(targetOrId) : targetOrId;
            if (!target) return;
            target.dataset.tooltip = text;
            target.removeAttribute('title');
        }

        function clearTooltip(targetOrId) {
            const target = typeof targetOrId === 'string' ? el(targetOrId) : targetOrId;
            if (!target) return;
            delete target.dataset.tooltip;
            target.removeAttribute('title');
        }

        function clearTooltipTimer() {
            if (tooltipTimer) {
                clearTimeout(tooltipTimer);
                tooltipTimer = undefined;
            }
        }

        function hideTooltip() {
            clearTooltipTimer();
            tooltipTarget = undefined;
            const tooltip = el('tooltip');
            tooltip.classList.remove('visible');
            tooltip.hidden = true;
        }

        function showTooltipFor(target) {
            const text = target.dataset.tooltip;
            if (!text) return;
            const tooltip = el('tooltip');
            tooltip.textContent = text;
            tooltip.hidden = false;
            tooltip.classList.remove('visible');
            const rect = target.getBoundingClientRect();
            const pad = 8;
            const tooltipWidth = tooltip.offsetWidth;
            const tooltipHeight = tooltip.offsetHeight;
            const maxLeft = Math.max(pad, window.innerWidth - tooltipWidth - pad);
            const left = clamp(rect.left, pad, maxLeft);
            let top = rect.bottom + 6;
            if (top + tooltipHeight > window.innerHeight - pad) {
                top = Math.max(pad, rect.top - tooltipHeight - 6);
            }
            tooltip.style.left = `${left}px`;
            tooltip.style.top = `${top}px`;
            requestAnimationFrame(() => tooltip.classList.add('visible'));
        }

        function installFastTooltips() {
            document.addEventListener('mouseover', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (!target) return;
                tooltipTarget = target;
                clearTooltipTimer();
                tooltipTimer = setTimeout(() => {
                    if (tooltipTarget === target) {
                        showTooltipFor(target);
                    }
                }, tooltipDelayMs);
            });
            document.addEventListener('mouseout', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (!target) return;
                if (event.relatedTarget && target.contains(event.relatedTarget)) return;
                hideTooltip();
            });
            document.addEventListener('focusin', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (target) {
                    tooltipTarget = target;
                    showTooltipFor(target);
                }
            });
            document.addEventListener('focusout', () => hideTooltip());
            document.addEventListener('pointerdown', event => {
                if (event.target.closest?.('[data-tooltip]')) {
                    hideTooltip();
                }
            });
            document.addEventListener('click', event => {
                if (event.target.closest?.('[data-tooltip]')) {
                    hideTooltip();
                }
            });
            document.addEventListener('scroll', () => hideTooltip(), true);
            window.addEventListener('resize', () => hideTooltip());
        }

        function modeDescription(modeValue) {
            const mode = modeValue || el('syncMode').value;
            if (plan.direction === 'export') {
                return mode === 'trueUp'
                    ? 'Export every workbook module to the selected folder, then delete stale .bas/.cls module files that no longer exist in the workbook.'
                    : 'Export every workbook module to the selected folder. XLIDE will create missing module files and update changed files, but will not delete stale files.';
            }
            return mode === 'trueUpStandardClass'
                ? 'Import/update selected .bas/.cls files, then delete workbook-only standard/class modules missing from the folder. New standard/class modules can be created; existing document modules and UserForm .cls code-behind are updated on name match; document modules and UserForm code-behind are never created or deleted by this mode.'
                : 'Import/update selected .bas/.cls files without deleting workbook modules. New standard/class modules can be created; existing document modules and UserForm .cls code-behind are updated on name match; missing document modules and UserForm code-behind are skipped because XLIDE cannot create them directly.';
        }

        function settingsSourceLabel(source) {
            if (source === 'workbook') return 'Workbook override';
            if (source === 'session') return 'Current session';
            if (source === 'machine') return 'VS Code machine setting';
            if (source === 'unknown') return 'Unknown';
            return 'Built-in default';
        }

        function folderSourceLabel(source) {
            if (source === 'workbook') return 'Workbook sidecar';
            if (source === 'session') return 'Current session';
            return 'Not saved';
        }

        function settingsPathDescription() {
            return plan.settingsPath ? ` Settings file: ${plan.settingsPath}` : '';
        }

        function updateModeTitle() {
            clearTooltip('syncMode');
            clearTooltip('modeLabel');
            clearTooltip('modeField');
        }

        function renderChrome() {
            el('title').textContent = plan.title;
            el('subtitle').textContent = `${plan.workbookPath} <-> ${plan.folderPath}${plan.exportMode ? '  [' + plan.exportMode + ']' : ''}`;
            el('folderLabel').textContent = plan.direction === 'export' ? 'Export folder' : 'Import folder';
            el('folderValue').textContent = plan.folderPath;
            el('folderSource').textContent = `Source: ${folderSourceLabel(plan.folderPathSource)}`;
            setTooltip('folderValue', plan.direction === 'export'
                ? `Folder XLIDE will compare against and write selected workbook modules into: ${plan.folderPath}`
                : `Folder XLIDE will compare against and import selected module files from: ${plan.folderPath}`);
            setTooltip('folderSource', `This folder is workbook-scoped.${settingsPathDescription()}`);
            setTooltip('chooseFolder', plan.direction === 'export'
                ? 'Choose the folder to compare with this workbook and receive exported module files.'
                : 'Choose the folder containing module files to compare with and import into this workbook.');
            el('modeField').classList.remove('hidden');
            const mode = el('syncMode');
            mode.innerHTML = '';
            if (plan.direction === 'export') {
                el('modeLabel').textContent = 'Export mode';
                mode.append(option('exportAll', 'Export All (No Deletes)', modeDescription('exportAll')));
                mode.append(option('trueUp', 'Export All + Delete Missing', modeDescription('trueUp')));
                mode.value = plan.exportMode || 'exportAll';
            } else {
                el('modeLabel').textContent = 'Import mode';
                mode.append(option('updateOnly', 'Import/Update (No Deletes)', modeDescription('updateOnly')));
                mode.append(option('trueUpStandardClass', 'Import/Update + Delete Missing', modeDescription('trueUpStandardClass')));
                mode.value = plan.importMode || 'updateOnly';
            }
            const modeSource = plan.direction === 'export' ? plan.exportModeSource : plan.importModeSource;
            el('modeSource').textContent = `Source: ${settingsSourceLabel(modeSource)}`;
            setTooltip('modeSource', `This mode uses a workbook override when present, otherwise the built-in default.${settingsPathDescription()}`);
            updateModeTitle();
            el('selectChanged').textContent = 'Select Pending';
            setTooltip('selectChanged', plan.direction === 'import'
                ? 'Select every pending import row that will create, update, or delete a workbook module under the current import mode.'
                : 'Select every pending export row that will create, overwrite, or remove files under the current export mode.');
            setTooltip('clear', 'Clear the current module selection without changing files.');
            setTooltip('apply', plan.direction === 'import'
                ? 'Apply the selected import changes to the workbook.'
                : 'Apply the selected export changes to the folder.');
            setTooltip('cancel', 'Close this preview without applying changes.');
            setTooltip('toggleHeaders', 'Toggle hidden VBA Attribute header lines in the diff preview and copy buttons.');
            if (shouldShowWarnings()) {
                el('warnings').classList.add('visible');
                el('warnings').textContent = plan.warnings.join('\n');
            } else {
                el('warnings').classList.remove('visible');
                el('warnings').textContent = '';
            }
        }

        function setPlan(nextPlan, message, autoSaveSettings) {
            plan = nextPlan;
            selected = selectedFromPlan();
            activeId = plan.items[0]?.id;
            applying = false;
            applied = false;
            el('apply').textContent = 'Apply Selected';
            el('result').textContent = message || '';
            renderChrome();
            renderList();
            // A refreshed plan can reuse the same item id with new content, so
            // the cached key must not suppress the rebuild.
            renderDiff(true);
            if (autoSaveSettings) {
                scheduleSettingsAutosave();
            }
        }

        function renderList() {
            const list = el('list');
            list.innerHTML = '';
            listNodes.clear();
            for (const item of plan.items) {
                const tone = statusTone(item);
                const row = document.createElement('div');
                row.className = `item status-${tone}`;
                row.dataset.id = item.id;
                const checkHit = document.createElement('div');
                checkHit.className = 'checkHit' + (!item.selectable ? ' disabled' : '');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.disabled = !item.selectable;
                checkHit.addEventListener('click', event => {
                    event.stopPropagation();
                    if (!item.selectable) return;
                    activeId = item.id;
                    if (selected.has(item.id)) selected.delete(item.id);
                    else selected.add(item.id);
                    syncListState();
                    renderDiff();
                });
                checkHit.append(checkbox);
                const text = document.createElement('div');
                text.className = 'itemText';
                const name = document.createElement('div');
                name.className = 'name';
                name.textContent = item.moduleName;
                const meta = document.createElement('div');
                meta.className = 'meta';
                meta.textContent = [item.relativeName, item.moduleType, item.warning].filter(Boolean).join(' | ');
                text.append(name, meta);
                const badge = document.createElement('span');
                badge.className = `badge ${tone}`;
                badge.textContent = item.detail || item.status;
                badge.title = item.warning || item.detail || item.status;
                row.append(checkHit, text, badge);
                row.addEventListener('click', () => {
                    activeId = item.id;
                    syncListState();
                    renderDiff();
                });
                list.append(row);
                listNodes.set(item.id, { row, checkbox });
            }
            syncListState();
        }

        /*
         * Selection and active-row changes only touch attributes on rows that
         * already exist - rebuilding the list on every click also rebuilt every
         * listener, and paired with a full diff rebuild made each click O(whole
         * workbook).
         */
        function syncListState() {
            for (const item of plan.items) {
                const node = listNodes.get(item.id);
                if (!node) continue;
                const isActive = item.id === activeId;
                node.row.classList.toggle('active', isActive);
                node.row.setAttribute('aria-selected', isActive ? 'true' : 'false');
                node.checkbox.checked = selected.has(item.id);
            }
            renderCounts();
        }

        function buildDiffRow(line) {
            const row = document.createElement('div');
            row.className = 'line ' + line.kind;
            const leftNo = document.createElement('div');
            leftNo.className = 'ln';
            leftNo.textContent = line.leftNumber || '';
            const left = document.createElement('pre');
            left.className = 'left';
            left.textContent = line.left;
            const rightNo = document.createElement('div');
            rightNo.className = 'ln';
            rightNo.textContent = line.rightNumber || '';
            const right = document.createElement('pre');
            right.className = 'right';
            right.textContent = line.right;
            row.append(leftNo, left, rightNo, right);
            return row;
        }

        /*
         * Column width and line height drive the height estimate for rows that
         * have not been rendered (and therefore never measured) yet. The probe
         * must lay out in flow: an absolutely positioned grid row shrink-to-fits
         * and reports a useless column width.
         */
        function measureDiffMetrics() {
            const probe = buildDiffRow({ kind: '', leftNumber: '1', left: 'M', rightNumber: '', right: '' });
            probe.style.position = 'static';
            probe.style.visibility = 'hidden';
            diffCanvas.append(probe);
            const pre = probe.querySelector('pre.left');
            const style = getComputedStyle(pre);
            const measurer = document.createElement('canvas').getContext('2d');
            measurer.font = style.fontSize + ' ' + style.fontFamily;
            const charWidth = measurer.measureText('M').width || 8;
            const usable = pre.clientWidth - 16;
            diffCharsPerCol = Math.max(8, Math.floor(usable / charWidth));
            diffLineHeight = Math.max(12, parseFloat(style.lineHeight) || 16);
            diffRowPadding = Math.max(0, probe.offsetHeight - diffLineHeight);
            probe.remove();
        }

        function estimateDiffHeight(line) {
            const wraps = Math.max(
                Math.ceil((line.left || '').length / diffCharsPerCol),
                Math.ceil((line.right || '').length / diffCharsPerCol),
                1);
            return wraps * diffLineHeight + diffRowPadding;
        }

        function recomputeDiffOffsets() {
            let top = 0;
            for (let i = 0; i < diffHeights.length; i++) {
                diffOffsets[i] = top;
                top += diffHeights[i];
            }
            diffOffsets.length = diffHeights.length;
            diffCanvas.style.height = top + 'px';
        }

        function firstDiffRowAt(scrollTop) {
            let lo = 0;
            let hi = diffHeights.length - 1;
            let best = 0;
            while (lo <= hi) {
                const mid = (lo + hi) >> 1;
                if (diffOffsets[mid] <= scrollTop) { best = mid; lo = mid + 1; }
                else { hi = mid - 1; }
            }
            return best;
        }

        function paintDiffWindow() {
            if (!diffCanvas || !diffLines.length) return;
            const diff = el('diff');
            const viewTop = diff.scrollTop;
            const viewBottom = viewTop + diff.clientHeight;
            const first = Math.max(0, firstDiffRowAt(viewTop) - DIFF_OVERSCAN);
            let last = first;
            while (last < diffLines.length && diffOffsets[last] < viewBottom) last++;
            last = Math.min(diffLines.length - 1, last + DIFF_OVERSCAN);

            for (const [index, node] of diffNodes) {
                if (index < first || index > last) {
                    node.remove();
                    diffNodes.delete(index);
                }
            }
            const fragment = document.createDocumentFragment();
            const fresh = [];
            for (let i = first; i <= last; i++) {
                if (diffNodes.has(i)) continue;
                const node = buildDiffRow(diffLines[i]);
                diffNodes.set(i, node);
                fragment.append(node);
                fresh.push(i);
            }
            diffCanvas.append(fragment);
            // Replace estimates with the real heights now that these rows exist.
            let corrected = false;
            for (const i of fresh) {
                const actual = diffNodes.get(i).offsetHeight;
                if (actual && Math.abs(actual - diffHeights[i]) > 0.5) {
                    diffHeights[i] = actual;
                    corrected = true;
                }
            }
            if (corrected) recomputeDiffOffsets();
            for (const [index, node] of diffNodes) {
                node.style.top = diffOffsets[index] + 'px';
            }
        }

        function mountDiffLines(lines) {
            const diff = el('diff');
            diff.innerHTML = '';
            diffNodes.clear();
            diffLines = lines || [];
            diffCanvas = document.createElement('div');
            diffCanvas.className = 'diffCanvas';
            diff.append(diffCanvas);
            diff.scrollTop = 0;
            measureDiffMetrics();
            diffHeights = diffLines.map(estimateDiffHeight);
            recomputeDiffOffsets();
            paintDiffWindow();
        }

        // A narrower diff pane rewraps every line, so estimates for rows that
        // are not currently rendered have to be rebuilt from the new width.
        function remeasureDiffLayout() {
            if (!diffCanvas || !diffLines.length) return;
            measureDiffMetrics();
            for (let i = 0; i < diffLines.length; i++) {
                diffHeights[i] = estimateDiffHeight(diffLines[i]);
            }
            for (const [, node] of diffNodes) node.remove();
            diffNodes.clear();
            recomputeDiffOffsets();
            paintDiffWindow();
        }

        function renderDiff(force) {
            const item = plan.items.find(candidate => candidate.id === activeId) || plan.items[0];
            const key = item ? `${item.id}|${showHeaders ? 1 : 0}` : '';
            // Toggling a checkbox on the row already being shown leaves the diff
            // identical; rebuilding it is the single most expensive thing this
            // view can do, so only rebuild when the shown content changes.
            if (!force && key === diffRenderedKey) return;
            diffRenderedKey = key;
            el('toggleHeaders').textContent = showHeaders ? 'Hide Headers in Diff' : 'Show Headers in Diff';
            el('toggleHeaders').setAttribute('aria-pressed', String(showHeaders));
            if (!item) {
                const diff = el('diff');
                diff.innerHTML = '';
                diffCanvas = null;
                diffLines = [];
                diffNodes.clear();
                el('leftTitle').textContent = '';
                el('rightTitle').textContent = '';
                el('copyLeft').disabled = true;
                el('copyRight').disabled = true;
                el('toggleHeaders').disabled = true;
                const empty = document.createElement('pre');
                empty.textContent = 'No module differences found for the current settings.';
                diff.append(empty);
                return;
            }
            const leftCode = showHeaders ? item.leftRawCode : item.leftCode;
            const rightCode = showHeaders ? item.rightRawCode : item.rightCode;
            el('leftTitle').textContent = item.leftTitle;
            el('rightTitle').textContent = item.rightTitle;
            el('copyLeft').disabled = !leftCode;
            el('copyRight').disabled = !rightCode;
            setTooltip('copyLeft', copyTooltip(item, 'left', Boolean(leftCode)));
            setTooltip('copyRight', copyTooltip(item, 'right', Boolean(rightCode)));
            el('toggleHeaders').disabled = false;
            mountDiffLines(showHeaders ? item.diffWithHeaders : item.diff);
        }

        function renderCounts() {
            const selectedItems = plan.items.filter(item => selected.has(item.id));
            const unsupported = selectedItems.filter(item => item.unsupportedDirectCreation).length;
            const settingsStatus = settingsSaving ? ' | auto-saving settings' : '';
            el('counts').textContent = `${selectedItems.length} selected${unsupported ? ' | ' + unsupported + ' will show skipping import warning' : ''}${settingsStatus}`;
            el('apply').disabled = applying || applied || selectedItems.length === 0;
            el('chooseFolder').disabled = applying;
            el('syncMode').disabled = applying;
        }

        function clearSettingsAutosave() {
            if (settingsSaveTimer) {
                clearTimeout(settingsSaveTimer);
                settingsSaveTimer = undefined;
            }
        }

        function scheduleSettingsAutosave() {
            clearSettingsAutosave();
            el('result').textContent = 'Settings changed. Auto-saving...';
            settingsSaveTimer = setTimeout(() => {
                settingsSaveTimer = undefined;
                settingsSaving = true;
                renderCounts();
                vscode.postMessage({ type: 'save-settings', quiet: true, ...currentSettings() });
            }, 650);
        }

        function refreshSettings() {
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Refreshing...';
            renderCounts();
            vscode.postMessage({ type: 'refresh-settings', autoSaveSettings: true, ...currentSettings() });
        }

        el('selectChanged').addEventListener('click', () => {
            selected.clear();
            for (const item of plan.items) {
                if (isRelevantItem(item)) selected.add(item.id);
            }
            syncListState();
        });
        el('clear').addEventListener('click', () => {
            selected.clear();
            syncListState();
        });
        el('diff').addEventListener('scroll', paintDiffWindow, { passive: true });
        new ResizeObserver(() => {
            if (diffResizeFrame) return;
            diffResizeFrame = requestAnimationFrame(() => {
                diffResizeFrame = 0;
                remeasureDiffLayout();
            });
        }).observe(el('diff'));
        el('copyLeft').addEventListener('click', event => {
            event.stopPropagation();
            vscode.postMessage({ type: 'copy-code', itemId: activeId, side: 'left', showHeaders });
        });
        el('copyRight').addEventListener('click', event => {
            event.stopPropagation();
            vscode.postMessage({ type: 'copy-code', itemId: activeId, side: 'right', showHeaders });
        });
        document.addEventListener('contextmenu', event => {
            event.preventDefault();
        });
        el('toggleHeaders').addEventListener('click', () => {
            showHeaders = !showHeaders;
            renderDiff();
        });
        el('chooseFolder').addEventListener('click', () => {
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Choosing folder...';
            renderCounts();
            vscode.postMessage({ type: 'choose-folder', autoSaveSettings: true, ...currentSettings() });
        });
        el('syncMode').addEventListener('change', () => {
            updateModeTitle();
            refreshSettings();
        });
        el('cancel').addEventListener('click', () => {
            clearSettingsAutosave();
            vscode.postMessage({ type: 'cancel' });
        });
        el('apply').addEventListener('click', () => {
            if (applied) return;
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Applying...';
            renderCounts();
            vscode.postMessage({ type: 'apply', selectedIds: Array.from(selected) });
        });

        window.addEventListener('message', event => {
            const message = event.data;
            if (message.type === 'applying') {
                applying = true;
                el('result').textContent = 'Applying...';
                renderCounts();
            } else if (message.type === 'refreshing') {
                applying = true;
                el('result').textContent = message.message || 'Refreshing...';
                renderCounts();
            } else if (message.type === 'ready') {
                applying = false;
                el('result').textContent = '';
                renderCounts();
            } else if (message.type === 'plan') {
                setPlan(message.plan, message.message || 'Settings updated. Review the refreshed diff before applying.', message.autoSaveSettings === true);
            } else if (message.type === 'saving-settings') {
                settingsSaving = true;
                el('result').textContent = 'Saving settings...';
                renderCounts();
            } else if (message.type === 'settings-saved') {
                settingsSaving = false;
                el('result').textContent = message.quiet ? 'Settings auto-saved.' : message.result.summary;
                renderCounts();
            } else if (message.type === 'copied') {
                el('result').textContent = message.side === 'left' ? 'Left code copied.' : 'Right code copied.';
            } else if (message.type === 'applied') {
                applying = false;
                applied = true;
                el('result').textContent = message.result.summary;
                el('apply').textContent = 'Applied';
                renderCounts();
            } else if (message.type === 'error') {
                applying = false;
                settingsSaving = false;
                el('result').textContent = message.error;
                renderCounts();
            }
        });

        selected = selectedFromPlan();
        installSplitter();
        installFastTooltips();
        renderChrome();
        renderList();
        renderDiff();
