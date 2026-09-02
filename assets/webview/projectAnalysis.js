        const vscode = acquireVsCodeApi();
        {{toastScript}}
        let model = {{modelJson}};
        const severityIds = ['error', 'warning', 'information'];
        let analysisSettingsKey = model.analysisSettingsKey;
        let visibleSeverities = new Set(normalizeSeverityList(model.visibleSeverities ?? severityIds));
        let activeModule = 'all';
        let showHiddenItems = false;
        let sortKey = 'severity';
        let sortDirection = 'asc';
        let settingsOpen = false;
        let rows = Array.from(document.querySelectorAll('.problemRow'));
        const sortHeaders = Array.from(document.querySelectorAll('[data-sort]'));
        const table = document.querySelector('.table');
        const statsSection = document.querySelector('.stats');
        const moduleList = document.querySelector('.moduleList');
        const projectSubtitle = document.querySelector('header .subtle');
        const showHiddenButton = document.querySelector('[data-show-hidden]');
        const visibleCount = document.getElementById('visibleCount');
        const contextMenu = document.getElementById('rowContextMenu');
        const quickFixMenuItem = document.getElementById('quickFixMenuItem');
        const quickFixSubmenu = document.getElementById('quickFixSubmenu');
        const suppressionDivider = document.getElementById('suppressionDivider');
        const trackingDivider = document.getElementById('trackingDivider');
        const trackingProjectAction = document.getElementById('trackingProjectAction');
        const trackingGlobalAction = document.getElementById('trackingGlobalAction');
        const settingsDialog = document.getElementById('analysisSettingsDialog');
        const settingsSource = document.querySelector('.settingsSource');
        const settingsResetButton = document.querySelector('[data-reset-analysis]');
        const projectUntrackedRulesContainer = document.getElementById('projectUntrackedRules');
        let contextRow = null;

        function applyModel(next) {
            model = next;
            if (next.analysisSettingsKey !== analysisSettingsKey) {
                analysisSettingsKey = next.analysisSettingsKey;
                visibleSeverities = new Set(normalizeSeverityList(next.visibleSeverities ?? severityIds));
            }
            if (projectSubtitle) {
                projectSubtitle.textContent = next.projectName;
            }
            renderStats();
            renderModuleFilters();
            renderRows();
            renderProjectUntrackedRules();
            reattachContextRow();
            sortRows();
            syncSortHeaders();
            syncModuleFilterButtons();
            syncSeverityFilterButtons();
            syncHiddenToggleButton();
            updateRows();
        }

        function element(tag, className, text) {
            const node = document.createElement(tag);
            if (className) {
                node.className = className;
            }
            if (text !== undefined) {
                node.textContent = text;
            }
            return node;
        }

        function renderStats() {
            statsSection.textContent = '';
            for (const [value, label] of [
                [model.errorCount, 'Errors'],
                [model.warningCount, 'Warnings'],
                [model.informationCount, 'Information'],
                [model.suppressedCount, 'Suppressed'],
                [model.untrackedCount, 'Untracked'],
                [model.moduleCount, 'Modules'],
            ]) {
                const stat = element('div', 'stat');
                stat.appendChild(element('strong', undefined, String(value)));
                stat.appendChild(element('span', undefined, label));
                statsSection.appendChild(stat);
            }
        }

        function buildModuleFilterButton(filter, labelNode, count) {
            const button = element('button', 'moduleFilter');
            button.type = 'button';
            button.dataset.moduleFilter = filter;
            button.appendChild(labelNode);
            button.appendChild(element('span', undefined, String(count)));
            return button;
        }

        function renderModuleFilters() {
            moduleList.textContent = '';
            moduleList.appendChild(buildModuleFilterButton(
                'all',
                document.createTextNode('All modules '),
                model.totalProblems,
            ));
            for (const group of model.groups) {
                const identity = element('span', 'moduleIdentity');
                const icon = element('span', 'moduleIcon', group.moduleIcon);
                icon.title = group.moduleTypeLabel;
                identity.appendChild(icon);
                identity.appendChild(element('span', 'moduleName', group.moduleName));
                moduleList.appendChild(buildModuleFilterButton(group.moduleName, identity, group.total));
            }
        }

        function renderRows() {
            for (const row of rows) {
                row.remove();
            }
            const empty = table.querySelector('.empty');
            if (empty) {
                empty.remove();
            }
            rows = model.rows.map(buildProblemRow);
            if (rows.length === 0) {
                table.appendChild(element('div', 'empty', 'No analysis findings.'));
                return;
            }
            for (const row of rows) {
                table.appendChild(row);
            }
        }

        function buildProblemRow(row) {
            const button = element('button', 'problemRow severity-' + row.severity);
            button.type = 'button';
            button.dataset.openIndex = String(row.index);
            button.dataset.suppressed = row.suppressed ? 'yes' : 'no';
            button.dataset.status = row.statusKey;
            button.dataset.module = row.moduleName;
            button.dataset.moduleType = row.moduleType;
            button.dataset.moduleOrder = String(row.moduleOrder);
            button.dataset.severity = row.severity;
            button.dataset.compile = row.vbeCompileEquivalent ? 'yes' : 'no';
            button.dataset.line = String(row.line);
            button.dataset.column = String(row.column);
            button.dataset.endColumn = String(row.endColumn);
            button.dataset.rule = row.rule;
            button.dataset.ruleCode = row.ruleCode;
            button.dataset.message = row.message;
            button.dataset.evidence = row.evidence;
            button.dataset.quickFixes = JSON.stringify(row.quickFixTitles ?? []);
            button.dataset.suppressionScopes = JSON.stringify(row.suppressionScopes ?? []);
            button.dataset.tracked = row.tracked ? 'yes' : 'no';
            button.dataset.trackingSource = row.trackingSource;
            button.appendChild(element('span', 'cell severity', row.severity));
            button.appendChild(element('span', 'cell status', row.statusLabel));
            button.appendChild(element('span', 'cell location', row.location));
            button.appendChild(element('span', 'cell code', row.rule));
            button.appendChild(element('span', 'cell kind', row.evidence));
            button.appendChild(element('span', 'cell message', row.message));
            return button;
        }

        function renderProjectUntrackedRules() {
            settingsSource.textContent = 'Source: ' + (model.rulesSourceIsProject ? 'File settings' : 'No file override');
            settingsResetButton.disabled = !model.rulesSourceIsProject;
            projectUntrackedRulesContainer.textContent = '';
            const rules = model.projectUntrackedRules ?? [];
            if (rules.length === 0) {
                projectUntrackedRulesContainer.appendChild(
                    element('div', 'settingsEmpty', 'No project rules are manually untracked.'),
                );
                return;
            }
            const rulesTable = element('table', 'settingsTable');
            const head = element('thead');
            const headRow = element('tr');
            for (const [className, label] of [
                ['settingsTableCode', 'Rule'],
                [undefined, 'Title'],
                ['settingsTableAction', 'Action'],
            ]) {
                const cell = element('th', className, label);
                cell.scope = 'col';
                headRow.appendChild(cell);
            }
            head.appendChild(headRow);
            rulesTable.appendChild(head);
            const body = element('tbody');
            for (const rule of rules) {
                const ruleRow = element('tr');
                ruleRow.appendChild(element('td', 'settingsTableCode', rule.code));
                ruleRow.appendChild(element('td', undefined, rule.title));
                const action = element('td', 'settingsTableAction');
                const track = element('button', 'secondaryButton', 'Track');
                track.type = 'button';
                track.dataset.settingsTrackRuleCode = rule.code;
                action.appendChild(track);
                ruleRow.appendChild(action);
                body.appendChild(ruleRow);
            }
            rulesTable.appendChild(body);
            projectUntrackedRulesContainer.appendChild(rulesTable);
        }

        function reattachContextRow() {
            if (!contextRow) {
                return;
            }
            const previous = contextRow.dataset;
            const match = rows.find((row) =>
                row.dataset.module === previous.module &&
                row.dataset.line === previous.line &&
                row.dataset.column === previous.column &&
                row.dataset.ruleCode === previous.ruleCode);
            if (match) {
                contextRow = match;
            } else {
                hideContextMenu();
            }
        }

        function rowMatchesModule(row) {
            return activeModule === 'all' || row.dataset.module === activeModule;
        }

        function updateRows() {
            let count = 0;
            let shownHidden = 0;
            for (const row of rows) {
                const moduleVisible = rowMatchesModule(row);
                const severityVisible = visibleSeverities.has(row.dataset.severity);
                const untracked = row.dataset.tracked === 'no';
                const suppressed = row.dataset.suppressed === 'yes';
                const hiddenByStatus = untracked || suppressed;
                const visible = moduleVisible &&
                    severityVisible &&
                    (!hiddenByStatus || showHiddenItems);
                const untrackedVisible = visible && untracked;
                const suppressedVisible = visible && suppressed && !untracked;
                row.hidden = !visible;
                row.classList.toggle('hiddenVisible', untrackedVisible);
                row.classList.toggle('suppressedVisible', suppressedVisible);
                if (visible && !hiddenByStatus) {
                    count += 1;
                }
                if (visible && hiddenByStatus) {
                    shownHidden += 1;
                }
            }
            const details = [];
            if (showHiddenItems) {
                details.push(`${shownHidden} non-tracked/suppressed visible`);
            }
            visibleCount.textContent = details.length > 0
                ? `${count} shown, ${details.join(', ')}`
                : `${count} shown`;
        }

        function sortRows() {
            const direction = sortDirection === 'asc' ? 1 : -1;
            rows.sort((left, right) => direction * compareRows(left, right, sortKey));
            for (const row of rows) {
                table.appendChild(row);
            }
        }

        function compareRows(left, right, key) {
            if (key === 'severity') {
                const severityOrder = { error: 0, warning: 1, information: 2 };
                return compareNumber(severityOrder[left.dataset.severity] ?? 4, severityOrder[right.dataset.severity] ?? 4)
                    || compareLocation(left, right);
            }
            if (key === 'status') {
                const statusOrder = { tracked: 0, untracked: 1, suppressed: 2 };
                return compareNumber(statusOrder[left.dataset.status] ?? 9, statusOrder[right.dataset.status] ?? 9)
                    || compareLocation(left, right);
            }
            if (key === 'rule') {
                return compareText(left.dataset.rule, right.dataset.rule) || compareLocation(left, right);
            }
            if (key === 'message') {
                return compareText(left.dataset.message, right.dataset.message) || compareLocation(left, right);
            }
            if (key === 'evidence') {
                return compareText(left.dataset.evidence, right.dataset.evidence) || compareLocation(left, right);
            }
            return compareLocation(left, right);
        }

        function compareLocation(left, right) {
            return compareNumber(Number(left.dataset.moduleOrder), Number(right.dataset.moduleOrder))
                || compareText(left.dataset.module, right.dataset.module)
                || compareNumber(Number(left.dataset.line), Number(right.dataset.line))
                || compareNumber(Number(left.dataset.column), Number(right.dataset.column));
        }

        function compareNumber(left, right) {
            return left === right ? 0 : left < right ? -1 : 1;
        }

        function compareText(left, right) {
            return String(left ?? '').localeCompare(String(right ?? ''));
        }

        function setActive(buttons, activeButton) {
            for (const button of buttons) {
                button.classList.toggle('active', button === activeButton);
            }
        }

        function syncHiddenToggleButton() {
            showHiddenButton?.classList.toggle('active', showHiddenItems);
            showHiddenButton?.setAttribute('aria-pressed', showHiddenItems ? 'true' : 'false');
        }

        function syncSeverityFilterButtons() {
            for (const button of document.querySelectorAll('[data-severity-toggle]')) {
                const active = visibleSeverities.has(button.dataset.severityToggle);
                button.classList.toggle('active', active);
                button.setAttribute('aria-pressed', active ? 'true' : 'false');
            }
        }

        function syncSortHeaders() {
            for (const header of sortHeaders) {
                const active = header.dataset.sort === sortKey;
                const sortValue = active
                    ? (sortDirection === 'asc' ? 'ascending' : 'descending')
                    : 'none';
                const label = header.dataset.sortLabel ?? header.textContent?.trim() ?? 'column';
                const nextDirection = active && sortDirection === 'asc' ? 'descending' : 'ascending';
                header.setAttribute('aria-sort', sortValue);
                header.title = active
                    ? `Sorted ${sortValue}. Click to sort ${nextDirection}.`
                    : `Sort by ${label}`;
            }
        }

        function syncModuleFilterButtons() {
            const buttons = Array.from(document.querySelectorAll('[data-module-filter]'));
            let activeButton = buttons.find(button => button.dataset.moduleFilter === activeModule);
            if (!activeButton) {
                activeModule = 'all';
                activeButton = buttons.find(button => button.dataset.moduleFilter === 'all');
            }
            if (activeButton) {
                setActive(buttons, activeButton);
            }
        }

        function syncSettingsDialog() {
            if (settingsDialog) {
                settingsDialog.hidden = !settingsOpen;
            }
        }

        function normalizeSeverityList(value) {
            return Array.isArray(value)
                ? value.filter(item => severityIds.includes(item))
                : severityIds;
        }

        function hideContextMenu() {
            contextMenu.hidden = true;
            quickFixSubmenu.hidden = true;
            contextRow = null;
        }

        function setSettingsOpen(open) {
            settingsOpen = open;
            syncSettingsDialog();
        }

        function showContextMenu(row, x, y) {
            contextRow = row;
            const quickFixButton = contextMenu.querySelector('[data-context-action="quickFix"]');
            const fixes = quickFixTitlesForRow(row);
            if (quickFixButton) {
                const enabled = fixes.length > 0;
                quickFixButton.disabled = !enabled;
                quickFixButton.setAttribute('aria-disabled', enabled ? 'false' : 'true');
                quickFixMenuItem?.classList.toggle('hasQuickFixes', enabled);
            }
            renderQuickFixSubmenu(fixes);
            const suppressionScopes = suppressionScopesForRow(row);
            const rowSuppressed = row.dataset.suppressed === 'yes';
            let visibleSuppressActions = 0;
            for (const button of contextMenu.querySelectorAll('[data-suppress-scope]')) {
                const visible = suppressionScopes.has(button.dataset.suppressScope);
                button.hidden = !visible;
                button.disabled = rowSuppressed;
                button.setAttribute('aria-disabled', rowSuppressed ? 'true' : 'false');
                if (visible) {
                    visibleSuppressActions += 1;
                }
            }
            if (suppressionDivider) {
                suppressionDivider.hidden = visibleSuppressActions === 0;
            }
            const visibleTrackingActions = syncTrackingActions(row);
            if (trackingDivider) {
                trackingDivider.hidden = visibleTrackingActions === 0;
            }
            contextMenu.hidden = false;
            const rect = contextMenu.getBoundingClientRect();
            contextMenu.style.left = Math.min(x, window.innerWidth - rect.width - 8) + 'px';
            contextMenu.style.top = Math.min(y, window.innerHeight - rect.height - 8) + 'px';
        }

        function contextProblemIndex() {
            return contextRow ? Number(contextRow.dataset.openIndex) : undefined;
        }

        function contextProblemSuppressed() {
            return contextRow?.dataset.suppressed === 'yes';
        }

        function contextProblemIdentity() {
            if (!contextRow) { return {}; }
            const d = contextRow.dataset;
            return {
                moduleName: d.module,
                line: Number(d.line),
                column: Number(d.column),
                endColumn: Number(d.endColumn),
                code: d.ruleCode,
            };
        }

        function contextProblemTracked() {
            return contextRow?.dataset.tracked !== 'no';
        }

        function syncTrackingActions(row) {
            const hasRuleCode = String(row.dataset.ruleCode ?? '').trim().length > 0;
            const tracked = row.dataset.tracked !== 'no';
            const source = row.dataset.trackingSource === 'project' ? 'project' : 'global';
            return configureTrackingAction(trackingProjectAction, {
                hidden: !hasRuleCode || (!tracked && source !== 'project'),
                label: tracked ? 'Untrack In Project' : 'Track In Project',
            }) + configureTrackingAction(trackingGlobalAction, {
                hidden: !hasRuleCode || (!tracked && source !== 'global'),
                label: tracked ? 'Untrack Globally' : 'Track Globally',
            });
        }

        function configureTrackingAction(button, options) {
            if (!button) {
                return 0;
            }
            button.hidden = options.hidden;
            button.textContent = options.label;
            return options.hidden ? 0 : 1;
        }

        function quickFixTitlesForRow(row) {
            try {
                const parsed = JSON.parse(row.dataset.quickFixes ?? '[]');
                return Array.isArray(parsed) ? parsed.filter(item => typeof item === 'string') : [];
            } catch {
                return [];
            }
        }

        function suppressionScopesForRow(row) {
            try {
                const parsed = JSON.parse(row.dataset.suppressionScopes ?? '[]');
                return new Set(
                    Array.isArray(parsed)
                        ? parsed.filter(item => item === 'block' || item === 'member' || item === 'module')
                        : []
                );
            } catch {
                return new Set();
            }
        }

        function renderQuickFixSubmenu(fixes) {
            quickFixSubmenu.innerHTML = '';
            quickFixSubmenu.hidden = fixes.length === 0;
            fixes.forEach((title, index) => {
                const button = document.createElement('button');
                button.type = 'button';
                button.dataset.contextAction = 'applyQuickFix';
                button.dataset.fixIndex = String(index);
                button.textContent = title;
                quickFixSubmenu.appendChild(button);
            });
        }

        function postOpenProblem(row) {
            vscode.postMessage({
                type: 'openProblem',
                index: Number(row.dataset.openIndex),
                suppressed: row.dataset.suppressed === 'yes',
                moduleName: row.dataset.module,
                moduleType: row.dataset.moduleType,
                line: Number(row.dataset.line),
                column: Number(row.dataset.column),
                endColumn: Number(row.dataset.endColumn),
                severity: row.dataset.severity,
                code: row.dataset.ruleCode,
                message: row.dataset.message,
            });
        }

        document.addEventListener('click', (event) => {
            if (event.button === 2) {
                return;
            }
            const contextButton = event.target.closest?.('[data-context-action]');
            if (contextButton) {
                const action = contextButton.dataset.contextAction;
                if (action === 'quickFix') {
                    return;
                }
                const index = contextProblemIndex();
                const suppressed = contextProblemSuppressed();
                const currentlyTracked = contextProblemTracked();
                // Stable identity of the right-clicked finding, so the host can
                // reject the action if a background refresh shifted the indices.
                const identity = contextProblemIdentity();
                hideContextMenu();
                if (typeof index !== 'number' || Number.isNaN(index)) {
                    return;
                }
                if (action === 'applyQuickFix') {
                    vscode.postMessage({
                        type: 'quickFixProblem',
                        index,
                        suppressed,
                        ...identity,
                        fixIndex: Number(contextButton.dataset.fixIndex ?? 0),
                    });
                } else if (action === 'askCopilot') {
                    vscode.postMessage({ type: 'askCopilot', index, suppressed, ...identity });
                } else if (action === 'setRuleTrackingProject' || action === 'setRuleTrackingGlobal') {
                    vscode.postMessage({
                        type: 'setRuleTracking',
                        index,
                        suppressed,
                        ...identity,
                        tracked: !currentlyTracked,
                        trackingScope: action === 'setRuleTrackingGlobal' ? 'global' : 'project',
                    });
                } else if (action === 'suppressBlock') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, ...identity, scope: 'block' });
                } else if (action === 'suppressMember') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, ...identity, scope: 'member' });
                } else if (action === 'suppressModule') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, ...identity, scope: 'module' });
                }
                return;
            }
            hideContextMenu();
            const settingsButton = event.target.closest?.('#analysisSettings');
            if (settingsButton) {
                setSettingsOpen(true);
                return;
            }
            const closeSettingsButton = event.target.closest?.('#closeAnalysisSettings');
            if (closeSettingsButton) {
                setSettingsOpen(false);
                return;
            }
            if (event.target === settingsDialog) {
                setSettingsOpen(false);
                return;
            }
            const resetAnalysisButton = event.target.closest?.('[data-reset-analysis]');
            if (resetAnalysisButton) {
                if (resetAnalysisButton.disabled) {
                    return;
                }
                const scope = resetAnalysisButton.dataset.resetAnalysis;
                if (scope === 'rules') {
                    vscode.postMessage({ type: 'resetAnalysisRuleTracking' });
                }
                return;
            }
            const settingsTrackRule = event.target.closest?.('[data-settings-track-rule-code]');
            if (settingsTrackRule) {
                const code = settingsTrackRule.dataset.settingsTrackRuleCode;
                vscode.postMessage({
                    type: 'setRuleTracking',
                    code,
                    tracked: true,
                    trackingScope: 'project',
                });
                return;
            }
            const filterButton = event.target.closest?.('[data-severity-toggle]');
            if (filterButton) {
                const id = filterButton.dataset.severityToggle;
                if (visibleSeverities.has(id)) {
                    visibleSeverities.delete(id);
                } else {
                    visibleSeverities.add(id);
                }
                syncSeverityFilterButtons();
                updateRows();
                return;
            }
            const sortButton = event.target.closest?.('[data-sort]');
            if (sortButton) {
                const nextKey = sortButton.dataset.sort;
                if (sortKey === nextKey) {
                    sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    sortKey = nextKey;
                    sortDirection = 'asc';
                }
                sortRows();
                syncSortHeaders();
                updateRows();
                return;
            }
            const showHiddenButton = event.target.closest?.('[data-show-hidden]');
            if (showHiddenButton) {
                showHiddenItems = !showHiddenItems;
                syncHiddenToggleButton();
                updateRows();
                return;
            }
            const moduleButton = event.target.closest?.('[data-module-filter]');
            if (moduleButton) {
                activeModule = moduleButton.dataset.moduleFilter;
                setActive(document.querySelectorAll('[data-module-filter]'), moduleButton);
                updateRows();
                return;
            }
            const problemRow = event.target.closest?.('[data-open-index]');
            if (problemRow) {
                postOpenProblem(problemRow);
            }
        });

        document.addEventListener('contextmenu', (event) => {
            const problemRow = event.target.closest?.('[data-open-index]');
            if (!problemRow) {
                event.preventDefault();
                hideContextMenu();
                return;
            }
            event.preventDefault();
            showContextMenu(problemRow, event.clientX, event.clientY);
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                if (settingsOpen) {
                    setSettingsOpen(false);
                    return;
                }
                hideContextMenu();
            }
        });

        window.addEventListener('scroll', hideContextMenu, true);

        document.getElementById('copyReport').addEventListener('click', () => {
            vscode.postMessage({ type: 'copyReport' });
        });
        document.getElementById('copyJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'copyJson' });
        });
        document.getElementById('exportReport').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportReport' });
        });
        document.getElementById('exportJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportJson' });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'model') {
                applyModel(event.data.model);
            } else if (event.data?.type === 'copied') {
                showToast('Copied');
            } else if (event.data?.type === 'exported') {
                showToast('Exported');
            } else if (event.data?.type === 'suppressed') {
                showToast('Analysis ignore directive inserted');
            } else if (event.data?.type === 'quickFixed') {
                showToast('Quick fix applied');
            } else if (event.data?.type === 'quickFixUnavailable') {
                showToast('No quick fix available');
            } else if (event.data?.type === 'ruleTrackingChanged') {
                const code = String(event.data.code ?? '').toLowerCase();
                const tracked = event.data.tracked === true;
                const scope = event.data.scope === 'global' ? 'globally' : 'in this file';
                showToast(code
                    ? (tracked ? 'Tracked ' : 'Untracked ') + code + ' ' + scope
                    : 'Rule tracking updated');
            } else if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE action failed');
            }
        });

        sortRows();
        syncSortHeaders();
        syncModuleFilterButtons();
        syncSeverityFilterButtons();
        syncHiddenToggleButton();
        syncSettingsDialog();
        updateRows();
