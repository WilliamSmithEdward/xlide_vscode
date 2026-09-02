        const vscode = acquireVsCodeApi();
        {{toastScript}}
        const projectPath = {{projectPathJson}};
        const tagNames = {{tagNamesJson}};
        const testIds = {{testIdsJson}};
        const canRun = {{canRunJson}};
        const hasTags = {{hasTagsJson}};
        const hasTests = {{hasTestsJson}};
        const hasLastFailed = {{hasLastFailedJson}};
        let filterState = initialFilterState();
        let running = false;

        function initialFilterState() {
            const saved = vscode.getState?.();
            if (saved?.projectPath === projectPath) {
                return {
                    projectPath,
                    includeTags: reconcileTags(saved.includeTags),
                    excludeTags: reconcileTags(saved.excludeTags),
                    selectedTestIds: Array.isArray(saved.selectedTestIds)
                        ? reconcileTestIds(saved.selectedTestIds)
                        : [...testIds],
                    failFast: Boolean(saved.failFast),
                };
            }
            return {
                projectPath,
                includeTags: [...tagNames],
                excludeTags: [],
                selectedTestIds: [...testIds],
                failFast: false,
            };
        }

        function reconcileTags(value) {
            if (!Array.isArray(value)) {
                return [];
            }
            const available = new Set(tagNames);
            return [...new Set(value.map((tag) => String(tag).trim()).filter((tag) => available.has(tag)))];
        }

        function reconcileTestIds(value) {
            if (!Array.isArray(value)) {
                return [];
            }
            const available = new Set(testIds);
            return [...new Set(value.map((id) => String(id).trim()).filter((id) => available.has(id)))];
        }

        function saveFilterState() {
            vscode.setState?.(filterState);
        }

        function setRunning(next) {
            running = next;
            syncFilterUi();
        }

        function syncFilterUi() {
            document.querySelectorAll('input[data-filter-kind]').forEach((input) => {
                const list = input.dataset.filterKind === 'exclude' ? filterState.excludeTags : filterState.includeTags;
                input.checked = list.includes(input.dataset.tag);
            });
            document.querySelectorAll('input[data-test-id]').forEach((input) => {
                input.checked = filterState.selectedTestIds.includes(input.dataset.testId);
            });
            const failFast = document.getElementById('failFast');
            if (failFast) {
                failFast.checked = filterState.failFast;
            }
            const summary = document.getElementById('filterSummary');
            if (summary) {
                const includeCount = filterState.includeTags.length;
                const excludeCount = filterState.excludeTags.length;
                summary.textContent = includeCount + ' include, ' + excludeCount + ' exclude';
            }
            const selectedSummary = document.getElementById('selectedTestSummary');
            if (selectedSummary) {
                const count = filterState.selectedTestIds.length;
                selectedSummary.textContent = count + ' selected, ' + testIds.length + ' discovered';
            }
            document.querySelectorAll('[data-filter-action], [data-test-action], input[data-filter-kind], input[data-test-id], #failFast').forEach((control) => {
                control.disabled = running;
            });
            const runAll = document.querySelector('button[data-action="runAll"]');
            if (runAll && canRun) {
                runAll.disabled = running;
            }
            const runSelected = document.querySelector('button[data-action="runSelected"]');
            if (runSelected && canRun && hasTests) {
                runSelected.disabled = running || filterState.selectedTestIds.length === 0;
            }
            document.querySelectorAll('button[data-action="runCurrentModule"], button[data-action="runCurrentTest"]').forEach((button) => {
                if (canRun) {
                    button.disabled = running;
                }
            });
            const runWithFilters = document.querySelector('button[data-action="runWithFilters"]');
            if (runWithFilters && canRun && hasTags) {
                runWithFilters.disabled = running ||
                    (filterState.includeTags.length === 0 && filterState.excludeTags.length === 0);
            }
            const rerunFailed = document.querySelector('button[data-action="rerunFailed"]');
            if (rerunFailed && canRun && hasLastFailed) {
                rerunFailed.disabled = running;
            }
        }

        function setFilterTags(kind, values) {
            const selected = [...new Set(values.filter((tag) => tagNames.includes(tag)))];
            if (kind === 'exclude') {
                filterState.excludeTags = selected;
                filterState.includeTags = filterState.includeTags.filter((tag) => !selected.includes(tag));
            } else {
                filterState.includeTags = selected;
                filterState.excludeTags = filterState.excludeTags.filter((tag) => !selected.includes(tag));
            }
            saveFilterState();
            syncFilterUi();
        }

        function setSelectedTestIds(values) {
            filterState.selectedTestIds = reconcileTestIds(values);
            saveFilterState();
            syncFilterUi();
        }

        function toggleSelectedTest(id, checked) {
            const next = checked
                ? [...new Set([...filterState.selectedTestIds, id])]
                : filterState.selectedTestIds.filter((candidate) => candidate !== id);
            setSelectedTestIds(next);
        }

        function toggleFilterTag(kind, tag, checked) {
            const list = kind === 'exclude' ? filterState.excludeTags : filterState.includeTags;
            const next = checked
                ? [...new Set([...list, tag])]
                : list.filter((candidate) => candidate !== tag);
            setFilterTags(kind, next);
        }

        document.addEventListener('change', (event) => {
            const input = event.target.closest?.('input[data-filter-kind]');
            if (input) {
                toggleFilterTag(input.dataset.filterKind, input.dataset.tag, input.checked);
                return;
            }
            if (event.target?.id === 'failFast') {
                filterState.failFast = Boolean(event.target.checked);
                saveFilterState();
                return;
            }
            const testInput = event.target.closest?.('input[data-test-id]');
            if (testInput) {
                toggleSelectedTest(testInput.dataset.testId, testInput.checked);
            }
        });

        document.addEventListener('click', (event) => {
            const testButton = event.target.closest?.('[data-test-action]');
            if (testButton) {
                if (testButton.dataset.testAction === 'selectAll') {
                    setSelectedTestIds(testIds);
                } else if (testButton.dataset.testAction === 'clear') {
                    setSelectedTestIds([]);
                }
                return;
            }
            const filterButton = event.target.closest?.('[data-filter-action]');
            if (filterButton) {
                const action = filterButton.dataset.filterAction;
                const kind = filterButton.dataset.filterKind;
                if (action === 'selectAll') {
                    setFilterTags(kind, tagNames);
                } else if (action === 'clear') {
                    setFilterTags(kind, []);
                }
                return;
            }
            const button = event.target.closest?.('[data-action]');
            if (!button || button.disabled) {
                return;
            }
            if (button.dataset.action === 'runAll') {
                setRunning(true);
                vscode.postMessage({ type: 'runAll' });
                return;
            }
            if (button.dataset.action === 'runSelected') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runSelected',
                    testIds: filterState.selectedTestIds,
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'runCurrentModule') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runCurrentModule',
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'runCurrentTest') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runCurrentTest',
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'runWithFilters') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runWithFilters',
                    includeTags: filterState.includeTags,
                    excludeTags: filterState.excludeTags,
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'rerunFailed') {
                setRunning(true);
                vscode.postMessage({ type: 'rerunFailed' });
                return;
            }
            vscode.postMessage({ type: button.dataset.action });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                setRunning(false);
                showToast(event.data.error || 'XLIDE test action failed');
            } else if (event.data?.type === 'refreshed') {
                setRunning(false);
                showToast('Test support refreshed');
            }
        });

        syncFilterUi();
