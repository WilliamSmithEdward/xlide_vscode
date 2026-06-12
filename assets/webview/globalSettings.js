        const vscode = acquireVsCodeApi();
        const toast = document.getElementById('toast');
        let toastTimer;

        restoreGlobalSettingsState();

        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE settings update failed');
            }
        });

        document.addEventListener('change', (event) => {
            const target = event.target;
            if (!(target instanceof HTMLElement)) {
                return;
            }
            if (target.matches('input[data-setting-kind="boolean"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.checked === true
                });
                return;
            }
            if (target.matches('input[data-setting-kind="text"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.value
                });
                return;
            }
            if (target.matches('select[data-setting-kind="enum"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.value
                });
                return;
            }
            if (target.matches('input[data-severity-filter]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: 'analysis.visibleSeverities',
                    value: checkedValues('input[data-severity-filter]:checked')
                });
                return;
            }
            if (target.matches('input[data-rule-untracked]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: 'analysis.untrackedRules',
                    value: checkedValues('input[data-rule-untracked]:checked')
                });
                return;
            }
            if (target.matches('select[data-rule-severity]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'setRuleSeverityOverride',
                    code: target.dataset.ruleCode,
                    severity: target.value
                });
            }
        });

        document.addEventListener('input', (event) => {
            const search = event.target.closest?.('#ruleSearch');
            if (!search) {
                return;
            }
            applyRuleSearch(search.value);
            persistGlobalSettingsState();
        });

        document.addEventListener('click', (event) => {
            const button = event.target.closest?.('button[data-reset-setting]');
            if (!button) {
                return;
            }
            persistGlobalSettingsState();
            vscode.postMessage({
                type: 'resetSetting',
                key: button.dataset.resetSetting
            });
        });

        function checkedValues(selector) {
            return Array.from(document.querySelectorAll(selector)).map((input) => input.value);
        }

        function applyRuleSearch(value) {
            const query = value.trim().toLowerCase();
            for (const row of document.querySelectorAll('[data-rule-row]')) {
                row.hidden = query.length > 0 && !row.dataset.search.includes(query);
            }
        }

        function persistGlobalSettingsState() {
            vscode.setState({
                pageScrollTop: window.scrollY,
                ruleSearch: document.getElementById('ruleSearch')?.value ?? '',
                ruleListScrollTop: document.getElementById('ruleList')?.scrollTop ?? 0,
                overrideListScrollTop: document.getElementById('overrideList')?.scrollTop ?? 0
            });
        }

        function restoreGlobalSettingsState() {
            const state = typeof vscode.getState === 'function' ? vscode.getState() || {} : {};
            const ruleSearch = document.getElementById('ruleSearch');
            if (ruleSearch && typeof state.ruleSearch === 'string') {
                ruleSearch.value = state.ruleSearch;
                applyRuleSearch(state.ruleSearch);
            }
            requestAnimationFrame(() => {
                if (typeof state.pageScrollTop === 'number') {
                    window.scrollTo(0, state.pageScrollTop);
                }
                const ruleList = document.getElementById('ruleList');
                if (ruleList && typeof state.ruleListScrollTop === 'number') {
                    ruleList.scrollTop = state.ruleListScrollTop;
                }
                const overrideList = document.getElementById('overrideList');
                if (overrideList && typeof state.overrideListScrollTop === 'number') {
                    overrideList.scrollTop = state.overrideListScrollTop;
                }
            });
        }
