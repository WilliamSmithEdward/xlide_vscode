        const vscode = acquireVsCodeApi();
        {{toastScript}}
        let running = false;

        function setRunning(next) {
            running = next;
            document.querySelectorAll('button[data-action="rerunFailed"]').forEach((button) => {
                button.disabled = running;
            });
        }

        document.addEventListener('click', (event) => {
            const testLink = event.target.closest?.('[data-open-test-index]');
            if (testLink) {
                vscode.postMessage({
                    type: 'openTest',
                    index: Number(testLink.dataset.openTestIndex),
                });
                return;
            }
            const button = event.target.closest?.('button[data-action="rerunFailed"]');
            if (!button || button.disabled) {
                return;
            }
            setRunning(true);
            vscode.postMessage({ type: 'rerunFailed' });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                setRunning(false);
                showToast(event.data.error || 'XLIDE test action failed');
            } else if (event.data?.type === 'rerunComplete') {
                setRunning(false);
            } else if (event.data?.type === 'setRunning') {
                setRunning(Boolean(event.data.running));
            }
        });
