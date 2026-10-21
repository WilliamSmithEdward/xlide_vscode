(() => {
    const vscode = acquireVsCodeApi();
    const el = id => document.getElementById(id);
    const projects = new Map(), groups = new Map(), rows = new Map();
    let worker, url, count = 0, complete = false, busy = false, incomplete = false;
    const send = data => vscode.postMessage(data);
    function resetWorker() {
        if (worker) { worker.terminate(); worker = undefined; }
        if (url) { URL.revokeObjectURL(url); url = undefined; }
    }
    function clearPreview() {
        for (const row of rows.values()) { row.after.remove(); }
        el('apply').hidden = true;
    }
    function controls() {
        el('stop').disabled = !busy;
        el('selected').disabled = busy || !complete || ![...rows.values()].some(row => row.check.checked);
        el('all').disabled = busy || !complete || !count || incomplete;
    }
    function clearResults() {
        el('results').replaceChildren(); projects.clear(); groups.clear(); rows.clear();
        count = 0; complete = false; clearPreview(); controls();
    }
    function invalidate() {
        send({ type: 'invalidate' }); busy = false; clearResults();
        el('status').textContent = 'Search again to update results.';
    }
    function preview(ids, all = false) {
        clearPreview(); busy = true; controls();
        send({ type: 'preview', ids, all, replacement: el('replacement').value });
    }
    el('search').onclick = () => {
        clearResults(); el('errors').textContent = ''; busy = true; controls();
        send({ type: 'search', pattern: el('pattern').value, scope: el('scope').value,
            matchCase: el('matchCase').checked, wholeWord: el('wholeWord').checked, regex: el('regex').checked });
    };
    el('pattern').onkeydown = event => { if (event.key === 'Enter') { el('search').click(); } };
    for (const id of ['pattern', 'scope', 'matchCase', 'wholeWord', 'regex']) { el(id).addEventListener('input', invalidate); }
    el('replacement').oninput = () => { clearPreview(); send({ type: 'replacementChanged' }); };
    el('stop').onclick = () => { send({ type: 'stop' }); };
    el('selected').onclick = () => preview([...rows].filter(([, row]) => row.check.checked).map(([id]) => id));
    el('all').onclick = () => preview([], true);
    el('apply').onclick = () => { busy = true; controls(); el('apply').hidden = true; send({ type: 'apply' }); };
    window.addEventListener('message', event => {
        if (event.origin !== window.origin) { return; }
        const { data } = event;
        if (!data || typeof data.type !== 'string') { return; }
        if (data.type === 'resetWorker') { resetWorker(); }
        if (data.type === 'match') {
            if (!worker) {
                url = URL.createObjectURL(new Blob([searchWorkerSource], { type: 'application/javascript' }));
                worker = new Worker(url);
                worker.onmessage = ({ data }) => send({ type: 'matched', ...data });
                worker.onerror = () => { send({ type: 'matched', id: data.id, error: 'Search worker failed.' }); resetWorker(); };
            }
            worker.postMessage(data);
        }
        if (data.type === 'scope') {
            const project = el('scope').querySelector('[value=project]');
            const module = el('scope').querySelector('[value=module]');
            project.disabled = !data.projectAvailable;
            module.disabled = !data.moduleAvailable;
            project.textContent = data.projectAvailable ? `Project: ${data.project.split(/[\\/]/).pop()}` : 'Selected project';
            project.title = data.project;
            module.textContent = data.moduleAvailable ? `Module: ${data.module}` : 'Selected module';
            el('scope').value = data.selected;
            el('pattern').focus();
        }
        if (data.type === 'busy') { busy = true; clearResults(); el('status').textContent = 'Searching…'; }
        if (data.type === 'results') {
            for (const match of data.rows) {
                let project = projects.get(match.projectPath);
                if (!project) {
                    project = document.createElement('details'); project.open = true;
                    const heading = document.createElement('summary'); heading.textContent = match.project; heading.title = match.projectPath;
                    project.append(heading); el('results').append(project); projects.set(match.projectPath, project);
                }
                const key = JSON.stringify([match.projectPath, match.module]);
                let group = groups.get(key);
                if (!group) {
                    const details = document.createElement('details'); details.open = true;
                    const summary = document.createElement('summary');
                    summary.textContent = match.module;
                    details.append(summary); project.append(details); groups.set(key, details); group = details;
                }
                const row = document.createElement('div'); row.className = 'match';
                const check = document.createElement('input'); check.type = 'checkbox'; check.setAttribute('aria-label', `Select ${match.module} line ${match.line}`);
                check.onchange = () => { clearPreview(); send({ type: 'replacementChanged' }); controls(); };
                const open = document.createElement('button'); open.className = 'link'; open.textContent = `Line ${match.line}`;
                open.onclick = () => send({ type: 'open', id: match.id });
                const one = document.createElement('button'); one.textContent = 'Preview replacement';
                one.onclick = () => { if (!busy && complete) { preview([match.id]); } };
                const snippet = document.createElement('div'); snippet.className = 'snippet';
                const highlight = document.createElement('mark'); highlight.textContent = match.text.slice(match.highlightStart, match.highlightEnd);
                snippet.append(document.createTextNode(match.text.slice(0, match.highlightStart)), highlight,
                    document.createTextNode(match.text.slice(match.highlightEnd)));
                const after = document.createElement('div'); after.className = 'replacement';
                row.append(check, open, document.createTextNode(' '), one, snippet); group.append(row);
                rows.set(match.id, { check, after, row }); count++;
            }
            el('status').textContent = `${count} match(es) found; searching…`;
        }
        if (data.type === 'done') {
            busy = false; complete = true; incomplete = data.incomplete;
            el('status').textContent = `${data.count} match(es).${incomplete ? ' Search incomplete; Replace All is disabled. Narrow the scope or select individual results.' : ''}`;
            el('errors').textContent = data.errors.join('\n');
        }
        if (data.type === 'preview') {
            busy = false;
            for (const change of data.rows) {
                const row = rows.get(change.id);
                if (row) { row.after.textContent = `Replace match with: ${change.text || '(empty text)'}`; row.row.append(row.after); }
            }
            el('status').textContent = `Preview: ${data.count} replacement(s) in ${data.modules} module(s). Review before applying.`;
            el('apply').hidden = false;
        }
        if (data.type === 'error') { busy = false; el('errors').textContent = data.text; el('apply').hidden = true; }
        if (data.type === 'idle') { busy = false; el('apply').hidden = !data.hasPreview; }
        if (data.type === 'stopped') { busy = false; clearResults(); el('status').textContent = 'Search stopped. Run Search to refresh.'; }
        if (data.type === 'applied') { busy = false; clearResults(); el('status').textContent = `${data.count} replacement(s) applied to documents. Normal Save / Auto Save settings apply; use Undo to revert.`; }
        controls();
    });
    window.addEventListener('unload', resetWorker);
    send({ type: 'ready' });
})();
