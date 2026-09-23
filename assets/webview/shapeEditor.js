        const vscode = acquireVsCodeApi();
        const model = {{modelJson}};
        const $ = (id) => document.getElementById(id);
        const form = $('form');
        const buttons = ['save', 'cancel', 'delete'].map($);

        const HOSTS = { excel: 'Excel', word: 'Word', powerpoint: 'PowerPoint' };
        const FONTS = ['Aptos', 'Aptos Display', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New',
            'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
        const DASH_LABELS = {
            solid: 'Solid', dot: 'Round dot', dash: 'Dash', lgDash: 'Long dash', dashDot: 'Dash dot',
            lgDashDot: 'Long dash dot', lgDashDotDot: 'Long dash dot dot', sysDash: 'Square dot dash',
            sysDot: 'Square dot', sysDashDot: 'Square dash dot', sysDashDotDot: 'Square dash dot dot',
        };
        const LOOK_TYPES = {
            fill: [['auto', "Automatic, from the shape's style"], ['none', 'No fill'], ['solid', 'Solid color'], ['other', 'Gradient, picture or pattern (kept as it is)']],
            line: [['auto', "Automatic, from the shape's style"], ['none', 'No outline'], ['solid', 'Solid line'], ['other', 'Gradient or pattern (kept as it is)']],
        };

        function option(select, value, label) {
            const item = document.createElement('option');
            item.value = value;
            item.textContent = label;
            select.appendChild(item);
            return item;
        }

        function show(id, visible) {
            $(id).hidden = !visible;
        }

        /** The radio choices for a fill or outline: automatic and other only when the shape has them now. */
        function lookRadios(kind, current) {
            const holder = $(kind + 'Types');
            holder.textContent = '';
            for (const [value, label] of LOOK_TYPES[kind]) {
                if ((value === 'auto' || value === 'other') && current !== value) { continue; }
                const row = document.createElement('div');
                const input = document.createElement('input');
                input.type = 'radio';
                input.name = kind + 'Type';
                input.id = kind + 'Type-' + value;
                input.value = value;
                input.checked = value === current;
                const text = document.createElement('label');
                text.htmlFor = input.id;
                text.textContent = label;
                row.append(input, text);
                holder.appendChild(row);
            }
        }

        function checkedRadio(name) {
            const checked = form.querySelector('input[name="' + name + '"]:checked');
            return checked ? checked.value : '';
        }

        function setRadio(name, value) {
            const input = form.querySelector('input[name="' + name + '"][value="' + value + '"]');
            if (input) { input.checked = true; }
        }

        /** The fields the shape has: its own, or, while adding, those of the type chosen. */
        function currentFields() {
            if (model.mode === 'add' && model.fieldsByType) {
                return model.fieldsByType[$('type').value] || model.fields;
            }
            return model.fields;
        }

        function applyFields() {
            const fields = currentFields();
            const box = fields.position === 'box' || fields.position === 'size';
            show('rangeField', fields.position === 'range');
            show('boxFields', box);
            show('leftField', fields.position === 'box');
            show('topField', fields.position === 'box');
            show('rotationField', fields.rotation);
            show('positionSection', fields.position !== 'none' || fields.rotation || Boolean(fields.positionNote));
            $('positionNote').textContent = fields.positionNote || '';
            show('textSection', fields.text);
            show('fontSection', fields.font);
            show('fillSection', fields.fill);
            show('lineSection', fields.line);
            show('arrangeSection', model.mode === 'edit' && (fields.zOrder || Boolean(fields.zOrderNote)));
            $('zOrder').disabled = !fields.zOrder;
            $('stackInfo').textContent = fields.zOrder
                ? stackText()
                : (fields.zOrderNote || '');
            show('macroField', fields.macro);
            $('macroNote').textContent = fields.macro ? '' : (fields.macroNote || '');
            show('macroSection', fields.macro || Boolean(fields.macroNote));
            show('linkedCellField', fields.linkedCell);
            show('inputRangeField', fields.inputRange);
            show('controlSection', fields.linkedCell || fields.inputRange);
            $('delete').hidden = !fields.delete;
            const kind = model.mode === 'add' ? $('type').value : (model.shape && model.shape.kind);
            $('textLabel').textContent = ['button', 'checkBox', 'optionButton', 'label', 'groupBox'].includes(kind) ? 'Caption' : 'Text';
        }

        function stackText() {
            const at = model.shape && model.shape.zOrder;
            if (!at || !model.stackCount) { return ''; }
            return at === 1
                ? 'At the back, of ' + model.stackCount + ' shapes that stack here.'
                : at === model.stackCount
                    ? 'In front of the others, of ' + model.stackCount + ' shapes that stack here.'
                    : 'Number ' + at + ' from the back, of ' + model.stackCount + ' shapes that stack here.';
        }

        function fillMacros(current) {
            const select = $('macro');
            select.textContent = '';
            option(select, '', '(none)');
            const known = new Set();
            for (const macro of model.macros) {
                known.add(macro.macro.toLowerCase());
                option(select, macro.macro, macro.macro === macro.proc ? macro.proc + '  (' + macro.module + ')' : macro.macro);
            }
            if (current && !known.has(current.toLowerCase())) {
                option(select, current, current + '  (not a Public Sub in this project)');
            }
            select.value = current;
            $('macroHint').textContent = model.macros.length === 0
                ? 'The project has no Public Sub without required parameters to run yet.'
                : '';
            updateGoTo();
        }

        function updateGoTo() {
            $('goToMacro').disabled = !$('macro').value;
        }

        function load() {
            const v = model.values;
            const hostName = HOSTS[model.host] || model.host;
            if (model.mode === 'add') {
                $('title').textContent = 'New shape';
                $('subtitle').textContent = 'On ' + model.surface + ' in ' + model.fileName + ' (' + hostName + ')';
                for (const type of model.types) { option($('type'), type.value, type.label); }
                $('type').value = v.type;
                $('save').textContent = 'Add Shape';
                $('nameHint').textContent = 'Leave empty for the name ' + hostName + ' would give it.';
            } else {
                const kind = model.kindLabel || model.shape.kind;
                $('title').textContent = model.shape.name;
                $('subtitle').textContent = kind.charAt(0).toUpperCase() + kind.slice(1) + ' on ' + model.surface + ' in ' + model.fileName;
            }
            show('typeField', model.mode === 'add');
            show('surfaceField', model.mode === 'add' && Array.isArray(model.surfaces) && model.surfaces.length > 1);
            for (const surface of model.surfaces || []) { option($('surface'), surface, surface); }
            if (model.surfaces) { $('surface').value = v.surface; }

            for (const font of FONTS) { option($('fontNames'), font, font); }
            for (const dash of model.dashes) { option($('lineDash'), dash, DASH_LABELS[dash] || dash); }
            for (const id of ['name', 'altText', 'range', 'left', 'top', 'width', 'height', 'rotation', 'text',
                'linkedCell', 'inputRange', 'fontName', 'fontSize', 'fillTransparency', 'lineWeight']) {
                $(id).value = v[id];
            }
            for (const id of ['hidden', 'fontBold', 'fontItalic', 'fontUnderline', 'fontColorFromStyle']) {
                $(id).checked = v[id];
            }
            $('fontColor').value = v.fontColor.toLowerCase();
            $('fillColor').value = v.fillColor.toLowerCase();
            $('lineColor').value = v.lineColor.toLowerCase();
            $('lineDash').value = v.lineDash;
            $('zOrder').value = '';
            lookRadios('fill', v.fillType);
            lookRadios('line', v.lineType);
            fillMacros(v.macro);
            applyFields();
        }

        function values() {
            const out = {};
            for (const id of ['type', 'surface', 'name', 'altText', 'range', 'left', 'top', 'width', 'height', 'rotation',
                'text', 'macro', 'linkedCell', 'inputRange', 'zOrder', 'fontName', 'fontSize', 'fontColor',
                'fillColor', 'fillTransparency', 'lineColor', 'lineWeight', 'lineDash']) {
                out[id] = $(id).value;
            }
            for (const id of ['hidden', 'fontBold', 'fontItalic', 'fontUnderline', 'fontColorFromStyle']) {
                out[id] = $(id).checked;
            }
            out.fillType = checkedRadio('fillType') || model.values.fillType;
            out.lineType = checkedRadio('lineType') || model.values.lineType;
            return out;
        }

        function clearErrors() {
            for (const node of form.querySelectorAll('.fieldError')) { node.remove(); }
            for (const node of form.querySelectorAll('[aria-invalid="true"]')) {
                node.removeAttribute('aria-invalid');
                node.removeAttribute('aria-describedby');
            }
            $('errorSummary').hidden = true;
        }

        /** Errors by field, next to each field and listed at the top with links to them. */
        function showErrors(errors, general) {
            clearErrors();
            const list = $('errorList');
            list.textContent = '';
            for (const [key, message] of Object.entries(errors || {})) {
                const input = $(key) || form.querySelector('input[name="' + key + '"]');
                if (input) {
                    const note = document.createElement('p');
                    note.className = 'fieldError';
                    note.id = key + 'Error';
                    note.textContent = message;
                    const field = input.closest('.field') || input.parentElement;
                    field.appendChild(note);
                    input.setAttribute('aria-invalid', 'true');
                    input.setAttribute('aria-describedby', note.id);
                }
                const item = document.createElement('li');
                const link = document.createElement('a');
                link.href = '#' + key;
                link.textContent = message;
                link.addEventListener('click', (event) => {
                    event.preventDefault();
                    if (input) { input.focus(); }
                });
                item.appendChild(link);
                list.appendChild(item);
            }
            $('errorSummaryText').textContent = general || 'Fix these to save:';
            $('errorSummary').hidden = false;
            $('errorSummary').focus();
        }

        function busy(text) {
            for (const button of buttons) { button.disabled = Boolean(text); }
            $('status').textContent = text || '';
        }

        form.addEventListener('submit', (event) => {
            event.preventDefault();
            clearErrors();
            busy(model.mode === 'add' ? 'Adding...' : 'Saving...');
            vscode.postMessage({ type: 'save', values: values() });
        });
        $('cancel').addEventListener('click', () => vscode.postMessage({ type: 'cancel' }));
        $('delete').addEventListener('click', () => vscode.postMessage({ type: 'delete' }));
        $('goToMacro').addEventListener('click', () => vscode.postMessage({ type: 'goToMacro', macro: $('macro').value }));
        $('macro').addEventListener('change', updateGoTo);
        $('type').addEventListener('change', applyFields);

        // Changing a color, weight or dash means a solid fill or line.
        for (const id of ['fillColor', 'fillTransparency']) {
            $(id).addEventListener('input', () => setRadio('fillType', 'solid'));
        }
        for (const id of ['lineColor', 'lineWeight', 'lineDash']) {
            $(id).addEventListener('input', () => setRadio('lineType', 'solid'));
        }
        $('fontColor').addEventListener('input', () => { $('fontColorFromStyle').checked = false; });
        // An error clears as soon as its field is changed.
        form.addEventListener('input', (event) => {
            const target = event.target;
            if (target && target.getAttribute && target.getAttribute('aria-invalid') === 'true') {
                target.removeAttribute('aria-invalid');
                const note = $(target.id + 'Error');
                if (note) { note.remove(); }
            }
        });

        window.addEventListener('message', (event) => {
            const message = event.data;
            if (message.type === 'invalid') {
                busy('');
                showErrors(message.errors);
            } else if (message.type === 'error') {
                busy('');
                showErrors({}, message.error);
            } else if (message.type === 'idle') {
                busy('');
            }
        });

        load();
