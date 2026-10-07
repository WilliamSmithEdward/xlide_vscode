import * as vscode from 'vscode';
import { collectVbaColors, vbaColorPresentations, vbaColorNameShadowedAt } from './vbaColors';

export class VbaColorProvider implements vscode.DocumentColorProvider {
	provideDocumentColors(document: vscode.TextDocument, token: vscode.CancellationToken): vscode.ColorInformation[] {
		if (token.isCancellationRequested) { return []; }
		return collectVbaColors(document.getText(), () => token.isCancellationRequested).map(c => new vscode.ColorInformation(
			new vscode.Range(document.positionAt(c.start), document.positionAt(c.end)),
			new vscode.Color(c.red / 255, c.green / 255, c.blue / 255, 1),
		));
	}

	provideColorPresentations(color: vscode.Color, context: { document: vscode.TextDocument; range: vscode.Range }): vscode.ColorPresentation[] {
		const source = context.document.getText();
		const offset = context.document.offsetAt(context.range.start);
		const rgbShadowed = vbaColorNameShadowedAt(source, 'RGB', offset);
		const labels = vbaColorPresentations(color.red, color.green, color.blue);
		if (rgbShadowed) {
			if (vbaColorNameShadowedAt(source, 'VBA', offset)) { labels.shift(); }
			else { labels[0] = `VBA.${labels[0]}`; }
		}
		return labels.map(label => {
			const presentation = new vscode.ColorPresentation(label);
			presentation.textEdit = vscode.TextEdit.replace(context.range, label);
			return presentation;
		});
	}
}
