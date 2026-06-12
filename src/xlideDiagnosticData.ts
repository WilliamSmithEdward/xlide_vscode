// Shared carrier for analyzer-provided diagnostic data on vscode.Diagnostic
// instances: the live diagnostics engine attaches the analyzer's
// VbaDiagnosticData under a module-private symbol and the code-action
// provider reads it back when building quick fixes.

import * as vscode from 'vscode';
import { type VbaDiagnosticData } from './analyzer';

export const XLIDE_DIAGNOSTIC_DATA = Symbol('xlideDiagnosticData');

export type XlideDiagnosticWithData = vscode.Diagnostic & {
    [XLIDE_DIAGNOSTIC_DATA]?: VbaDiagnosticData;
};
