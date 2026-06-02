import * as path from 'path';
import type * as vscode from 'vscode';

export interface XlideOutputLogEntry {
    timestamp: string;
    line: string;
}

const MAX_OUTPUT_LOG_ENTRIES = 250;
const WINDOWS_PATH_RE = /[A-Za-z]:\\[^\s'")]+/g;
const FILE_URI_RE = /\b(?:file|xlide-vba):\/\/[^\s'")]+/g;
const POSIX_PATH_RE = /(?:^|[\s'"])(\/(?:Users|home|var|tmp|mnt)\/[^\s'")]+)/g;

const outputLog: XlideOutputLogEntry[] = [];

export function recordXlideOutputLine(line: string, now = new Date()): void {
    for (const part of String(line).split(/\r\n|\r|\n/)) {
        outputLog.push({
            timestamp: now.toISOString(),
            line: redactSupportLogLine(part),
        });
    }
    if (outputLog.length > MAX_OUTPUT_LOG_ENTRIES) {
        outputLog.splice(0, outputLog.length - MAX_OUTPUT_LOG_ENTRIES);
    }
}

export function recentXlideOutputLog(limit = 50): XlideOutputLogEntry[] {
    return outputLog.slice(-Math.max(0, limit));
}

export function clearXlideOutputLog(): void {
    outputLog.length = 0;
}

export function redactSupportLogLine(line: string): string {
    return line
        .replace(FILE_URI_RE, '<redacted-uri>')
        .replace(WINDOWS_PATH_RE, redactPathMatch)
        .replace(POSIX_PATH_RE, (full, captured: string) =>
            full.replace(captured, redactPathMatch(captured)),
        );
}

export function createRecordedOutputChannel(channel: vscode.OutputChannel): vscode.OutputChannel {
    return new RecordedOutputChannel(channel);
}

function redactPathMatch(value: string): string {
    const ext = path.extname(value);
    return ext ? `<redacted>${ext}` : '<redacted>';
}

class RecordedOutputChannel implements vscode.OutputChannel {
    constructor(private readonly channel: vscode.OutputChannel) {}

    get name(): string {
        return this.channel.name;
    }

    append(value: string): void {
        this.channel.append(value);
    }

    appendLine(value: string): void {
        recordXlideOutputLine(value);
        this.channel.appendLine(value);
    }

    replace(value: string): void {
        clearXlideOutputLog();
        recordXlideOutputLine(value);
        this.channel.replace(value);
    }

    clear(): void {
        clearXlideOutputLog();
        this.channel.clear();
    }

    show(column?: vscode.ViewColumn, preserveFocus?: boolean): void;
    show(preserveFocus?: boolean): void;
    show(columnOrPreserveFocus?: vscode.ViewColumn | boolean, preserveFocus?: boolean): void {
        if (typeof columnOrPreserveFocus === 'boolean') {
            this.channel.show(columnOrPreserveFocus);
            return;
        }
        this.channel.show(columnOrPreserveFocus, preserveFocus);
    }

    hide(): void {
        this.channel.hide();
    }

    dispose(): void {
        this.channel.dispose();
    }
}
