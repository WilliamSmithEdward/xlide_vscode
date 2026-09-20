// Highlight for agent writes nobody has kept or reverted yet, and for
// modules that changed since the last commit.
//
// A tree row's label can only be coloured through a file decoration, and a
// file decoration only reaches a row that carries a `resourceUri`. The rows
// get one in a scheme nothing else decorates, so git, problems and file-icon
// themes never touch them, and the provider below is the only one that
// answers for it. A module row is coloured and badged while its review is
// pending, or marked `M`/`A` in git's own colours when it differs from HEAD;
// a project row is coloured and carries the count either way, so a collapsed
// tree still says where the edits are. An agent edit outranks a git mark on
// the same row: it is the one waiting for a decision.

import * as vscode from 'vscode';
import {
    hasPendingAgentReview,
    onDidChangePendingAgentReviews,
    pendingAgentReviewModules,
} from './xlideAgentDiff';
import type { GitChangeMarksSource } from './gitChangeMarks';

export const AGENT_REVIEW_DECORATION_SCHEME = 'xlide-agent-review';

/** The colour an agent edit awaiting review is drawn in; themable in package.json. */
export const AGENT_REVIEW_COLOR_ID = 'xlide.agentEdit.foreground';

type DecorationTarget =
    | { kind: 'project'; filePath: string }
    | { kind: 'module'; filePath: string; moduleName: string };

/** The URI a project row carries, whether or not it has anything pending. */
export function projectDecorationUri(filePath: string): vscode.Uri {
    return decorationUri({ kind: 'project', filePath });
}

/** The URI a module row carries while its agent edit awaits review. */
export function moduleDecorationUri(filePath: string, moduleName: string): vscode.Uri {
    return decorationUri({ kind: 'module', filePath, moduleName });
}

/**
 * The identity rides in the path, base64url-encoded, so a Windows path's
 * colons and separators can never be read as URI structure, and two modules
 * never share a key that ignores the query.
 */
function decorationUri(target: DecorationTarget): vscode.Uri {
    const payload = target.kind === 'module'
        ? [target.filePath, target.moduleName]
        : [target.filePath];
    const encoded = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
    return vscode.Uri.from({
        scheme: AGENT_REVIEW_DECORATION_SCHEME,
        path: `/${target.kind}/${encoded}`,
    });
}

/** Reads a decoration URI back; undefined for any URI this module did not build. */
export function parseDecorationUri(uri: vscode.Uri): DecorationTarget | undefined {
    if (uri.scheme !== AGENT_REVIEW_DECORATION_SCHEME) {
        return undefined;
    }
    const match = /^\/(project|module)\/([A-Za-z0-9_-]+)$/.exec(uri.path);
    if (!match) {
        return undefined;
    }
    let payload: unknown;
    try {
        payload = JSON.parse(Buffer.from(match[2], 'base64url').toString('utf8'));
    } catch {
        return undefined;
    }
    if (!Array.isArray(payload) || !payload.every((part) => typeof part === 'string')) {
        return undefined;
    }
    if (match[1] === 'module') {
        return payload.length === 2
            ? { kind: 'module', filePath: payload[0], moduleName: payload[1] }
            : undefined;
    }
    return payload.length === 1 ? { kind: 'project', filePath: payload[0] } : undefined;
}

/** A decoration badge holds at most two characters. */
export function pendingCountBadge(count: number): string {
    return count > 9 ? '9+' : String(count);
}

/** A count badge holds two characters. */
export function changedCountBadge(count: number): string {
    return count > 99 ? '99' : String(count);
}

export class AgentReviewDecorationProvider implements vscode.FileDecorationProvider, vscode.Disposable {
    private readonly _emitter = new vscode.EventEmitter<undefined>();
    readonly onDidChangeFileDecorations = this._emitter.event;
    private readonly _subscriptions: vscode.Disposable[];

    constructor(private readonly _gitMarks?: GitChangeMarksSource) {
        // A change can recolour a project row and every module in it, and the
        // rows carrying this scheme are few, so every one is asked again rather
        // than working out which URIs a given change reached.
        this._subscriptions = [
            onDidChangePendingAgentReviews(() => this._emitter.fire(undefined)),
            ...(_gitMarks ? [_gitMarks.onDidChange(() => this._emitter.fire(undefined))] : []),
        ];
    }

    provideFileDecoration(uri: vscode.Uri): vscode.FileDecoration | undefined {
        const target = parseDecorationUri(uri);
        if (!target) {
            return undefined;
        }
        return this._agentDecoration(target) ?? this._gitDecoration(target);
    }

    private _agentDecoration(target: DecorationTarget): vscode.FileDecoration | undefined {
        const color = new vscode.ThemeColor(AGENT_REVIEW_COLOR_ID);
        if (target.kind === 'module') {
            if (!hasPendingAgentReview(target.filePath, target.moduleName)) {
                return undefined;
            }
            return {
                badge: 'AI',
                color,
                tooltip: 'An agent edited this module. Keep or revert the change.',
            };
        }
        const count = pendingAgentReviewModules(target.filePath).length;
        if (count === 0) {
            return undefined;
        }
        return {
            badge: pendingCountBadge(count),
            color,
            tooltip: count === 1
                ? '1 agent edit awaiting review'
                : `${count} agent edits awaiting review`,
        };
    }

    /** The Explorer's own `M` and `A`, in the Explorer's own colours. */
    private _gitDecoration(target: DecorationTarget): vscode.FileDecoration | undefined {
        const marks = this._gitMarks?.marksFor(target.filePath);
        if (!marks) {
            return undefined;
        }
        if (target.kind === 'module') {
            const kind = marks.byModule.get(target.moduleName.toLowerCase());
            if (kind === 'modified') {
                return {
                    badge: 'M',
                    color: new vscode.ThemeColor('gitDecoration.modifiedResourceForeground'),
                    tooltip: 'Modified since the last commit',
                };
            }
            if (kind === 'added') {
                return {
                    badge: 'A',
                    color: new vscode.ThemeColor('gitDecoration.addedResourceForeground'),
                    tooltip: 'Not in the last commit',
                };
            }
            return undefined;
        }
        const count = marks.byModule.size + marks.removed;
        if (count === 0) {
            return undefined;
        }
        return {
            badge: changedCountBadge(count),
            color: new vscode.ThemeColor('gitDecoration.modifiedResourceForeground'),
            tooltip: count === 1 ? '1 module changed since the last commit' : `${count} modules changed since the last commit`,
        };
    }

    dispose(): void {
        for (const subscription of this._subscriptions) {
            subscription.dispose();
        }
        this._emitter.dispose();
    }
}
