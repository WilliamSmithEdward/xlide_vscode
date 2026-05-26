// Minimal ambient type declarations for the `vsls` package.
//
// The `vsls` npm package ships its API in a `.ts` file (`node_modules/vsls/vscode.ts`)
// without a corresponding `.d.ts`. Under `moduleResolution: Node16`, TypeScript
// doesn't resolve `.ts` files inside `node_modules`, so we declare the subset of
// the API surface XLIDE actually uses.

declare module 'vsls/vscode' {
    import * as vscode from 'vscode';

    export const extensionId: string;

    export function getApi(callingExtensionId?: string): Promise<LiveShare | null>;

    export enum Role {
        None = 0,
        Host = 1,
        Guest = 2,
    }

    export interface Session {
        readonly id: string | null;
        readonly role: Role;
    }

    export interface SessionChangeEvent {
        readonly session: Session;
    }

    export interface SharedService {
        readonly isServiceAvailable: boolean;
        readonly onDidChangeIsServiceAvailable: vscode.Event<boolean>;
        onRequest(name: string, handler: (args: unknown[]) => unknown | Promise<unknown>): void;
        onNotify(name: string, handler: (args: object) => void): void;
        notify(name: string, args: object): void;
    }

    export interface SharedServiceProxy {
        readonly isServiceAvailable: boolean;
        readonly onDidChangeIsServiceAvailable: vscode.Event<boolean>;
        onNotify(name: string, handler: (args: object) => void): void;
        request(name: string, args: unknown[]): Promise<unknown>;
        notify(name: string, args: object): void;
    }

    export interface LiveShare {
        readonly session: Session;
        readonly onDidChangeSession: vscode.Event<SessionChangeEvent>;
        shareService(name: string): Promise<SharedService | null>;
        unshareService(name: string): Promise<void>;
        getSharedService(name: string): Promise<SharedServiceProxy | null>;
    }
}
