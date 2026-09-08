// What kind of module a file on its own is.
//
// A module inside a project is described by the project: the container says
// which of its modules are standard, class, document or designer, and every
// rule that keys off module kind gets a settled answer. A file no project
// claims has to say so itself, and it does: the extension the VBE exports it
// under, plus the `Attribute VB_*` header it writes at the top. Reading
// neither left every loose file analyzed as a standard module, so `Me` and a
// document's event handlers reported against code that is correct (issue #73).

import {
    classifyDocumentType,
    classifyModuleType,
    type DocumentType,
    type ModuleType,
} from './vba/projectService';

export interface LooseModuleMetadata {
    moduleType: ModuleType;
    documentType: DocumentType | undefined;
}

/**
 * The module type and, for a document module, which document it is behind.
 *
 * `fileName` is the file's own name, extension included; `source` is its text,
 * which only has to reach past the `Attribute VB_*` header for the answer to
 * be exact.
 */
export function looseModuleMetadata(fileName: string, source: string): LooseModuleMetadata {
    const moduleName = moduleNameFromFileName(fileName);
    const moduleType = looseModuleType(fileName, moduleName, source);
    return {
        moduleType,
        documentType: moduleType === 'document'
            ? classifyDocumentType(moduleName, source)
            : undefined,
    };
}

function looseModuleType(fileName: string, moduleName: string, source: string): ModuleType {
    switch (extensionOf(fileName)) {
        // A .bas is a standard module whatever it is called. Classifying by
        // content would make a standard module named `Sheet1` a document.
        case 'bas':
            return 'standard';
        case 'frm':
            return 'userform';
        case 'ctl':
            return 'usercontrol';
        case 'pag':
            return 'propertypage';
        case 'dsr':
            return 'designer';
        // A .cls is a class, a document module, or a form, never a standard
        // module - the same upgrade the import planner applies.
        case 'cls': {
            const classified = classifyModuleType(moduleName, source);
            return classified === 'standard' ? 'class' : classified;
        }
        default:
            return classifyModuleType(moduleName, source);
    }
}

function extensionOf(fileName: string): string {
    const dot = fileName.lastIndexOf('.');
    return dot < 0 ? '' : fileName.slice(dot + 1).toLowerCase();
}

function moduleNameFromFileName(fileName: string): string {
    return fileName.replace(/\.[^.]+$/, '') || fileName;
}
