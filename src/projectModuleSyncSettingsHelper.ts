import * as path from 'node:path';


/**
 * Resolves the path to the folder used for exporting, importing, and previewing modules, accounting for both absolute and relative paths.<br>
 * This ensures that, if a relative path is used, it is interpreted as relative to the project file
 * @param projectPath It is the absolute path, including the filename, of the project
 * @param folderPath It is the path — absolute or relative — to the folder that needs to be resolved
 * @returns Absolute path to the export, import, or preview folder
 */
export function resolveProjectModuleSyncFolderPath(
    projectPath: string,
    folderPath?: string
): string | undefined
{
    return folderPath ? 
        path.isAbsolute(folderPath) ? 
            folderPath : 
            path.resolve(path.dirname(projectPath), folderPath) :
        undefined;
}
