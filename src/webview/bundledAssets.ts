// The webview templates, carried inside the browser bundle.
//
// A desktop build reads assets/webview/** from the installed extension
// directory at runtime (see extensionAssets.ts) and leaves this empty. A
// browser has no extension directory, so the web build fills this in at
// bundle time: webBuild.js replaces the contents of this module with the real
// files. Declaring it here rather than as a virtual module keeps the desktop
// type-check honest - both builds compile the same source.

/** Extension-root-relative path ('assets/webview/x.html') to LF-normalized text. */
export const bundledTextAssets: Record<string, string> = {};
