// esbuild injects this module into the browser bundle, which rewrites every
// free `Buffer` reference in the engine to this export. Nothing imports it by
// hand; see the `inject` entry in webBuild.js.

export { WebBuffer as Buffer } from './webBuffer';
