// Analysis worker entry point: bundled separately (out/analysisWorker.js) and
// spawned by the extension host so the full analysis pass never blocks the
// extension-host event loop. All behavior lives in AnalysisWorkerState; this
// file only wires the message port. Must never import 'vscode'.

import { parentPort } from 'worker_threads';
import { AnalysisWorkerState } from './analysisWorkerLogic';
import type { AnalysisWorkerRequest } from './analysisWorkerProtocol';

const state = new AnalysisWorkerState();

parentPort?.on('message', (request: AnalysisWorkerRequest) => {
	const response = state.handle(request);
	if (response) {
		parentPort?.postMessage(response);
	}
});
