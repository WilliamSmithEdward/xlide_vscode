// The extension registers every host's object model at load; the tests run
// the analyzer as the extension does. vbaHostModelRegistration.test.ts checks
// what an embedder gets without this.
import { registerBuiltInHostModels } from '../src/analyzer/host/builtInHostModels';

registerBuiltInHostModels();
