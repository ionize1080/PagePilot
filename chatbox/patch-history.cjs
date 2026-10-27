const fs = require('node:fs');
const path = require('node:path');
const file = 'dist/renderer/js/index.TlLbGKvw.js';
function apply(patch, append) {
  patch(file, 'preserveReasoning:nu.provider===ModelProviderEnum.DeepSeek}',
    'preserveReasoning:nu.provider===ModelProviderEnum.DeepSeek||pagepilotHistory.isDeepSeek(Du.modelId,Du.options?.apiHost)}');
  // Keep the pinned upstream converter for content types and error formatting.
  patch(file, 'async function emitAssistantMessages(eu,uu,au,tu){',
    'async function emitAssistantMessages(eu,uu,au,tu){return pagepilotHistory.emitAssistantMessages(eu,uu,au,tu,convertAssistantContentParts,stringifyErrorResult);}\nasync function pagepilotOriginalEmitAssistantMessages(eu,uu,au,tu){');
  append(file, '\n' + fs.readFileSync(path.join(__dirname, 'conversation-history.js'), 'utf8'));
}
module.exports = { apply };
