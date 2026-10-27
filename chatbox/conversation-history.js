// Included in the pinned Chatbox renderer by prepare.cjs. No stored history is rewritten.
const pagepilotHistory = (() => {
  function isDeepSeek(modelId, apiHost) {
    if (/(?:^|[\s/:._-])deepseek(?:$|[\s/:._-])/i.test(String(modelId || ''))) return true;
    try { return new URL(apiHost).hostname.toLowerCase() === 'api.deepseek.com'; }
    catch { return false; }
  }

  async function emitAssistantMessages(parts, dependencies, output, options, convert, stringifyError) {
    // Adjacent calls belong to one assistant response. Splitting them loses the
    // reasoning attached to that response and creates invented assistant turns.
    let segment = [];
    let calls = [];
    async function flush() {
      const converted = await convert(segment, dependencies, options);
      if (converted.length) output.push({ role: 'assistant', content: converted });
      if (calls.length) output.push({ role: 'tool', content: calls.map(call => ({
        type: 'tool-result', toolCallId: call.toolCallId, toolName: call.toolName,
        output: call.state === 'error'
          ? { type: 'error-text', value: stringifyError(call.result) }
          : { type: 'json', value: call.result ?? null },
      })) });
      segment = []; calls = [];
    }
    for (const part of parts) {
      if (part.type === 'tool-call' && part.state === 'call') continue;
      if (calls.length && part.type !== 'tool-call') await flush();
      segment.push(part);
      if (part.type === 'tool-call') calls.push(part);
    }
    await flush();
  }
  return { isDeepSeek, emitAssistantMessages };
})();
