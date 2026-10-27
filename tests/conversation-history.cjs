// Exercise the shipped stream processor, history converter and SDK serializer.
// The local endpoint enforces DeepSeek's documented reasoning/tool history contract.
const fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const assert = require('node:assert/strict'), http = require('node:http');
const root = path.resolve(__dirname, '..');
const bundle = fs.readFileSync((process.argv[2] && !process.argv[2].startsWith('--') ? process.argv[2] : undefined) || path.join(root, 'dist/Chatbox/native/dist/renderer/js/index.TlLbGKvw.js'), 'utf8');
const sdk = fs.readFileSync(path.join(root, 'dist/Chatbox/native/dist/renderer/js/vendor-ai.CMAD2Bsd.js'), 'utf8');
function range(text, start, end) {
  const a = text.indexOf(start), b = text.indexOf(end, a + start.length);
  assert(a >= 0 && b > a, `Pinned code boundaries: ${start}`);
  return text.slice(a, b);
}
const context = vm.createContext({ URL, Date, BaseError: Error, lodashExports: { compact: a => a.filter(Boolean) },
  createModelDependencies: async () => ({}), getMessageText: m => m.contentParts.filter(p => p.type === 'text').map(p => p.text).join('\n') });
let code = range(bundle, 'function stringifyErrorResult', 'function injectModelSystemPrompt') +
  range(bundle, 'function createInitialState', 'const DEFAULT_LINES=') +
  range(sdk, 'function ur(e)', 'function fp(');
if (bundle.includes('const pagepilotHistory =')) code += bundle.slice(bundle.indexOf('const pagepilotHistory ='));
const flag = bundle.match(/preserveReasoning:(nu\.provider===ModelProviderEnum\.DeepSeek[^}]*)(?=\})/);
assert(flag, 'Production provider detection found');
code += `\nfunction preserve(modelId, apiHost, provider='custom'){const nu={provider},Du={modelId,options:{apiHost}},ModelProviderEnum={DeepSeek:'deepseek'};return ${flag[1]};}`;
vm.runInContext(code, context);
const evaluate = expression => vm.runInContext(expression, context);
const send = async (url, messages) => {
  const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({model:'deepseek-flash',tools:[{type:'function',function:{name:'pagepilot_context'}}],messages}) });
  return { status: response.status, body: await response.json() };
};
const server = http.createServer(async (req, res) => {
  let body = ''; for await (const chunk of req) body += chunk;
  const { messages } = JSON.parse(body); const pending = new Set(); let error;
  for (const message of messages) {
    if (message.role === 'assistant') {
      if (pending.size) error = 'Unanswered tool calls';
      if (!message.reasoning_content) error = 'Missing reasoning_content in assistant history';
      for (const call of message.tool_calls || []) pending.add(call.id);
    } else if (message.role === 'tool') {
      if (!pending.delete(message.tool_call_id)) error = 'Unpaired tool result';
    } else if (pending.size) error = 'New user turn before tool results';
  }
  if (pending.size) error = 'Missing tool results';
  res.writeHead(error ? 400 : 200, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(error ? { error } : { ok: true }));
});
(async () => {
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const url = `http://127.0.0.1:${server.address().port}/chat/completions`;
  context.history = [{ role: 'user', contentParts: [{type:'text',text:'Read the synthetic document'}] }];
  const serialize = () => evaluate(`(async()=>FO(await convertToModelMessages(history,{modelSupportVision:false,preserveReasoning:preserve('deepseek-flash','https://gateway.example/v1')})))()`);
  assert.equal((await send(url, await serialize())).status, 200);
  // Feed real stream chunks, including two parallel tools, through production code.
  context.chunks = [
    {type:'reasoning-delta',text:'Synthetic reasoning before tools.'},
    {type:'tool-call',toolCallId:'call_a',toolName:'pagepilot_context',input:{}},
    {type:'tool-call',toolCallId:'call_b',toolName:'pagepilot_capabilities',input:{}},
    {type:'tool-result',toolCallId:'call_b',output:{version:'test'}},
    {type:'tool-result',toolCallId:'call_a',output:{name:'synthetic.docx'}},
    {type:'reasoning-delta',text:'Synthetic reasoning after tools.'},
    {type:'text-delta',text:'Read complete.'},
  ];
  await evaluate(`(async()=>{let state=createInitialState();for(const chunk of chunks)state=(await processStreamChunk(chunk,state,{})).state;history.push({role:'assistant',contentParts:state.contentParts},{role:'user',contentParts:[{type:'text',text:'Continue'}]});})()`);
  const second = await send(url, await serialize());
  if (process.argv.includes('--expect-bug')) {
    assert.equal(second.status, 400); console.log('REPRODUCED: turn 1 HTTP 200; turn 2 HTTP 400:', second.body.error); return;
  }
  assert.equal(second.status, 200, JSON.stringify(second.body));
  let wire = await serialize();
  assert.equal(wire.filter(m => m.role === 'assistant').length, 2);
  assert.equal(wire.find(m => m.tool_calls)?.tool_calls.length, 2);
  assert.equal(wire[1].reasoning_content, 'Synthetic reasoning before tools.');
  assert.equal(wire[4].reasoning_content, 'Synthetic reasoning after tools.');
  // Persistence/reload must preserve history, independent of the in-memory Word map.
  context.history = JSON.parse(JSON.stringify(context.history));
  context.history.push({role:'assistant',contentParts:[{type:'reasoning',text:'Third-turn evidence.'},{type:'text',text:'Still ready.'}]},{role:'user',contentParts:[{type:'text',text:'Third request'}]});
  assert.equal((await send(url, await serialize())).status, 200);
  for (const id of ['deepseek-chat','deepseek-reasoner','deepseek/deepseek-v3.2','vendor/DeepSeek-V4']) assert(evaluate(`preserve(${JSON.stringify(id)},'https://gateway.example/v1')`));
  assert(evaluate(`preserve('custom-alias','https://api.deepseek.com/v1')`));
  for (const id of ['grok-4','mistral-large','gpt-4o','notdeepseek-fake']) assert.equal(evaluate(`preserve(${JSON.stringify(id)},'https://gateway.example/v1')`), false);
  assert.equal(evaluate(`preserve('alias','https://api.deepseek.com.attacker.example')`), false);
  assert(evaluate(`preserve('alias',undefined,'deepseek')`));
  context.history = [{role:'assistant',contentParts:[{type:'reasoning',text:'Do not leak to other providers.'},{type:'text',text:'Answer'},{type:'tool-call',state:'call',toolCallId:'pending',toolName:'x',args:{}},{type:'tool-call',state:'error',toolCallId:'failed',toolName:'x',args:{},result:{message:'Denied'}}]}];
  wire = await evaluate(`(async()=>FO(await convertToModelMessages(history,{modelSupportVision:false,preserveReasoning:preserve('grok-4')})))()`);
  assert(!('reasoning_content' in wire[0])); assert.equal(wire[0].tool_calls.length,1); assert.equal(wire[1].content,'Denied');
  console.log('PASS: three HTTP turns, persisted history, parallel tools, exact reasoning, provider isolation, failed/pending tools');
  const keyIndex = process.argv.indexOf('--live-key');
  const liveIndex = process.argv.indexOf('--live-config');
  if (liveIndex !== -1 || keyIndex !== -1) {
    // Explicit opt-in. Credentials stay in memory; use only a synthetic tool/document.
    const settings = keyIndex !== -1 ? {apiKey:fs.readFileSync(process.argv[keyIndex+1],'utf8').match(/sk-[A-Za-z0-9_-]+/)[0],models:[{modelId:'deepseek-flash'}]} : JSON.parse(fs.readFileSync(process.argv[liveIndex + 1], 'utf8')).settings.providers.deepseek;
    assert(settings.apiKey, 'Configured DeepSeek credential required');
    const model = settings.models[0].modelId;
    context.history = [];
    for (let turn = 1; turn <= 3; turn++) {
      context.history.push({role:'user',contentParts:[{type:'text',text:`Test turn ${turn}. Call pagepilot_context once for the synthetic document, then answer only its name. Always call the tool once this turn. If it fails, report the failure and do not retry.`}]});
      const state = evaluate('createInitialState()');
      let completed = false, calls = 0;
      for (let step = 0; step < 4; step++) {
        context.current = [...context.history, ...(state.contentParts.length ? [{role:'assistant',contentParts:state.contentParts}] : [])];
        context.liveModel = model;
        const messages = await evaluate(`(async()=>FO(await convertToModelMessages(current,{modelSupportVision:false,preserveReasoning:preserve(liveModel,'https://gateway.example/v1')})))()`);
        const response = await fetch('https://api.deepseek.com/chat/completions', {
          method:'POST', headers:{'Content-Type':'application/json',Authorization:'Bearer '+settings.apiKey},
          body:JSON.stringify({model, messages, stream:false, max_tokens:1500, thinking:{type:'enabled'}, tools:[{type:'function',function:{name:'pagepilot_context',description:'Returns the synthetic test document name.',parameters:{type:'object',properties:{},additionalProperties:false}}}]}),
          signal:AbortSignal.timeout(90000),
        });
        const payload = await response.json();
        assert.equal(response.status,200,`Live turn ${turn}, step ${step}, HTTP ${response.status}; code=${payload.error?.code || 'unknown'}`);
        const message = payload.choices[0].message;
        const chunks = [];
        if(message.reasoning_content)chunks.push({type:'reasoning-delta',text:message.reasoning_content});
        if(message.content)chunks.push({type:'text-delta',text:message.content});
        for(const call of message.tool_calls || [])chunks.push({type:'tool-call',toolCallId:call.id,toolName:call.function.name,input:JSON.parse(call.function.arguments)});
        for(const call of message.tool_calls || []){calls++;chunks.push({type:'tool-result',toolCallId:call.id,output:turn===1?Object.assign(new Error('Synthetic tool rejection'),{name:'MCPClientError',cause:{message:'Synthetic tool rejection'}}):{name:'PagePilot synthetic regression document'}});}
        context.liveState=state; context.liveChunks=chunks;
        Object.assign(state,await evaluate(`(async()=>{let s=liveState;for(const c of liveChunks)s=(await processStreamChunk(c,s,{})).state;return s;})()`));
        if(!message.tool_calls?.length){assert.equal(payload.choices[0].finish_reason,'stop');completed=true;break;}
      }
      assert(completed && calls > 0, `Live turn ${turn} completed with a tool call`);
      context.history.push({role:'assistant',contentParts:state.contentParts});
      if(turn===2)context.history=JSON.parse(JSON.stringify(context.history));
      console.log(`PASS live DeepSeek compatible history: turn ${turn}, ${calls} synthetic tool calls`);
    }
  }
})().catch(error => { console.error(error); process.exitCode = 1; }).finally(() => server.close());
