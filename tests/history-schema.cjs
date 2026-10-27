// Validate through the exact shipped AI SDK schema, including live Error instances.
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');
const sdk=fs.readFileSync(path.join(root,'dist/Chatbox/native/dist/renderer/js/vendor-ai.CMAD2Bsd.js'),'utf8')
 .replace(/import\{[^}]+\}from"[^"]+";/g,'').replace(/export\{[^}]+\};/g,'');
const c=vm.createContext({TransformStream,ReadableStream,WritableStream,FormData,DOMException,btoa,atob,URL,TextDecoder,TextEncoder,AbortController,AbortSignal,Headers,Request,Response,fetch,Buffer,process,console,setTimeout,clearTimeout,G$:x=>x});
vm.runInContext(sdk,c);
const b=fs.readFileSync(path.join(root,'dist/Chatbox/native/dist/renderer/js/index.TlLbGKvw.js'),'utf8');
c.lodashExports={compact:a=>a.filter(Boolean)};c.createModelDependencies=async()=>({});c.getMessageText=m=>m.contentParts.filter(p=>p.type==='text').map(p=>p.text).join('\n');
vm.runInContext(b.slice(b.indexOf('function stringifyErrorResult'),b.indexOf('function injectModelSystemPrompt'))+b.slice(b.indexOf('const pagepilotHistory =')),c);
(async()=>{
 // Upstream catches MCP rejections and returns the Error as a successful value.
 const error=Object.assign(new Error('Synthetic document binding rejected'),{name:'MCPClientError',cause:{code:0,message:'Synthetic document binding rejected'},code:0});
 c.history=[{role:'user',contentParts:[{type:'text',text:'Read a synthetic document'}]},
 {role:'assistant',contentParts:[{type:'reasoning',text:'Read first.'},{type:'tool-call',toolCallId:'bad',toolName:'read',args:{},state:'result',result:error},
 {type:'reasoning',text:'The tool failed.'},{type:'text',text:'Activate the correct document.'}]},
 {role:'user',contentParts:[{type:'text',text:'Continue'}]}];
 const validate=async()=>{c.messages=await vm.runInContext('convertToModelMessages(history,{preserveReasoning:true,modelSupportVision:false})',c);await vm.runInContext('Mx({messages})',c);return c.messages;};
 // Establish the precise regression independently of the patched converter.
 c.original=[{role:'tool',content:[{type:'tool-result',toolCallId:'bad',toolName:'read',output:{type:'json',value:error}}]}];
 await assert.rejects(vm.runInContext('Mx({messages:original})',c),/ModelMessage/);
 let messages=await validate();assert.equal(messages[2].content[0].output.type,'error-text');assert.match(messages[2].content[0].output.value,/binding rejected/);
 c.history=JSON.parse(JSON.stringify(c.history));messages=await validate();assert.equal(messages[2].content[0].output.type,'error-text');
 for(const result of [new Error('denied'),{array:[undefined,NaN,Infinity],date:new Date(0)},null]){c.history[1].contentParts[1].result=result;await validate();}
 const cycle={};cycle.self=cycle;c.history[1].contentParts[1].result=cycle;messages=await validate();assert.equal(messages[2].content[0].output.type,'error-text');
 assert(b.includes('try{return await iu?.(su,lu)}catch(cu){throw cu}'),'MCP rejects instead of resolving errors');
 if(process.argv[2]){c.history=JSON.parse(fs.readFileSync(process.argv[2],'utf8')).messages;await validate();console.log('PASS private session replay through shipped SDK (content not logged)');}
 console.log('PASS shipped SDK: reproduced old failure; live errors, persisted errors, JSON values, cyclic outputs, MCP rejection propagation');
})().catch(e=>{console.error(e.message);process.exitCode=1;});
