// Regression for the prior Folio/Qwen trailing-system compatibility failure.
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync('dist/Chatbox/native/dist/renderer/js/index.TlLbGKvw.js','utf8');
function extract(name){const start=source.indexOf('window.'+name+'=');assert(start>=0);const end=source.indexOf('\n};',start);assert(end>start);return source.slice(start,end+4);}
let document={name:'First.docx',path:'C:\\Synthetic\\First.docx'},workspace={workspace_id:'a'.repeat(32),sourceCount:2};
const scope={window:{pagepilot:{context:async()=>({...document})},pagepilotFiles:{context:async()=>({...workspace})},pagepilotWorkspaceSession:()=> 'qa'},mcpController:{getAvailableTools:()=>({mcp__pagepilot_context:{},mcp__workspace_list:{}})},setTimeout,Map};
vm.createContext(scope);vm.runInContext('const pagepilotContexts=new Map();'+['pagepilotExpand','pagepilotCapture','pagepilotCaptureWorkspace','pagepilotExpandWorkspace'].map(extract).join('\n'),scope);
const roles=messages=>messages.forEach((m,i)=>assert(!(i>0&&m.role==='system'),'No trailing system message'));
const expand=messages=>scope.window.pagepilotExpandWorkspace(scope.window.pagepilotExpand(messages));
(async()=>{
 const image={type:'image',storageKey:'synthetic-image'},raw=Object.freeze({id:'first',role:'user',contentParts:Object.freeze([{type:'text',text:'分析资料'},image])});
 const first=await scope.window.pagepilotCapture(await scope.window.pagepilotCaptureWorkspace(raw));
 const system={role:'system',contentParts:[{type:'text',text:'Original rules'}]},history=[system,first];
 const result=expand(history);roles(result);assert.equal(result[0],system);assert.equal(result[1].contentParts[1],image);assert.equal(raw.contentParts.length,2);
 assert(JSON.stringify(result[1]).includes('First.docx'));assert(JSON.stringify(result[1]).includes('a'.repeat(32)));
 document={name:'Second.docx',path:'C:\\Synthetic\\Second.docx'};workspace={workspace_id:'b'.repeat(32),sourceCount:1};
 const second=await scope.window.pagepilotCapture(await scope.window.pagepilotCaptureWorkspace({id:'second',role:'user',contentParts:[{type:'text',text:'继续'}]}));
 const call={role:'assistant',contentParts:[{type:'tool-call',toolCallId:'t1',toolName:'workspace_read',input:{}}]},tool={role:'tool',contentParts:[{type:'tool-result',toolCallId:'t1',output:'reference'}]};
 const continued=expand([...history,{role:'assistant',contentParts:[{type:'text',text:'done'}]},second,call,tool]);
 roles(continued);assert.equal(continued[1],first);assert.equal(continued.at(-2),call);assert.equal(continued.at(-1),tool);
 assert(JSON.stringify(continued[3]).includes('Second.docx'));assert(JSON.stringify(continued[3]).includes('b'.repeat(32)));assert(!JSON.stringify(continued[3]).includes('First.docx'));
 assert.equal(expand([system])[0],system);
 console.log('PASS production Word/workspace context: new and existing chats, fresh document and workspace, no trailing system, images and tool pairing retained, original messages unchanged');
})().catch(e=>{console.error(e);process.exitCode=1});
