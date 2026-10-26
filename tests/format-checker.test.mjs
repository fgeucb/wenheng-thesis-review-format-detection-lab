import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import {createFormatService,formatConfig,summarizeFormatReport} from '../src/format-checker.mjs';
import {fixtureRepo} from './helpers.mjs';

const report=()=>({findings:[{rule_id:'067',severity:'错误'}],counts_by_severity:{错误:1,警告:2,提醒:3},unknowns:['无法确定页码'],report_only:['公式只报告'],skipped:['语义判断'],timings:{总耗时:12}});
const bytes=Buffer.from([0x50,0x4b,0x03,0x04,1,2,3,4]);
function upload(){const form=new FormData();form.set('template',new Blob([bytes]),'模板.docx');form.set('document',new Blob([bytes]),'论文.docx');return form;}
async function writeResult(task,config,{onMessage}){onMessage('正在写入批注');await fs.writeFile(task.output_path,bytes);await fs.writeFile(task.report_path,JSON.stringify(report()));}
async function fixture(context,runner=writeResult){
 const base=await fixtureRepo(),service=createFormatService(base.repo,base.dir,{config:{configured:true},runner});
 context.after(async()=>{await service.shutdown();await base.cleanup();});return {...base,service};
}

test('format configuration rejects a missing checker and report summaries retain incomplete checks',()=>{
 assert.equal(formatConfig({FORMAT_CHECKER_ROOT:'missing-format-checker'}).configured,false);
 assert.deepEqual(summarizeFormatReport(report()),{issues:6,errors:1,warnings:2,hints:3,unknown:1,report_only:1,skipped:1,seconds:12,limited:true});
 assert.throws(()=>summarizeFormatReport({findings:[]}),/报告结构无效/);
});

test('format tasks persist, restrict ownership, download results and reuse stored inputs',async context=>{
 const current=await fixture(context),owner=current.user.id;
 const submitted=await current.service.submit(owner,upload());await current.service.idle();
 const completed=await current.service.get(submitted.task_id,owner);assert.equal(completed.status,'succeeded');assert.equal(completed.summary.issues,6);
 assert.equal(completed.template.name,'模板.docx');assert.equal(completed.output_name,'论文_带批注.docx');assert.equal(completed.output_path,undefined);
 assert.equal((await current.service.latest(owner)).task_id,submitted.task_id);assert.equal(await current.service.latest('other-user'),null);
 await assert.rejects(current.service.get(submitted.task_id,'other-user'),{status:404});
 await assert.rejects(current.service.download(submitted.task_id,'other-user'),{status:404});
 assert.deepEqual(await fs.readFile((await current.service.download(submitted.task_id,owner)).path),bytes);
 const reuse=new FormData();reuse.set('source_task_id',submitted.task_id);
 await assert.rejects(current.service.submit('other-user',reuse),{status:404});
 const repeated=await current.service.submit(owner,reuse);await current.service.idle();assert.notEqual(repeated.task_id,submitted.task_id);
 assert.equal((await current.service.get(repeated.task_id,owner)).status,'succeeded');
 const stored=await current.repo.one('SELECT * FROM format_tasks WHERE id=?',[repeated.task_id]);assert.deepEqual(await fs.readFile(stored.template_path),bytes);
});

test('format upload validation rejects missing, duplicate and non-DOCX inputs',async context=>{
 const current=await fixture(context),owner=current.user.id;
 await assert.rejects(current.service.submit(owner,new FormData()),{status:400});
 const invalid=upload();invalid.set('template',new Blob(['not a docx']),'模板.docx');await assert.rejects(current.service.submit(owner,invalid),{status:400});
 const extension=upload();extension.set('template',new Blob([bytes]),'模板.pdf');await assert.rejects(current.service.submit(owner,extension),{status:400});
 const duplicate=upload();duplicate.append('template',new Blob([bytes]),'其他.docx');await assert.rejects(current.service.submit(owner,duplicate),{status:400});
 assert.equal(await current.service.latest(owner),null);
});

test('format queue rejects concurrent uploads from one account',async context=>{
 const gate=Promise.withResolvers();context.after(()=>gate.resolve());
 const current=await fixture(context,async(...args)=>{await gate.promise;await writeResult(...args);});
 try{
  const results=await Promise.allSettled([current.service.submit(current.user.id,upload()),current.service.submit(current.user.id,upload())]);
  assert.equal(results.filter(result=>result.status==='fulfilled').length,1);assert.equal(results.find(result=>result.status==='rejected').reason.status,409);
  const accepted=results.find(result=>result.status==='fulfilled').value;await assert.rejects(current.service.download(accepted.task_id,current.user.id),{status:409});
 }finally{gate.resolve();await current.service.idle();}
});

test('format queue failure allows retry and shutdown records interruption',async context=>{
 let attempt=0;
 const current=await fixture(context,async(...args)=>{if(!attempt++)throw Error('Word 不可用');await writeResult(...args);});
 const failed=await current.service.submit(current.user.id,upload());await current.service.idle();assert.equal((await current.service.get(failed.task_id,current.user.id)).status,'failed');
 const retry=new FormData();retry.set('source_task_id',failed.task_id);const done=await current.service.submit(current.user.id,retry);await current.service.idle();
 assert.equal((await current.service.get(done.task_id,current.user.id)).status,'succeeded');
 const started=Promise.withResolvers();
 const waiting=createFormatService(current.repo,current.dir,{config:{configured:true},runner:async(task,config,{signal})=>{started.resolve();await new Promise((resolve,reject)=>{signal.addEventListener('abort',()=>reject(Error('中断')),{once:true});});}});
 context.after(()=>waiting.shutdown());const running=await waiting.submit(current.user.id,upload());await started.promise;await waiting.shutdown();
 assert.equal((await waiting.get(running.task_id,current.user.id)).status,'interrupted');
});

test('format queue lock serializes two service instances and recovers interrupted tasks',async context=>{
 const current=await fixture(context),owner=current.user.id;
 const initial=await current.service.submit(owner,upload());await current.service.idle();
 await current.repo.q("UPDATE format_tasks SET status='running' WHERE id=?",[initial.task_id]);
 const secondOwner=crypto.randomUUID();await current.repo.createUser({id:secondOwner,username:'format_second',phone:'',credentials:{hash:'00',salt:'00'},created_at:new Date().toISOString()});
 const gate=Promise.withResolvers(),started=Promise.withResolvers();
 const first=createFormatService(current.repo,current.dir,{config:{configured:true},runner:async(...args)=>{started.resolve();await gate.promise;await writeResult(...args);}});
 const second=createFormatService(current.repo,current.dir,{config:{configured:true},runner:async()=>{throw Error('queue lock was bypassed');}});
 context.after(async()=>{gate.resolve();await first.shutdown();await second.shutdown();});
 try{
  await first.submit(secondOwner,upload());await started.promise;
  assert.equal((await first.get(initial.task_id,owner)).status,'interrupted');
  const queued=await second.submit(owner,upload());await second.idle();assert.equal((await second.get(queued.task_id,owner)).status,'queued');
  gate.resolve();await first.idle();assert.equal((await second.get(queued.task_id,owner)).status,'succeeded');
 }finally{gate.resolve();await first.idle();}
});
