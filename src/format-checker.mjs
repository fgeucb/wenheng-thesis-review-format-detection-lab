import fs from 'node:fs/promises';
import {existsSync} from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';

const root=fileURLToPath(new URL('..',import.meta.url));
const fail=(status,message)=>Object.assign(Error(message),{status});
const now=()=>new Date().toISOString();
export const formatFileLimit=30*1024*1024;
export const docxMime='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export function formatConfig(env=process.env){
 const checkerRoot=path.resolve(env.FORMAT_CHECKER_ROOT||path.join(root,'..','thesis-format-checker','thesis-format-checker'));
 const python=env.FORMAT_CHECKER_PYTHON||path.join(checkerRoot,'.venv-run','Scripts','python.exe');
 const seconds=Number(env.FORMAT_CHECKER_TIMEOUT_SECONDS||1800);
 let reason='';
 if(process.platform!=='win32')reason='当前检测器需要 Windows 和桌面版 Microsoft Word。';
 else if(!existsSync(path.join(checkerRoot,'src','thesis_format_checker','cli.py')))reason='未找到格式检测程序，请检查服务端配置。';
 else if(!existsSync(python))reason='未找到格式检测所需的 Python 环境，请检查服务端配置。';
 else if(!Number.isSafeInteger(seconds)||seconds<1)reason='格式检测超时配置无效。';
 return {checkerRoot,python,timeoutMs:seconds*1000,configured:!reason,reason};
}
export function summarizeFormatReport(report){
 if(!report||typeof report!=='object'||!Array.isArray(report.findings)||!Array.isArray(report.unknowns)||!Array.isArray(report.report_only)||!Array.isArray(report.skipped))throw Error('检测报告结构无效');
 const counts=report.counts_by_severity||{};
 const number=value=>Number.isFinite(Number(value))?Math.max(0,Number(value)):0;
 return {issues:Object.values(counts).reduce((total,value)=>total+number(value),0),errors:number(counts['错误']),warnings:number(counts['警告']),hints:number(counts['提醒']),unknown:report.unknowns.length,report_only:report.report_only.length,skipped:report.skipped.length,seconds:number(report.timings?.['总耗时']),limited:report.unknowns.length>0||report.report_only.length>0||report.skipped.length>0};
}
export class FormatRepository{
 constructor(repo){this.repo=repo;}
 get(id,owner){return this.repo.one('SELECT * FROM format_tasks WHERE id=? AND owner_id=?',[id,owner]);}
 latest(owner){return this.repo.one('SELECT * FROM format_tasks WHERE owner_id=? ORDER BY created_at DESC,id DESC LIMIT 1',[owner]);}
 insert(task){return this.repo.transaction(async tx=>{
  await tx.one('SELECT id FROM users WHERE id=? FOR UPDATE',[task.owner_id]);
  if(await tx.one("SELECT id FROM format_tasks WHERE owner_id=? AND status IN ('queued','running') LIMIT 1",[task.owner_id]))throw fail(409,'已有格式检测任务正在处理，请等待完成。');
  await tx.insert('format_tasks',task);
 });}
 update(id,values){const keys=Object.keys(values);return this.repo.q(`UPDATE format_tasks SET ${keys.map(key=>key+'=?').join(',')} WHERE id=?`,[...keys.map(key=>values[key]),id]);}
 next(){return this.repo.one("SELECT * FROM format_tasks WHERE status='queued' ORDER BY created_at,id LIMIT 1");}
 interrupt(){return this.repo.q("UPDATE format_tasks SET status='interrupted',error=?,message=?,finished_at=? WHERE status='running'",['检测进程已中断，可点击重新检测。','检测已中断',now()]);}
 async acquire(){
  const connection=await this.repo.pool.getConnection();
  try{
   const [[row]]=await connection.query("SELECT GET_LOCK(CONCAT('wenheng-format-',DATABASE()),0) acquired");
   if(!row.acquired){connection.release();return null;}
   return async()=>{try{await connection.query("SELECT RELEASE_LOCK(CONCAT('wenheng-format-',DATABASE()))");}finally{connection.release();}};
  }catch(error){connection.release();throw error;}
 }
}
export async function runFormatWorker(task,config,{signal,onMessage}){
 const directory=path.dirname(task.output_path),log=await fs.open(path.join(directory,'worker.log'),'w');
 const args=['-B','-X','utf8','-u',path.join(root,'scripts','format-worker.py'),'--checker-root',config.checkerRoot,'--template',task.template_path,'--document',task.document_path,'--output',task.output_path,'--report-json',task.report_path,'--cache-dir',path.join(directory,'..','cache',task.owner_id)];
 try{
  await new Promise((resolve,reject)=>{
   const child=spawn(config.python,args,{cwd:config.checkerRoot,windowsHide:true,stdio:['ignore','pipe','pipe'],env:{...process.env,PYTHONDONTWRITEBYTECODE:'1',PYTHONUTF8:'1'}});
   let pending='',logged=0,stopped=false,reason=null;
   const stop=message=>{reason=message;child.kill('SIGKILL');};
   const abort=()=>stop('检测进程已中断，可重新检测。');
   const timer=setTimeout(()=>stop('格式检测超时，请检查文档后重新检测。'),config.timeoutMs);
   signal?.addEventListener('abort',abort,{once:true});
   if(signal?.aborted)abort();
   const clean=()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort);};
   const record=text=>{if(logged>=1024*1024)return;const bytes=Buffer.from(text).subarray(0,1024*1024-logged);logged+=bytes.length;log.write(bytes).catch(()=>{});};
   child.stdout.setEncoding('utf8');child.stderr.setEncoding('utf8');
   child.stdout.on('data',text=>{
    record(text);pending+=text;
    let boundary;while((boundary=pending.indexOf('\n'))>=0){const line=pending.slice(0,boundary);pending=pending.slice(boundary+1);try{const event=JSON.parse(line);if(event.type==='stage'&&typeof event.message==='string')onMessage(event.message.slice(0,200));if(event.type==='error'&&typeof event.message==='string')reason=event.message.slice(0,500);}catch{}}
    if(pending.length>65536)pending='';
   });
   child.stderr.on('data',record);
   child.on('error',()=>{if(stopped)return;stopped=true;clean();reject(Error('无法启动格式检测程序，请检查 Python 环境和服务端配置。'));});
   child.on('close',code=>{if(stopped)return;stopped=true;clean();if(reason||code!==0)reject(Error(reason||'格式检测程序执行失败，请检查文档是否损坏或受密码保护。'));else resolve();});
  });
 }finally{await log.close();}
}
export function createFormatService(repo,dataDir,{store=new FormatRepository(repo),config=formatConfig(),runner=runFormatWorker}={}){
 const base=path.resolve(dataDir,'format');let pumping=null,timer=null,stopping=false,controller=null;
 function configuration(){return {configured:config.configured,message:config.reason,max_file_bytes:formatFileLimit};}
 function publicTask(task){
  if(!task)return null;
  return {task_id:task.id,status:task.status,template:{name:task.template_name,size:task.template_size},paper:{name:task.document_name,size:task.document_size},output_name:task.output_name,message:task.message,error:task.error,created_at:task.created_at,started_at:task.started_at,finished_at:task.finished_at,summary:task.summary_json?JSON.parse(task.summary_json):null};
 }
 async function owned(id,owner){const task=await store.get(id,owner);if(!task)throw fail(404,'格式检测任务不存在。');return task;}
 async function submit(owner,form){
  if(!config.configured)throw fail(503,config.reason);
  const sourceId=form.get('source_task_id');
  if(sourceId!==null&&typeof sourceId!=='string')throw fail(400,'无效的格式检测任务。');
  const source=sourceId?await owned(sourceId,owner):null,inputs={};
  for(const [kind,label] of [['template','模板论文'],['document','待检测论文']]){
   const file=form.get(kind);
   if(file&&typeof file.arrayBuffer==='function'){
    if(form.getAll(kind).length!==1||!file.size||file.size>formatFileLimit||!/\.docx$/i.test(file.name))throw fail(400,label+'请选择一份不超过 30 MB 的 .docx 文档。');
    const bytes=Buffer.from(await file.arrayBuffer());
    if(bytes.length<4||bytes.readUInt32LE(0)!==0x04034b50)throw fail(400,label+'不是有效的 DOCX 文档。');
    const filename=path.basename(file.name.replace(/\\/g,'/')).replace(/[\u0000-\u001f\u007f]/g,'').slice(0,240);
    inputs[kind]={name:filename,size:bytes.length,bytes};
   }else if(source&&!file){inputs[kind]={name:source[kind+'_name'],size:source[kind+'_size'],path:source[kind+'_path']};}
   else throw fail(400,'请选择'+label+'。');
  }
  const id=crypto.randomUUID(),directory=path.join(base,id);
  await fs.mkdir(directory,{recursive:true});
  try{
   for(const [kind,input] of Object.entries(inputs)){const target=path.join(directory,kind+'.docx');if(input.bytes)await fs.writeFile(target,input.bytes,{flag:'wx'});else await fs.copyFile(input.path,target);}
   const task={id,owner_id:owner,status:'queued',template_name:inputs.template.name,document_name:inputs.document.name,template_size:inputs.template.size,document_size:inputs.document.size,template_path:path.join(directory,'template.docx'),document_path:path.join(directory,'document.docx'),output_path:path.join(directory,'annotated.docx'),report_path:path.join(directory,'report.json'),output_name:inputs.document.name.replace(/\.docx$/i,'')+'_带批注.docx',message:'已提交，等待格式检测',error:null,summary_json:null,created_at:now(),started_at:null,finished_at:null};
   await store.insert(task);kick();return publicTask(task);
  }catch(error){if(path.dirname(path.resolve(directory))===base&&path.basename(directory)===id)await fs.rm(directory,{recursive:true,force:true});if(error.code==='ENOENT')throw fail(404,'原任务的文档已不存在，请重新选择两份文件。');throw error;}
 }
 async function execute(task){
  controller=new AbortController();
  await store.update(task.id,{status:'running',started_at:now(),message:'正在检查文档并启动格式检测',error:null});
  let updates=Promise.resolve();
  try{
   if(!config.configured)throw Error(config.reason);
   await runner(task,config,{signal:controller.signal,onMessage:message=>{updates=updates.then(()=>store.update(task.id,{message})).catch(()=>{});}});
   await updates;
   const output=await fs.stat(task.output_path);if(!output.size)throw Error('检测未生成有效的批注文档。');
   const report=JSON.parse(await fs.readFile(task.report_path,'utf8')),summary=summarizeFormatReport(report);
   await store.update(task.id,{status:'succeeded',message:summary.limited?'检测完成，部分项目需人工核查':'检测完成',summary_json:JSON.stringify(summary),finished_at:now()});
  }catch(error){await updates;await store.update(task.id,{status:stopping?'interrupted':'failed',error:error.message?.slice(0,500)||'格式检测失败。',message:stopping?'检测已中断':'检测失败',finished_at:now()});}
  finally{controller=null;}
 }
 async function drain(){
  const release=await store.acquire();if(!release)return;
  try{await store.interrupt();let task;while(!stopping&&(task=await store.next()))await execute(task);}finally{await release();}
 }
 function kick(){if(!stopping&&!pumping)pumping=drain().catch(()=>console.error('format queue: 数据库或检测队列暂时不可用')).finally(()=>{pumping=null;});return pumping;}
 return {
  configuration,submit,
  get:async(id,owner)=>publicTask(await owned(id,owner)),
  latest:async owner=>publicTask(await store.latest(owner)),
  async download(id,owner){const task=await owned(id,owner);if(task.status!=='succeeded')throw fail(409,'检测尚未成功完成，暂时无法下载。');try{await fs.access(task.output_path);}catch{throw fail(404,'检测结果文件已不存在，请重新检测。');}return {path:task.output_path,name:task.output_name};},
  start(){if(!timer){kick();timer=setInterval(kick,5000);timer.unref();}},
  async idle(){while(pumping)await pumping;},
  async shutdown(){stopping=true;clearInterval(timer);controller?.abort();if(pumping)await pumping;}
 };
}
