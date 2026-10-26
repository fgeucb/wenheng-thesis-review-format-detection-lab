import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import {execFile,spawn} from 'node:child_process';
import {once} from 'node:events';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';
import {formatConfig} from '../src/format-checker.mjs';
import {testDatabaseConfig,testDatabaseEnv} from './helpers.mjs';
import {dropDatabase} from '../src/database.mjs';

const root=fileURLToPath(new URL('..',import.meta.url)),execute=promisify(execFile);
const digest=bytes=>crypto.createHash('sha256').update(bytes).digest('hex');
const generate=String.raw`
import sys
from pathlib import Path
from docx import Document
from docx.shared import Mm, Pt
from docx.oxml.ns import qn
folder = Path(sys.argv[1])
for name, template in [("template.docx", True), ("document.docx", False)]:
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = Mm(25)
    section.left_margin = Mm(30 if template else 20)
    section.right_margin = Mm(30)
    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    document.add_paragraph("论文格式检测功能验收样例")
    document.add_paragraph("摘要")
    document.add_paragraph("本文件为自动生成的虚构测试材料，用于验证格式检测和批注下载。")
    document.add_paragraph("关键词：格式检测；测试文档")
    document.add_heading("第一章 绪论", level=1)
    for index in range(8):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.first_line_indent = Pt(24 if template else 0)
        paragraph.paragraph_format.line_spacing = 1.5 if template else 1.0
        paragraph.paragraph_format.space_after = Pt(6 if template else 0)
        paragraph._p.get_or_add_pPr().get_or_add_ind().set(qn("w:firstLineChars"), "200" if template else "0")
        run = paragraph.add_run(f"这是第{index + 1}段用于格式检测的虚构正文，包含中文内容和 English research 123456。")
        run.font.name = "Times New Roman" if template else "Arial"
        run.font.size = Pt(12 if template else 10)
        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体" if template else "微软雅黑")
    document.save(folder / name)
`;
const verify=String.raw`
import json, sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree
from docx import Document
source, output = map(Path, sys.argv[1:])
assert [paragraph.text for paragraph in Document(source).paragraphs] == [paragraph.text for paragraph in Document(output).paragraphs]
with ZipFile(source) as before, ZipFile(output) as after:
    comments = ElementTree.fromstring(after.read("word/comments.xml"))
    assert len(comments) > 0
    for name in before.namelist():
        if name.startswith(("word/styles", "word/numbering", "word/theme/", "word/media/")):
            assert before.read(name) == after.read(name), name
    print(json.dumps({"comments": len(comments), "body_text_preserved": True, "styles_and_media_preserved": True}))
`;

test('format HTTP/browser workflow uses real Word and preserves document content',{
 skip:process.platform!=='win32'||process.env.RUN_WORD_INTEGRATION!=='1',timeout:300000
},async()=>{
 const config=formatConfig();assert(config.configured,config.reason);
 const directory=await fs.mkdtemp(path.join(os.tmpdir(),'wenheng-format-http-'));
 const output=path.join(root,'test-results','format-integration');await fs.mkdir(output,{recursive:true});
 const database=testDatabaseConfig();let server,browser,base,cookie,task,diagnostics='';
 const hashes=await Promise.all(['public/index.html','public/css/theme.css'].map(async file=>digest(await fs.readFile(path.join(root,file)))));
 try{
  await execute(config.python,['-B','-X','utf8','-c',generate,output],{windowsHide:true});
  const inputNames=['template.docx','document.docx'],inputHashes=await Promise.all(inputNames.map(async name=>digest(await fs.readFile(path.join(output,name)))));
  server=spawn(process.execPath,['src/server.mjs'],{cwd:root,windowsHide:true,stdio:['ignore','pipe','pipe'],env:{...process.env,...testDatabaseEnv(database),REVIEW_SKIP_ENV:'1',REVIEW_HOST:'127.0.0.1',REVIEW_PORT:'0',REVIEW_DATA_DIR:directory,FORMAT_CHECKER_ROOT:config.checkerRoot,FORMAT_CHECKER_PYTHON:config.python,FORMAT_CHECKER_TIMEOUT_SECONDS:'180',MINERU_TOKEN:'',CODEX_API_KEY:'',OPENAI_API_KEY:'',DEEPSEEK_API_KEY:''}});
  server.stderr.on('data',chunk=>{diagnostics+=chunk;});
  base=await new Promise((resolve,reject)=>{
   const timeout=setTimeout(()=>reject(Error('格式检测测试服务启动超时：'+diagnostics)),15000);let text='';
   server.stdout.on('data',chunk=>{text+=chunk;const match=text.match(/http:\/\/127\.0\.0\.1:\d+/);if(match){clearTimeout(timeout);resolve(match[0]);}});
   server.once('error',error=>{clearTimeout(timeout);reject(error);});server.once('exit',()=>{clearTimeout(timeout);reject(Error(diagnostics));});
  });
  const registration=await fetch(base+'/api/auth/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'format_acceptance',password:crypto.randomUUID()})});
  assert.equal(registration.status,200);cookie=registration.headers.get('set-cookie').split(';')[0];
  const request=(url,options={})=>fetch(base+url,{...options,headers:{Cookie:cookie,...options.headers}});
  assert.equal((await fetch(base+'/api/format/config')).status,401);
  assert.equal((await (await request('/api/format/config')).json()).configured,true);
  browser=await chromium.launch({headless:true});const context=await browser.newContext({viewport:{width:1907,height:1000}});
  await context.addCookies([{name:'review_session',value:cookie.split('=')[1],url:base}]);
  const page=await context.newPage(),errors=[];page.on('pageerror',error=>errors.push(error.message));
  await page.goto(base+'/#format');await page.waitForFunction(()=>document.getElementById('format-hint').textContent.includes('选择两份'));
  await page.screenshot({path:path.join(output,'format-before.png'),fullPage:true});
  await page.locator('#format-template-file').setInputFiles(path.join(output,'template.docx'));
  await page.locator('#format-paper-file').setInputFiles(path.join(output,'document.docx'));
  await page.locator('#format-start').click();
  await page.waitForFunction(()=>['排队中','检测中','部分项目需核查','检测完成','检测失败'].includes(document.getElementById('format-result-status').textContent));
  await page.reload();await page.waitForFunction(()=>document.getElementById('format-template-name').textContent==='template.docx');
  const deadline=Date.now()+210000;
  do{
   task=(await (await request('/api/format/tasks/latest')).json()).task;
   if(task&&['succeeded','failed','interrupted'].includes(task.status))break;
   await new Promise(resolve=>setTimeout(resolve,1000));
  }while(Date.now()<deadline);
  if(task){for(const name of ['worker.log','report.json']){try{await fs.copyFile(path.join(directory,'format',task.task_id,name),path.join(output,name));}catch(error){if(error.code!=='ENOENT')throw error;}}}
  assert.equal(task?.status,'succeeded',JSON.stringify(task)+'\n'+diagnostics);assert(task.summary.issues>0);
  const report=JSON.parse(await fs.readFile(path.join(output,'report.json'),'utf8'));assert(report.timings['Word段落属性与分页']>0);
  const workerLog=await fs.readFile(path.join(output,'worker.log'),'utf8');assert.match(workerLog,/正在读取 Word 格式与分页信息/);
  await page.waitForFunction(()=>!document.getElementById('format-download').disabled);
  await page.screenshot({path:path.join(output,'format-complete.png'),fullPage:true});
  const pendingDownload=page.waitForEvent('download');await page.locator('#format-download').click();const download=await pendingDownload;
  assert.equal(download.suggestedFilename(),'document_带批注.docx');await download.saveAs(path.join(output,'annotated.docx'));
  const verified=JSON.parse((await execute(config.python,['-B','-X','utf8','-c',verify,path.join(output,'document.docx'),path.join(output,'annotated.docx')],{windowsHide:true})).stdout);
  assert.deepEqual(await Promise.all(inputNames.map(async name=>digest(await fs.readFile(path.join(output,name))))),inputHashes);
  await page.reload();await page.waitForFunction(()=>!document.getElementById('format-download').disabled);
  const other=await fetch(base+'/api/auth/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'format_other',password:crypto.randomUUID()})});
  assert.equal(other.status,200);const otherCookie=other.headers.get('set-cookie').split(';')[0];
  for(const suffix of ['', '/download'])assert.equal((await fetch(base+'/api/format/tasks/'+task.task_id+suffix,{headers:{Cookie:otherCookie}})).status,404);
  const invalid=new FormData();invalid.set('template',new Blob([Buffer.from([0x50,0x4b,0x03,0x04])]),'broken.docx');invalid.set('document',new Blob([await fs.readFile(path.join(output,'document.docx'))]),'document.docx');
  const submitted=await request('/api/format/tasks',{method:'POST',body:invalid});assert.equal(submitted.status,202);const broken=(await submitted.json()).task;
  let failed;
  for(let attempt=0;attempt<30;attempt++){failed=(await (await request('/api/format/tasks/'+broken.task_id)).json()).task;if(failed.status==='failed')break;await new Promise(resolve=>setTimeout(resolve,500));}
  assert.equal(failed.status,'failed');assert.match(failed.error,/DOCX|文档/);assert.equal((await request('/api/format/tasks/'+broken.task_id+'/download')).status,409);
  assert.deepEqual(errors,[]);
  assert.deepEqual(await Promise.all(['public/index.html','public/css/theme.css'].map(async file=>digest(await fs.readFile(path.join(root,file))))),hashes);
  await fs.writeFile(path.join(output,'verification.json'),JSON.stringify({task_id:task.task_id,status:task.status,summary:task.summary,...verified,source_files_unchanged:true,html_css_unchanged:true,browser_errors:errors,checks:['browser upload','real Word detection','reload during task','native comments download','reload completed task','account isolation','invalid DOCX rejection']},null,2));
 }finally{
  if(browser)await browser.close();
  if(server&&server.exitCode===null&&server.signalCode===null){server.kill('SIGINT');await once(server,'exit');}
  await dropDatabase(database);
  assert(path.dirname(path.resolve(directory))===path.resolve(os.tmpdir())&&path.basename(directory).startsWith('wenheng-format-http-'));await fs.rm(directory,{recursive:true,force:true});
 }
});
