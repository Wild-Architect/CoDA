const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const projects = require('../desktop/edessb-projects');
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'coda-edessb-'));
  t.after(async () => {
    assert.ok(path.resolve(root).startsWith(path.resolve(os.tmpdir()) + path.sep));
    await fs.rm(root, {recursive:true, force:true});
  });
  const pdf = path.join(root, 'source.pdf');
  await fs.writeFile(pdf, '%PDF-1.4 test content');
  return {root, pdf};
}
const payload = (identifier, title = 'Освіти9-кориг2025') => ({identifier, title, organization:'Організація'});
test('projects sort by descending numeric identifiers including large values', async t => {
  const {root} = await fixture(t);
  for (const id of ['1','10','2','99999999999999999999','100000000000000000000']) await projects.createProject(root, payload(id));
  assert.deepEqual((await projects.listProjects(root)).map(p=>p.identifier), ['100000000000000000000','99999999999999999999','10','2','1']);
});
test('rename preserves PDFs, links and revision operations', async t => {
  const {root,pdf} = await fixture(t);
  const p = await projects.createProject(root,payload('1'));
  await projects.addRevision(root,p.id,{type:'design-task',sourcePath:pdf,url:'https://e-construction.gov.ua/test'});
  const oldFolder = await projects.getDocumentsFolder(root,'1');
  assert.equal(path.basename(oldFolder),'1_Освіти9-кориг2025');
  await projects.renameProject(root,p.id,'Нова назва');
  await assert.rejects(fs.stat(oldFolder),{code:'ENOENT'});
  const revision = await projects.getRevision(root,p.id,'design-task',1);
  assert.equal(path.basename(path.dirname(revision.filePath)),'1_Нова назва');
  assert.equal(await fs.readFile(revision.filePath,'utf8'),'%PDF-1.4 test content');
  assert.equal(revision.revision.url,'https://e-construction.gov.ua/test');
  await projects.updateRevision(root,p.id,{type:'design-task',number:1,sourcePath:pdf});
  await projects.addRevision(root,p.id,{type:'design-task',sourcePath:pdf});
  await projects.deleteRevision(root,p.id,'design-task',1);
  assert.ok((await projects.getRevision(root,p.id,'design-task',2)).filePath);
  const folder=await projects.getDocumentsFolder(root,'1');
  await projects.deleteProject(root,p.id);
  await assert.rejects(fs.stat(folder),{code:'ENOENT'});
});
test('legacy numeric folders migrate without losing extra files and remain stable', async t => {
  const {root,pdf}=await fixture(t);
  const p=await projects.createProject(root,payload('2'));
  await projects.addRevision(root,p.id,{type:'design-task',sourcePath:pdf});
  const folder=await projects.getDocumentsFolder(root,'2');
  await fs.writeFile(path.join(folder,'extra.txt'),'keep');
  const legacy=path.join(root,'edessb_documents','2');
  await fs.rename(folder,legacy);
  for(let i=0;i<2;i++) assert.equal((await projects.listProjects(root))[0].documents[0].revisions[0].fileExists,true);
  assert.equal(await fs.readFile(path.join(folder,'extra.txt'),'utf8'),'keep');
  await assert.rejects(fs.stat(legacy),{code:'ENOENT'});
});
test('invalid title and occupied target preserve original project and documents', async t => {
  const {root,pdf}=await fixture(t);
  const p=await projects.createProject(root,payload('3'));
  await projects.addRevision(root,p.id,{type:'design-task',sourcePath:pdf});
  await assert.rejects(projects.renameProject(root,p.id,'  '));
  const target=path.join(root,'edessb_documents','3_Зайнято');
  await fs.mkdir(target);await fs.writeFile(path.join(target,'other.txt'),'untouched');
  await assert.rejects(projects.renameProject(root,p.id,'Зайнято'));
  assert.equal((await projects.listProjects(root))[0].title,p.title);
  assert.ok((await projects.getRevision(root,p.id,'design-task',1)).filePath);
  assert.equal(await fs.readFile(path.join(target,'other.txt'),'utf8'),'untouched');
});
test('Windows-invalid characters are sanitized only in folder names', async t => {
  const {root}=await fixture(t);
  const p=await projects.createProject(root,payload('4'));
  const title='Школа / корпус: А?';
  await projects.renameProject(root,p.id,title);
  assert.equal(path.basename(await projects.getDocumentsFolder(root,'4')),'4_Школа _ корпус_ А_');
  assert.equal((await projects.listProjects(root))[0].title,title);
  await projects.renameProject(root,p.id,'школа / корпус: а?');
  assert.equal(path.basename(await projects.getDocumentsFolder(root,'4')),'4_школа _ корпус_ а_');
});

test('metadata save failure rolls folder back and keeps original metadata', async t => {
  const {root,pdf}=await fixture(t);
  const p=await projects.createProject(root,payload('5'));
  await projects.addRevision(root,p.id,{type:'design-task',sourcePath:pdf});
  const rename=fs.rename;
  fs.rename=async (source,target)=>{
    if(path.basename(target)==='project.json') throw new Error('simulated metadata failure');
    return rename(source,target);
  };
  try { await assert.rejects(projects.renameProject(root,p.id,'Нова назва'),/simulated metadata failure/); }
  finally { fs.rename=rename; }
  assert.equal((await projects.listProjects(root))[0].title,p.title);
  assert.equal(await fs.readFile((await projects.getRevision(root,p.id,'design-task',1)).filePath,'utf8'),'%PDF-1.4 test content');
  await assert.rejects(fs.stat(path.join(root,'edessb_documents','5_Нова назва')),{code:'ENOENT'});
});
