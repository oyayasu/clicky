import assert from 'node:assert/strict';
import {test} from 'node:test';
import fs from 'node:fs';
import {createDocx} from '../src/docx.js';
import {readZip} from '../src/docx-import.js';

test('App entrypoint and editor controls exist',()=>{
 const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
 for(const id of ['editor','tabList','newTab','highlightToggle','autoHighlight','fontPickerBtn','themeBtn','imageFile','assetLock','assetSnap','exportTop'])assert.match(html,new RegExp(`id="${id}"`));
});

test('DOCX creates a ZIP signature',async()=>{
 const blob=createDocx({tabs:[]});
 const bytes=new Uint8Array(await blob.arrayBuffer());
 assert.equal(String.fromCharCode(...bytes.subarray(0,4)),'PK\x03\x04');
 const text=new TextDecoder().decode(bytes);
 for(const filename of ['[Content_Types].xml','word/document.xml','word/styles.xml'])assert.ok(text.includes(filename));
});

test('Tauri contains Windows/macOS release prerequisites',()=>{
 const tauri=JSON.parse(fs.readFileSync(new URL('../src-tauri/tauri.conf.json',import.meta.url),'utf8'));
 assert.equal(tauri.productName,'Clicky');
 assert.equal(tauri.build.frontendDist,'../dist');
 assert.equal(tauri.bundle.active,true);
 const pipeline=fs.readFileSync(new URL('../.github/workflows/build.yml',import.meta.url),'utf8');
 assert.match(pipeline,/windows-latest/);assert.match(pipeline,/macos-latest/);
});

test('Library, comment threads, proofreading and shortcuts expose interactive controls',()=>{
 const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
 for(const id of ['libraryView','documentList','folderList','libraryGrid','libraryList','librarySearch','libraryNewFolder','libraryNewDocument','libraryMore','workspaceImport','commentsPanel','commentComposer','commentButton','commentDraft','saveComment','commentsList','commentsButton'])assert.match(html,new RegExp(`id="${id}"`));
 const app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
 for(const functionName of ['renderLibrary','openDocument','beginComment','commitComment','renderComments','openProofing','openDictionary','handleKeydown','exportLibraryBackup','importLibraryBackup'])assert.match(app,new RegExp(`function ${functionName}\\(`));
});

test('Storage has independent documents, folders and backup metadata',()=>{
 const storage=fs.readFileSync(new URL('../src/storage.js',import.meta.url),'utf8');
 for(const store of ['document_index','documents','folders','revisions','preferences'])assert.match(storage,new RegExp(`['"]${store}['"]`));
 assert.match(storage,/initializeLibrary/);assert.match(storage,/loadDocument/);assert.match(storage,/deleteFolder/);
 const conf=JSON.parse(fs.readFileSync(new URL('../src-tauri/tauri.conf.json',import.meta.url),'utf8'));
 assert.equal(conf.version,'0.5.0');assert.equal(conf.app.security.csp['script-src'],"'self'");
});


test('DOCX import reads actual Word XML from locally generated archive',async()=>{
 const word=createDocx({title:'Verificação',tabs:[]});
 const archive=await readZip(word);
 const xml=new TextDecoder().decode(await archive.get('word/document.xml'));
 assert.match(xml,/<w:document/);
 assert.match(xml,/<w:body>/);
});

test('New interface has no C. previews and the icon has alpha',()=>{
 const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
 const app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
 assert.match(html,/id="libraryOpenDocx"/);
 assert.match(html,/id="searchMenus"/);
 assert.doesNotMatch(app,/doc-art-mark.*c\./);
 assert.match(app,/function lockSvg/);
 assert.match(app,/function drawSnapGuides/);
 const png=fs.readFileSync(new URL('../src-tauri/icons/clicky-master.png',import.meta.url));
 assert.equal(png.subarray(1,4).toString(),'PNG');
});


test('New documents start with only Guia 1 and contextual hint',()=>{
 const app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
 assert.match(app,/activeId:'guia-1',tabs:\[\{id:'guia-1',name:'Guia 1',html:'<p><\/p>',assets:\[\]\}\]/);
 assert.doesNotMatch(app,/name:'Original',html:/);
 assert.match(app,/if\(state\.tabs\.length===1\)\{const hint=document\.createElement\('p'\)/);
 assert.match(app,/Os títulos que forem adicionados ao documento aparecerão aqui\./);
 assert.match(app,/state\.tabs\.length\+1/);
});


test('Reference geometry uses centered page and independent tabs sidebar',()=>{
 const css=fs.readFileSync(new URL('../src/style.css',import.meta.url),'utf8');
 const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
 for(const field of ['--accent:#c02950','.page-stage{min-width:684px','.page-ruler{width:684px;left:calc(50% - 16px)','.paper{width:684px;left:-16px','.editor{padding:82px 82px 82px','.sidebar{width:263px!important'])assert.ok(css.includes(field),field);
 for(const id of ['zoomSelect','favoriteButton','cloudSaveBtn','moveToFolderBtn','revisionsTop','sidebarBack','proofTop','paintFormat','winMin','winMax','winClose'])assert.match(html,new RegExp(`id="${id}"`));
 const app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
 for(const name of ['initNativeWindowButtons','setPageZoom','beginFormatPainting','applyPaintedFormat'])assert.ok(app.includes(`function ${name}(`),name);
});


test('0.5 includes no browser dialogs, contextual navigation and local language control',()=>{
 const app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
 const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8');
 for(const marker of ['function appDialog','function selectionIsBackward','function scheduleOutline','function applyInterfaceLanguage','function renderTabs','function askLink'])assert.ok(app.includes(marker),marker);
 assert.doesNotMatch(app,/\b(?:window\.)?(?:alert|prompt|confirm)\s*\(/);
 for(const id of ['homeWinMin','homeWinMax','homeWinClose','floatLookup','dialogShade','dialogInput'])assert.match(html,new RegExp(`id="${id}"`));
 assert.doesNotMatch(html,/id="themeImport"|id="themePicture"/);
});