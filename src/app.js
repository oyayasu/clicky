import {initializeLibrary,listLibrary,loadDocument,save,saveFolder,deleteFolder,removeDocument,revisions,cleanupRevisions,getPreference,setPreference} from './storage.js';
import {createDocx} from './docx.js';
import {readDocx} from './docx-import.js';
const $=id=>document.getElementById(id);
const uid=()=>globalThis.crypto?.randomUUID?.()||'id'+Date.now().toString(36)+Math.random().toString(36).slice(2);
const clamp=(v,min,max)=>Math.max(min,Math.min(max,v));
const fonts=['Cambria','Georgia','Times New Roman','Arial','Helvetica','Verdana','Inter','Roboto','Open Sans','Google Sans','Montserrat','Poppins','Lato','Lora','Lexend','Merriweather','Garamond','Palatino Linotype','Baskerville','Book Antiqua','Trebuchet MS','Tahoma','Courier New','Consolas','Comic Sans MS'];
const highlightColors=['#ffdb75','#fcb6c8','#ffb4a5','#b3e7b0','#a8daf2','#cdbaff','#dadde1','#fff2a3','#f7c889','#e6d0b3','#ffc9ed','#c7e5db'];
const presets=[{name:'Papel',accent:'#c02950',surface:'#f7f9fc',panel:'#f8fafd',text:'#202124',paper:'#ffffff'},{name:'Grafite',accent:'#b787db',surface:'#22252d',panel:'#2d313b',text:'#f1f2f7',paper:'#fffdfa'},{name:'Rosa chá',accent:'#c44875',surface:'#f7eef1',panel:'#fff8fa',text:'#513540',paper:'#fffdfb'},{name:'Oliva',accent:'#708e69',surface:'#edf0e7',panel:'#f9fbf4',text:'#364237',paper:'#fffef8'},{name:'Areia',accent:'#b88059',surface:'#f3eee6',panel:'#fffaf2',text:'#554638',paper:'#fffdfa'},{name:'Azul',accent:'#4778b1',surface:'#edf2fa',panel:'#f9fbff',text:'#293a56',paper:'#ffffff'}];
const defaultState=()=>({id:uid(),folderId:null,language:'pt-BR',spellcheckEnabled:true,version:2,title:'Documento sem título',activeId:'guia-1',tabs:[{id:'guia-1',name:'Guia 1',html:'<p></p>',assets:[]}],font:'Arial',fontSize:11,fonts:[],highlight:'#ffdb75',autoHighlight:true,snap:true,ruler:true,theme:{...presets[0],image:null}});
let state=defaultState(),currentAsset=null,activeSelection=null,saveTimer=null,notificationTimer=null,dragState=null,saving=false,dirty=false,lastSnapshot=0,appReady=false;
let library={documents:[],folders:[]},libraryFolder='all',libraryLayout='grid',dictionary=[],pendingCommentRange=null,pendingCommentQuote='';
const currentTab=()=>state.tabs.find(t=>t.id===state.activeId)||state.tabs[0];
function lockSvg(locked){return locked?'<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></svg>':'<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0"/></svg>'}
function calculateSnap(pos,others=[]){
 const guides=[];let {x,y,width,height}=pos;const threshold=23,axis=(p,targets)=>{let best=null;for(const t of targets){const d=Math.abs(p-t);if(d<=threshold&&(!best||d<best.d))best={d,value:t}}return best};
 const px=$('paper')?.offsetWidth||816,py=$('paper')?.offsetHeight||1056;
 const xTargets=[0,px/2,px,90,px-90,...others.flatMap(a=>[a.x,a.x+a.width/2,a.x+a.width])];
 const yTargets=[0,py/2,py,96,py-96,...others.flatMap(a=>[a.y,a.y+a.height/2,a.y+a.height])];
 const xSnap=[{from:x,offset:0},{from:x+width/2,offset:width/2},{from:x+width,offset:width}].map(q=>({...axis(q.from,xTargets),offset:q.offset})).filter(v=>Number.isFinite(v.d)).sort((a,b)=>a.d-b.d)[0];
 const ySnap=[{from:y,offset:0},{from:y+height/2,offset:height/2},{from:y+height,offset:height}].map(q=>({...axis(q.from,yTargets),offset:q.offset})).filter(v=>Number.isFinite(v.d)).sort((a,b)=>a.d-b.d)[0];
 if(xSnap){x=xSnap.value-xSnap.offset;guides.push({type:'vertical',position:xSnap.value})}
 if(ySnap){y=ySnap.value-ySnap.offset;guides.push({type:'horizontal',position:ySnap.value})}
 // Diagonal magnets cross the page center and origin; used when no Y alignment is active.
 if(!ySnap){for(const origin of [{x:px/2,y:py/2},{x:90,y:96},...others.map(a=>({x:a.x+a.width/2,y:a.y+a.height/2}))]){
  const centerX=x+width/2,centerY=y+height/2;
  for(const slope of [-1,1]){const lineY=origin.y+slope*(centerX-origin.x);if(Math.abs(centerY-lineY)<threshold){y+=lineY-centerY;guides.push({type:slope===1?'diag-up':'diag-down',origin});break}}if(guides.some(g=>g.type.startsWith('diag')))break;
 }}
 return {x:Math.round(x),y:Math.round(y),guides};
}
function drawSnapGuides(guides,asset){
 let layer=$('snapGuides');if(!layer){layer=document.createElement('div');layer.id='snapGuides';layer.className='snap-guides';$('paper').append(layer)}layer.innerHTML='';
 const centerX=asset.x+asset.width/2,centerY=asset.y+asset.height/2;
 function add(type,position,regular=false){const div=document.createElement('div');div.className='snap-line '+type+(regular?' regular-guide':'');if(type==='vertical')div.style.left=position+'px';else if(type==='horizontal')div.style.top=position+'px';else{div.style.left=(centerX-750)+'px';div.style.top=centerY+'px'}layer.append(div)}
 // Horizontal and both diagonals stay visible throughout the drag, similar to a design canvas.
 add('horizontal',centerY,true);add('diag-up',null,true);add('diag-down',null,true);
 for(const guide of guides)add(guide.type,guide.position||centerY);
}
function clearSnapGuides(){const layer=$('snapGuides');if(layer)layer.innerHTML=''}

function notify(message){const t=$('toast');t.textContent=message;t.classList.remove('hidden');clearTimeout(notificationTimer);notificationTimer=setTimeout(()=>t.classList.add('hidden'),2900)}
function setSaveState(message){$('saveStatus').innerHTML=`<span class="mini-dot"></span> ${message}`}
function escapeHTML(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function writeEditor(){currentTab().html=$('editor').innerHTML}
function scheduleSave(){if(!appReady)return;dirty=true;setSaveState('Salvando...');clearTimeout(saveTimer);saveTimer=setTimeout(()=>void persist(),550)}
async function persist(force=false){if(!appReady||(!dirty&&!force))return;if(saving){scheduleSave();return;}clearTimeout(saveTimer);writeEditor();saving=true;dirty=false;try{const snapshot=Date.now()-lastSnapshot>5*60*1000;await save(state,snapshot);if(snapshot){lastSnapshot=Date.now();void cleanupRevisions(15,state.id).catch(console.warn)}setSaveState('Salvo')}catch(err){console.error(err);dirty=true;setSaveState('Falha ao salvar');notify('Falha no salvamento local. Use Arquivo > Salvar como para criar uma cópia.')}finally{saving=false;if(dirty)scheduleSave()}}
function syncControls(){$('docTitle').value=state.title;$('breadcrumbDoc').textContent=state.title;$('breadcrumbTab').textContent=currentTab().name;$('autoHighlight').checked=state.autoHighlight;$('autoHighlightPopup').checked=state.autoHighlight;$('assetSnap').checked=state.snap;$('highlightCustom').value=state.highlight;$('currentFont').textContent=state.font;$('fontSize').value=state.fontSize;$('highlightToggle').style.background=state.highlight;$('highlightPaletteBtn').style.background=state.highlight;$('highlightToggle').style.color='#23252b';$('highlightPaletteBtn').style.color='#23252b';$('ruler').classList.toggle('hidden',!state.ruler);$('editor').lang=state.language||'pt-BR';$('editor').spellcheck=state.spellcheckEnabled!==false;const favorite=$('favoriteButton');if(favorite){favorite.classList.toggle('is-favorite',!!state.starred);favorite.setAttribute('aria-pressed',String(!!state.starred));favorite.title=state.starred?'Remover dos favoritos':'Adicionar aos favoritos'}}
function renderTabs(){$('tabList').innerHTML='';state.tabs.forEach(tab=>{const b=document.createElement('button');b.className=`tab-item ${state.activeId===tab.id?'active':''}`;b.innerHTML=`<span class="tab-icon">▤</span><span class="tab-name">${escapeHTML(tab.name)}</span><span class="tab-more" title="Opções da guia">⋮</span>`;b.addEventListener('click',e=>{if(e.target.closest('.tab-more')){openMenu([{text:'Renomear guia',action:()=>renameTab(tab.id)},{text:'Duplicar guia',action:()=>duplicateTab(tab.id)},{text:'Excluir guia',action:()=>deleteTab(tab.id)}],e.clientX,e.clientY);return}switchTab(tab.id)});$('tabList').appendChild(b)});if(state.tabs.length===1){const hint=document.createElement('p');hint.className='tab-empty-hint';hint.textContent='Os títulos que forem adicionados ao documento aparecerão aqui.';$('tabList').appendChild(hint)}}
function switchTab(id){if(!state.tabs.some(t=>t.id===id))return;writeEditor();hideMenus();currentAsset=null;closeInspector();state.activeId=id;$('editor').innerHTML=currentTab().html||'';renderTabs();renderAssets();syncControls();updateWordCount();renderComments();scheduleSave()}
function newTab(){writeEditor();const id=uid();state.tabs.push({id,name:`Guia ${state.tabs.length+1}`,html:'<p></p>',assets:[]});switchTab(id);notify('Guia criada')}
function renameTab(id){const tab=state.tabs.find(t=>t.id===id);if(!tab)return;const name=prompt('Novo nome da guia:',tab.name);if(!name?.trim())return;tab.name=name.trim();renderTabs();syncControls();scheduleSave()}
function duplicateTab(id){writeEditor();const tab=state.tabs.find(t=>t.id===id);if(!tab)return;const clone=structuredClone(tab);clone.id=uid();clone.name+=' · cópia';clone.assets=clone.assets.map(a=>({...a,id:uid()}));state.tabs.push(clone);switchTab(clone.id)}
function deleteTab(id){if(state.tabs.length===1){notify('Mantenha ao menos uma guia');return;}const tab=state.tabs.find(t=>t.id===id);if(!confirm(`Excluir a guia “${tab?.name}”? Exporte um backup se precisar guardá-la.`))return;state.tabs=state.tabs.filter(t=>t.id!==id);if(id===state.activeId){state.activeId=state.tabs[0].id;$('editor').innerHTML=currentTab().html||'';renderAssets()}renderTabs();syncControls();scheduleSave()}
function setTheme(t){state.theme={...t};const root=document.documentElement;root.style.setProperty('--accent',t.accent);root.style.setProperty('--surface',t.surface);root.style.setProperty('--panel',t.panel);root.style.setProperty('--text',t.text);root.style.setProperty('--paper',t.paper);root.style.setProperty('--border',t.surface=== '#22252d'?'#434754':'#e4e7ed');root.style.setProperty('--muted',t.surface==='#22252d'?'#a9b1bf':'#7b8492');root.style.setProperty('--background-image',t.image?`linear-gradient(#ffffff55,#ffffff55),url("${t.image}")`:'none');const bodyBg=t.surface==='#22252d'?'dark':'light';root.style.colorScheme=bodyBg;}
function chooseFont(name){if(activeSelection&&restoreSelection()){document.execCommand('fontName',false,name)}state.font=name;$('currentFont').textContent=name;$('editor').style.fontFamily=`"${name}",serif`;writeEditor();scheduleSave();closeModal();$('editor').focus()}
function applyTextFont(name){chooseFont(name)}
function activateTab(){const tab=currentTab();$('editor').innerHTML=tab.html||'';$('editor').style.fontFamily=`"${state.font}",serif`;renderAssets();updateWordCount();renderComments()}
function updateWordCount(){const text=$('editor').innerText||'';const words=text.trim()?text.trim().split(/\s+/).length:0;$('wordCounter').textContent=`${words} palavras · ${text.length} caracteres`}
function getValidSelection(){const selection=window.getSelection();if(!selection||selection.isCollapsed||!selection.rangeCount)return null;const range=selection.getRangeAt(0);if(!$('editor').contains(range.commonAncestorContainer))return null;return range}
function preserveSelection(){const range=getValidSelection();if(range)activeSelection=range.cloneRange()}
function restoreSelection(){if(!activeSelection)return false;const selection=window.getSelection();try{selection.removeAllRanges();selection.addRange(activeSelection.cloneRange());return true}catch{return false}}
function editorCommand(command,value){$('editor').focus();restoreSelection();document.execCommand('styleWithCSS',false,true);document.execCommand(command,false,value);writeEditor();updateWordCount();scheduleSave()}
function applyHighlight(color=state.highlight){const range=getValidSelection()||(restoreSelection()?getValidSelection():null);if(!range){notify('Selecione um texto primeiro');return}document.execCommand('styleWithCSS',false,true);document.execCommand('hiliteColor',false,color);preserveSelection();writeEditor();scheduleSave();}
function removeHighlight(){const range=getValidSelection()||(restoreSelection()?getValidSelection():null);if(!range)return;document.execCommand('hiliteColor',false,'transparent');writeEditor();scheduleSave()}
function autoMark(){const range=getValidSelection();if(!range)return;preserveSelection();if(state.autoHighlight){applyHighlight(state.highlight);$('selectionMenu').classList.add('hidden')}}
function openMenu(items,x,y){hideMenus();const menu=$('menuPopup');menu.innerHTML='';items.forEach(item=>{const b=document.createElement('button');b.textContent=item.text;b.addEventListener('click',()=>{menu.classList.add('hidden');item.action()});menu.appendChild(b)});positionPopup(menu,x,y)}
function positionPopup(el,x,y){el.classList.remove('hidden');requestAnimationFrame(()=>{el.style.left=`${clamp(x,8,innerWidth-el.offsetWidth-8)}px`;el.style.top=`${clamp(y,8,innerHeight-el.offsetHeight-8)}px`})}
function hideMenus(){$('menuPopup').classList.add('hidden');$('selectionMenu').classList.add('hidden');$('highlightPopup').classList.add('hidden')}
function openHighlightPopup(x,y){hideMenus();renderSwatches();positionPopup($('highlightPopup'),x,y)}
function renderSwatches(){const box=$('highlightSwatches');box.innerHTML='';highlightColors.forEach(c=>{const b=document.createElement('button');b.style.backgroundColor=c;b.className=state.highlight===c?'chosen':'';b.title=c;b.addEventListener('click',()=>{state.highlight=c;syncControls();if(getValidSelection()||activeSelection)applyHighlight(c);renderSwatches();scheduleSave()});box.appendChild(b)})}
function addNote(){restoreSelection();const range=getValidSelection();if(!range){notify('Selecione o trecho da nota');return}const note=prompt('Nota para o trecho selecionado:');if(!note?.trim())return;const fragment=range.extractContents();const node=document.createElement('span');node.className='note-mark';node.title=note.trim();node.dataset.note=note.trim();node.style.borderBottom='2px dotted #ba5474';node.append(fragment);range.insertNode(node);writeEditor();scheduleSave();notify('Nota adicionada; passe o mouse para ler')}
function installFontFaces(){document.querySelectorAll('style[data-custom-font]').forEach(n=>n.remove());(state.fonts||[]).filter(font=>/^[a-zA-ZÀ-ÿ0-9 _.()\-]{1,80}$/.test(font.name)&&/^data:(font\/[^;,]+|application\/(x-font-[^;,]+|octet-stream));base64,/.test(font.data)).forEach(font=>{const style=document.createElement('style');style.dataset.customFont='1';style.textContent=`@font-face{font-family:"${font.name.replaceAll('"','')}";src:url("${font.data}");font-display:swap}`;document.head.appendChild(style)})}
function readData(file){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=()=>reject(reader.error);reader.readAsDataURL(file)})}
function download(blob,filename){const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=filename;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),8000)}
function safeFilename(){return(state.title||'Documento').replace(/[<>:"/\\|?*\x00-\x1F]/g,'').trim().slice(0,75)||'Documento'}
function exportNative(){writeEditor();download(new Blob([JSON.stringify({...state,exportedAt:new Date().toISOString()},null,2)],{type:'application/json'}),`${safeFilename()}.clicky`);notify('Backup Clicky exportado')}
function exportTxt(){writeEditor();let lines=[];for(const t of state.tabs){const node=document.createElement('div');node.innerHTML=t.html;lines.push(`${t.name.toUpperCase()}\n\n${node.innerText||node.textContent}\n`)}download(new Blob([lines.join('\n-------------------\n\n')],{type:'text/plain;charset=utf-8'}),`${safeFilename()}.txt`);notify('TXT exportado')}
function exportHtml(){writeEditor();let sections=state.tabs.map(t=>`<section><h1>${escapeHTML(t.name)}</h1>${t.html||''}${(t.assets||[]).map(a=>`<img src="${a.data}" width="${a.width}" alt="Imagem adicionada"/>`).join('')}</section>`).join('<hr/>');const html=`<!doctype html><html lang="pt-BR"><head><meta charset="UTF-8"><title>${escapeHTML(state.title)}</title></head><body style="max-width:800px;margin:45px auto;font-family:${escapeHTML(state.font)},serif;line-height:1.5">${sections}</body></html>`;download(new Blob([html],{type:'text/html;charset=utf-8'}),`${safeFilename()}.html`)}
function exportWord(){try{writeEditor();download(createDocx(state),`${safeFilename()}.docx`);notify('DOCX exportado (imagens inseridas ao final de cada guia)')}catch(err){console.error(err);notify('Erro ao gerar DOCX: '+err.message)}}
function showExports(x,y){openMenu([{text:'Salvar como documento Word (.docx)',action:()=>void saveAs('docx')},{text:'Salvar como PDF',action:()=>window.print()},{text:'Salvar como texto (.txt)',action:()=>void saveAs('txt')},{text:'Salvar projeto editável (.clicky)',action:()=>void saveAs('clicky')},{text:'Salvar página web (.html)',action:exportHtml}],x,y)}
function sanitizeHtml(html){
 const root=new DOMParser().parseFromString(`<div id="clean-root">${html||''}</div>`,'text/html').getElementById('clean-root');
 const tags=new Set(['P','DIV','BR','H1','H2','H3','H4','H5','H6','TABLE','TBODY','TR','TD','TH','SPAN','B','STRONG','I','EM','U','S','STRIKE','DEL','SUB','SUP','MARK','UL','OL','LI','BLOCKQUOTE','A','FONT']);
 const blocked=new Set(['SCRIPT','STYLE','IFRAME','OBJECT','EMBED','SVG','MATH','FORM','INPUT','BUTTON','VIDEO','AUDIO','META','LINK']);
 for(const el of [...root.querySelectorAll('*')].reverse()){
  if(blocked.has(el.tagName)){el.remove();continue}
  if(!tags.has(el.tagName)){el.replaceWith(...el.childNodes);continue}
  for(const attr of [...el.attributes]){
   const key=attr.name.toLowerCase(),value=attr.value;
   if(key==='style'){
    const tmp=document.createElement('span');tmp.style.cssText=value;
    const safe=['color','background-color','font-family','font-size','font-weight','font-style','text-align','text-decoration','text-decoration-line'].map(k=>[k,tmp.style.getPropertyValue(k)]).filter(([k,v])=>v&&!/url\s*\(|expression\s*\(|[;{}<>]/i.test(v)).map(([k,v])=>`${k}:${v}`).join(';');
    if(safe)el.setAttribute('style',safe);else el.removeAttribute('style');
   }else if(key==='href'&&el.tagName==='A'&&/^(https?:\/\/|mailto:|#)/i.test(value)){} 
   else if(key==='title'&&el.tagName==='SPAN')el.setAttribute('title',value.slice(0,500));
   else if(key==='data-note'&&el.tagName==='SPAN')el.setAttribute('data-note',value.slice(0,500));
    else if(key==='data-comment-id'&&el.tagName==='SPAN'&&/^[a-zA-Z0-9_-]{5,100}$/.test(value)){}
   else if(key==='class'&&el.tagName==='SPAN'&&['note-mark','comment-anchor'].includes(value)){}
   else if(key==='color'&&el.tagName==='FONT'&&/^#[0-9a-f]{6}$/i.test(value)){}
   else el.removeAttribute(attr.name);
  }
 }
 return root.innerHTML;
}
function normalizeImportedState(s){
 const safe=structuredClone(s);
 safe.tabs=safe.tabs.map(t=>({...t,name:String(t.name||'Guia').slice(0,100),html:sanitizeHtml(t.html),comments:Array.isArray(t.comments)?t.comments.filter(c=>typeof c.id==='string'&&typeof c.text==='string').map(c=>({id:c.id.slice(0,100),quote:String(c.quote||'').slice(0,250),text:c.text.slice(0,10000),createdAt:c.createdAt||new Date().toISOString(),resolved:!!c.resolved,replies:Array.isArray(c.replies)?c.replies.filter(r=>typeof r.text==='string').map(r=>({text:r.text.slice(0,10000),createdAt:r.createdAt||new Date().toISOString()})):[]})):[],assets:(Array.isArray(t.assets)?t.assets:[]).filter(a=>/^data:image\/(png|jpeg|webp|gif);base64,/.test(a.data)&&a.data.length<28000000).map(a=>({...a,id:String(a.id||uid()),width:clamp(+a.width||260,35,1600),height:clamp(+a.height||200,1,9000),x:clamp(+a.x||0,-2000,3000),y:clamp(+a.y||0,-2000,9000),locked:!!a.locked}))}));
 safe.title=String(safe.title||'Documento').slice(0,150);
 safe.fonts=Array.isArray(safe.fonts)?safe.fonts.filter(f=>typeof f.name==='string'&&typeof f.data==='string'&&f.data.length<8000000):[];
 safe.theme={...presets[0],...safe.theme};
 for(const k of ['accent','surface','panel','text','paper'])if(!/^#[\da-f]{6}$/i.test(safe.theme[k]))safe.theme[k]=presets[0][k];
 if(safe.theme.image&&(!/^data:image\/(png|gif|jpeg|webp);base64,/.test(safe.theme.image)||safe.theme.image.length>12000000))safe.theme.image=null;
 return safe;
}
async function importNative(file){if(!file)return;try{const data=JSON.parse(await file.text());if(!Array.isArray(data.tabs)||!data.tabs.length||!data.tabs.every(t=>typeof t.id==='string'&&typeof t.html==='string'&&Array.isArray(t.assets)))throw new Error('Formato do documento incompatível');if(!confirm('Importar este arquivo como um novo documento Clicky?'))return;writeEditor();state=normalizeImportedState({...defaultState(),...data,id:uid(),folderId:libraryFolder==='all'||libraryFolder==='root'?null:libraryFolder});state.activeId=state.tabs.some(t=>t.id===state.activeId)?state.activeId:state.tabs[0].id;installFontFaces();setTheme(state.theme);renderTabs();activateTab();syncControls();closeInspector();scheduleSave();await persist(true);await refreshLibrary();showEditor();notify('Documento importado como novo projeto')}catch(err){notify('Falha ao importar arquivo: '+err.message)}}
function openModal(title,html){hideMenus();$('modalTitle').textContent=title;$('modalBody').innerHTML=html;$('modalShade').classList.remove('hidden')}
function closeModal(){$('modalShade').classList.add('hidden')}
function openThemes(){const cards=presets.map(t=>`<button class="theme-card" data-preset="${escapeHTML(t.name)}"><span class="theme-circle" style="background:${t.surface};box-shadow:inset 0 0 0 6px ${t.accent}"></span>${escapeHTML(t.name)}</button>`).join('');openModal('Aparência e temas',`<section class="modal-section"><h3>Temas prontos</h3><div class="palette-grid">${cards}</div></section><section class="modal-section"><h3>Personalizar cores</h3><div class="form-grid"><label>Cor principal <input id="themeAccent" type="color" value="${state.theme.accent}"/></label><label>Fundo do aplicativo <input id="themeSurface" type="color" value="${state.theme.surface}"/></label><label>Barra e painéis <input id="themePanel" type="color" value="${state.theme.panel}"/></label><label>Cor do texto da interface <input id="themeText" type="color" value="${state.theme.text}"/></label><label>Cor do papel <input id="themePaper" type="color" value="${state.theme.paper}"/></label></div><div class="modal-actions"><button id="themeImageButton" class="pill-button">Imagem de fundo</button><button id="removeImageButton" class="pill-button">Remover fundo</button></div></section><section class="modal-section"><h3>Compartilhar tema</h3><div class="modal-actions"><button id="exportTheme" class="pill-button primary">Exportar tema</button><button id="importThemeBtn" class="pill-button">Importar tema</button></div></section>`);document.querySelectorAll('[data-preset]').forEach(btn=>btn.addEventListener('click',()=>{const theme=presets.find(t=>t.name===btn.dataset.preset);if(!theme)return;setTheme({...theme,image:null});scheduleSave();openThemes()}));[['themeAccent','accent'],['themeSurface','surface'],['themePanel','panel'],['themeText','text'],['themePaper','paper']].forEach(([id,key])=>$(id).addEventListener('input',e=>{setTheme({...state.theme,[key]:e.target.value});scheduleSave()}));$('themeImageButton').onclick=()=>$('themePicture').click();$('removeImageButton').onclick=()=>{setTheme({...state.theme,image:null});scheduleSave()};$('exportTheme').onclick=()=>{download(new Blob([JSON.stringify({...state.theme,format:'clicky-theme-v1'},null,2)],{type:'application/json'}),`${state.theme.name||'Meu tema'}.clicky-theme.json`);notify('Tema exportado')};$('importThemeBtn').onclick=()=>$('themeImport').click()}
function openFonts(){const choices=[...new Set([...fonts,...(state.fonts||[]).map(f=>f.name)])];openModal('Biblioteca de fontes',`<div class="modal-section"><input class="font-search" id="fontSearch" placeholder="Pesquisar fontes..."/><div id="fontOptions" class="font-options"></div></div><div class="modal-actions"><button id="fontImportBtn" class="pill-button primary">+ Importar arquivos .ttf / .otf / .woff</button></div><p>Fontes importadas ficam disponíveis neste dispositivo. Para compartilhar com outro computador, exporte um backup Clicky.</p>`);function fill(query=''){const list=$('fontOptions');list.innerHTML='';choices.filter(f=>f.toLowerCase().includes(query.toLowerCase())).forEach(f=>{const item=document.createElement('button');item.className='font-option';item.innerHTML=`<span style="font-family:${escapeHTML(f)},serif">${escapeHTML(f)} · Aa Bb Cc 123</span><small>${f===state.font?'Em uso':'Selecionar'}</small>`;item.onclick=()=>applyTextFont(f);list.appendChild(item)})}fill();$('fontSearch').addEventListener('input',e=>fill(e.target.value));$('fontImportBtn').onclick=()=>$('fontFile').click()}
function openSettings(){openModal('Preferências',`<section class="modal-section"><h3>Escrita</h3><label class="inspect-row">Destaque automático após selecionar <input type="checkbox" id="prefMark" ${state.autoHighlight?'checked':''}/></label><label class="inspect-row">Magnetismo para imagens <input type="checkbox" id="prefSnap" ${state.snap?'checked':''}/></label><label class="inspect-row">Mostrar régua da página <input type="checkbox" id="prefRuler" ${state.ruler?'checked':''}/></label></section><section class="modal-section"><h3>Proteção do conteúdo</h3><p>Os documentos são salvos automaticamente neste dispositivo com IndexedDB. Também é recomendado exportar periodicamente um arquivo .clicky como backup externo.</p><div class="modal-actions"><button id="backupNow" class="pill-button primary">Exportar backup agora</button><button id="revisionsBtn" class="pill-button">Ver versões de segurança</button></div></section>`);$('prefMark').onchange=e=>{state.autoHighlight=e.target.checked;syncControls();scheduleSave()};$('prefSnap').onchange=e=>{state.snap=e.target.checked;syncControls();scheduleSave()};$('prefRuler').onchange=e=>{state.ruler=e.target.checked;syncControls();scheduleSave()};$('backupNow').onclick=exportNative;$('revisionsBtn').onclick=()=>void showRevisions()}
async function showRevisions(){const all=(await revisions(state.id)).slice(0,15);openModal('Versões de segurança',`<p>O Clicky mantém cópias locais periódicas. Essas cópias não substituem o backup exportado.</p><div id="revisionsList" class="font-options"></div>`);const list=$('revisionsList');for(const r of all){const item=document.createElement('button');item.className='font-option';item.textContent=`${new Date(r.createdAt).toLocaleString('pt-BR')} · ${r.state.title}`;item.onclick=()=>{if(confirm('Restaurar esta versão? As alterações atuais serão substituídas.')){state=r.state;setTheme(state.theme);installFontFaces();renderTabs();activateTab();syncControls();closeModal();scheduleSave()}};list.appendChild(item)}if(!all.length)list.textContent='Ainda não existem versões anteriores.'}
function assetById(id){return currentTab().assets.find(a=>a.id===id)}
function closeInspector(){currentAsset=null;$('assetInspector').classList.add('hidden');document.querySelectorAll('.asset.selected').forEach(el=>el.classList.remove('selected'))}
function selectAsset(id){currentAsset=id;const a=assetById(id);if(!a)return;document.querySelectorAll('.asset').forEach(n=>n.classList.toggle('selected',n.dataset.id===id));$('assetInspector').classList.remove('hidden');$('inspectorPreview').innerHTML='';const img=document.createElement('img');img.src=a.data;$('inspectorPreview').append(img);$('assetWidth').value=a.width;$('assetLock').checked=!!a.locked;$('assetSnap').checked=state.snap}
function renderAssets(){const layer=$('attachmentsLayer');layer.innerHTML='';for(const asset of currentTab().assets){const div=document.createElement('div');div.className=`asset ${asset.locked?'locked':''} ${currentAsset===asset.id?'selected':''}`;div.dataset.id=asset.id;div.style.left=`${asset.x}px`;div.style.top=`${asset.y}px`;div.style.width=`${asset.width}px`;const img=document.createElement('img');img.src=asset.data;img.draggable=false;const lock=document.createElement('button');lock.className='lock-bubble';lock.innerHTML=lockSvg(asset.locked);lock.setAttribute('aria-label',asset.locked?'Desbloquear imagem':'Bloquear imagem');lock.title=asset.locked?'Desbloquear imagem':'Bloquear imagem';lock.onclick=e=>{e.stopPropagation();asset.locked=!asset.locked;renderAssets();selectAsset(asset.id);scheduleSave()};const handle=document.createElement('span');handle.className='resize-handle';div.append(img,lock,handle);div.addEventListener('pointerdown',e=>{if(e.target===lock)return;selectAsset(asset.id);if(e.button!==0||asset.locked)return;const resizing=e.target===handle;const initial={x:asset.x,y:asset.y,width:asset.width,clientX:e.clientX,clientY:e.clientY};dragState={id:asset.id,resizing,initial};div.setPointerCapture(e.pointerId);e.preventDefault()});div.addEventListener('pointermove',e=>{if(!dragState||dragState.id!==asset.id||asset.locked)return;const d=dragState.initial;const dx=e.clientX-d.clientX,dy=e.clientY-d.clientY;if(dragState.resizing){asset.width=clamp(Math.round(d.width+dx),35,1600);div.style.width=`${asset.width}px`}else{let x=d.x+dx,y=d.y+dy;if(state.snap){const result=calculateSnap({x,y,width:asset.width,height:asset.height},currentTab().assets.filter(a=>a.id!==asset.id));x=result.x;y=result.y;drawSnapGuides(result.guides,{x,y,width:asset.width,height:asset.height})}else clearSnapGuides();asset.x=Math.round(x);asset.y=Math.round(y);div.style.left=`${asset.x}px`;div.style.top=`${asset.y}px`}});div.addEventListener('pointerup',e=>{if(!dragState)return;dragState=null;clearSnapGuides();if(div.hasPointerCapture(e.pointerId))div.releasePointerCapture(e.pointerId);$('assetWidth').value=asset.width;scheduleSave()});div.addEventListener('pointercancel',()=>{dragState=null;clearSnapGuides();scheduleSave()});layer.appendChild(div)}}
async function addImage(file){if(!file)return;if(file.size>18*1024*1024){notify('Imagem muito grande: limite de 18 MB nesta versão');return}try{const data=await readData(file);const dims=await new Promise((resolve,reject)=>{const image=new Image();image.onload=()=>resolve({width:image.naturalWidth,height:image.naturalHeight});image.onerror=reject;image.src=data});const width=clamp(dims.width,90,380);const asset={id:uid(),name:file.name,data,x:100,y:150,width,height:Math.round(dims.height*width/dims.width),locked:false};currentTab().assets.push(asset);renderAssets();selectAsset(asset.id);scheduleSave();notify('Imagem inserida. Arraste ou use o cadeado.')}catch(err){console.error(err);notify('Não foi possível inserir esta imagem')}}
function menuFor(name,x,y){
 const actions={
  file:[['Meus documentos',()=>void showLibrary()],['Novo documento',()=>void newDocument()],['Nova pasta',()=>void newFolder()],['Abrir arquivo (.docx ou .clicky)',()=>$('nativeImport').click()],['Fazer uma cópia',()=>void duplicateDocument()],['Mover para pasta',()=>moveDocumentDialog()],['Renomear',()=>{$('docTitle').focus();$('docTitle').select()}],['Salvar como…',()=>showExports(x,y)],['Imprimir',()=>window.print()]],
  edit:[['Desfazer · Ctrl+Z',()=>editorCommand('undo')],['Refazer · Ctrl+Shift+Z',()=>editorCommand('redo')],['Recortar · Ctrl+X',()=>editorCommand('cut')],['Copiar · Ctrl+C',()=>editorCommand('copy')],['Colar',()=>notify('Use Ctrl/Cmd+V para colar')],['Selecionar tudo · Ctrl+A',()=>editorCommand('selectAll')],['Localizar e substituir · Ctrl+H',openFindReplace]],
  view:[['Mostrar / ocultar guias',toggleSidebar],['Mostrar / ocultar régua',()=>{state.ruler=!state.ruler;syncControls();scheduleSave()}],['Comentários',toggleComments],['Biblioteca de documentos',()=>void showLibrary()],['Personalizar aparência',openThemes]],
  insert:[['Imagem ou GIF',()=>$('imageFile').click()],['Comentar · Ctrl+Alt+M',beginComment],['Nova guia',newTab],['Link · Ctrl+K',()=>{const url=prompt('Endereço do link:','https://');if(url)editorCommand('createLink',url)}],['Linha horizontal',()=>editorCommand('insertHorizontalRule')]],
  format:[['Biblioteca de fontes',openFonts],['Negrito · Ctrl+B',()=>editorCommand('bold')],['Itálico · Ctrl+I',()=>editorCommand('italic')],['Sublinhado · Ctrl+U',()=>editorCommand('underline')],['Grifo rápido',()=>openHighlightPopup(x,y)],['Texto normal',()=>editorCommand('formatBlock','p')],['Limpar formatação',()=>editorCommand('removeFormat')]],
  tools:[['Ortografia e idioma',openProofing],['Revisão rápida',openReview],['Dicionário pessoal',openDictionary],['Contagem de palavras',openWordCount],['Preferências',openSettings],['Versões de segurança',()=>void showRevisions()]],
  extensions:[['Fontes personalizadas',openFonts],['Temas personalizados',openThemes],['Revisão e dicionário',openProofing]],
  help:[['Pesquisar nos menus · Alt+/',openCommandSearch],['Atalhos do teclado · Ctrl+/',openShortcutHelp],['Sobre o aplicativo',()=>openModal('Sobre o aplicativo',`<p>Editor local e open source. Versão 0.3.0.</p><p>Seus documentos são armazenados no dispositivo. Comentários são locais; sincronização e colaboração não estão implementadas.</p>`)] ]
 };openMenu((actions[name]||[]).map(([text,action])=>({text,action})),x,y)
}
function toggleSidebar(){const collapsed=$('sidebar').classList.toggle('collapsed');$('stage-top').classList.toggle('sidebar-hidden',collapsed)}
let copiedFormat=null;
function beginFormatPainting(){
 const selection=window.getSelection();const node=selection?.anchorNode;
 const anchor=node?.nodeType===Node.TEXT_NODE?node.parentElement:node;
 if(!anchor||!$('editor').contains(anchor)){notify('Selecione primeiro um trecho formatado.');return}
 const computed=getComputedStyle(anchor);
 copiedFormat={fontFamily:computed.fontFamily,fontSize:computed.fontSize,fontWeight:computed.fontWeight,fontStyle:computed.fontStyle,color:computed.color,textDecoration:computed.textDecorationLine};
 $('paintFormat').classList.add('paint-active');notify('Selecione o texto que receberá a formatação.');
}
function applyPaintedFormat(){
 if(!copiedFormat)return;
 const range=getValidSelection();if(!range)return;
 const selection=window.getSelection();const text=range.toString();
 if(!text.trim())return;
 const span=document.createElement('span');Object.assign(span.style,copiedFormat);
 try{range.surroundContents(span)}catch{
  const fragment=range.extractContents();span.appendChild(fragment);range.insertNode(span)
 }
 selection?.removeAllRanges();copiedFormat=null;$('paintFormat').classList.remove('paint-active');writeEditor();scheduleSave();
}
function setPageZoom(value){const zoom=clamp(value,75,150);$('paper').style.zoom=String(zoom/100);$('ruler').style.zoom=String(zoom/100);$('zoomSelect').value=String(zoom);$('pageStage').style.minHeight=`${Math.ceil(zoom/100*1090+40)}px`;}
function initNativeWindowButtons(){if(!window.__TAURI__?.core?.invoke)return;document.documentElement.classList.add('desktop-app');for(const [id,action] of [['winMin','minimize_window'],['winMax','toggle_maximize_window'],['winClose','close_window']]){$(id).onclick=()=>window.__TAURI__.core.invoke(action).catch(error=>notify('Não foi possível alterar a janela: '+error));}const bar=$('windowDragArea');bar.addEventListener('pointerdown',e=>{if(e.button===0&&!e.target.closest('button,input,select,textarea,a,[contenteditable]'))window.__TAURI__.core.invoke('start_dragging_window').catch(()=>{})});}
async function newDocument(folderId=null){await persist(true);state=defaultState();if(typeof folderId==='string')state.folderId=folderId;else if(!folderId&&libraryFolder!=='all'&&libraryFolder!=='root')state.folderId=libraryFolder;installFontFaces();setTheme(state.theme);renderTabs();activateTab();syncControls();closeInspector();await save(state,true);await refreshLibrary();showEditor();notify('Novo documento criado')}
function wire(){
 wireLibraryAndFeatures();
 $('docTitle').addEventListener('input',e=>{state.title=e.target.value||'Sem título';$('breadcrumbDoc').textContent=state.title;scheduleSave()});
 $('newTab').onclick=newTab;$('sidebarHideBtn').onclick=toggleSidebar;$('sidebarShowBtn').onclick=toggleSidebar;$('sidebarBack').onclick=toggleSidebar;$('importDocBtn').onclick=()=>$('nativeImport').click();$('newTab').title='Adicionar guia';
 document.querySelectorAll('[data-menu]').forEach(b=>b.addEventListener('click',e=>{const r=b.getBoundingClientRect();menuFor(b.dataset.menu,r.left,r.bottom+6)}));
 document.querySelectorAll('[data-action]').forEach(b=>b.addEventListener('mousedown',e=>e.preventDefault()));document.querySelectorAll('[data-action]').forEach(b=>b.addEventListener('click',()=>{const a=b.dataset.action;const actionMap={undo:()=>editorCommand('undo'),redo:()=>editorCommand('redo'),print:()=>window.print(),bold:()=>editorCommand('bold'),italic:()=>editorCommand('italic'),underline:()=>editorCommand('underline'),smaller:()=>setSize((+($('fontSize').value)||12)-1),bigger:()=>setSize((+($('fontSize').value)||12)+1),link:()=>{const url=prompt('Cole o endereço do link:','https://');if(url)editorCommand('createLink',url)},alignLeft:()=>editorCommand('justifyLeft'),alignCenter:()=>editorCommand('justifyCenter'),alignRight:()=>editorCommand('justifyRight'),bullet:()=>editorCommand('insertUnorderedList'),numbered:()=>editorCommand('insertOrderedList'),indent:()=>editorCommand('indent'),outdent:()=>editorCommand('outdent'),clearFormat:()=>editorCommand('removeFormat')};actionMap[a]?.()}));
 $('paragraphType').onchange=e=>editorCommand('formatBlock',e.target.value);$('fontPickerBtn').onclick=openFonts;$('fontSize').onchange=e=>setSize(+e.target.value);$('textColor').oninput=e=>editorCommand('foreColor',e.target.value);
 $('highlightToggle').addEventListener('mousedown',e=>e.preventDefault());$('highlightToggle').onclick=()=>{restoreSelection();applyHighlight()};$('highlightPaletteBtn').onclick=e=>{const r=e.target.getBoundingClientRect();openHighlightPopup(r.left,r.bottom+8)};$('highlightCustom').oninput=e=>{state.highlight=e.target.value;syncControls();renderSwatches();scheduleSave()};
 $('autoHighlight').onchange=e=>{state.autoHighlight=e.target.checked;syncControls();scheduleSave()};$('autoHighlightPopup').onchange=e=>{state.autoHighlight=e.target.checked;syncControls();scheduleSave()};$('themeBtn').onclick=openThemes;$('settingsBtn').onclick=openSettings;$('proofTop').onclick=openProofing;$('paintFormat').onclick=beginFormatPainting;$('revisionsTop').onclick=()=>void showRevisions();$('moveToFolderBtn').onclick=moveDocumentDialog;$('cloudSaveBtn').onclick=()=>notify('Salvo automaticamente neste dispositivo. Use Arquivo > Salvar como para escolher outra pasta.');$('favoriteButton').onclick=()=>{state.starred=!state.starred;syncControls();scheduleSave()};$('zoomSelect').onchange=e=>setPageZoom(Number(e.target.value)||100);initNativeWindowButtons();$('imageBtn').onclick=()=>$('imageFile').click();$('exportTop').onclick=e=>{const r=e.target.getBoundingClientRect();showExports(r.left,r.bottom+5)};$('searchMenus').onclick=openCommandSearch;
 $('toggleRuler').onclick=()=>{state.ruler=!state.ruler;syncControls();scheduleSave()};$('imageFile').onchange=async e=>{await addImage(e.target.files?.[0]);e.target.value=''};
 $('themePicture').onchange=async e=>{const f=e.target.files?.[0];if(f){if(f.size>8*1024*1024){notify('Fundo muito grande. Use uma imagem menor que 8 MB.');return}setTheme({...state.theme,image:await readData(f)});scheduleSave();notify('Imagem de fundo aplicada')}e.target.value=''};
 $('fontFile').onchange=async e=>{for(const f of [...e.target.files]){if(f.size>5*1024*1024){notify(`Fonte ${f.name} excede 5 MB`);continue}const name=f.name.replace(/\.(ttf|otf|woff2?)$/i,'').replace(/[_-]/g,' ');const data=await readData(f);state.fonts=state.fonts.filter(x=>x.name!==name);state.fonts.push({name,data})}installFontFaces();scheduleSave();openFonts();e.target.value='';notify('Fontes importadas')};
 $('nativeImport').onchange=async e=>{const file=e.target.files?.[0];if(file){if(/\.docx$/i.test(file.name))await importWord(file);else await importNative(file)}e.target.value=''};
 $('themeImport').onchange=async e=>{const file=e.target.files?.[0];if(!file)return;try{const t=JSON.parse(await file.text());if(!['accent','surface','panel','text','paper'].every(k=>/^#[a-f\d]{6}$/i.test(t[k])))throw new Error('As cores do tema são inválidas');if(t.image&&!/^data:image\/(png|gif|jpeg|webp);base64,/.test(t.image))throw new Error('Imagem de fundo inválida');setTheme({...t});scheduleSave();openThemes();notify('Tema importado')}catch(err){notify('Não foi possível ler o tema: '+err.message)}e.target.value=''};
 $('closeModal').onclick=closeModal;$('modalShade').onclick=e=>{if(e.target===$('modalShade'))closeModal()};
 $('editor').addEventListener('input',()=>{writeEditor();updateWordCount();scheduleSave()});$('editor').addEventListener('paste',e=>{const html=e.clipboardData?.getData('text/html');if(!html)return;e.preventDefault();document.execCommand('insertHTML',false,sanitizeHtml(html));scheduleSave()});$('editor').addEventListener('mouseup',()=>setTimeout(()=>{if(copiedFormat)applyPaintedFormat();else autoMark()},0));$('editor').addEventListener('keyup',e=>{if(e.key==='Shift'||e.key.startsWith('Arrow'))preserveSelection()});
 $('editor').addEventListener('click',e=>{const id=e.target.closest('[data-comment-id]')?.dataset.commentId;if(id){toggleComments(true);focusComment(id)}});$('editor').addEventListener('contextmenu',e=>{const range=getValidSelection();if(!range)return;e.preventDefault();preserveSelection();hideMenus();positionPopup($('selectionMenu'),e.clientX,e.clientY)});
 $('floatHighlight').onclick=()=>{hideMenus();restoreSelection();applyHighlight()};$('floatColor').onclick=e=>{const r=e.target.getBoundingClientRect();openHighlightPopup(r.left,r.bottom+4)};$('floatComment').onclick=()=>{hideMenus();beginComment()};$('floatCopy').onclick=()=>{restoreSelection();document.execCommand('copy');hideMenus();notify('Texto copiado')};$('floatNote').onclick=()=>{restoreSelection();hideMenus();addNote()};$('floatRemove').onclick=()=>{restoreSelection();removeHighlight();hideMenus()};
 $('closeInspector').onclick=closeInspector;$('assetWidth').onchange=e=>{const a=assetById(currentAsset);if(!a)return;const old=a.width;a.width=clamp(+e.target.value||old,35,1600);a.height=Math.round(a.height*a.width/old);renderAssets();selectAsset(a.id);scheduleSave()};$('assetLock').onchange=e=>{const a=assetById(currentAsset);if(!a)return;a.locked=e.target.checked;renderAssets();selectAsset(a.id);scheduleSave()};$('assetSnap').onchange=e=>{state.snap=e.target.checked;scheduleSave()};$('assetDelete').onclick=()=>{if(!currentAsset)return;currentTab().assets=currentTab().assets.filter(a=>a.id!==currentAsset);closeInspector();renderAssets();scheduleSave()};
 $('scrollSurface').addEventListener('pointerdown',e=>{if(e.target===$('paper')||e.target===$('scrollSurface')||e.target===$('pageStage'))closeInspector()});
 document.addEventListener('pointerdown',e=>{if(!e.target.closest('.floating-menu,.color-popup,.menu-popup,.highlight-more,[data-menu],.export-button'))hideMenus()});
 document.addEventListener('keydown',handleKeydown);
 window.addEventListener('pagehide',()=>{writeEditor();void persist(true)});document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden')void persist(true)});
}
function setSize(size){size=clamp(size,7,96);state.fontSize=size;$('fontSize').value=size;editorCommand('fontSize',3);const selection=window.getSelection();if(selection?.rangeCount){const node=selection.anchorNode?.parentElement;const el=node?.closest('span[style*="font-size"],font[size]');if(el){el.style.fontSize=`${size}pt`;el.removeAttribute('size')}}$('editor').style.fontSize=`${size*4/3}px`;scheduleSave()}
// --- Clicky 0.2: personal library, comment threads, proofreading and shortcuts ---
async function refreshLibrary(){library=await listLibrary();renderLibrary()}
function showEditor(){closeModal();$('libraryView').classList.add('hidden');$('app').classList.remove('hidden');$('editor').focus();}
async function showLibrary(){await persist(true);closeModal();closeInspector();await refreshLibrary();$('app').classList.add('hidden');$('libraryView').classList.remove('hidden');}
async function openDocument(id){
 if(id===state.id){showEditor();return}
 await persist(true);
 const doc=await loadDocument(id);if(!doc){notify('Documento não encontrado');return}
 state=normalizeImportedState({...defaultState(),...doc});
 if(!state.tabs.some(t=>t.id===state.activeId))state.activeId=state.tabs[0].id;
 installFontFaces();setTheme(state.theme);renderTabs();activateTab();syncControls();closeInspector();$('commentsPanel').classList.add('hidden');
 showEditor();setSaveState('Salvo');
}
function setLibraryFolder(value){libraryFolder=value;renderLibrary()}
function renderLibrary(){
 if(!$('documentList'))return;
 const folderName=id=>library.folders.find(f=>f.id===id)?.name||'Sem pasta';
 const folderButtons=$('folderList');folderButtons.innerHTML='';
 for(const folder of [...library.folders].sort((a,b)=>a.name.localeCompare(b.name,'pt-BR'))){
  const row=document.createElement('div');row.className='folder-row';
  const b=document.createElement('button');b.className='folder-link'+(libraryFolder===folder.id?' active':'');b.innerHTML=`<span class="folder-dot" style="background:${/^#[a-f0-9]{6}$/i.test(folder.color)?folder.color:'#ad94c8'}"></span> <span class="folder-ellipsis">${escapeHTML(folder.name)}</span>`;
  b.onclick=()=>setLibraryFolder(folder.id);
  const more=document.createElement('button');more.className='folder-more';more.textContent='⋯';more.title=`Opções de ${folder.name}`;more.onclick=e=>{const r=e.currentTarget.getBoundingClientRect();openMenu([{text:'Alterar nome e cor',action:()=>openFolderEditor(folder)},{text:'Excluir pasta (conservar documentos)',action:()=>void removeFolder(folder)}],r.left,r.bottom+3)};
  row.append(b,more);folderButtons.append(row);
 }
 $('allDocuments').classList.toggle('active',libraryFolder==='all');$('rootDocuments').classList.toggle('active',libraryFolder==='root');
 const filtered=library.documents.filter(d=>libraryFolder==='all'||(libraryFolder==='root'?!d.folderId:d.folderId===libraryFolder))
 .filter(d=>String(d.title).toLowerCase().includes(($('librarySearch').value||'').trim().toLowerCase()))
 .sort((a,b)=>new Date(b.updatedAt||0)-new Date(a.updatedAt||0));
 $('libraryHeading').textContent=libraryFolder==='all'?'Todos os documentos':libraryFolder==='root'?'Sem pasta':folderName(libraryFolder);
 $('libraryCount').textContent=`${filtered.length} documento${filtered.length!==1?'s':''}`;
 $('libraryGrid').classList.toggle('selected',libraryLayout==='grid');$('libraryList').classList.toggle('selected',libraryLayout==='list');
 const list=$('documentList');list.className=libraryLayout==='grid'?'document-grid':'document-list';list.innerHTML='';
 if(!filtered.length){list.innerHTML='<div class="library-empty"><div class="empty-icon">▤</div><h3>Nenhum documento aqui</h3><p>Crie um documento ou mova um projeto para esta pasta.</p></div>';return}
 for(const doc of filtered){
  const card=document.createElement('div');card.className='document-card';
  const hasFolder=!!doc.folderId;const folder=library.folders.find(f=>f.id===doc.folderId);
  const updated=new Date(doc.updatedAt||Date.now()).toLocaleDateString('pt-BR',{day:'2-digit',month:'short',year:'numeric'});
  card.innerHTML=`<div class="doc-art"><div class="doc-art-page" aria-hidden="true"><div class="doc-art-preview"></div></div></div><div class="doc-card-info"><strong>${escapeHTML(doc.title||'Sem título')}</strong><span>${updated}${hasFolder?` · ${escapeHTML(folder?.name||'Sem pasta')}`:''}</span></div><button class="doc-more" title="Opções do documento">⋯</button>`;
  card.addEventListener('click',e=>{if(e.target.closest('.doc-more'))return;void openDocument(doc.id)});
  card.querySelector('.doc-more').onclick=e=>{e.stopPropagation();const r=e.currentTarget.getBoundingClientRect();openMenu([{text:'Abrir documento',action:()=>void openDocument(doc.id)},{text:'Renomear',action:()=>void renameDocumentInLibrary(doc)},{text:'Mover para pasta',action:()=>moveDocumentDialog(doc.id)},{text:'Duplicar',action:()=>void duplicateDocument(doc.id)},{text:'Excluir documento',action:()=>void trashDocument(doc)}],r.left,r.bottom+5)};
  if(folder)card.style.setProperty('--folder-accent',folder.color);
  list.append(card);
  void loadDocument(doc.id).then(saved=>{if(!card.isConnected||!saved)return;const tab=saved.tabs?.[0];const temp=document.createElement('div');temp.innerHTML=tab?.html||'';const text=(temp.textContent||'').replace(/\s+/g,' ').trim();const preview=card.querySelector('.doc-art-preview');if(preview)preview.textContent=text.slice(0,420)}).catch(()=>{});
 }
}

async function exportLibraryBackup(){
 await persist(true);await refreshLibrary();const docs=[];
 for(const meta of library.documents){const doc=await loadDocument(meta.id);if(doc)docs.push(doc)}
 const backup={format:'clicky-library-v1',createdAt:new Date().toISOString(),documents:docs,folders:library.folders,preferences:{libraryLayout,dictionary}};
 download(new Blob([JSON.stringify(backup)],{type:'application/json'}),`Clicky-Biblioteca-${new Date().toISOString().slice(0,10)}.clicky-library`);
 notify(`${docs.length} documentos incluídos no backup`)
}
async function importLibraryBackup(file){
 if(!file)return;
 if(file.size>160*1024*1024){notify('Backup maior que 160 MB; importe em lotes menores');return}
 try{
  const data=JSON.parse(await file.text());
  if(data.format!=='clicky-library-v1'||!Array.isArray(data.documents)||!Array.isArray(data.folders)||data.documents.length>2000||data.folders.length>1000)throw new Error('Arquivo de biblioteca inválido');
  if(!confirm(`Importar ${data.documents.length} documentos e ${data.folders.length} pastas? A biblioteca atual será preservada; os itens importados serão acrescentados.`))return;
  const folderMap=new Map();
  for(const f of data.folders){if(typeof f.id!=='string'||typeof f.name!=='string')continue;const newId=uid();folderMap.set(f.id,newId);await saveFolder({id:newId,name:f.name.slice(0,80),color:/^#[a-f0-9]{6}$/i.test(f.color)?f.color:'#b584d1'})}
  let count=0;
  for(const d of data.documents){if(!d||!Array.isArray(d.tabs)||!d.tabs.length||typeof d.title!=='string')continue;
   const clean=normalizeImportedState({...defaultState(),...d});clean.id=uid();clean.folderId=folderMap.get(d.folderId)||null;
   if(!clean.tabs.length)continue;await save(clean);count++;
  }
  await refreshLibrary();notify(`${count} documentos importados`)
 }catch(err){console.error(err);notify('Não foi possível importar o backup: '+err.message)}
}

async function trashDocument(meta){
 if(!confirm(`Excluir “${meta.title}”? Faça backup antes; essa ação não pode ser desfeita.`))return;
 await removeDocument(meta.id);if(state.id===meta.id){const next=library.documents.find(d=>d.id!==meta.id);state=next?(await loadDocument(next.id)):defaultState();if(!next)await save(state,true);renderTabs();activateTab();syncControls()}
 await refreshLibrary();notify('Documento excluído')
}
async function renameDocumentInLibrary(meta){const title=prompt('Nome do documento:',meta.title);if(!title?.trim())return;const doc=await loadDocument(meta.id);if(!doc)return;doc.title=title.trim().slice(0,150);await save(doc);if(state.id===doc.id){state.title=doc.title;syncControls()}await refreshLibrary()}
async function duplicateDocument(id=state.id){
 if(state.id===id)await persist(true);const src=await loadDocument(id);if(!src)return;
 const copy=structuredClone(src);copy.id=uid();copy.title=(src.title||'Documento')+' · cópia';
 copy.tabs=copy.tabs.map(t=>{const ids=new Map((t.comments||[]).map(c=>[c.id,uid()]));let html=t.html||'';for(const [oldId,newId] of ids)html=html.replaceAll(`data-comment-id=\"${oldId}\"`,`data-comment-id=\"${newId}\"`);return {...t,id:uid(),html,comments:(t.comments||[]).map(c=>({...c,id:ids.get(c.id)})),assets:(t.assets||[]).map(a=>({...a,id:uid()}))}});
 copy.activeId=copy.tabs[0].id;await save(copy,true);await refreshLibrary();notify('Cópia criada');if(!$('libraryView').classList.contains('hidden'))return;await openDocument(copy.id)
}
function newFolder(){openFolderEditor(null)}
function openFolderEditor(folder){openModal(folder?'Personalizar pasta':'Nova pasta',`<section class="modal-section"><label class="library-label">Nome da pasta <input id="folderNameInput" maxlength="80" placeholder="Ex.: Roteiros" value="${escapeHTML(folder?.name||'')}"/></label><label class="library-label">Cor da pasta <input id="folderColorInput" type="color" value="${/^#[a-f\d]{6}$/i.test(folder?.color||'')?folder.color:'#b584d1'}"/></label><div class="modal-actions"><button id="saveFolderInput" class="pill-button primary">${folder?'Salvar alterações':'Criar pasta'}</button></div></section>`);
 $('saveFolderInput').onclick=async()=>{const name=$('folderNameInput').value.trim();if(!name){notify('Digite um nome para a pasta');return}await saveFolder({id:folder?.id||uid(),name:name.slice(0,80),color:$('folderColorInput').value});closeModal();await refreshLibrary();notify(folder?'Pasta atualizada':'Pasta criada')};$('folderNameInput').focus()}
async function removeFolder(folder){if(!confirm(`Excluir a pasta “${folder.name}”? Os documentos serão mantidos em “Sem pasta”.`))return;await deleteFolder(folder.id);if(libraryFolder===folder.id)libraryFolder='root';if(state.folderId===folder.id)state.folderId=null;await refreshLibrary();notify('Pasta removida. Documentos preservados.')}
function moveDocumentDialog(id=state.id){const docMeta=library.documents.find(d=>d.id===id);openModal('Mover para pasta',`<section class="modal-section"><p>Documento: <b>${escapeHTML(docMeta?.title||state.title)}</b></p><div id="moveFolderChoices" class="folder-choices"></div></section>`);
 const choices=$('moveFolderChoices');for(const [folderId,name,color] of [[null,'Sem pasta','#aeb6c4'],...library.folders.map(f=>[f.id,f.name,f.color])]){const btn=document.createElement('button');btn.className='move-folder-choice';btn.innerHTML=`<span class="folder-dot" style="background:${color}"></span>${escapeHTML(name)}`;btn.onclick=async()=>{const doc=await loadDocument(id);if(!doc)return;doc.folderId=folderId;await save(doc);if(state.id===id)state.folderId=folderId;closeModal();await refreshLibrary();notify('Documento movido')};choices.append(btn)}
}
function wireLibraryAndFeatures(){
 $('homeButton').onclick=()=>void showLibrary();$('libraryNewDocument').onclick=()=>void newDocument();$('libraryOpenDocx').onclick=()=>$('nativeImport').click();$('libraryNewFolder').onclick=newFolder;
 $('libraryMore').onclick=e=>{const r=e.currentTarget.getBoundingClientRect();openMenu([{text:'Exportar biblioteca completa',action:()=>void exportLibraryBackup()},{text:'Importar backup da biblioteca',action:()=>$('workspaceImport').click()},{text:'Abrir projeto Clicky (.clicky)',action:()=>$('nativeImport').click()}],r.left,r.bottom+6)};
 $('workspaceImport').onchange=async e=>{await importLibraryBackup(e.target.files?.[0]);e.target.value=''};
 $('librarySearch').oninput=renderLibrary;$('allDocuments').onclick=()=>setLibraryFolder('all');$('rootDocuments').onclick=()=>setLibraryFolder('root');
 $('libraryGrid').onclick=()=>void changeLayout('grid');$('libraryList').onclick=()=>void changeLayout('list');
 $('commentButton').addEventListener('mousedown',e=>e.preventDefault());$('commentButton').onclick=beginComment;
 $('commentsButton').onclick=()=>toggleComments();$('closeComments').onclick=()=>toggleComments(false);
 $('newCommentSide').onclick=beginComment;$('saveComment').onclick=commitComment;$('cancelComment').onclick=()=>{$('commentComposer').classList.add('hidden');pendingCommentRange=null};
 $('commentsList').addEventListener('click',async e=>{
  const card=e.target.closest('[data-thread]');if(!card)return;const comment=currentTab().comments.find(c=>c.id===card.dataset.thread);if(!comment)return;
  if(e.target.closest('[data-resolve]')){comment.resolved=!comment.resolved;writeEditor();renderComments();scheduleSave();return}
  if(e.target.closest('[data-delete]')){if(!confirm('Excluir comentário e suas respostas?'))return;currentTab().comments=currentTab().comments.filter(c=>c.id!==comment.id);$('editor').querySelectorAll('[data-comment-id]').forEach(el=>{if(el.dataset.commentId===comment.id)el.replaceWith(...el.childNodes)});writeEditor();renderComments();scheduleSave();return}
  if(e.target.closest('[data-reply]')){const input=card.querySelector('.reply-draft');if(!input.value.trim())return;comment.replies.push({text:input.value.trim().slice(0,10000),createdAt:new Date().toISOString()});renderComments();scheduleSave();return}
 });
}
async function changeLayout(layout){libraryLayout=layout;await setPreference('libraryLayout',layout);renderLibrary()}
function toggleComments(force){const panel=$('commentsPanel');const show=typeof force==='boolean'?force:panel.classList.contains('hidden');panel.classList.toggle('hidden',!show);if(show){closeInspector();renderComments()}}
function focusComment(id){const el=$('commentsList').querySelector(`[data-thread="${CSS.escape(id)}"]`);if(el){el.scrollIntoView({block:'nearest'});el.classList.add('focused-comment');setTimeout(()=>el.classList.remove('focused-comment'),1500)}}
function beginComment(){if(!$('libraryView').classList.contains('hidden'))return;const range=getValidSelection()||(restoreSelection()?getValidSelection():null);pendingCommentRange=range?.cloneRange()||null;pendingCommentQuote=range?.toString().trim().slice(0,250)||'';toggleComments(true);$('commentComposer').classList.remove('hidden');$('commentQuote').textContent=pendingCommentQuote?`“${pendingCommentQuote}”`:'Comentário geral desta guia';$('commentDraft').value='';$('commentDraft').focus()}
function commitComment(){const text=$('commentDraft').value.trim();if(!text){notify('Escreva o comentário');return}const id=uid();let quote=pendingCommentQuote;
 if(pendingCommentRange){try{const r=pendingCommentRange;if($('editor').contains(r.commonAncestorContainer)&&!r.collapsed){const span=document.createElement('span');span.className='comment-anchor';span.dataset.commentId=id;span.append(r.extractContents());r.insertNode(span);quote=span.textContent.slice(0,250);writeEditor()}}catch(err){console.warn('Âncora não disponível',err)}}
 currentTab().comments ||= [];currentTab().comments.push({id,quote,text:text.slice(0,10000),createdAt:new Date().toISOString(),resolved:false,replies:[]});pendingCommentRange=null;$('commentComposer').classList.add('hidden');renderComments();scheduleSave();notify('Comentário adicionado')
}
function renderComments(){const list=$('commentsList');if(!list)return;list.innerHTML='';for(const c of currentTab().comments||[]){const card=document.createElement('article');card.className='comment-card'+(c.resolved?' is-resolved':'');card.dataset.thread=c.id;
 const replies=(c.replies||[]).map(r=>`<div class="comment-reply"><span>Resposta · ${new Date(r.createdAt).toLocaleDateString('pt-BR')}</span><p>${escapeHTML(r.text)}</p></div>`).join('');
 card.innerHTML=`<div class="comment-author"><span class="comment-avatar">C</span><strong>Você</strong><small>${new Date(c.createdAt).toLocaleDateString('pt-BR')}</small></div>${c.quote?`<div class="comment-quote">${escapeHTML(c.quote)}</div>`:''}<p class="comment-text">${escapeHTML(c.text)}</p>${replies}<div class="comment-tools"><button data-resolve>${c.resolved?'Reabrir':'Resolver'}</button><button data-delete>Excluir</button></div>${c.resolved?'':`<div class="reply-wrap"><input class="reply-draft" placeholder="Responder..." maxlength="10000"/><button data-reply>↗</button></div>`}`;
 list.append(card)}if(!currentTab().comments?.length)list.innerHTML='<div class="comment-empty">Os comentários desta guia aparecerão aqui. Selecione um trecho e pressione Ctrl+Alt+M para começar.</div>';}
const supportedLanguages={'pt-BR':'Português (Brasil)','en-US':'English (US)','es-ES':'Español','ja-JP':'日本語'};
function openProofing(){openModal('Ortografia e idioma',`<section class="modal-section"><h3>Idioma do documento</h3><p>A verificação ortográfica utiliza os recursos do sistema/WebView. A disponibilidade de correções e sugestões varia conforme o Windows ou macOS.</p><label class="library-label">Idioma <select id="proofLanguage">${Object.entries(supportedLanguages).map(([code,label])=>`<option value="${code}" ${state.language===code?'selected':''}>${label}</option>`).join('')}</select></label><label class="inspect-row">Verificação ortográfica do sistema <input id="proofToggle" type="checkbox" ${$('editor').spellcheck?'checked':''}/></label><div class="modal-actions"><button id="reviewButton" class="pill-button primary">Revisar texto</button><button id="dictionaryButton" class="pill-button">Dicionário pessoal</button></div></section>`);
 $('proofLanguage').onchange=e=>{state.language=e.target.value;$('editor').lang=state.language;scheduleSave()};$('proofToggle').onchange=e=>{state.spellcheckEnabled=e.target.checked;$('editor').spellcheck=e.target.checked;scheduleSave();notify(e.target.checked?'Verificação do sistema ativada':'Verificação do sistema desativada')};$('reviewButton').onclick=openReview;$('dictionaryButton').onclick=openDictionary;}
function getReviewHints(){if(state.language!=='pt-BR')return [];const value=$('editor').innerText||'';const hints=[];const patterns=[[/\bconcerteza\b/giu,'com certeza','Ortografia'],[/\bderrepente\b/giu,'de repente','Ortografia'],[/\bexcessão\b/giu,'exceção','Ortografia'],[/\bmenas\b/giu,'menos','Concordância'],[/\bseje\b/giu,'seja','Conjugação']];
 for(const [reg,fix,kind] of patterns){for(const found of value.matchAll(reg)){if(dictionary.some(x=>x.toLocaleLowerCase()===found[0].toLocaleLowerCase()))continue;hints.push({original:found[0],replacement:fix,kind})}}
 for(const found of value.matchAll(/\b([\p{L}]{2,})\s+\1\b/giu))hints.push({original:found[0],replacement:found[1],kind:'Repetição'});
 return hints.slice(0,50)}
function replaceFirstInEditor(original,replacement){const walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode;const position=n.data.toLocaleLowerCase().indexOf(original.toLocaleLowerCase());if(position<0)continue;n.replaceData(position,original.length,replacement);writeEditor();updateWordCount();scheduleSave();return true}return false}
function openReview(){const hints=getReviewHints();openModal('Revisão rápida',`<section class="modal-section"><p>Esta revisão complementar verifica alguns erros comuns, repetições e concordâncias simples. A revisão completa depende do corretor ortográfico do sistema; não substitui uma revisão gramatical avançada.</p><h3>${hints.length} sugest${hints.length===1?'ão':'ões'} encontrada${hints.length===1?'':'s'}</h3><div id="reviewSuggestions" class="review-list"></div><div class="modal-actions"><button id="proofPrefs" class="pill-button">Idioma e ortografia</button></div></section>`);
 const list=$('reviewSuggestions');if(!hints.length)list.innerHTML='<p>Não foram encontrados erros entre os padrões cobertos pela revisão rápida.</p>';
 for(const hint of hints){const row=document.createElement('div');row.className='review-item';row.innerHTML=`<div><small>${escapeHTML(hint.kind)}</small><strong>${escapeHTML(hint.original)} → ${escapeHTML(hint.replacement)}</strong></div><button class="pill-button">Corrigir</button>`;row.querySelector('button').onclick=()=>{if(replaceFirstInEditor(hint.original,hint.replacement))openReview();else notify('Trecho já foi alterado')};list.append(row)}$('proofPrefs').onclick=openProofing;}
function openDictionary(){const selected=(getValidSelection()?.toString()||activeSelection?.toString()||'').trim().split(/\s+/).slice(0,1)[0]||'';
 openModal('Dicionário',`<section class="modal-section"><h3>Consultar uma palavra</h3><p>A consulta de definições abre o Wikcionário no navegador e requer internet. Nenhum texto é enviado automaticamente.</p><div class="dictionary-search"><input id="dictionaryQuery" maxlength="100" value="${escapeHTML(selected)}" placeholder="Digite uma palavra"/><button id="dictionaryLookup" class="pill-button primary">Consultar</button></div></section><section class="modal-section"><h3>Meu dicionário pessoal</h3><p>Armazena palavras aceitas no revisor rápido do Clicky. Ainda não modifica o corretor nativo do sistema.</p><div class="dictionary-search"><input id="dictionaryAdd" maxlength="80" placeholder="Adicionar termo"/><button id="dictionarySave" class="pill-button">Adicionar</button></div><div class="custom-words" id="dictionaryWords"></div></section>`);
 $('dictionaryLookup').onclick=()=>{const w=$('dictionaryQuery').value.trim();if(!w)return;window.open('https://pt.wiktionary.org/wiki/'+encodeURIComponent(w),'_blank','noopener,noreferrer')};
 $('dictionarySave').onclick=async()=>{const val=$('dictionaryAdd').value.trim();if(!val||val.length>80||dictionary.includes(val))return;dictionary.push(val);dictionary.sort((a,b)=>a.localeCompare(b,'pt-BR'));await setPreference('customDictionary',dictionary);openDictionary()};
 const container=$('dictionaryWords');for(const word of dictionary){const chip=document.createElement('button');chip.className='word-chip';chip.textContent=word+' ×';chip.title='Remover do dicionário';chip.onclick=async()=>{dictionary=dictionary.filter(w=>w!==word);await setPreference('customDictionary',dictionary);openDictionary()};container.append(chip)}if(!dictionary.length)container.textContent='Nenhuma palavra adicionada.';
}
function openWordCount(){const value=$('editor').innerText||'';const words=value.trim()?value.trim().split(/\s+/).length:0;openModal('Contagem de palavras',`<section class="modal-section"><div class="word-stats"><div><strong>${words}</strong><span>Palavras</span></div><div><strong>${value.length}</strong><span>Caracteres</span></div><div><strong>${value.replace(/\s/g,'').length}</strong><span>Sem espaços</span></div></div></section>`)}
function openFindReplace(){openModal('Localizar e substituir',`<section class="modal-section"><div class="form-grid"><label>Localizar <input id="findQuery" placeholder="Texto a procurar"/></label><label>Substituir por <input id="replaceQuery" placeholder="Novo texto"/></label></div><div class="modal-actions"><button class="pill-button" id="findNext">Localizar</button><button class="pill-button primary" id="replaceAll">Substituir todas</button></div><p id="findCount" class="caption-muted">A pesquisa considera o texto da guia ativa.</p></section>`);
 $('findNext').onclick=()=>{const needle=$('findQuery').value;if(!needle)return;const walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT);while(walker.nextNode()){const node=walker.currentNode;const start=node.textContent.toLocaleLowerCase().indexOf(needle.toLocaleLowerCase());if(start<0)continue;const range=document.createRange();range.setStart(node,start);range.setEnd(node,start+needle.length);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(range);node.parentElement?.scrollIntoView({block:'center'});$('findCount').textContent='Primeira ocorrência selecionada';return}$('findCount').textContent='Nenhuma ocorrência encontrada'};
 $('replaceAll').onclick=()=>{const needle=$('findQuery').value;if(!needle)return;const replacement=$('replaceQuery').value;let count=0;const walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT);while(walker.nextNode()){const node=walker.currentNode;const parts=node.data.split(new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'));if(parts.length>1){count+=parts.length-1;node.data=parts.join(replacement)}}writeEditor();updateWordCount();scheduleSave();$('findCount').textContent=`${count} ocorrência${count===1?'':'s'} substituída${count===1?'':'s'}`};$('findQuery').focus();}
const shortcuts=[
 ['Negrito','Mod+B'],['Itálico','Mod+I'],['Sublinhado','Mod+U'],['Desfazer','Mod+Z'],['Refazer','Mod+Shift+Z'],['Selecionar tudo','Mod+A'],['Copiar / recortar / colar','Mod+C / X / V'],['Colar sem formatação','Mod+Shift+V'],['Link','Mod+K'],['Imprimir / PDF','Mod+P'],['Abrir documento Word ou projeto','Mod+O'],['Salvar','Mod+S'],['Salvar cópia editável','Mod+Shift+S'],['Novo documento','Mod+N'],['Comentário','Mod+Alt+M'],['Painel de comentários','Mod+Alt+Shift+A'],['Localizar','Mod+F'],['Localizar e substituir','Mod+H'],['Contagem de palavras','Mod+Shift+C'],['Alinhar à esquerda','Mod+Shift+L'],['Centralizar','Mod+Shift+E'],['Alinhar à direita','Mod+Shift+R'],['Justificar','Mod+Shift+J'],['Lista numerada','Mod+Shift+7'],['Marcadores','Mod+Shift+8'],['Título 1–3','Mod+Alt+1 a 3'],['Limpar formatação','Mod+\\'],['Aumentar / diminuir fonte','Mod+Shift+. / ,'],['Dicionário','Mod+Shift+Y'],['Nova guia','Shift+F11'],['Refazer (Windows)','Mod+Y'],['Guias 1–9','Mod+1 a 9'],['Pesquisar nos menus','Alt+/'],['Lista de atalhos','Mod+/'] ];
function openShortcutHelp(){openModal('Atalhos de teclado',`<section class="modal-section"><p>Usamos Ctrl no Windows e ⌘ no macOS. Alt corresponde a Option no Mac. Apenas os atalhos abaixo estão implementados no Clicky. Outros do Google Docs dependerão das funções correspondentes.</p><div id="shortcutTable" class="shortcuts-table">${shortcuts.map(([label,keys])=>`<div><span>${escapeHTML(label)}</span><kbd>${escapeHTML(keys.replaceAll('Mod',navigator.platform?.includes('Mac')?'⌘':'Ctrl').replaceAll('Alt',navigator.platform?.includes('Mac')?'Option':'Alt'))}</kbd></div>`).join('')}</div></section>`)}
function openCommandSearch(){openModal('Pesquisar nos menus',`<section class="modal-section"><input id="commandQuery" class="font-search" placeholder="Pesquisar um comando..." autofocus/><div id="commandResults" class="command-results"></div></section>`);const commands=[['Novo documento',()=>void newDocument()],['Comentários',toggleComments],['Adicionar comentário',beginComment],['Revisão rápida',openReview],['Dicionário',openDictionary],['Pastas e documentos',()=>void showLibrary()],['Salvar como Word',()=>void saveAs('docx')],['Localizar e substituir',openFindReplace],['Temas',openThemes],['Biblioteca de fontes',openFonts],['Atalhos de teclado',openShortcutHelp]];
 function draw(){const list=$('commandResults');list.innerHTML='';const q=$('commandQuery').value.toLocaleLowerCase();for(const [name,action] of commands.filter(c=>c[0].toLocaleLowerCase().includes(q))){const btn=document.createElement('button');btn.textContent=name;btn.onclick=()=>{closeModal();action()};list.append(btn)}}$('commandQuery').oninput=draw;draw();$('commandQuery').focus()}

async function saveAs(format='docx'){
 writeEditor();const name=safeFilename();let blob,ext;
 if(format==='docx'){blob=createDocx(state);ext='docx'}
 else if(format==='txt'){const all=state.tabs.map(t=>{const div=document.createElement('div');div.innerHTML=t.html;return t.name+'\n'+div.textContent}).join('\n\n');blob=new Blob([all],{type:'text/plain;charset=utf-8'});ext='txt'}
 else{blob=new Blob([JSON.stringify({...state,exportedAt:new Date().toISOString()},null,2)],{type:'application/json'});ext='clicky'}
 const fileName=`${name}.${ext}`;
 try{
  const t=globalThis.__TAURI__?.core;
  if(t?.invoke){const bytes=[...new Uint8Array(await blob.arrayBuffer())];const result=await t.invoke('save_file_dialog',{filename:fileName,bytes});if(result)notify(`Salvo em ${result}`);return}
  if(window.showSaveFilePicker){const handle=await window.showSaveFilePicker({suggestedName:fileName,types:[{description:ext.toUpperCase(),accept:{'application/octet-stream':['.'+ext]}}]});const output=await handle.createWritable();await output.write(blob);await output.close();notify('Arquivo salvo na pasta escolhida');return}
  download(blob,fileName);notify('Arquivo baixado. Para escolher a pasta, configure os downloads do navegador.');
 }catch(err){if(err?.name==='AbortError'||/cancelled|canceled/i.test(String(err)))return;console.error(err);notify('Não foi possível salvar: '+String(err).slice(0,140))}
}
async function importWord(file){if(!file)return;try{
 const result=await readDocx(file);
 await persist(true);
 const next=defaultState();next.title=result.title;next.folderId=libraryFolder==='all'||libraryFolder==='root'?null:libraryFolder;
 next.tabs=[{id:'original',name:'Documento',html:sanitizeHtml(result.html),assets:result.assets.map(a=>({...a,id:uid()})),comments:[]}];next.activeId='original';
 state=next;setTheme(state.theme);renderTabs();activateTab();syncControls();closeInspector();await save(state,true);await refreshLibrary();showEditor();
 notify(result.warning||'Documento Word aberto. Salve como .docx ou .clicky quando terminar.');
 }catch(err){console.error(err);notify(`Não foi possível abrir o DOCX: ${err.message}`)}
}
function handleKeydown(e){
 if(e.shiftKey&&e.key==='F11'){e.preventDefault();if($('libraryView').classList.contains('hidden'))newTab();return}
 const mod=e.ctrlKey||e.metaKey,k=e.key.toLowerCase(),code=e.code,libOpen=!$('libraryView').classList.contains('hidden');
 if(e.key==='Escape'){hideMenus();closeModal();if(!$('commentComposer').classList.contains('hidden'))$('commentComposer').classList.add('hidden');return}
 if(e.altKey&&!mod&&(k==='/'||code==='Slash')){e.preventDefault();openCommandSearch();return}
 if(e.altKey&&!mod&&!e.shiftKey&&'fevioth'.includes(k)){const mapping={f:'file',e:'edit',v:'view',i:'insert',o:'format',t:'tools',h:'help'};const button=document.querySelector(`[data-menu=\"${mapping[k]}\"]`);if(button&&$('libraryView').classList.contains('hidden')){e.preventDefault();button.click();return}}
 if(!mod)return;
 if(k==='/'||code==='Slash'&&!e.shiftKey){e.preventDefault();openShortcutHelp();return}
 if(k==='n'&&!e.shiftKey){e.preventDefault();void newDocument();return}
 if(k==='o'&&!e.shiftKey){e.preventDefault();$('nativeImport').click();return}
 if(libOpen)return;
 if(e.altKey&&k==='m'&&!e.shiftKey){e.preventDefault();beginComment();return}
 if(e.altKey&&e.shiftKey&&k==='a'){e.preventDefault();toggleComments();return}
 if(e.altKey&&!e.shiftKey&&/^[1-3]$/.test(k)){e.preventDefault();editorCommand('formatBlock','h'+k);return}
 if(k==='s'){e.preventDefault();if(e.shiftKey)void saveAs('clicky');else void saveAs('docx');return}
 if(k==='p'){e.preventDefault();window.print();return}
 if(k==='k'&&!e.altKey){e.preventDefault();const url=prompt('URL do link:','https://');if(url)editorCommand('createLink',url);return}
 if(k==='h'){e.preventDefault();openFindReplace();return}
 if(k==='f'){e.preventDefault();openFindReplace();return}
 if(k==='c'&&e.shiftKey){e.preventDefault();openWordCount();return}
 if(k==='y'&&!e.shiftKey&&!e.altKey){e.preventDefault();editorCommand('redo');return}
 if(k==='y'&&e.shiftKey){e.preventDefault();openDictionary();return}
 if(e.shiftKey&&!e.altKey){
  const alignment={l:'justifyLeft',e:'justifyCenter',r:'justifyRight',j:'justifyFull'}[k];if(alignment){e.preventDefault();editorCommand(alignment);return}
  if(code==='Digit7'){e.preventDefault();editorCommand('insertOrderedList');return}
  if(code==='Digit8'){e.preventDefault();editorCommand('insertUnorderedList');return}
  if(code==='Period'){e.preventDefault();setSize((+state.fontSize||12)+1);return}
  if(code==='Comma'){e.preventDefault();setSize((+state.fontSize||12)-1);return}
  if(k==='z'){e.preventDefault();editorCommand('redo');return}
 }
 if(k==='\\'||code==='Backslash'&&!e.shiftKey){e.preventDefault();editorCommand('removeFormat');return}
 if(e.altKey||e.shiftKey)return;
 if('123456789'.includes(k)){e.preventDefault();const tab=state.tabs[+k-1];if(tab)switchTab(tab.id)}
}

async function start(){
 wire();
 try{
  await initializeLibrary(defaultState);
  await refreshLibrary();
  dictionary=await getPreference('customDictionary',[]);
  libraryLayout=await getPreference('libraryLayout','grid');
  const first=library.documents[0];
  if(first){const stored=await loadDocument(first.id);if(stored)state=normalizeImportedState({...defaultState(),...stored});}
 }catch(err){console.error(err);notify('Não foi possível inicializar a biblioteca local')}
 state.tabs=state.tabs.map(t=>({...t,comments:Array.isArray(t.comments)?t.comments:[],assets:Array.isArray(t.assets)?t.assets:[]}));
 if(!state.tabs.some(t=>t.id===state.activeId))state.activeId=state.tabs[0].id;
 installFontFaces();setTheme(state.theme);renderTabs();activateTab();syncControls();appReady=true;setSaveState('Salvo');
 renderLibrary();$('libraryView').classList.remove('hidden');$('app').classList.add('hidden');
}
void start();