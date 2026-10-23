/** Local .docx reader; no uploads or remote requests. Supports common WordprocessingML. */
const utf8=new TextDecoder('utf-8');
const encode=(s)=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const children=(el,name)=>[...el?.children||[]].filter(x=>x.localName===name);
const first=(el,name)=>children(el,name)[0];
const attr=(el,name)=>el?.getAttribute('w:'+name)||el?.getAttribute(name)||'';
function readU16(v,o){return v.getUint16(o,true)}
function readU32(v,o){return v.getUint32(o,true)}
function decodeXml(bytes){return new DOMParser().parseFromString(utf8.decode(bytes),'application/xml')}
/** Reads ZIP central directory and rejects traversal, oversized entries and unsupported compression. */
export async function readZip(file){
 const buf=await file.arrayBuffer(),view=new DataView(buf),bytes=new Uint8Array(buf),end=buf.byteLength;
 if(end>80*1024*1024)throw Error('Documento maior que 80 MB.');
 let eocd=-1;for(let i=end-22;i>=Math.max(0,end-65557);i--){if(readU32(view,i)===0x06054b50){eocd=i;break}}
 if(eocd<0)throw Error('Arquivo DOCX inválido (ZIP).');
 const count=readU16(view,eocd+10),offset=readU32(view,eocd+16);
 if(count>3500)throw Error('Documento com muitas entradas.');
 const zip=new Map();let p=offset,total=0;
 for(let i=0;i<count;i++){
  if(p+46>end||readU32(view,p)!==0x02014b50)throw Error('Índice ZIP corrompido.');
  const compression=readU16(view,p+10),compressed=readU32(view,p+20),uncompressed=readU32(view,p+24),nameLen=readU16(view,p+28),extra=readU16(view,p+30),comment=readU16(view,p+32),local=readU32(view,p+42);
  const name=utf8.decode(bytes.subarray(p+46,p+46+nameLen));p+=46+nameLen+extra+comment;
  if(name.startsWith('/')||name.split('/').includes('..'))continue;
  total+=uncompressed;if(total>130*1024*1024||uncompressed>30*1024*1024)throw Error('Documento descompactado muito grande.');
  if(local+30>end||readU32(view,local)!==0x04034b50)throw Error('Entrada ZIP inválida.');
  const start=local+30+readU16(view,local+26)+readU16(view,local+28);
  if(start+compressed>end)throw Error('Dados ZIP truncados.');
  zip.set(name,{compression,compressed,uncompressed,start});
 }
 async function get(name){const f=zip.get(name);if(!f)return null;const b=bytes.slice(f.start,f.start+f.compressed);if(f.compression===0)return b;if(f.compression!==8)throw Error('Método de compressão não suportado.');
  if(!globalThis.DecompressionStream)throw Error('Este sistema não suporta leitura ZIP local.');
  const stream=new Blob([b]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const decompressed=new Uint8Array(await new Response(stream).arrayBuffer());if(decompressed.length!==f.uncompressed)throw Error('Falha na verificação do tamanho descompactado.');return decompressed;
 }
 return {get,has:n=>zip.has(n)};
}
function parseStyles(doc){const map=new Map();for(const s of doc?.getElementsByTagNameNS('*','style')||[]){const id=attr(s,'styleId');const n=first(s,'name');const type=attr(s,'type');map.set(id,{name:attr(n,'val')||id,type});}return map}
function runToHtml(r){
 let txt='';for(const node of r.children){if(node.localName==='t'||node.localName==='delText')txt+=encode(node.textContent);else if(node.localName==='tab')txt+='&emsp;';else if(node.localName==='br')txt+='<br>';}
 if(!txt)return '';
 const props=first(r,'rPr'),css=[];let start='',end='';
 const color=attr(first(props,'color'),'val');if(/^[A-Fa-f0-9]{6}$/.test(color))css.push(`color:#${color}`);
 const size=Number(attr(first(props,'sz'),'val'));if(size>=12&&size<=192)css.push(`font-size:${size/2}pt`);
 const family=attr(first(props,'rFonts'),'ascii');if(/^[\w\s-]{2,60}$/.test(family))css.push(`font-family:'${family}'`);
 const hl=attr(first(props,'highlight'),'val');const colors={yellow:'#ffff00',green:'#00ff00',cyan:'#00ffff',magenta:'#ff00ff',red:'#ff0000',blue:'#0000ff',lightGray:'#d3d3d3',darkYellow:'#d3a500'};if(colors[hl])css.push(`background-color:${colors[hl]}`);
 if(first(props,'b')&&attr(first(props,'b'),'val')!=='0'){start+='<strong>';end='</strong>'+end}
 if(first(props,'i')&&attr(first(props,'i'),'val')!=='0'){start+='<em>';end='</em>'+end}
 if(first(props,'u')){start+='<u>';end='</u>'+end}
 if(first(props,'strike')){start+='<s>';end='</s>'+end}
 return css.length?`<span style="${css.join(';')}">${start}${txt}${end}</span>`:start+txt+end;
}
function paragraph(node,styles){const pPr=first(node,'pPr');const style=attr(first(pPr,'pStyle'),'val');const styleName=styles.get(style)?.name||style;
 let tag='p';const heading=styleName.match(/heading\s*([1-6])/i);if(heading)tag='h'+heading[1];else if(/title/i.test(styleName))tag='h1';
 const align=attr(first(pPr,'jc'),'val'),css=[];if(['center','right','both','left'].includes(align))css.push(`text-align:${align==='both'?'justify':align}`);
 const indent=Number(attr(first(pPr,'ind'),'left'));if(indent>0&&indent<8000)css.push(`margin-left:${Math.round(indent/15)}px`);
 const parts=[];for(const ch of node.children){if(ch.localName==='r')parts.push(runToHtml(ch));else if(ch.localName==='hyperlink')parts.push(children(ch,'r').map(runToHtml).join(''));else if(ch.localName==='ins')parts.push(children(ch,'r').map(runToHtml).join(''));}
 const inside=parts.join('')||'<br>';return `<${tag}${css.length?` style="${css.join(';')}"`:''}>${inside}</${tag}>`;
}
function tableToHtml(table,styles){const rows=[];for(const row of children(table,'tr')){const cells=children(row,'tc').map(cell=>`<td>${children(cell,'p').map(p=>paragraph(p,styles)).join('')||'<p><br></p>'}</td>`);rows.push(`<tr>${cells.join('')}</tr>`)}return `<table><tbody>${rows.join('')}</tbody></table>`}
export async function readDocx(file){
 if(!/\.docx$/i.test(file.name))throw Error('Selecione um arquivo .docx');
 const zip=await readZip(file),docData=await zip.get('word/document.xml');if(!docData)throw Error('Arquivo Word sem document.xml.');
 const doc=decodeXml(docData);if(doc.getElementsByTagName('parsererror').length)throw Error('XML Word inválido.');
 const styles=zip.has('word/styles.xml')?parseStyles(decodeXml(await zip.get('word/styles.xml'))):new Map();
 const body=doc.getElementsByTagNameNS('*','body')[0];if(!body)throw Error('Conteúdo Word não encontrado.');
 let html='';for(const el of body.children){if(el.localName==='p')html+=paragraph(el,styles);else if(el.localName==='tbl')html+=tableToHtml(el,styles)}
 // Pictures are extracted separately into free-floating attachments. Word's exact anchors are not yet retained.
 const assets=[];let count=0;
 for(const path of ['png','jpeg','jpg','gif','webp']){
  for(let n=1;n<=200;n++){const names=[`word/media/image${n}.${path}`,`word/media/image${n+1}.${path}`];let found=null;for(const name of names){if(zip.has(name)){found=name;break}}if(!found)continue;
   if(assets.some(a=>a.path===found))continue;const data=await zip.get(found);if(!data||data.length>12*1024*1024)continue;
   let binary='';for(let j=0;j<data.length;j+=8192)binary+=String.fromCharCode(...data.subarray(j,j+8192));
   const mime=path==='jpg'?'jpeg':path;
   assets.push({path:found,data:`data:image/${mime};base64,${btoa(binary)}`,x:90+count%3*30,y:160+count*180,width:280,height:200,locked:false});count++;if(count>=60)break;
  }
 }
 return {title:file.name.replace(/\.docx$/i,''),html:html||'<p><br></p>',assets,warning:assets.length?'Imagens extraídas; a posição original não é preservada.':''};
}