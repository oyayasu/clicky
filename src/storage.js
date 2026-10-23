/* Clicky v0.2 – offline-first IndexedDB with migration of legacy single-document workspaces. */
const NAME = 'clicky-workspace';
const VERSION = 2;
let dbPromise;
function database() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(NAME, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      for (const name of ['workspace', 'revisions', 'documents', 'document_index', 'folders', 'preferences']) {
        if (!db.objectStoreNames.contains(name)) db.createObjectStore(name, name === 'revisions' ? { keyPath: 'id' } : undefined);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}
function requestResult(req) { return new Promise((resolve, reject) => { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); }); }
function txDone(tx) { return new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onabort = () => reject(tx.error || new Error('Transação cancelada')); tx.onerror = () => reject(tx.error); }); }
export async function load() { const db = await database(); return requestResult(db.transaction('workspace','readonly').objectStore('workspace').get('main')); }
export async function initializeLibrary(fallback) {
  const db = await database();
  const prefs = await requestResult(db.transaction('preferences','readonly').objectStore('preferences').get('seeded'));
  if (prefs) return;
  const old = await load();
  const first = old?.tabs?.length ? structuredClone(old) : fallback();
  first.id ||= crypto.randomUUID();
  first.folderId ||= null;
  const now = new Date().toISOString();
  const tx = db.transaction(['documents','document_index','preferences'], 'readwrite');
  tx.objectStore('documents').put(first,first.id);
  tx.objectStore('document_index').put({id:first.id,title:first.title,folderId:first.folderId,updatedAt:now,createdAt:now},first.id);
  tx.objectStore('preferences').put(true,'seeded');
  await txDone(tx);
}
export async function listLibrary() {
  const db=await database();const tx=db.transaction(['document_index','folders'],'readonly');
  const [documents,folders]=await Promise.all([requestResult(tx.objectStore('document_index').getAll()),requestResult(tx.objectStore('folders').getAll())]);
  return {documents,folders};
}
export async function loadDocument(id) { const db=await database(); return requestResult(db.transaction('documents','readonly').objectStore('documents').get(id)); }
export async function save(workspace,snapshot=false) {
  if (!workspace.id) throw new Error('Documento sem identificador');
  const db=await database();const tx=db.transaction(['documents','document_index','revisions'],'readwrite');
  const now=new Date().toISOString();const snapshotCopy=structuredClone(workspace);
  tx.objectStore('documents').put(snapshotCopy,workspace.id);
  const previous=await requestResult(tx.objectStore('document_index').get(workspace.id));
  tx.objectStore('document_index').put({id:workspace.id,title:workspace.title||'Sem título',folderId:workspace.folderId||null,createdAt:previous?.createdAt||now,updatedAt:now},workspace.id);
  if(snapshot)tx.objectStore('revisions').put({id:`${workspace.id}:${Date.now()}`,documentId:workspace.id,createdAt:now,state:snapshotCopy});
  await txDone(tx);
}
export async function saveFolder(folder) { const db=await database();const tx=db.transaction('folders','readwrite');tx.objectStore('folders').put(structuredClone(folder),folder.id);await txDone(tx); }
export async function deleteFolder(id) {
  const db=await database();const tx=db.transaction(['folders','document_index','documents'],'readwrite');
  tx.objectStore('folders').delete(id);
  const documents=await requestResult(tx.objectStore('document_index').getAll());
  for (const meta of documents.filter(m=>m.folderId===id)) {
    meta.folderId=null;tx.objectStore('document_index').put(meta,meta.id);
    const doc=await requestResult(tx.objectStore('documents').get(meta.id));
    if (doc) { doc.folderId=null;tx.objectStore('documents').put(doc,doc.id); }
  }
  await txDone(tx);
}
export async function removeDocument(id) {const db=await database();const tx=db.transaction(['documents','document_index'],'readwrite');tx.objectStore('documents').delete(id);tx.objectStore('document_index').delete(id);await txDone(tx);}
export async function revisions(documentId) {
  const db=await database();const all=await requestResult(db.transaction('revisions','readonly').objectStore('revisions').getAll());
  return all.filter(r=>!documentId||r.documentId===documentId||(!r.documentId&&r.state?.id===documentId)).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
}
export async function cleanupRevisions(max=15,documentId) {
  const db=await database();const all=await revisions(documentId);if(all.length<=max)return;
  const tx=db.transaction('revisions','readwrite');all.slice(max).forEach(r=>tx.objectStore('revisions').delete(r.id));await txDone(tx);
}
export async function getPreference(key,fallback=null){const db=await database();return (await requestResult(db.transaction('preferences','readonly').objectStore('preferences').get(key)))??fallback;}
export async function setPreference(key,value){const db=await database();const tx=db.transaction('preferences','readwrite');tx.objectStore('preferences').put(structuredClone(value),key);await txDone(tx);}
