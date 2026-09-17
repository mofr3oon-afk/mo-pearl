// Optional local-only image storage: no uploads; resilient when IndexedDB is unavailable.
const DB = 'invoice-maker-v13-draft';
const STORE = 'documents';
function openDb() {
  return new Promise((resolve,reject)=>{
    if (!('indexedDB' in window)) return reject(new Error('IndexedDB unavailable'));
    const request=indexedDB.open(DB,1);
    request.onupgradeneeded=()=>request.result.createObjectStore(STORE);
    request.onerror=()=>reject(request.error);
    request.onsuccess=()=>resolve(request.result);
  });
}
export async function readDraftMedia(){
  const db=await openDb();
  try{return await new Promise((resolve,reject)=>{const tx=db.transaction(STORE,'readonly');const req=tx.objectStore(STORE).get('current');req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error);});}
  finally{db.close();}
}
export async function writeDraftMedia(media){
  const db=await openDb();
  try{return await new Promise((resolve,reject)=>{const tx=db.transaction(STORE,'readwrite');tx.objectStore(STORE).put(media,'current');tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);});}
  finally{db.close();}
}
