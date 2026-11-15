// Sổ điểm cá nhân – chạy được khi mất mạng
const C='sdcn-v1';
const CORE=['./','index.html','manifest.webmanifest','icon.svg',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET')return;
  if(u.origin===location.origin){
    if(u.pathname.endsWith('vnedu.js'))return; // nút dấu trang luôn lấy bản mới
    // có mạng: lấy bản mới rồi cất; mất mạng: dùng bản đã cất
    const key=u.origin+u.pathname; // bỏ ?k=… để không cất trùng
    e.respondWith(fetch(e.request).then(r=>{if(r.ok&&e.request.mode!=='no-cors'){const cp=r.clone();caches.open(C).then(c=>c.put(key,cp))}return r})
      .catch(()=>caches.match(key).then(r=>r||caches.match('index.html'))));
  }else if(u.hostname==='cdnjs.cloudflare.com'){
    e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(r2=>{const cp=r2.clone();caches.open(C).then(c=>c.put(e.request,cp));return r2})));
  }
});
