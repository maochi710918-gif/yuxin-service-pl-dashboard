// 產生加密版網頁：node build_secure.mjs <輸出路徑>
// 密碼從環境變數 WEB_PASSWORD 讀取（不寫進任何檔案）。
import { readFileSync, writeFileSync } from 'node:fs';
const { subtle } = globalThis.crypto;

const PW = process.env.WEB_PASSWORD;
if (!PW) { console.error('請先設定環境變數 WEB_PASSWORD'); process.exit(1); }
const out = process.argv[2] || 'index.secure.html';
const ITER = 600000;

const html = readFileSync(new URL('./app.html', import.meta.url), 'utf8');
const data = readFileSync(new URL('./data.json', import.meta.url), 'utf8');
const start = html.search(/<script>\r?\nconst D = /);
const end = html.indexOf('</script>', start);
if (start < 0 || end < 0) throw new Error('找不到主程式區塊');
const appJs = html.slice(start + '<script>'.length, end).replace('/*__DATA__*/null', data);

const salt = crypto.getRandomValues(new Uint8Array(16));
const iv = crypto.getRandomValues(new Uint8Array(12));
const base = await subtle.importKey('raw', new TextEncoder().encode(PW), 'PBKDF2', false, ['deriveKey']);
const key = await subtle.deriveKey({ name: 'PBKDF2', salt, iterations: ITER, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(appJs)));
const b64 = u => Buffer.from(u).toString('base64');

const lock = `<div id="lock" style="position:fixed;inset:0;z-index:999;background:var(--bg);display:flex;align-items:center;justify-content:center;padding:16px">
 <form id="lockf" class="card" style="width:min(360px,100%);text-align:center">
  <div style="font-size:30px">🔒</div><h2 style="margin:6px 0 2px">服務部損益經營戰情</h2>
  <div class="muted" style="margin-bottom:14px">內部資料，請輸入密碼</div>
  <input id="lockpw" type="password" autocomplete="current-password" placeholder="密碼" style="width:100%;border:1px solid var(--line-2);border-radius:10px;padding:10px 12px;background:var(--surface);font-size:15px">
  <button class="btn pri" style="width:100%;margin-top:10px;padding:10px">進入</button>
  <div id="lockmsg" class="bad-t" style="margin-top:8px;font-size:13px;min-height:18px"></div>
 </form></div>
<script>
(()=>{const P={s:"${b64(salt)}",i:"${b64(iv)}",c:"${b64(ct)}",n:${ITER}};
const u=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
async function run(pw){const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt:u(P.s),iterations:P.n,hash:'SHA-256'},await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']),{name:'AES-GCM',length:256},false,['decrypt']);
 const js=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:u(P.i)},k,u(P.c)));
 try{sessionStorage.setItem('svcdash-pw',pw)}catch(e){}
 document.getElementById('lock').remove(); const s=document.createElement('script'); s.textContent=js; document.body.appendChild(s);}
const f=document.getElementById('lockf'), m=document.getElementById('lockmsg');
f.onsubmit=async e=>{e.preventDefault(); m.textContent='驗證中…'; try{await run(document.getElementById('lockpw').value);}catch(err){m.textContent='密碼錯誤';}};
let saved=null; try{saved=sessionStorage.getItem('svcdash-pw')}catch(e){}
if(saved) run(saved).catch(()=>{try{sessionStorage.removeItem('svcdash-pw')}catch(e){}});
document.getElementById('lockpw').focus();})();
</script>`;

writeFileSync(out, html.slice(0, start) + lock + html.slice(end + '</script>'.length));
console.log('written', out, (ct.length / 1024).toFixed(0) + ' KB encrypted');
