import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const root = path.dirname(fileURLToPath(import.meta.url));
const source = relative => fs.readFileSync(path.join(root, relative));
const text = relative => source(relative).toString('utf8');
const destination = path.join(root, 'dist');
fs.mkdirSync(path.join(destination, 'offline'), { recursive: true });
fs.copyFileSync(path.join(root, 'CNAME'), path.join(destination, 'CNAME'));
const logo = 'data:image/png;base64,' + source('src/mightora-logo.png').toString('base64');

const onlineHead = `<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/mightora/shared-ui@main/shared.css">
<script>(function(){const nativeFetch=window.fetch;window.fetch=function(input,init){const url=typeof input==='string'?input:input&&input.url||'';if(/footer\\.ya?ml(?:\\?|$)/i.test(url)){return nativeFetch.call(this,input,init).then(async response=>{if(!response.ok)return response;const body=(await response.text()).replace(/\\t/g,'  ').replace(/^( +)-\\s*$/gm,'$1-');return new Response(body,{status:response.status,statusText:response.statusText,headers:response.headers});});}return nativeFetch.call(this,input,init);};})();</script>`;
const onlineScripts = `<script src="https://cdn.jsdelivr.net/npm/js-yaml@4.1.0/dist/js-yaml.min.js" defer></script>
<script src="https://cdn.jsdelivr.net/gh/mightora/shared-ui@main/components.js" defer></script>`;
const offlineHead = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; base-uri 'none'; form-action 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; connect-src 'none'; font-src data:; object-src 'none'">`;

function html(offline) {
  let result = text('src/index.template.html');
  result = result.replace('<!-- ONLINE_HEAD -->', offline ? offlineHead : onlineHead)
    .replace('/* APP_CSS */', text('src/styles.css'))
    .replace('__LOGO__', logo)
    .replace('/* CORE_JS */', text('src/core.js'))
    .replace('/* APP_JS */', text('src/app.js'))
    .replace('<!-- ONLINE_SCRIPTS -->', offline ? '' : onlineScripts);
  if (offline) {
    result = result.replace('<meta name="robots" content="index, follow">','<meta name="robots" content="noindex">')
      .replace('<link rel="canonical" href="https://sharefoundry.mightora.io/">','');
  }
  return result;
}
fs.writeFileSync(path.join(destination, 'index.html'), html(false));
const offlineHtml = html(true);
fs.writeFileSync(path.join(destination, 'offline/sharepoint-maker-toolkit.html'), offlineHtml);
const offlineReadme = `Mightora SharePoint Maker Toolkit – Offline Edition\n\nOpen sharepoint-maker-toolkit.html directly in a browser on Windows, macOS or Linux. No web server is necessary.\n\nThe application uses no network endpoints and does not need installation. You can create list schemas, formatting JSON and portable project backups. Your browser may restrict localStorage on file:// pages; export your projects as JSON to retain them.\n\nOnly exported files leave the browser, and only when you choose to download them. The Microsoft Graph request body is a template: deployment to SharePoint is not performed by this tool.\n\nTo inspect network requests, press F12, open Network, enable Preserve log, then reload. An independent audit is more complete than the built-in resource timing summary.\n`;

// Minimal standards-compliant ZIP writer (no build dependencies).
const crcTable = Array.from({ length: 256 }, (_, i) => { let c=i;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0; });
function crc32(data) { let c=0xffffffff;for(const byte of data)c=crcTable[(c^byte)&255]^(c>>>8);return(c^0xffffffff)>>>0; }
function makeZip(entries) {
  let offset=0;const local=[];const central=[];
  for(const [filename, dataText] of entries) {
    const name=Buffer.from(filename);const data=Buffer.from(dataText);const payload=zlib.deflateRawSync(data,{level:9});const checksum=crc32(data);
    const localHeader=Buffer.alloc(30);localHeader.writeUInt32LE(0x04034b50,0);localHeader.writeUInt16LE(20,4);localHeader.writeUInt16LE(8,8);localHeader.writeUInt32LE(checksum,14);localHeader.writeUInt32LE(payload.length,18);localHeader.writeUInt32LE(data.length,22);localHeader.writeUInt16LE(name.length,26);
    local.push(localHeader,name,payload);
    const centralHeader=Buffer.alloc(46);centralHeader.writeUInt32LE(0x02014b50,0);centralHeader.writeUInt16LE(20,4);centralHeader.writeUInt16LE(20,6);centralHeader.writeUInt16LE(8,10);centralHeader.writeUInt32LE(checksum,16);centralHeader.writeUInt32LE(payload.length,20);centralHeader.writeUInt32LE(data.length,24);centralHeader.writeUInt16LE(name.length,28);centralHeader.writeUInt32LE(offset,42);
    central.push(centralHeader,name);offset+=localHeader.length+name.length+payload.length;
  }
  const directory=Buffer.concat(central);const end=Buffer.alloc(22);end.writeUInt32LE(0x06054b50,0);end.writeUInt16LE(entries.length,8);end.writeUInt16LE(entries.length,10);end.writeUInt32LE(directory.length,12);end.writeUInt32LE(offset,16);
  return Buffer.concat([...local,directory,end]);
}
fs.writeFileSync(path.join(destination,'offline/sharepoint-maker-toolkit.zip'),makeZip([
  ['sharepoint-maker-toolkit.html',offlineHtml],['README.txt',offlineReadme],['LICENSE',text('LICENSE')]
]));
console.log('Built online and self-contained offline editions.');
