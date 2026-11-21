// Test-only HTTPS server for the add-in page.
// Serves addin/ at https://localhost:3000 and injects a probe that reports what happens
// inside the real PowerPoint webview back to /probe, which is appended to a log file.
//   node server.js <addinDir> <pfxFile> <logFile>
const https = require('https');
const fs = require('fs');
const path = require('path');

const [rootArg, pfx, logFile] = process.argv.slice(2);
const root = path.resolve(rootArg);
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.png': 'image/png', '.xml': 'text/xml', '.css': 'text/css',
  '.ps1': 'text/plain; charset=utf-8', '.sh': 'text/plain; charset=utf-8' };

function log(o) {
  const line = JSON.stringify(Object.assign({ t: new Date().toISOString() }, o));
  fs.appendFileSync(logFile, line + '\n');
  console.log(line);
}

// Runs inside the add-in page (host). ES5 only, so it also works in legacy webviews.
const PROBE = `<script>(function(){
  function send(o){try{var x=new XMLHttpRequest();x.open('POST','/probe',true);x.setRequestHeader('Content-Type','text/plain');x.send(JSON.stringify(o));}catch(e){}}
  send({ev:'page-load',ua:navigator.userAgent,w:window.innerWidth,h:window.innerHeight,hasOffice:!!window.Office,srcdoc:('srcdoc' in document.createElement('iframe'))});
  window.addEventListener('error',function(e){send({ev:'host-error',msg:String(e.message),src:e.filename,line:e.lineno});});
  // Runs inside the deck (sandboxed iframe). Reports readiness and real user input.
  var BEACON='<scr'+'ipt>(function(){function p(o){o.__probe=1;try{parent.postMessage(o,"*")}catch(e){}}'
    +'function st(){var on=document.querySelector(".slide.on");return {slide:window.__LIVE_SLIDE__,on:on&&on.id,w:innerWidth,h:innerHeight}}'
    +'addEventListener("error",function(e){p({ev:"deck-error",msg:String(e.message),line:e.lineno})});'
    +'addEventListener("load",function(){setTimeout(function(){var s=st();s.ev="deck-ready";s.title=document.title;s.fonts=(document.fonts&&document.fonts.status)||"";s.ff=getComputedStyle(document.body).fontFamily;p(s)},600)});'
    +'document.addEventListener("click",function(e){var t=e.target,s=st();s.ev="deck-click";s.x=e.clientX;s.y=e.clientY;s.tag=t.tagName;s.id=t.id||"";s.cls=String(t.className||"").slice(0,40);s.trusted=e.isTrusted;p(s)},true);'
    +'document.addEventListener("keydown",function(e){var s=st();s.ev="deck-key";s.key=e.key;s.trusted=e.isTrusted;p(s)},true);'
    +'document.addEventListener("input",function(e){var s=st();s.ev="deck-input";s.id=e.target.id||"";s.value=String(e.target.value).slice(0,40);p(s)},true);'
    +'})();</scr'+'ipt>';
  try{
    var d=Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype,'srcdoc');
    if(d&&d.set){Object.defineProperty(HTMLIFrameElement.prototype,'srcdoc',{configurable:true,get:d.get,set:function(v){
      send({ev:'srcdoc-set',len:v.length,marker:(/__LIVE_SLIDE__=(\\d+)/.exec(v)||[])[1]||null});
      d.set.call(this,v+BEACON);}});}
    else send({ev:'no-srcdoc-setter'});
  }catch(e){send({ev:'hook-error',msg:String(e)});}
  window.addEventListener('message',function(e){if(e.data&&e.data.__probe){send(e.data);}});
  if(window.Office&&Office.onReady){Office.onReady(function(info){
    var k={};try{var s=Office.context.document.settings;k={htmlLen:(s.get('html')||'').length,slide:s.get('slide'),width:s.get('width'),height:s.get('height')};}catch(e){k={err:String(e)};}
    var mode=null;try{mode=Office.context.document.mode;}catch(e){}
    send({ev:'office-ready',host:info&&info.host,platform:info&&info.platform,settings:k,mode:mode});
    try{Office.context.document.getActiveViewAsync(function(r){send({ev:'active-view',view:r.value,status:r.status});});}catch(e){send({ev:'active-view-error',msg:String(e)});}
    try{Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged,function(a){send({ev:'view-changed',view:a.activeView});});}catch(e){}
  });}
})();</script>`;

const server = https.createServer({ pfx: fs.readFileSync(pfx), passphrase: 'livehtml' }, (req, res) => {
  const u = new URL(req.url, 'https://localhost:3000');
  if (u.pathname === '/mark') { log({ mark: u.searchParams.get('l') }); res.writeHead(204); res.end(); return; }
  if (u.pathname === '/probe') {
    let body = '';
    req.on('data', (c) => { body += c; });
    req.on('end', () => {
      let o; try { o = JSON.parse(body); } catch (e) { o = { raw: body.slice(0, 200) }; }
      log(Object.assign({ probe: true }, o));
      res.writeHead(204); res.end();
    });
    return;
  }
  const rel = u.pathname === '/' ? '/index.html' : decodeURIComponent(u.pathname);
  const file = path.join(root, path.normalize(rel));
  log({ req: req.method + ' ' + u.pathname + u.search, ua: (req.headers['user-agent'] || '').slice(0, 160) });
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.writeHead(404); res.end('not found'); return;
  }
  let data = fs.readFileSync(file);
  if (rel === '/index.html' && !process.env.NO_PROBE) {
    data = Buffer.from(data.toString('utf8').replace('</head>', PROBE + '</head>'), 'utf8');
  }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
  res.end(data);
});
// Connection-level events: a client that rejects our certificate never produces an HTTP request.
server.on('connection', (sock) => log({ tcp: 'connection', from: sock.remoteAddress }));
server.on('secureConnection', (sock) => log({ tls: 'ok', proto: sock.getProtocol(), sni: sock.servername || null }));
server.on('tlsClientError', (err, sock) => log({ tls: 'clientError', msg: String(err && err.message), from: sock && sock.remoteAddress }));
server.listen(3000, () => log({ listening: 3000, root }));
