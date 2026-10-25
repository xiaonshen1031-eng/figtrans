document.documentElement.classList.add('js');
const $=(s,root=document)=>root.querySelector(s);
const $$=(s,root=document)=>[...root.querySelectorAll(s)];
const reduceMotion=matchMedia('(prefers-reduced-motion: reduce)').matches;
let toastTimer;
function toast(message){const box=$('#toast');box.textContent=message;box.classList.add('visible');clearTimeout(toastTimer);toastTimer=setTimeout(()=>box.classList.remove('visible'),2600);}
if('IntersectionObserver' in window){const observer=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){entry.target.classList.add('is-visible');observer.unobserve(entry.target)}}),{threshold:.08});$$('.reveal').forEach(el=>observer.observe(el));}else{$$('.reveal').forEach(el=>el.classList.add('is-visible'));}

// This original browser illustration explains attachment. It does not convert uploads.
const svg=$('#diagram'),node=$('#editable-node');
let nodePosition={x:286,y:172},drag=null,nodeMoved=false;
function paintNode(){const{x,y}=nodePosition;node.setAttribute('transform',`translate(${x} ${y})`);const cy=y+38;$('#edge-in').setAttribute('d',`M199 210C245 210 ${x-35} ${cy} ${x} ${cy}`);$('#edge-pptx').setAttribute('d',`M${x+152} ${cy}C${x+195} ${cy} 466 142 506 142`);$('#edge-visio').setAttribute('d',`M${x+152} ${cy}C${x+195} ${cy} 466 280 506 280`);$('#inspect-node-label').setAttribute('x',x-4);$('#inspect-node-label').setAttribute('y',y-23);}
function svgPoint(event){const point=new DOMPoint(event.clientX,event.clientY);return point.matrixTransform(svg.getScreenCTM().inverse());}
node.addEventListener('pointerdown',event=>{if(event.button!==0)return;const p=svgPoint(event);drag={dx:p.x-nodePosition.x,dy:p.y-nodePosition.y};node.setPointerCapture(event.pointerId);node.classList.add('is-dragging');event.preventDefault();});
node.addEventListener('pointermove',event=>{if(!drag)return;const p=svgPoint(event);nodePosition={x:Math.max(245,Math.min(340,p.x-drag.dx)),y:Math.max(128,Math.min(218,p.y-drag.dy))};paintNode();nodeMoved=true;$('#canvas-note').textContent='关系仍在。修改可以继续。';});
function releaseNode(){drag=null;node.classList.remove('is-dragging');}
node.addEventListener('pointerup',releaseNode);node.addEventListener('pointercancel',releaseNode);
$('#move-node').addEventListener('click',()=>{nodeMoved=!nodeMoved;nodePosition=nodeMoved?{x:309,y:137}:{x:286,y:172};paintNode();$('#canvas-note').textContent=nodeMoved?'节点已移动，连接仍保持跟随。':'拖动绿色节点，连接会跟随。';});
$('#inspect-toggle').addEventListener('click',event=>{const active=event.currentTarget.getAttribute('aria-pressed')!=='true';event.currentTarget.setAttribute('aria-pressed',active);$('#workshop').classList.toggle('is-inspecting',active);$('#canvas-filename').textContent=active?'figure.objects':'figure.editable';});
$$('[data-format]').forEach(button=>button.addEventListener('click',()=>{const format=button.dataset.format;$$('[data-format]').forEach(b=>b.setAttribute('aria-pressed',b===button));$$('.output-pptx').forEach(el=>el.classList.toggle('muted-output',format==='visio'));$$('.output-visio').forEach(el=>el.classList.toggle('muted-output',format==='pptx'));$('#canvas-state').textContent={both:'2 native formats',pptx:'.pptx / PowerPoint',visio:'.vsdx / Visio'}[format];}));

const video=$('#demo-video'),videoPlay=$('#video-play');
videoPlay.addEventListener('click',async()=>{try{videoPlay.hidden=true;await video.play();}catch{videoPlay.hidden=false;toast('点击视频自带的播放按钮即可观看。');}});
video.addEventListener('play',()=>videoPlay.hidden=true);
video.addEventListener('ended',()=>videoPlay.hidden=false);
video.addEventListener('error',()=>{videoPlay.hidden=true;toast('视频暂时无法播放，可下载演示文件查看。');});
$$('.watch-link').forEach(a=>a.addEventListener('click',()=>{video.preload='metadata';}));

$$('[data-copy]').forEach(button=>button.addEventListener('click',async()=>{const original=button.innerHTML;try{if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(button.dataset.copy);}else{const area=document.createElement('textarea');area.value=button.dataset.copy;area.style.position='fixed';area.style.opacity='0';document.body.append(area);area.select();const ok=document.execCommand('copy');area.remove();if(!ok)throw Error('copy');}button.textContent='已复制 ✓';toast('已复制，可以粘贴使用。');setTimeout(()=>button.innerHTML=original,2000);}catch{toast('复制未成功，请选中文本后手动复制。');}}));
$$('.download-cta').forEach(a=>a.addEventListener('click',()=>toast('正在下载 FigTrans v1.1.0 · MIT 开源包')));

let sourceIndex=null,sourceAbort=null;
const sourceDialog=$('#source-dialog');
async function fetchText(url){const response=await fetch(url);if(!response.ok)throw Error(`HTTP ${response.status}`);return response.text();}
async function showSource(file){if(sourceAbort)sourceAbort.abort();sourceAbort=new AbortController();const name=file.path||file.name;$('#source-filename').textContent=name;$('#source-content').textContent='正在读取源码…';$('#source-file-download').href=file.url;$$('#source-files button').forEach(b=>b.setAttribute('aria-current',b.dataset.path===name));$('#source-files button[aria-current="true"]')?.scrollIntoView({block:'nearest',inline:'nearest',behavior:'instant'});try{const res=await fetch(file.url,{signal:sourceAbort.signal});if(!res.ok)throw Error('source');$('#source-content').textContent=await res.text();$('.source-code pre').scrollTop=0;}catch(error){if(error.name!=='AbortError')$('#source-content').textContent='源码暂时读取失败，可使用上方链接下载文件。';}}
async function openSource(){sourceDialog.showModal();try{if(!sourceIndex){const response=await fetch('source/index.json');if(!response.ok)throw Error('index');const index=await response.json();sourceIndex=Array.isArray(index)?index:index.files;const box=$('#source-files');box.replaceChildren();sourceIndex.forEach(file=>{const button=document.createElement('button');button.type='button';button.dataset.path=file.path||file.name;button.textContent=(file.path||file.name).replace(/^senior-figtrans\//,'');button.addEventListener('click',()=>showSource(file));box.append(button);});}const first=sourceIndex.find(f=>(f.path||f.name).endsWith('SKILL.md'))||sourceIndex[0];await showSource(first);}catch{$('#source-content').textContent='暂时无法加载源码目录。可直接下载完整 ZIP 查看所有源码。';}}
$$('.source-open').forEach(button=>button.addEventListener('click',openSource));
$('#install-guide').addEventListener('click',async()=>{const dialog=$('#guide-dialog');dialog.showModal();try{$('#guide-content').textContent=await fetchText('docs/install.md');}catch{$('#guide-content').textContent='说明暂时无法加载，请使用下方链接下载。';}});
$$('.dialog-close').forEach(button=>button.addEventListener('click',()=>button.closest('dialog').close()));
$$('dialog').forEach(dialog=>{dialog.addEventListener('click',event=>{if(event.target===dialog){const r=dialog.getBoundingClientRect();if(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom)dialog.close();}});dialog.addEventListener('close',()=>{document.body.style.overflow='';});const observer=new MutationObserver(()=>{if(dialog.open)document.body.style.overflow='hidden';});observer.observe(dialog,{attributes:true,attributeFilter:['open']});});
fetch('downloads/release.json').then(r=>{if(!r.ok)throw Error();return r.json()}).then(release=>{const bytes=release.zip_bytes||release.bytes||release.archive_bytes||release.download?.bytes;if(bytes)$('#release-size').textContent=`ZIP · ${(bytes/1024).toFixed(1)} KB · 源码与许可完整保留`;}).catch(()=>{});
paintNode();
