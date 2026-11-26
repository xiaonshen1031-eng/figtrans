'use strict';

const find = selector => document.querySelector(selector);
const all = selector => Array.from(document.querySelectorAll(selector));
const revision = '?v=130';

const toast = find('#toast');
let toastTimer;
function announce(message) {
  toast.textContent = message;
  toast.classList.add('visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('visible'), 2500);
}

const workshop = find('#workshop');
const svg = find('#diagram');
const node = find('#editable-node');
const state = find('#canvas-state');
const originalPosition = {x: 286, y: 172};
const shiftedPosition = {x: 309, y: 137};
let position = originalPosition;
let drag = null;

function drawNode() {
  const {x, y} = position;
  const center = y + 38;
  node.setAttribute('transform', 'translate(' + x + ' ' + y + ')');
  find('#edge-in').setAttribute('d', 'M199 210C245 210 ' + (x - 35) + ' ' + center + ' ' + x + ' ' + center);
  for (const [id, end] of [['#edge-pptx', 142], ['#edge-visio', 280]]) {
    find(id).setAttribute('d', 'M' + (x + 152) + ' ' + center + 'C' + (x + 195) + ' ' + center + ' 466 ' + end + ' 506 ' + end);
  }
  const label = find('#inspect-node-label');
  label.setAttribute('x', x - 4);
  label.setAttribute('y', y - 23);
}

function svgPoint(event) {
  return new DOMPoint(event.clientX, event.clientY).matrixTransform(svg.getScreenCTM().inverse());
}

node.addEventListener('pointerdown', event => {
  if (!event.isPrimary || event.button !== 0) return;
  const p = svgPoint(event);
  drag = {x: p.x - position.x, y: p.y - position.y};
  node.setPointerCapture(event.pointerId);
  node.classList.add('is-dragging');
});
node.addEventListener('pointermove', event => {
  if (!drag || !node.hasPointerCapture(event.pointerId)) return;
  const p = svgPoint(event);
  position = {
    x: Math.max(245, Math.min(340, p.x - drag.x)),
    y: Math.max(128, Math.min(218, p.y - drag.y))
  };
  drawNode();
});
node.addEventListener('lostpointercapture', () => {
  drag = null;
  node.classList.remove('is-dragging');
  state.textContent = '节点移动完成，连接端点保持附着。';
});
find('#move-node').addEventListener('click', () => {
  const atDefault = position.x === originalPosition.x && position.y === originalPosition.y;
  position = atDefault ? shiftedPosition : originalPosition;
  drawNode();
  state.textContent = atDefault ? '节点已移动，连接端点保持附着。' : '节点回到原位。';
});
find('#inspect-toggle').addEventListener('click', event => {
  const active = workshop.classList.toggle('is-inspecting');
  event.currentTarget.setAttribute('aria-pressed', String(active));
});
const formatButtons = all('[data-format]');
formatButtons.forEach(button => button.addEventListener('click', () => {
  const format = button.dataset.format;
  formatButtons.forEach(item => item.setAttribute('aria-pressed', String(item === button)));
  all('.output-pptx').forEach(path => path.classList.toggle('muted-output', format === 'visio'));
  all('.output-visio').forEach(path => path.classList.toggle('muted-output', format === 'pptx'));
}));
drawNode();

const comparison = find('#compare-stage');
const slider = find('#comparison-split');
function setSplit(raw) {
  const split = Math.round(Math.min(100, Math.max(0, raw)));
  comparison.style.setProperty('--split', split + '%');
  slider.value = String(split);
  slider.setAttribute('aria-valuetext', '参考图占画面 ' + split + '%');
}
slider.addEventListener('input', () => setSplit(Number(slider.value)));
comparison.classList.add('is-interactive');
function pointerSplit(event) {
  const frame = comparison.getBoundingClientRect();
  if (frame.width) setSplit((event.clientX - frame.left) / frame.width * 100);
}
comparison.addEventListener('pointerdown', event => {
  if (!event.isPrimary || event.button !== 0) return;
  comparison.setPointerCapture(event.pointerId);
  pointerSplit(event);
});
comparison.addEventListener('pointermove', event => {
  if (comparison.hasPointerCapture(event.pointerId)) pointerSplit(event);
});

const timelapse = find('#timelapse');
const video = find('#demo-video');
function openTimelapse() {
  if (location.hash === '#timelapse') timelapse.open = true;
}
all('a[href="#timelapse"]').forEach(link => link.addEventListener('click', () => { timelapse.open = true; }));
window.addEventListener('hashchange', openTimelapse);
openTimelapse();
timelapse.addEventListener('toggle', () => {
  if (!timelapse.open) video.pause();
});

const copyButton = find('[data-copy-prompt]');
copyButton.addEventListener('click', async () => {
  const prompt = find('.prompt-box code');
  try {
    await navigator.clipboard.writeText(prompt.textContent);
    announce('已复制请求。');
  } catch {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(prompt);
    selection.removeAllRanges();
    selection.addRange(range);
    announce('已选中请求，可以手动复制。');
  }
});

async function fetchResource(url, options) {
  const response = await fetch(url, options);
  if (!response.ok) throw new Error('HTTP ' + response.status);
  return response;
}

const dialog = find('#source-dialog');
const files = find('#source-files');
const content = find('#source-content');
let opener;
let indexedFiles;
let sourceRequest;

async function showSource(file) {
  sourceRequest?.abort();
  const request = new AbortController();
  sourceRequest = request;
  find('#source-filename').textContent = file.path;
  find('#source-file-download').href = file.url + revision;
  content.textContent = '正在读取源码…';
  all('#source-files button').forEach(button => button.setAttribute('aria-current', String(button.dataset.path === file.path)));
  files.querySelector('button[aria-current="true"]')?.scrollIntoView({block: 'nearest', inline: 'nearest'});
  try {
    const response = await fetchResource(file.url + revision, {signal: request.signal});
    const code = await response.text();
    if (sourceRequest === request) {
      content.textContent = code;
      find('.source-code pre').scrollTop = 0;
    }
  } catch (error) {
    if (error.name !== 'AbortError' && sourceRequest === request) content.textContent = '源码暂时读取失败，可直接下载此文件。';
  }
}

async function openSource(button) {
  if (dialog.open) return;
  opener = button;
  dialog.showModal();
  find('#close-source').focus();
  try {
    if (!indexedFiles) {
      const index = await (await fetchResource('source/index.json' + revision)).json();
      const valid = index.files.filter(file => {
        const url = new URL(file.url, document.baseURI);
        const base = new URL('source/senior-figtrans/', document.baseURI);
        return url.origin === base.origin && url.pathname.startsWith(base.pathname) && !url.search && !url.hash;
      });
      if (!valid.length) throw new Error('Empty source index');
      indexedFiles = valid;
      files.replaceChildren(...valid.map(file => {
        const item = document.createElement('button');
        item.type = 'button';
        item.dataset.path = file.path;
        item.textContent = file.path;
        item.addEventListener('click', () => showSource(file));
        return item;
      }));
    }
    if (dialog.open) await showSource(indexedFiles.find(file => file.path === 'SKILL.md') || indexedFiles[0]);
  } catch {
    content.textContent = '源码目录暂时无法读取。可下载技能包查看。';
  }
}
all('.source-open').forEach(button => button.addEventListener('click', () => openSource(button)));
find('#close-source').addEventListener('click', () => dialog.close());
dialog.addEventListener('close', () => {
  sourceRequest?.abort();
  opener?.focus();
});

fetchResource('downloads/release.json' + revision)
  .then(response => response.json())
  .then(release => {
    const currentDownload = find('.header-download').getAttribute('href');
    if (release.download_url === currentDownload && Number.isFinite(release.bytes)) {
      find('#release-size').textContent = 'ZIP · ' + Math.round(release.bytes / 1024) + ' KiB · MIT';
    }
  })
  .catch(() => {});
