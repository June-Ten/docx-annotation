<template>
  <div class="app-shell" :class="{ 'is-fullscreen': isFullscreen }">
    <!-- 未上传：欢迎页 -->
    <section v-if="!preview.html" class="welcome">
      <div class="welcome-card">
        <div class="welcome-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none">
            <path d="M7 3h7l5 5v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z" stroke="currentColor" stroke-width="1.6" />
            <path d="M14 3v5h5M9 13h6M9 17h4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" />
          </svg>
        </div>
        <h1>DOCX 批注预览</h1>
        <p>上传 Word 文档，划词添加批注，并导出带批注的文件。</p>
        <label class="btn btn-primary welcome-upload">
          <input ref="fileInputRef" type="file" accept=".docx" @change="handleFileChange" />
          选择 DOCX 文件
        </label>
      </div>
    </section>

    <!-- 已上传：文档工作台 -->
    <template v-else>
      <header class="doc-header">
        <div class="doc-title">
          <span class="file-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" fill="none">
              <path d="M7 3h7l5 5v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z" stroke="currentColor" stroke-width="1.6" />
              <path d="M14 3v5h5" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" />
            </svg>
          </span>
          <span class="file-name">{{ preview.fileName }}</span>
        </div>
        <div class="header-actions">
          <button class="btn btn-ghost" type="button" :disabled="exporting" @click="exportDocx">
            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
              <path d="M12 4v10m0 0 4-4m-4 4-4-4M5 20h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
            {{ exporting ? '生成中...' : '下载' }}
          </button>
          <button class="btn btn-ghost" type="button" @click="toggleFullscreen">
            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
              <path d="M8 4H4v4M20 4h-4v4M4 20v-4h4M20 20h-4v-4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
            </svg>
            全屏
          </button>
          <button class="btn btn-primary" type="button" @click="shareDocument">
            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
              <circle cx="18" cy="5" r="2.5" stroke="currentColor" stroke-width="1.6" />
              <circle cx="6" cy="12" r="2.5" stroke="currentColor" stroke-width="1.6" />
              <circle cx="18" cy="19" r="2.5" stroke="currentColor" stroke-width="1.6" />
              <path d="M8.2 11 15.8 6.5M8.2 13l7.6 4.5" stroke="currentColor" stroke-width="1.6" />
            </svg>
            分享
          </button>
        </div>
      </header>

      <div class="doc-toolbar">
        <div class="toolbar-group">
          <button class="icon-btn" type="button" :disabled="currentPage <= 1" @click="goToPage(currentPage - 1)">
            <svg viewBox="0 0 24 24" fill="none"><path d="M14 6l-6 6 6 6" stroke="currentColor" stroke-width="2" stroke-linecap="round" /></svg>
          </button>
          <span class="page-indicator">{{ currentPage }} / {{ totalPages }}</span>
          <button class="icon-btn" type="button" :disabled="currentPage >= totalPages" @click="goToPage(currentPage + 1)">
            <svg viewBox="0 0 24 24" fill="none"><path d="M10 6l6 6-6 6" stroke="currentColor" stroke-width="2" stroke-linecap="round" /></svg>
          </button>
        </div>

        <div class="toolbar-divider" />

        <div class="toolbar-group">
          <select v-model="zoomLabel" class="zoom-select" @change="applyZoomLabel">
            <option v-for="level in zoomLevels" :key="level" :value="level">{{ level }}</option>
          </select>
          <button class="toolbar-text-btn" type="button" @click="zoomIn">放大</button>
          <button class="toolbar-text-btn" type="button" @click="zoomOut">缩小</button>
          <button class="toolbar-text-btn" type="button" @click="fitWidth">适应宽度</button>
        </div>

        <div class="toolbar-divider" />

        <div class="toolbar-group view-modes">
          <button class="toolbar-text-btn" :class="{ active: viewMode === 'single' }" type="button" @click="viewMode = 'single'">单页</button>
          <button class="toolbar-text-btn" :class="{ active: viewMode === 'double' }" type="button" @click="viewMode = 'double'">双页</button>
        </div>
      </div>

      <div class="doc-workspace">
        <section ref="viewerColumnRef" class="viewer-column">
          <div ref="viewerScrollRef" class="viewer-scroll" @scroll="onViewerScroll">
            <div class="viewer-stage" :class="viewMode">
              <div class="doc-canvas" :style="canvasStyle">
                <div ref="previewRef" class="doc-paper" v-html="preview.html"></div>
              </div>
            </div>
          </div>
          <footer class="status-bar">
            <span>字数: {{ wordCount.toLocaleString() }}</span>
            <span>页数: {{ totalPages }}</span>
            <span>最后更新时间: {{ lastUpdatedLabel }}</span>
          </footer>
        </section>

        <aside v-if="sidebarOpen" class="comment-sidebar">
          <header class="sidebar-header">
            <h2>批注列表</h2>
            <button class="icon-btn subtle" type="button" title="关闭侧栏" @click="sidebarOpen = false">
              <svg viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6 6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round" /></svg>
            </button>
          </header>

          <div class="sidebar-filters">
            <select v-model="filterMode" class="filter-select">
              <option value="all">全部批注</option>
              <option value="mine">我的批注</option>
            </select>
            <select v-model="sortMode" class="filter-select">
              <option value="position">按位置</option>
              <option value="time">按时间</option>
            </select>
          </div>

          <div class="comment-list">
            <p v-if="displayedComments.length === 0" class="empty-comments">划词选中正文后，在下方输入批注并发送。</p>
            <article
              v-for="comment in displayedComments"
              :key="comment.id"
              class="comment-card"
              :class="{ active: activeCommentId === comment.id }"
              @click="focusComment(comment)"
            >
              <div class="comment-card-head">
                <span class="comment-badge" :style="badgeStyle(comment.colorIndex)">{{ comment.number }}</span>
                <div class="comment-meta">
                  <strong>{{ comment.author }}</strong>
                  <time>{{ formatTime(comment.createdAt) }}</time>
                </div>
              </div>
              <p class="comment-quote">「{{ comment.selectedText }}」</p>
              <p class="comment-body">{{ comment.comment }}</p>
              <div class="comment-actions">
                <button type="button" class="link-btn" @click.stop="replyTo(comment)">回复</button>
                <button type="button" class="link-btn danger" @click.stop="removeComment(comment.id)">删除</button>
              </div>
            </article>
          </div>

          <div v-if="preview.draft" class="selection-hint">
            已选：{{ preview.draft.selectedText }}
          </div>

          <footer class="sidebar-compose">
            <input v-model="author" class="author-input" placeholder="署名" />
            <div class="compose-row">
              <textarea
                v-model="preview.commentText"
                rows="2"
                placeholder="添加批注..."
                @keydown.enter.exact.prevent="sendComment"
              />
              <button class="btn btn-primary send-btn" type="button" :disabled="!canSendComment" @click="sendComment">
                发送
              </button>
            </div>
          </footer>
        </aside>

        <button v-else class="sidebar-toggle" type="button" @click="sidebarOpen = true">批注</button>
      </div>
    </template>

    <p v-if="message" class="toast" role="status">{{ message }}</p>
  </div>
</template>

<script setup>
import Highlighter from 'web-highlighter';
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from 'vue';

const JAVA_ASPOSE_API = import.meta.env.VITE_JAVA_ASPOSE_API ?? '/api/java-aspose';

const COMMENT_COLORS = [
  { bg: '#fef08a', badge: '#eab308' },
  { bg: '#bfdbfe', badge: '#3b82f6' },
  { bg: '#bbf7d0', badge: '#22c55e' },
  { bg: '#fbcfe8', badge: '#ec4899' }
];

const PAGE_HEIGHT = 1056;
const ZOOM_MAP = { '75%': 0.75, '100%': 1, '125%': 1.25, '150%': 1.5, '200%': 2 };

const author = ref('Reviewer');
const message = ref('');
const exporting = ref(false);
const sidebarOpen = ref(true);
const isFullscreen = ref(false);
const viewMode = ref('single');
const filterMode = ref('all');
const sortMode = ref('position');
const zoom = ref(1);
const zoomLabel = ref('100%');
const zoomLevels = Object.keys(ZOOM_MAP);
const currentPage = ref(1);
const totalPages = ref(1);
const wordCount = ref(0);
const activeCommentId = ref('');

const fileInputRef = ref(null);
const previewRef = ref(null);
const viewerScrollRef = ref(null);
const viewerColumnRef = ref(null);

let highlighter = null;

const preview = reactive({
  sourceFile: null,
  uploadId: '',
  fileName: '',
  html: '',
  loadedAt: null,
  comments: [],
  draft: null,
  commentText: ''
});

const canvasStyle = computed(() => ({
  transform: `scale(${zoom.value})`,
  transformOrigin: 'top center'
}));

const lastUpdatedLabel = computed(() => {
  if (!preview.loadedAt) return '--';
  return formatTime(preview.loadedAt);
});

const canSendComment = computed(() => Boolean(preview.draft && preview.commentText.trim()));

const displayedComments = computed(() => {
  let list = [...preview.comments];
  if (filterMode.value === 'mine') {
    const name = author.value.trim() || 'Reviewer';
    list = list.filter((item) => item.author === name);
  }
  if (sortMode.value === 'time') {
    list.sort((a, b) => b.createdAt - a.createdAt);
  } else {
    list.sort((a, b) => a.number - b.number);
  }
  return list;
});

watch(
  () => preview.html,
  async () => {
    await nextTick();
    updateDocumentMetrics();
  }
);

onMounted(() => {
  window.addEventListener('resize', updateDocumentMetrics);
});

async function handleFileChange(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  resetPreviewState(file);
  message.value = '正在上传 DOCX 并解析...';

  try {
    const formData = new FormData();
    formData.append('file', file);
    const response = await fetch(`${JAVA_ASPOSE_API}/docx/upload`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    const result = await response.json();
    preview.uploadId = result.uploadId;
    preview.fileName = result.fileName || file.name;
    preview.html = result.html;
    preview.loadedAt = new Date();
    message.value = '';

    await nextTick();
    setupHighlighter();
    updateDocumentMetrics();
    goToPage(1);
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

function resetPreviewState(file) {
  destroyHighlighter();
  Object.assign(preview, {
    sourceFile: file,
    uploadId: '',
    fileName: file.name,
    html: '',
    loadedAt: null,
    comments: [],
    draft: null,
    commentText: ''
  });
  currentPage.value = 1;
  totalPages.value = 1;
  wordCount.value = 0;
  activeCommentId.value = '';
}

function setupHighlighter() {
  destroyHighlighter();
  const root = previewRef.value;
  if (!root) return;

  highlighter = new Highlighter({
    $root: root,
    exceptSelectors: ['button', 'input', 'textarea', 'select']
  });

  highlighter.on(Highlighter.event.CREATE, ({ sources }) => {
    const source = sources?.[0];
    const selectedText = source?.text?.replace(/\s+/g, ' ').trim();
    if (!source || !selectedText) return;

    preview.draft = {
      id: source.id,
      source,
      selectedText,
      occurrence: countOccurrenceBeforeSelection(selectedText)
    };
    preview.commentText = '';
    activeCommentId.value = source.id;
    sidebarOpen.value = true;
    message.value = '';
  });

  highlighter.run();
}

function destroyHighlighter() {
  highlighter = null;
}

function countOccurrenceBeforeSelection(selectedText) {
  const selection = window.getSelection();
  const root = previewRef.value;
  if (!selection || !root || selection.rangeCount === 0) return 0;

  const range = selection.getRangeAt(0);
  const beforeRange = document.createRange();
  beforeRange.selectNodeContents(root);
  beforeRange.setEnd(range.startContainer, range.startOffset);

  const beforeText = normalizeText(beforeRange.toString());
  const needle = normalizeText(selectedText);
  let count = 0;
  let index = beforeText.indexOf(needle);

  while (index >= 0) {
    count += 1;
    index = beforeText.indexOf(needle, index + needle.length);
  }

  return count;
}

function sendComment() {
  if (!canSendComment.value) return;

  const comment = {
    id: preview.draft.id,
    number: preview.comments.length + 1,
    colorIndex: preview.comments.length % COMMENT_COLORS.length,
    selectedText: preview.draft.selectedText,
    occurrence: preview.draft.occurrence,
    comment: preview.commentText.trim(),
    author: author.value.trim() || 'Reviewer',
    initials: getInitials(author.value),
    highlightSource: preview.draft.source,
    createdAt: Date.now()
  };

  preview.comments.push(comment);
  preview.draft = null;
  preview.commentText = '';
  activeCommentId.value = comment.id;

  nextTick(() => {
    styleHighlight(comment);
    updateDocumentMetrics();
  });
}

function styleHighlight(comment) {
  if (!highlighter?.getDoms) return;
  const nodes = highlighter.getDoms(comment.id) || [];
  const color = COMMENT_COLORS[comment.colorIndex];
  nodes.forEach((node) => {
    node.style.backgroundColor = color.bg;
    node.style.borderRadius = '3px';
    node.style.boxDecorationBreak = 'clone';
    node.style.setProperty('--marker-color', color.badge);
    node.dataset.commentNum = String(comment.number);
  });
}

function badgeStyle(colorIndex) {
  const color = COMMENT_COLORS[colorIndex % COMMENT_COLORS.length];
  return { backgroundColor: color.badge };
}

function focusComment(comment) {
  activeCommentId.value = comment.id;
  const nodes = highlighter?.getDoms?.(comment.id) || [];
  nodes[0]?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function replyTo(comment) {
  preview.draft = null;
  preview.commentText = `@${comment.author} `;
  message.value = '回复模式：补充内容后点击发送';
}

function removeComment(id) {
  preview.comments = preview.comments.filter((comment) => comment.id !== id);
  highlighter?.remove?.(id);
  if (activeCommentId.value === id) {
    activeCommentId.value = '';
  }
  renumberComments();
}

function renumberComments() {
  preview.comments.forEach((comment, index) => {
    comment.number = index + 1;
    comment.colorIndex = index % COMMENT_COLORS.length;
    styleHighlight(comment);
  });
}

function updateDocumentMetrics() {
  const paper = previewRef.value;
  if (!paper) return;

  const text = paper.textContent || '';
  wordCount.value = text.replace(/\s+/g, '').length;
  const height = paper.getBoundingClientRect().height / zoom.value;
  totalPages.value = Math.max(1, Math.ceil(height / PAGE_HEIGHT));
}

function onViewerScroll() {
  const scrollEl = viewerScrollRef.value;
  const paper = previewRef.value;
  if (!scrollEl || !paper) return;

  const paperTop = paper.offsetTop * zoom.value;
  const offset = scrollEl.scrollTop - paperTop;
  const page = Math.floor(Math.max(0, offset) / (PAGE_HEIGHT * zoom.value)) + 1;
  currentPage.value = Math.min(totalPages.value, Math.max(1, page));
}

function goToPage(page) {
  const scrollEl = viewerScrollRef.value;
  const paper = previewRef.value;
  if (!scrollEl || !paper) return;

  const target = Math.min(totalPages.value, Math.max(1, page));
  currentPage.value = target;
  const paperTop = paper.offsetTop * zoom.value;
  scrollEl.scrollTo({
    top: paperTop + (target - 1) * PAGE_HEIGHT * zoom.value,
    behavior: 'smooth'
  });
}

function zoomIn() {
  setZoom(zoom.value + 0.25);
}

function zoomOut() {
  setZoom(zoom.value - 0.25);
}

function setZoom(value) {
  zoom.value = Math.min(2, Math.max(0.5, value));
  const matched = Object.entries(ZOOM_MAP).find(([, v]) => Math.abs(v - zoom.value) < 0.01);
  zoomLabel.value = matched?.[0] ?? `${Math.round(zoom.value * 100)}%`;
  nextTick(updateDocumentMetrics);
}

function applyZoomLabel() {
  setZoom(ZOOM_MAP[zoomLabel.value] ?? 1);
}

function fitWidth() {
  const scrollEl = viewerScrollRef.value;
  const paper = previewRef.value;
  if (!scrollEl || !paper) return;

  const padding = viewMode.value === 'double' ? 120 : 80;
  const available = scrollEl.clientWidth - padding;
  const paperWidth = paper.scrollWidth || 816;
  setZoom(available / paperWidth);
}

function toggleFullscreen() {
  const el = viewerColumnRef.value;
  if (!el) return;

  if (!document.fullscreenElement) {
    el.requestFullscreen?.();
    isFullscreen.value = true;
  } else {
    document.exitFullscreen?.();
    isFullscreen.value = false;
  }
}

async function shareDocument() {
  const url = window.location.href;
  try {
    await navigator.clipboard.writeText(url);
    message.value = '链接已复制到剪贴板';
  } catch {
    message.value = url;
  }
  setTimeout(() => {
    if (message.value === '链接已复制到剪贴板' || message.value === url) {
      message.value = '';
    }
  }, 2500);
}

async function exportDocx() {
  if (!preview.sourceFile) {
    message.value = '请先上传 DOCX';
    return;
  }
  if (exporting.value) return;

  exporting.value = true;
  message.value = '';

  try {
    if (canSendComment.value) {
      sendComment();
    }
    if (preview.comments.length === 0) {
      message.value = '请先划词并添加批注';
      return;
    }

    const response = preview.uploadId
      ? await fetch(`${JAVA_ASPOSE_API}/docx/${preview.uploadId}/comments`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ comments: serializeComments() })
        })
      : await fetch(`${JAVA_ASPOSE_API}/comments/generate`, {
          method: 'POST',
          body: createExportFormData()
        });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = preview.fileName.replace(/\.docx$/i, '-comments.docx');
    link.click();
    URL.revokeObjectURL(url);
    message.value = '批注文件已下载';
  } catch (error) {
    message.value = error.message || '生成失败';
  } finally {
    exporting.value = false;
  }
}

function serializeComments() {
  return preview.comments.map(({ selectedText, occurrence, comment, author, initials }) => ({
    selectedText,
    occurrence,
    comment,
    author,
    initials
  }));
}

function createExportFormData() {
  const formData = new FormData();
  formData.append('file', preview.sourceFile);
  formData.append('comments', JSON.stringify(serializeComments()));
  return formData;
}

function formatTime(value) {
  const date = new Date(value);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const h = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return `${y}-${m}-${d} ${h}:${min}`;
}

function normalizeText(value) {
  return value.replace(/\s+/g, ' ').trim();
}

function getInitials(value) {
  const normalized = value.trim();
  if (!normalized) return 'RV';
  return normalized
    .split(/\s+/)
    .map((part) => part[0])
    .join('')
    .slice(0, 9)
    .toUpperCase();
}

async function readError(response) {
  const text = await response.text();
  try {
    const payload = JSON.parse(text);
    return payload.message || response.statusText || '生成失败';
  } catch {
    return text || response.statusText || '生成失败';
  }
}

onBeforeUnmount(() => {
  window.removeEventListener('resize', updateDocumentMetrics);
  destroyHighlighter();
});
</script>
