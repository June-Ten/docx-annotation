<template>
  <main class="page">
    <section class="toolbar">
      <div>
        <p class="eyebrow">DOCX Comment Editor</p>
        <h1>DOCX 划词高亮批注</h1>
        <p>mammoth、Python Aspose、Java Aspose 使用独立预览区域，web-highlighter 负责划词高亮，后端生成带 Word 批注的文件。</p>
      </div>

      <div class="upload-actions">
        <label class="upload-button">
          <input type="file" accept=".docx" @change="handleMammothFileChange" />
          mammoth 前端预览
        </label>
        <label class="upload-button aspose">
          <input type="file" accept=".docx" @change="handleAsposeFileChange" />
          Python Aspose 后端预览
        </label>
        <label class="upload-button java-aspose">
          <input type="file" accept=".docx" @change="handleJavaAsposeFileChange" />
          Java Aspose 后端预览
        </label>
      </div>
    </section>

    <section v-if="hasAnyDocument" class="workspace">
      <article class="preview-card">
        <header>
          <div>
            <strong>{{ currentPreview.fileName || '尚未上传当前 Tab 文件' }}</strong>
            <span>{{ currentPreview.label }}。拖拽选择正文，选中后会自动高亮，再在右侧填写批注。</span>
          </div>

          <div class="preview-tabs">
            <button type="button" :class="{ active: activeTab === 'mammoth' }" @click="switchTab('mammoth')">
              mammoth 预览
            </button>
            <button type="button" :class="{ active: activeTab === 'aspose' }" @click="switchTab('aspose')">
              Python Aspose 预览
            </button>
            <button type="button" :class="{ active: activeTab === 'javaAspose' }" @click="switchTab('javaAspose')">
              Java Aspose 预览
            </button>
          </div>
        </header>

        <div v-show="activeTab === 'mammoth'" ref="mammothPreviewRef" class="doc-preview" v-html="previews.mammoth.html"></div>
        <div v-show="activeTab === 'aspose'" ref="asposePreviewRef" class="doc-preview" v-html="previews.aspose.html"></div>
        <div v-show="activeTab === 'javaAspose'" ref="javaAsposePreviewRef" class="doc-preview" v-html="previews.javaAspose.html"></div>

        <div v-if="!currentPreview.html" class="tab-empty">
          当前 Tab 还没有预览内容，请使用上方对应入口上传 DOCX。
        </div>
      </article>

      <aside class="panel">
        <section class="current">
          <h2>当前选区</h2>
          <p v-if="!currentPreview.draft" class="muted">请先在当前 Tab 的预览区划词。</p>
          <blockquote v-else>{{ currentPreview.draft.selectedText }}</blockquote>
        </section>

        <label class="field">
          批注内容
          <textarea v-model="currentPreview.commentText" rows="5" placeholder="请输入批注内容"></textarea>
        </label>

        <label class="field">
          作者
          <input v-model="author" placeholder="Reviewer" />
        </label>

        <button type="button" :disabled="!canSaveComment" @click="saveComment">
          保存当前批注
        </button>

        <section class="comments">
          <h2>批注列表</h2>
          <p v-if="currentPreview.comments.length === 0" class="muted">当前 Tab 暂无批注。</p>
          <article v-for="comment in currentPreview.comments" :key="comment.id" class="comment-item">
            <strong>{{ comment.selectedText }}</strong>
            <p>{{ comment.comment }}</p>
            <small>第 {{ comment.occurrence + 1 }} 次出现 · {{ comment.author }}</small>
            <button type="button" class="text-button" @click="removeComment(comment.id)">删除</button>
          </article>
        </section>

        <div class="export-actions">
          <button class="primary" type="button" @click="exportDocx('python-docx')">
            {{ exporting === 'python-docx' ? '生成中...' : '生成批注文件（python-docx）' }}
          </button>
          <button class="primary aspose" type="button" @click="exportDocx('aspose')">
            {{ exporting === 'aspose' ? '生成中...' : '生成批注文件（Aspose）' }}
          </button>
          <button class="primary java-aspose" type="button" @click="exportDocx('java-aspose')">
            {{ exporting === 'java-aspose' ? '生成中...' : '生成批注文件（Java Aspose）' }}
          </button>
        </div>
      </aside>
    </section>

    <section v-else class="empty-state">
      <p>请用 mammoth 或 Aspose 入口上传一个 `.docx` 文件。</p>
    </section>

    <p v-if="message" class="message">{{ message }}</p>
  </main>
</template>

<script setup>
import Highlighter from 'web-highlighter';
import { computed, nextTick, onBeforeUnmount, reactive, ref } from 'vue';
import mammoth from 'mammoth/mammoth.browser';

// 默认走 Vite 代理：
// /api/python-docx -> http://localhost:5000
// /api/aspose -> http://localhost:5001
// /api/java-aspose -> http://localhost:8080
// 如需直连后端，可在 .env 中设置 VITE_API_BASE。
const API_BASE = import.meta.env.VITE_API_BASE ?? '';

const activeTab = ref('mammoth');
const author = ref('Reviewer');
const message = ref('');
const exporting = ref('');

const mammothPreviewRef = ref(null);
const asposePreviewRef = ref(null);
const javaAsposePreviewRef = ref(null);
const previewRefs = {
  mammoth: mammothPreviewRef,
  aspose: asposePreviewRef,
  javaAspose: javaAsposePreviewRef
};

const highlighters = {
  mammoth: null,
  aspose: null,
  javaAspose: null
};

const previews = reactive({
  mammoth: createPreviewState('mammoth 前端预览'),
  aspose: createPreviewState('Python Aspose 后端预览'),
  javaAspose: createPreviewState('Java Aspose 后端预览')
});

const currentPreview = computed(() => previews[activeTab.value]);
const hasAnyDocument = computed(() => Boolean(previews.mammoth.html || previews.aspose.html || previews.javaAspose.html));
const canSaveComment = computed(() => Boolean(currentPreview.value.draft && currentPreview.value.commentText.trim()));

function createPreviewState(label) {
  return {
    label,
    sourceFile: null,
    uploadId: '',
    fileName: '',
    html: '',
    comments: [],
    draft: null,
    commentText: ''
  };
}

async function handleMammothFileChange(event) {
  // 同一个 input 再次选择同名文件时，浏览器可能不触发 change，因此先清空 value。
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  activeTab.value = 'mammoth';
  resetPreviewState('mammoth', file);
  message.value = '正在使用 mammoth 解析 DOCX...';

  try {
    // mammoth 在浏览器端直接把 DOCX ArrayBuffer 转成 HTML，不需要先上传后端。
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    previews.mammoth.html = result.value;
    message.value = result.messages.length ? 'DOCX 已预览，部分格式可能被 mammoth 简化。' : '';

    // v-html 渲染是异步更新 DOM 的，等预览内容真实挂载后再初始化对应 Tab 的高亮器。
    await nextTick();
    setupHighlighter('mammoth');
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

async function handleAsposeFileChange(event) {
  // 同一个 input 再次选择同名文件时，浏览器可能不触发 change，因此先清空 value。
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  activeTab.value = 'aspose';
  resetPreviewState('aspose', file);
  message.value = '正在上传 DOCX 并使用 Aspose 解析...';

  try {
    // Aspose 预览走后端解析，前端只负责展示返回的 HTML。
    const formData = new FormData();
    formData.append('file', file);
    const response = await fetch(`${API_BASE}/api/aspose/docx/upload`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    const result = await response.json();
    previews.aspose.uploadId = result.uploadId;
    previews.aspose.fileName = result.fileName || file.name;
    previews.aspose.html = result.html;
    message.value = '';

    // v-html 渲染是异步更新 DOM 的，等预览内容真实挂载后再初始化对应 Tab 的高亮器。
    await nextTick();
    setupHighlighter('aspose');
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

async function handleJavaAsposeFileChange(event) {
  // 同一个 input 再次选择同名文件时，浏览器可能不触发 change，因此先清空 value。
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  activeTab.value = 'javaAspose';
  resetPreviewState('javaAspose', file);
  message.value = '正在上传 DOCX 并使用 Java Aspose 解析...';

  try {
    // Java Aspose 预览走 Spring Boot 后端解析，前端只负责展示返回的 HTML。
    const formData = new FormData();
    formData.append('file', file);
    const response = await fetch(`${API_BASE}/api/java-aspose/docx/upload`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    const result = await response.json();
    previews.javaAspose.uploadId = result.uploadId;
    previews.javaAspose.fileName = result.fileName || file.name;
    previews.javaAspose.html = result.html;
    message.value = '';

    // v-html 渲染是异步更新 DOM 的，等预览内容真实挂载后再初始化对应 Tab 的高亮器。
    await nextTick();
    setupHighlighter('javaAspose');
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

function resetPreviewState(mode, file) {
  destroyHighlighter(mode);
  Object.assign(previews[mode], {
    sourceFile: file,
    uploadId: '',
    fileName: file.name,
    html: '',
    comments: [],
    draft: null,
    commentText: ''
  });
}

function switchTab(mode) {
  activeTab.value = mode;
  message.value = '';
}

function setupHighlighter(mode) {
  destroyHighlighter(mode);
  const root = previewRefs[mode].value;
  if (!root) return;

  // 每个 Tab 有自己的 root 和 highlighter，互不复用预览 DOM。
  const highlighter = new Highlighter({
    $root: root,
    exceptSelectors: ['button', 'input', 'textarea']
  });

  highlighter.on(Highlighter.event.CREATE, ({ sources }) => {
    const source = sources?.[0];
    const selectedText = source?.text?.replace(/\s+/g, ' ').trim();
    if (!source || !selectedText) return;

    previews[mode].draft = {
      id: source.id,
      source,
      selectedText,
      // 后端按“文本 + 第几次出现”重新定位到原始 DOCX。
      occurrence: countOccurrenceBeforeSelection(mode, selectedText)
    };
    previews[mode].commentText = '';
    message.value = '';
  });

  highlighter.run();
  highlighters[mode] = highlighter;
}

function destroyHighlighter(mode) {
  // 当前库没有强制销毁逻辑；置空引用，等待新预览重新初始化。
  highlighters[mode] = null;
}

function countOccurrenceBeforeSelection(mode, selectedText) {
  // 统计当前选中文本在对应 Tab 预览区域里已经出现过几次，用于区分重复文本。
  const selection = window.getSelection();
  const root = previewRefs[mode].value;
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

function saveComment() {
  // 把当前 Tab 的选区、批注内容和 highlighter source 保存起来，后续统一提交后端。
  if (!canSaveComment.value) return;

  currentPreview.value.comments.push({
    id: currentPreview.value.draft.id,
    selectedText: currentPreview.value.draft.selectedText,
    occurrence: currentPreview.value.draft.occurrence,
    comment: currentPreview.value.commentText.trim(),
    author: author.value.trim() || 'Reviewer',
    initials: getInitials(author.value),
    highlightSource: currentPreview.value.draft.source
  });

  currentPreview.value.draft = null;
  currentPreview.value.commentText = '';
}

function removeComment(id) {
  // 这里只删除右侧批注数据；已绘制的高亮保留，便于用户知道曾经选中过哪里。
  currentPreview.value.comments = currentPreview.value.comments.filter((comment) => comment.id !== id);
}

async function exportDocx(backend) {
  const preview = currentPreview.value;
  if (!preview.sourceFile) {
    message.value = '请先在当前 Tab 上传 DOCX';
    return;
  }
  if (exporting.value) return;

  exporting.value = backend;
  message.value = '';

  try {
    // 如果用户还没点“保存当前批注”，导出前自动保存当前可用的草稿。
    if (canSaveComment.value) {
      saveComment();
    }
    if (preview.comments.length === 0) {
      message.value = '请先在当前 Tab 划词并填写批注内容';
      return;
    }

    const response =
      backend === 'aspose'
        ? await exportWithAsposeUpload(preview)
        : backend === 'java-aspose'
          ? await exportWithJavaAsposeUpload(preview)
        : await exportWithOriginalFile(backend, preview);

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    // 后端返回的是生成后的 docx blob，创建临时下载链接触发浏览器下载。
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = preview.fileName.replace(/\.docx$/i, `-${activeTab.value}-${backend}-comments.docx`);
    link.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    message.value = error.message || '生成失败';
  } finally {
    exporting.value = '';
  }
}

async function exportWithAsposeUpload(preview) {
  if (!preview.uploadId) {
    return exportWithOriginalFile('aspose', preview);
  }

  return fetch(`${API_BASE}/api/aspose/docx/${preview.uploadId}/comments`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ comments: preview.comments })
  });
}

async function exportWithJavaAsposeUpload(preview) {
  if (!preview.uploadId) {
    return exportWithOriginalFile('java-aspose', preview);
  }

  return fetch(`${API_BASE}/api/java-aspose/docx/${preview.uploadId}/comments`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ comments: preview.comments })
  });
}

async function exportWithOriginalFile(backend, preview) {
  // mammoth Tab 没有后端 uploadId，所以仍使用原始 DOCX + 批注 JSON 一次提交。
  const formData = new FormData();
  formData.append('file', preview.sourceFile);
  formData.append('comments', JSON.stringify(preview.comments));

  return fetch(`${API_BASE}/api/${backend}/comments/generate`, {
    method: 'POST',
    body: formData
  });
}

function normalizeText(value) {
  // 折叠空白，降低 HTML 预览文本和 Word 原文之间的空白差异。
  return value.replace(/\s+/g, ' ').trim();
}

function getInitials(value) {
  // Word 批注需要 initials 字段；没有作者时使用默认 RV。
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
  // 优先读取后端 JSON message；如果不是 JSON，就直接显示文本响应。
  const text = await response.text();
  try {
    const payload = JSON.parse(text);
    return payload.message || response.statusText || '生成失败';
  } catch {
    return text || response.statusText || '生成失败';
  }
}

onBeforeUnmount(() => {
  destroyHighlighter('mammoth');
  destroyHighlighter('aspose');
  destroyHighlighter('javaAspose');
});
</script>
