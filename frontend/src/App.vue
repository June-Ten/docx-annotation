<template>
  <main class="page">
    <section class="toolbar">
      <div>
        <p class="eyebrow">DOCX Comment Editor</p>
        <h1>DOCX 划词高亮批注</h1>
        <p>前端使用 mammoth.js 预览 DOCX，web-highlighter 负责划词高亮，Python 后端生成带 Word 批注的文件。</p>
      </div>

      <label class="upload-button">
        <input type="file" accept=".docx" @change="handleFileChange" />
        选择 DOCX
      </label>
    </section>

    <section v-if="fileName" class="workspace">
      <article class="preview-card">
        <header>
          <div>
            <strong>{{ fileName }}</strong>
            <span>拖拽选择正文，选中后会自动高亮，再在右侧填写批注。</span>
          </div>
        </header>

        <div ref="previewRef" class="doc-preview" v-html="html"></div>
      </article>

      <aside class="panel">
        <section class="current">
          <h2>当前选区</h2>
          <p v-if="!draft" class="muted">请先在预览区划词。</p>
          <blockquote v-else>{{ draft.selectedText }}</blockquote>
        </section>

        <label class="field">
          批注内容
          <textarea v-model="commentText" rows="5" placeholder="请输入批注内容"></textarea>
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
          <p v-if="comments.length === 0" class="muted">暂无批注。</p>
          <article v-for="comment in comments" :key="comment.id" class="comment-item">
            <strong>{{ comment.selectedText }}</strong>
            <p>{{ comment.comment }}</p>
            <small>第 {{ comment.occurrence + 1 }} 次出现 · {{ comment.author }}</small>
            <button type="button" class="text-button" @click="removeComment(comment.id)">删除</button>
          </article>
        </section>

        <div class="export-actions">
          <button
            class="primary"
            type="button"
            @click="exportDocx('python-docx')"
          >
            {{ exporting === 'python-docx' ? '生成中...' : '生成批注文件（python-docx）' }}
          </button>
          <button
            class="primary aspose"
            type="button"
            @click="exportDocx('aspose')"
          >
            {{ exporting === 'aspose' ? '生成中...' : '生成批注文件（Aspose）' }}
          </button>
        </div>
      </aside>
    </section>

    <section v-else class="empty-state">
      <p>上传一个 `.docx` 文件后开始预览和批注。</p>
    </section>

    <p v-if="message" class="message">{{ message }}</p>
  </main>
</template>

<script setup>
import Highlighter from 'web-highlighter';
import { computed, nextTick, onBeforeUnmount, ref } from 'vue';
import mammoth from 'mammoth/mammoth.browser';

// 默认走 Vite 代理：
// /api/python-docx -> http://localhost:5000
// /api/aspose -> http://localhost:5001
// 如需直连后端，可在 .env 中设置 VITE_API_BASE。
const API_BASE = import.meta.env.VITE_API_BASE ?? '';

// DOCX 预览容器和当前文件状态。
const previewRef = ref(null);
const sourceFile = ref(null);
const fileName = ref('');
const html = ref('');

// comments 是最终提交给后端的批注列表；draft 是当前刚划词但还未保存的批注。
const comments = ref([]);
const draft = ref(null);
const commentText = ref('');
const author = ref('Reviewer');
const message = ref('');
const exporting = ref('');

let highlighter = null;

// 必须已经划词并填写批注内容，才能保存一条批注。
const canSaveComment = computed(() => Boolean(draft.value && commentText.value.trim()));

async function handleFileChange(event) {
  // 同一个 input 再次选择同名文件时，浏览器可能不触发 change，因此先清空 value。
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  sourceFile.value = file;
  fileName.value = file.name;
  comments.value = [];
  draft.value = null;
  commentText.value = '';
  message.value = '正在解析 DOCX...';

  try {
    // mammoth 在浏览器端直接把 DOCX ArrayBuffer 转成 HTML，不需要先上传后端。
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    html.value = result.value;
    message.value = result.messages.length ? 'DOCX 已预览，部分格式可能被 mammoth 简化。' : '';

    // v-html 渲染是异步更新 DOM 的，等预览内容真实挂载后再初始化高亮器。
    await nextTick();
    setupHighlighter();
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

function setupHighlighter() {
  // 每次上传新文件都会重新渲染预览，需要重新绑定 highlighter。
  destroyHighlighter();
  if (!previewRef.value) return;

  // 只允许在预览区域内高亮，排除按钮、输入框等交互元素。
  highlighter = new Highlighter({
    $root: previewRef.value,
    exceptSelectors: ['button', 'input', 'textarea']
  });

  // 用户划词后，web-highlighter 会自动给选区加高亮，并返回可序列化的 source。
  highlighter.on(Highlighter.event.CREATE, ({ sources }) => {
    const source = sources?.[0];
    const selectedText = source?.text?.replace(/\s+/g, ' ').trim();
    if (!source || !selectedText) return;

    draft.value = {
      id: source.id,
      source,
      selectedText,
      // 后端按“文本 + 第几次出现”重新定位到原始 DOCX。
      occurrence: countOccurrenceBeforeSelection(selectedText)
    };
    commentText.value = '';
    message.value = '';
  });

  highlighter.run();
}

function destroyHighlighter() {
  // 当前库没有强制销毁逻辑；置空引用，等待新预览重新初始化。
  highlighter = null;
}

function countOccurrenceBeforeSelection(selectedText) {
  // 统计当前选中文本在预览区域里已经出现过几次，用于区分重复文本。
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

function saveComment() {
  // 把当前选区、批注内容和 highlighter source 保存起来，后续统一提交后端。
  if (!canSaveComment.value) return;

  comments.value.push({
    id: draft.value.id,
    selectedText: draft.value.selectedText,
    occurrence: draft.value.occurrence,
    comment: commentText.value.trim(),
    author: author.value.trim() || 'Reviewer',
    initials: getInitials(author.value),
    highlightSource: draft.value.source
  });

  draft.value = null;
  commentText.value = '';
}

function removeComment(id) {
  // 这里只删除右侧批注数据；已绘制的高亮保留，便于用户知道曾经选中过哪里。
  comments.value = comments.value.filter((comment) => comment.id !== id);
}

async function exportDocx(backend) {
  if (!sourceFile.value) return;
  if (exporting.value) return;

  exporting.value = backend;
  message.value = '';

  try {
    // 如果用户还没点“保存当前批注”，导出前自动保存当前可用的草稿。
    if (canSaveComment.value) {
      saveComment();
    }
    if (comments.value.length === 0) {
      message.value = '请先划词并填写批注内容';
      return;
    }

    // 后端需要原始 DOCX 文件和批注 JSON，使用 FormData 一次提交。
    const formData = new FormData();
    formData.append('file', sourceFile.value);
    formData.append('comments', JSON.stringify(comments.value));

    const response = await fetch(`${API_BASE}/api/${backend}/comments/generate`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    // 后端返回的是生成后的 docx blob，创建临时下载链接触发浏览器下载。
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName.value.replace(/\.docx$/i, `-${backend}-comments.docx`);
    link.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    message.value = error.message || '生成失败';
  } finally {
    exporting.value = '';
  }
}

function normalizeText(value) {
  // 折叠空白，降低 mammoth HTML 文本和 Word 原文之间的空白差异。
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
  destroyHighlighter();
});
</script>
