const appShell = document.querySelector('.app-shell');
const sidebar = document.querySelector('.sidebar');
const overlay = document.querySelector('.sidebar-overlay');
const mobileMenuToggle = document.querySelector('.mobile-menu-toggle');
const navItems = document.querySelectorAll('.nav-item');
const pageContent = document.querySelector('.page-content');
const startStudyingButton = document.querySelector('.primary-button');
const chatApp = document.querySelector('.chat-app');
const exerciseApp = document.getElementById('exercise-app');
const learningApp = document.getElementById('learning-app');
const settingsApp = document.getElementById('settings-app');
const exerciseChatToggle = document.getElementById('exercise-chat-toggle');
const exerciseChatResizer = document.getElementById('exercise-chat-resizer');
const exerciseChatPanel = document.getElementById('exercise-chat-panel');
const chatHistoryList = document.getElementById('chat-history-list');
const chatHistoryToggle = document.getElementById('chat-history-toggle');
const chatMessages = document.getElementById('chat-messages');
const chatForm = document.getElementById('chat-form');
const chatInput = document.getElementById('chat-input');
const aiServerStatus = document.getElementById('ai-server-status');
const suggestionRow = document.getElementById('chat-suggestions');
const activeChatTitle = document.getElementById('active-chat-title');
const renameModal = document.getElementById('rename-modal');
const deleteModal = document.getElementById('delete-modal');
const renameForm = document.getElementById('rename-form');
const renameInput = document.getElementById('rename-input');
const confirmDeleteBtn = document.getElementById('confirm-delete-btn');
const deleteModalText = document.getElementById('delete-modal-text');
const searchInput = document.getElementById('chat-search');
const historyButtons = document.querySelectorAll('.history-view');
const newChatBtn = document.querySelector('.new-chat-btn');
const bulkDeleteBtn = document.getElementById('bulk-delete-btn');
const selectAllBtn = document.getElementById('select-all-btn');

const STORAGE_KEY = 'formulaX-chat-history';
const CHAT_API_URL = window.FORMULAXL_CHAT_API_URL || 'http://127.0.0.1:8000/api/future-ai/chat';
const EXERCISE_CORRECTION_API_URL = window.FORMULAXL_EXERCISE_CORRECTION_API_URL
  || new URL('/api/future-ai/correct-exercise', CHAT_API_URL).href;
const EXERCISE_WORKBOOK_API_URL = window.FORMULAXL_EXERCISE_WORKBOOK_API_URL
  || new URL('/api/future-ai/generate-workbook', CHAT_API_URL).href;
const CHAT_API_OPENAPI_URL = new URL('/openapi.json', CHAT_API_URL).href;
let activeChatId = null;
let activeView = 'active';
let selectedChatIds = new Set();
let pendingDeleteChatId = null;
let pendingRenameChatId = null;
let isGenerating = false;
let generatingChatId = null;
let backendCheckInProgress = false;

const defaultConversation = {
  id: crypto.randomUUID(),
  title: 'Nova conversa',
  archived: false,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  messages: [
    {
      id: crypto.randomUUID(),
      role: 'assistant',
      content:
        'Olá! Sou o Axel AI. Posso te ajudar com fórmulas, funções, organização de planilhas, análise de dados e exercícios práticos de Excel.',
      createdAt: Date.now()
    }
  ]
};

let state = {
  chats: [defaultConversation]
};

function loadChats() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    return;
  }

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed.chats) && parsed.chats.length) {
      state = parsed;
      return;
    }
  } catch (error) {
    console.warn('Erro ao carregar histórico local:', error);
  }

  state = { chats: [defaultConversation] };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function saveChats() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function formatDateLabel(dateValue) {
  const date = new Date(dateValue);
  return date.toLocaleDateString('pt-BR', {
    day: '2-digit',
    month: 'short'
  });
}

function escapeHTML(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[character]);
}

function renderInlineMarkdown(value) {
  const codeSpans = [];
  let text = escapeHTML(value).replace(/`([^`\n]+)`/g, (_, code) => {
    const index = codeSpans.push(`<code>${code}</code>`) - 1;
    return `\u0000${index}\u0000`;
  });

  text = text
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/__(.+?)__/g, '<strong>$1</strong>')
    .replace(/(?<!\*)\*([^*\n]+)\*(?!\*)/g, '<em>$1</em>')
    .replace(/(?<!_)_([^_\n]+)_(?!_)/g, '<em>$1</em>');

  return text.replace(/\u0000(\d+)\u0000/g, (_, index) => codeSpans[Number(index)]);
}

function parseMarkdownTableRow(line) {
  const trimmed = line.trim().replace(/^\|+|\|+$/g, '');
  return trimmed.split(/\|+/).map((cell) => cell.trim());
}

function getMarkdownTableAlignment(separator) {
  if (separator.startsWith(':') && separator.endsWith(':')) return 'center';
  if (separator.endsWith(':')) return 'right';
  return 'left';
}

function renderAssistantMarkdown(value) {
  const lines = String(value).replace(/\r\n?/g, '\n').split('\n');
  const blocks = [];
  let paragraph = [];
  let listItems = [];
  let listType = null;
  let codeLines = null;

  const flushParagraph = () => {
    if (paragraph.length) {
      blocks.push(`<p>${paragraph.map(renderInlineMarkdown).join('<br>')}</p>`);
      paragraph = [];
    }
  };

  const flushList = () => {
    if (listItems.length) {
      const tag = listType === 'ordered' ? 'ol' : 'ul';
      blocks.push(`<${tag}>${listItems.map((item) => `<li>${renderInlineMarkdown(item)}</li>`).join('')}</${tag}>`);
      listItems = [];
      listType = null;
    }
  };

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const fenceMatch = line.match(/^\s*```/);
    if (fenceMatch) {
      flushParagraph();
      flushList();
      if (codeLines) {
        blocks.push(`<pre><code>${escapeHTML(codeLines.join('\n'))}</code></pre>`);
        codeLines = null;
      } else {
        codeLines = [];
      }
      continue;
    }

    if (codeLines) {
      codeLines.push(line);
      continue;
    }

    const separatorLine = lines[index + 1];
    const separatorCells = separatorLine && separatorLine.includes('|')
      ? parseMarkdownTableRow(separatorLine)
      : [];
    const isTableSeparator = separatorCells.length > 1
      && separatorCells.every((cell) => /^:?-{3,}:?$/.test(cell));

    if (line.includes('|') && isTableSeparator) {
      flushParagraph();
      flushList();
      const headers = parseMarkdownTableRow(line);
      const alignments = separatorCells.map(getMarkdownTableAlignment);
      const headerCells = headers.map((cell, cellIndex) => {
        const alignment = alignments[cellIndex] || 'left';
        return `<th scope="col" style="text-align:${alignment}">${renderInlineMarkdown(cell)}</th>`;
      }).join('');
      const rows = [];
      index += 2;

      while (index < lines.length && lines[index].trim() && lines[index].includes('|')) {
        const cells = parseMarkdownTableRow(lines[index]);
        rows.push(`<tr>${headers.map((_, cellIndex) => {
          const alignment = alignments[cellIndex] || 'left';
          return `<td style="text-align:${alignment}">${renderInlineMarkdown(cells[cellIndex] || '')}</td>`;
        }).join('')}</tr>`);
        index += 1;
      }

      index -= 1;
      blocks.push(
        `<div class="message-table-wrap"><table><thead><tr>${headerCells}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`
      );
      continue;
    }

    const headingMatch = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
    const unorderedMatch = line.match(/^\s*[-*+]\s+(.+)$/);
    const orderedMatch = line.match(/^\s*\d+[.)]\s+(.+)$/);

    if (!line.trim()) {
      flushParagraph();
      flushList();
      continue;
    }

    if (headingMatch) {
      flushParagraph();
      flushList();
      const level = headingMatch[1].length;
      blocks.push(`<h${level}>${renderInlineMarkdown(headingMatch[2])}</h${level}>`);
      continue;
    }

    if (unorderedMatch || orderedMatch) {
      flushParagraph();
      const nextListType = orderedMatch ? 'ordered' : 'unordered';
      if (listType && listType !== nextListType) flushList();
      listType = nextListType;
      listItems.push((unorderedMatch || orderedMatch)[1]);
      continue;
    }

    flushList();
    paragraph.push(line);
  }

  if (codeLines) blocks.push(`<pre><code>${escapeHTML(codeLines.join('\n'))}</code></pre>`);
  flushParagraph();
  flushList();
  return blocks.join('');
}

function getVisibleChats() {
  const query = searchInput.value.trim().toLowerCase();

  return state.chats.filter((chat) => {
    const matchesView = chat.archived === (activeView === 'archived');
    const matchesSearch = !query || chat.title.toLowerCase().includes(query) || chat.messages.some((message) => message.content.toLowerCase().includes(query));
    return matchesView && matchesSearch;
  }).sort((a, b) => b.updatedAt - a.updatedAt);
}

function createChat(title = 'Nova conversa') {
  const chat = {
    id: crypto.randomUUID(),
    title,
    archived: false,
    responseId: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    messages: [
      {
        id: crypto.randomUUID(),
        role: 'assistant',
        content: 'Olá! Sou o Axel AI. Como posso ajudar você hoje com Excel?',
        createdAt: Date.now()
      }
    ]
  };

  state.chats.unshift(chat);
  activeChatId = chat.id;
  saveChats();
  render();
}

function getActiveChat() {
  return state.chats.find((chat) => chat.id === activeChatId) || state.chats[0] || null;
}

function updateChatMeta(chatId) {
  const chat = state.chats.find((item) => item.id === chatId);
  if (!chat) return;
  chat.updatedAt = Date.now();
  saveChats();
}

function renderHistory() {
  const chats = getVisibleChats();

  if (!chats.length) {
    chatHistoryList.innerHTML = `
      <div class="empty-history">
        <p>Nenhuma conversa encontrada.</p>
      </div>
    `;
    return;
  }

  chatHistoryList.innerHTML = chats.map((chat) => {
    const lastMessage = chat.messages[chat.messages.length - 1];
    const selected = selectedChatIds.has(chat.id) ? 'selected' : '';
    const active = activeChatId === chat.id ? 'active' : '';

    return `
      <article class="chat-history-item ${selected} ${active}" data-chat-id="${escapeHTML(chat.id)}" tabindex="0" aria-label="Selecionar conversa ${escapeHTML(chat.title)}">
        <div class="history-check">
          <input type="checkbox" aria-label="Selecionar ${escapeHTML(chat.title)}" data-select-chat="${escapeHTML(chat.id)}" ${selected ? 'checked' : ''} />
        </div>
        <div class="history-content">
          <div class="history-header-row">
            <h4>${escapeHTML(chat.title)}</h4>
            <span>${formatDateLabel(chat.updatedAt)}</span>
          </div>
          <p>${lastMessage ? escapeHTML(lastMessage.content.slice(0, 80)) : 'Sem mensagens'}${lastMessage && lastMessage.content.length > 80 ? '…' : ''}</p>
        </div>
      </article>
    `;
  }).join('');

  chatHistoryList.querySelectorAll('.chat-history-item').forEach((item) => {
    item.addEventListener('click', (event) => {
      if (event.target.matches('input')) return;
      const chatId = item.dataset.chatId;
      activeChatId = chatId;
      render();
    });

    item.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        const chatId = item.dataset.chatId;
        activeChatId = chatId;
        render();
      }
    });
  });

  chatHistoryList.querySelectorAll('[data-select-chat]').forEach((checkbox) => {
    checkbox.addEventListener('change', (event) => {
      const chatId = event.target.dataset.selectChat;
      if (event.target.checked) {
        selectedChatIds.add(chatId);
      } else {
        selectedChatIds.delete(chatId);
      }
      renderBulkActions();
    });
  });
}

function renderBulkActions() {
  const hasSelection = selectedChatIds.size > 0;
  bulkDeleteBtn.disabled = !hasSelection;
  bulkDeleteBtn.classList.toggle('disabled', !hasSelection);
}

function renderMessages() {
  const chat = getActiveChat();
  if (!chat) {
    chatMessages.innerHTML = '<div class="empty-state">Nenhuma conversa aberta.</div>';
    activeChatTitle.textContent = 'Nova conversa';
    return;
  }

  activeChatTitle.textContent = chat.title;

  if (!chat.messages.length) {
    chatMessages.innerHTML = '<div class="empty-state">Comece a conversa com uma pergunta sobre Excel.</div>';
    return;
  }

  const messagesHTML = chat.messages.map((message) => `
    <div class="message-row ${message.role === 'user' ? 'user' : 'assistant'}">
      <div class="message-bubble ${message.role === 'user' ? 'user' : 'assistant'}">
        <div class="message-content">${message.role === 'assistant' ? renderAssistantMarkdown(message.content) : escapeHTML(message.content)}</div>
      </div>
    </div>
  `).join('');

  const typingIndicator = isGenerating && generatingChatId === chat.id
    ? '<div class="message-row assistant typing-row" role="status" aria-label="Axel AI está respondendo"><div class="typing-bubble"><span></span><span></span><span></span></div></div>'
    : '';
  chatMessages.innerHTML = `${messagesHTML}${typingIndicator}`;
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function syncSelectAllButton() {
  const visibleChats = getVisibleChats();
  const allVisibleSelected = visibleChats.length > 0 && visibleChats.every((chat) => selectedChatIds.has(chat.id));

  selectAllBtn.textContent = allVisibleSelected ? 'Desmarcar todos' : 'Selecionar todos';
  selectAllBtn.setAttribute('aria-pressed', String(allVisibleSelected));
}

function toggleSelectAllVisibleChats() {
  const visibleChats = getVisibleChats();
  const allVisibleSelected = visibleChats.length > 0 && visibleChats.every((chat) => selectedChatIds.has(chat.id));

  if (allVisibleSelected) {
    visibleChats.forEach((chat) => selectedChatIds.delete(chat.id));
  } else {
    visibleChats.forEach((chat) => selectedChatIds.add(chat.id));
  }

  render();
}

function render() {
  renderHistory();
  renderMessages();
  renderBulkActions();
  updateActionButtons();
  syncSelectAllButton();
  syncSuggestionVisibility();
}

function syncSuggestionVisibility() {
  const chat = getActiveChat();
  const hasUserMessage = !!chat && chat.messages.some((message) => message.role === 'user');

  if (suggestionRow) {
    suggestionRow.style.display = hasUserMessage ? 'none' : 'flex';
  }
}

function addUserMessage(content) {
  const chat = getActiveChat();
  if (!chat) return;

  const message = {
    id: crypto.randomUUID(),
    role: 'user',
    content,
    createdAt: Date.now()
  };

  chat.messages.push(message);
  chat.updatedAt = Date.now();
  saveChats();
  render();
}

function openRenameModal() {
  const chat = getActiveChat();
  if (!chat) return;
  pendingRenameChatId = chat.id;
  renameInput.value = chat.title;
  renameModal.classList.remove('hidden');
  renameModal.setAttribute('aria-hidden', 'false');
  renameInput.focus();
}

function closeRenameModal() {
  renameModal.classList.add('hidden');
  renameModal.setAttribute('aria-hidden', 'true');
  renameInput.value = '';
  pendingRenameChatId = null;
}

function openDeleteModal(chatId) {
  pendingDeleteChatId = chatId;
  const chat = state.chats.find((item) => item.id === chatId);
  deleteModalText.textContent = chat ? `Deseja excluir a conversa “${chat.title}”?` : 'Deseja excluir esta conversa?';
  deleteModal.classList.remove('hidden');
  deleteModal.setAttribute('aria-hidden', 'false');
}

function closeDeleteModal() {
  deleteModal.classList.add('hidden');
  deleteModal.setAttribute('aria-hidden', 'true');
  pendingDeleteChatId = null;
}

function deleteChat(chatId) {
  state.chats = state.chats.filter((chat) => chat.id !== chatId);
  selectedChatIds.delete(chatId);

  if (activeChatId === chatId) {
    activeChatId = state.chats[0]?.id || null;
  }

  saveChats();
  render();
}

function deleteSelectedChats() {
  if (!selectedChatIds.size) return;
  const ids = [...selectedChatIds];
  state.chats = state.chats.filter((chat) => !ids.includes(chat.id));

  if (activeChatId && ids.includes(activeChatId)) {
    activeChatId = state.chats[0]?.id || null;
  }

  selectedChatIds = new Set();
  saveChats();
  render();
}

function archiveChat(chatId) {
  const chat = state.chats.find((item) => item.id === chatId);
  if (!chat) return;
  chat.archived = true;
  chat.updatedAt = Date.now();
  if (activeChatId === chatId && activeView === 'active') {
    activeChatId = state.chats.find((item) => item.id !== chatId)?.id || null;
  }
  saveChats();
  render();
}

function restoreChat(chatId) {
  const chat = state.chats.find((item) => item.id === chatId);
  if (!chat) return;
  chat.archived = false;
  chat.updatedAt = Date.now();
  saveChats();
  render();
}

function updateActionButtons() {
  const chat = getActiveChat();
  const archiveBtn = document.querySelector('[data-action="archive"]');
  const renameBtn = document.querySelector('[data-action="rename"]');

  if (!chat) {
    archiveBtn.textContent = 'Arquivar';
    renameBtn.textContent = 'Renomear';
    return;
  }

  archiveBtn.textContent = chat.archived ? 'Restaurar' : 'Arquivar';
  renameBtn.textContent = 'Renomear';
}

function toggleChatView(view) {
  activeView = view;
  historyButtons.forEach((button) => {
    button.classList.toggle('is-active', button.dataset.view === activeView);
  });
  render();
}

async function sendChatMessage(text) {
  const message = text.trim();
  if (!message || isGenerating) return null;
  if (!getActiveChat()) {
    createChat('Nova conversa');
  }

  const chat = getActiveChat();
  if (!chat) return null;

  isGenerating = true;
  generatingChatId = chat.id;
  addUserMessage(message);
  try {
    return await sendMessageToBackend(chat.id, message);
  } finally {
    isGenerating = false;
    generatingChatId = null;
    render();
  }
}

function handleSendMessage(event) {
  event.preventDefault();
  const text = chatInput.value.trim();
  if (!text || isGenerating) return;
  chatInput.value = '';
  handleChatInput();
  sendChatMessage(text);
}

async function correctExerciseWorkbook(file, challengeDetails) {
  if (!(file instanceof File) || !file.size || !challengeDetails?.text?.trim() || isGenerating) return null;
  if (!getActiveChat()) createChat('Nova conversa');

  const chat = getActiveChat();
  if (!chat) return null;

  isGenerating = true;
  generatingChatId = chat.id;
  chat.messages.push({
    id: crypto.randomUUID(),
    role: 'user',
    content: `Enviei a planilha "${file.name}" para correção do desafio.`,
    createdAt: Date.now()
  });
  chat.updatedAt = Date.now();
  saveChats();
  render();

  try {
    const learningHistory = window.FormulaXLearning.getHistory();
    if (!learningHistory.ok) throw new Error(learningHistory.error);

    const formData = new FormData();
    formData.append('file', file, file.name);
    formData.append('challenge', challengeDetails.text);
    formData.append('challenge_id', challengeDetails.id);
    formData.append('challenge_title', challengeDetails.title);
    formData.append('difficulty', challengeDetails.difficulty);
    formData.append('functions', JSON.stringify(challengeDetails.functions));
    formData.append('previous_response_id', chat.responseId || '');
    formData.append('data_aprendizado', JSON.stringify({
      version: 1,
      corrections: learningHistory.corrections
    }));

    const response = await fetch(EXERCISE_CORRECTION_API_URL, {
      method: 'POST',
      body: formData
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || 'Não foi possível corrigir o exercício.');
    }
    if (typeof result.response !== 'string' || typeof result.response_id !== 'string') {
      throw new Error('O backend retornou uma resposta de correção inválida.');
    }

    const learningResult = result.learning_result;
    if (
      !learningResult
      || !['correct', 'incorrect', 'ungradable'].includes(learningResult.status)
      || typeof learningResult.feedback !== 'string'
    ) {
      throw new Error('O backend não retornou uma avaliação estruturada para o histórico de aprendizado.');
    }

    const recordedLearning = window.FormulaXLearning.recordCorrection({
      correctionId: crypto.randomUUID(),
      challengeId: challengeDetails.id,
      challengeTitle: challengeDetails.title,
      correctedAt: Date.now(),
      difficulty: challengeDetails.difficulty,
      functions: challengeDetails.functions,
      items: [{
        taskId: 'workbook-review',
        functionName: null,
        status: learningResult.status,
        feedback: learningResult.feedback,
        expected: typeof learningResult.expected === 'string' ? learningResult.expected : '',
        actual: typeof learningResult.actual === 'string' ? learningResult.actual : ''
      }]
    });
    if (!recordedLearning.ok) {
      throw new Error(`A correção foi recebida, mas não foi registrada no aprendizado: ${recordedLearning.error}`);
    }

    const currentChat = state.chats.find((item) => item.id === chat.id);
    if (!currentChat) return null;
    currentChat.responseId = result.response_id;
    currentChat.messages.push({
      id: crypto.randomUUID(),
      role: 'assistant',
      content: result.response,
      createdAt: Date.now()
    });
    currentChat.updatedAt = Date.now();
    saveChats();
    setBackendStatus(true);
    render();
    return { response: result.response, learningRecorded: true };
  } catch (error) {
    console.error('Não foi possível corrigir a planilha do exercício:', error);
    const isConnectionError = error instanceof TypeError || error.name === 'AbortError';
    if (isConnectionError) setBackendStatus(false);
    const currentChat = state.chats.find((item) => item.id === chat.id);
    if (!currentChat) return null;
    currentChat.messages.push({
      id: crypto.randomUUID(),
      role: 'assistant',
      content: isConnectionError
        ? 'Não foi possível conectar ao Axel AI. Verifique se o servidor está ativo e tente novamente.'
        : `Não foi possível corrigir a planilha: ${error.message}`,
      createdAt: Date.now()
    });
    currentChat.updatedAt = Date.now();
    saveChats();
    render();
    return null;
  } finally {
    isGenerating = false;
    generatingChatId = null;
    render();
  }
}

async function downloadExerciseWorkbook(challenge) {
  if (!challenge?.text?.trim()) {
    throw new Error('Gere um desafio antes de baixar a planilha.');
  }

  const response = await fetch(EXERCISE_WORKBOOK_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      challenge: challenge.text,
      title: challenge.title,
      functions: challenge.functions,
      difficulty: challenge.difficulty
    })
  });

  if (!response.ok) {
    let detail = 'Não foi possível gerar a planilha do desafio.';
    try {
      const result = await response.json();
      if (typeof result.detail === 'string') detail = result.detail;
    } catch (error) {
      console.error('O backend retornou um erro sem detalhes legíveis:', error);
    }
    throw new Error(detail);
  }

  const workbook = await response.blob();
  if (!workbook.size) throw new Error('O backend retornou uma planilha vazia.');
  const downloadUrl = URL.createObjectURL(workbook);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = 'desafio-formulaxl.xlsx';
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}

function setBackendStatus(isOnline) {
  aiServerStatus.classList.toggle('online', isOnline);
  aiServerStatus.classList.toggle('offline', !isOnline);
  aiServerStatus.classList.remove('checking');
  aiServerStatus.querySelector('.ai-server-status-text').textContent = isOnline
    ? 'Backend ativo'
    : 'IA fora do ar';
}

async function checkBackendStatus() {
  if (backendCheckInProgress) return;
  backendCheckInProgress = true;
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), 5000);

  try {
    const response = await fetch(CHAT_API_OPENAPI_URL, {
      cache: 'no-store',
      signal: controller.signal
    });
    if (!response.ok) {
      setBackendStatus(false);
      return;
    }

    const apiSchema = await response.json();
    const chatEndpointAvailable = Object.hasOwn(apiSchema.paths || {}, '/api/future-ai/chat');
    setBackendStatus(chatEndpointAvailable);
  } catch (error) {
    setBackendStatus(false);
  } finally {
    window.clearTimeout(timeoutId);
    backendCheckInProgress = false;
  }
}

async function sendMessageToBackend(chatId, message) {
  try {
    const learningHistory = window.FormulaXLearning.getHistory();
    if (!learningHistory.ok) {
      throw new Error(learningHistory.error);
    }

    const chat = state.chats.find((item) => item.id === chatId);
    if (!chat) return;

    const response = await fetch(CHAT_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message,
        data_aprendizado: {
          version: 1,
          corrections: learningHistory.corrections
        },
        previous_response_id: chat.responseId || null
      })
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || 'Não foi possível obter uma resposta da IA.');
    }
    if (typeof result.response !== 'string' || typeof result.response_id !== 'string') {
      throw new Error('O backend retornou uma resposta em formato inválido.');
    }

    setBackendStatus(true);
    const currentChat = state.chats.find((item) => item.id === chatId);
    if (!currentChat) return;
    currentChat.responseId = result.response_id;
    currentChat.messages.push({
      id: crypto.randomUUID(),
      role: 'assistant',
      content: result.response,
      createdAt: Date.now()
    });
    currentChat.updatedAt = Date.now();
    saveChats();
    render();
    return result.response;
  } catch (error) {
    console.error('Não foi possível enviar a mensagem ao Axel AI:', error);
    const isConnectionError = error instanceof TypeError || error.name === 'AbortError';
    if (isConnectionError) {
      setBackendStatus(false);
    }
    const currentChat = state.chats.find((item) => item.id === chatId);
    if (!currentChat) return;
    currentChat.messages.push({
      id: crypto.randomUUID(),
      role: 'assistant',
      content: isConnectionError
        ? 'Não foi possível conectar ao Axel AI. Verifique se o servidor está ativo e tente novamente.'
        : `Não consegui conectar ao Axel AI: ${error.message}`,
      createdAt: Date.now()
    });
    currentChat.updatedAt = Date.now();
    saveChats();
    render();
    return null;
  }
}

function handleChatInput() {
  chatInput.classList.toggle('has-line-break', /[\r\n]/.test(chatInput.value));
  chatInput.style.height = 'auto';
  chatInput.style.height = `${Math.min(chatInput.scrollHeight, 140)}px`;
}

function handleModalClose(event) {
  const button = event.target.closest('[data-close-modal]');
  if (!button) return;
  const type = button.dataset.closeModal;
  if (type === 'rename') closeRenameModal();
  if (type === 'delete') closeDeleteModal();
}

function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      closeRenameModal();
      closeDeleteModal();
    }
  });
}

newChatBtn.addEventListener('click', () => {
  createChat('Nova conversa');
});

bulkDeleteBtn.addEventListener('click', () => {
  if (selectedChatIds.size) {
    deleteSelectedChats();
  }
});

selectAllBtn.addEventListener('click', toggleSelectAllVisibleChats);

historyButtons.forEach((button) => {
  button.addEventListener('click', () => toggleChatView(button.dataset.view));
});

chatForm.addEventListener('submit', handleSendMessage);
chatInput.addEventListener('input', handleChatInput);
chatInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' && !event.ctrlKey && !event.shiftKey) {
    event.preventDefault();
    chatForm.requestSubmit();
  }
});
searchInput.addEventListener('input', render);

checkBackendStatus();
window.setInterval(checkBackendStatus, 15000);

window.FormulaXLChat = Object.freeze({
  sendMessage: sendChatMessage,
  correctExerciseWorkbook,
  downloadExerciseWorkbook,
  renderMarkdown: renderAssistantMarkdown
});

chatHistoryToggle.addEventListener('click', () => {
  const isCollapsed = chatApp.classList.toggle('chat-history-collapsed');
  const label = isCollapsed ? 'Abrir histórico' : 'Fechar histórico';
  chatHistoryToggle.setAttribute('aria-expanded', String(!isCollapsed));
  chatHistoryToggle.textContent = label;
});

document.querySelector('[data-action="rename"]').addEventListener('click', openRenameModal);
document.querySelector('[data-action="archive"]').addEventListener('click', () => {
  const chat = getActiveChat();
  if (!chat) return;
  if (chat.archived) {
    restoreChat(chat.id);
  } else {
    archiveChat(chat.id);
  }
});

renameForm.addEventListener('submit', (event) => {
  event.preventDefault();
  const value = renameInput.value.trim();
  const chat = state.chats.find((item) => item.id === pendingRenameChatId);
  if (!chat || !value) return;

  chat.title = value;
  chat.updatedAt = Date.now();
  saveChats();
  closeRenameModal();
  render();
});

confirmDeleteBtn.addEventListener('click', () => {
  if (pendingDeleteChatId) {
    deleteChat(pendingDeleteChatId);
  }
  closeDeleteModal();
});

document.body.addEventListener('click', (event) => {
  const closeBtn = event.target.closest('[data-close-modal]');
  if (closeBtn) {
    handleModalClose(event);
  }

  if (event.target === renameModal) closeRenameModal();
  if (event.target === deleteModal) closeDeleteModal();
});

chatHistoryList.addEventListener('contextmenu', (event) => {
  const item = event.target.closest('.chat-history-item');
  if (!item) return;
  event.preventDefault();
  const chatId = item.dataset.chatId;
  openDeleteModal(chatId);
});

document.querySelectorAll('.suggestion-chip').forEach((chip) => {
  chip.addEventListener('click', () => {
    const value = chip.dataset.suggestion;
    chatInput.value = value;
    handleChatInput();
    chatInput.focus();
  });
});

if (chatApp) {
  pageContent.style.display = 'block';
}

loadChats();
setupKeyboardShortcuts();
render();

function setMobileMenuOpen(isOpen) {
  appShell.classList.toggle('mobile-menu-open', isOpen);
  overlay.classList.toggle('visible', isOpen);
  mobileMenuToggle.setAttribute('aria-expanded', String(isOpen));
  mobileMenuToggle.setAttribute('aria-label', isOpen ? 'Fechar menu' : 'Abrir menu');
}

mobileMenuToggle.addEventListener('click', () => {
  setMobileMenuOpen(!appShell.classList.contains('mobile-menu-open'));
});

overlay.addEventListener('click', () => setMobileMenuOpen(false));

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && appShell.classList.contains('mobile-menu-open')) {
    setMobileMenuOpen(false);
    mobileMenuToggle.focus();
  }
});

function setVisibleSection(label) {
  const isHome = label === 'Início';
  const isAi = label === 'Axel AI';
  const isExercise = label === 'Exercícios';
  const isLearning = label === 'Meu aprendizado';
  const isSettings = label === 'Configurações';

  if (isAi && window.matchMedia('(max-width: 900px)').matches) {
    chatApp.classList.add('chat-history-collapsed');
    chatHistoryToggle.setAttribute('aria-expanded', 'false');
    chatHistoryToggle.textContent = 'Abrir histórico';
  }

  if (isExercise) {
    appShell.classList.add('exercise-chat-collapsed');
    if (exerciseChatToggle) {
      exerciseChatToggle.setAttribute('aria-expanded', 'false');
      exerciseChatToggle.setAttribute('aria-label', 'Abrir chat');
      exerciseChatToggle.querySelector('span').textContent = 'Abrir chat';
    }
  }

  appShell.classList.toggle('chat-view-active', isAi);
  appShell.classList.toggle('exercise-view-active', isExercise);

  navItems.forEach((item) => {
    const active = item.dataset.label === label;
    item.classList.toggle('is-active', active);
    item.setAttribute('aria-current', active ? 'page' : 'false');
  });

  if (pageContent) {
    pageContent.classList.toggle('hidden', !isHome);
  }

  if (chatApp) {
    const exerciseChatCollapsed = appShell.classList.contains('exercise-chat-collapsed');
    chatApp.classList.toggle('hidden', !isAi && (!isExercise || exerciseChatCollapsed));
  }

  if (exerciseChatResizer) {
    exerciseChatResizer.hidden = !isExercise || appShell.classList.contains('exercise-chat-collapsed');
  }

  if (exerciseApp) {
    exerciseApp.classList.toggle('hidden', !isExercise);
  }

  if (learningApp) {
    learningApp.classList.toggle('hidden', !isLearning);
  }

  if (settingsApp) {
    settingsApp.classList.toggle('hidden', !isSettings);
  }
}

function setExerciseChatWidth(width) {
  const mainPanel = document.querySelector('.main-panel');
  const mainPanelStyle = getComputedStyle(mainPanel);
  const horizontalPadding = parseFloat(mainPanelStyle.paddingLeft) + parseFloat(mainPanelStyle.paddingRight);
  const availableWidth = mainPanel.clientWidth - horizontalPadding;
  const minChatWidth = 280;
  const minExerciseWidth = 360;
  const maxChatWidth = Math.max(minChatWidth, Math.min(availableWidth * 0.6, availableWidth - minExerciseWidth - 24));
  const chatWidth = Math.round(Math.min(Math.max(width, minChatWidth), maxChatWidth));

  appShell.style.setProperty('--exercise-chat-width', `${chatWidth}px`);
  exerciseChatResizer.setAttribute('aria-valuenow', String(chatWidth));
  exerciseChatResizer.setAttribute('aria-valuemax', String(Math.round(maxChatWidth)));
}

if (exerciseChatToggle) {
  exerciseChatToggle.addEventListener('click', () => {
    const isCollapsed = appShell.classList.toggle('exercise-chat-collapsed');
    if (exerciseChatResizer) exerciseChatResizer.hidden = isCollapsed;
    exerciseChatToggle.setAttribute('aria-expanded', String(!isCollapsed));
    const actionLabel = isCollapsed ? 'Abrir chat' : 'Fechar chat';
    exerciseChatToggle.setAttribute('aria-label', actionLabel);
    exerciseChatToggle.querySelector('span').textContent = actionLabel;
    if (appShell.classList.contains('exercise-view-active') && chatApp) {
      chatApp.classList.toggle('hidden', isCollapsed);
    }
    if (!isCollapsed && exerciseChatResizer && exerciseChatPanel) {
      setExerciseChatWidth(exerciseChatPanel.getBoundingClientRect().width);
    }
  });
}

if (exerciseChatResizer && exerciseApp && exerciseChatPanel) {
  function getChatWidth() {
    return exerciseChatPanel.getBoundingClientRect().width;
  }

  exerciseChatResizer.addEventListener('pointerdown', (event) => {
    if (window.matchMedia('(max-width: 900px)').matches) return;
    event.preventDefault();
    exerciseChatResizer.setPointerCapture(event.pointerId);

    const onPointerMove = (moveEvent) => {
      const mainBounds = document.querySelector('.main-panel').getBoundingClientRect();
      setExerciseChatWidth(mainBounds.right - moveEvent.clientX - 32);
    };
    const onPointerUp = () => {
      exerciseChatResizer.removeEventListener('pointermove', onPointerMove);
      exerciseChatResizer.removeEventListener('pointerup', onPointerUp);
      exerciseChatResizer.removeEventListener('pointercancel', onPointerUp);
    };

    exerciseChatResizer.addEventListener('pointermove', onPointerMove);
    exerciseChatResizer.addEventListener('pointerup', onPointerUp);
    exerciseChatResizer.addEventListener('pointercancel', onPointerUp);
  });

  exerciseChatResizer.addEventListener('keydown', (event) => {
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const step = event.shiftKey ? 40 : 16;
    if (event.key === 'Home') {
      setExerciseChatWidth(Number(exerciseChatResizer.getAttribute('aria-valuemin')));
    } else if (event.key === 'End') {
      setExerciseChatWidth(Number(exerciseChatResizer.getAttribute('aria-valuemax')));
    } else {
      setExerciseChatWidth(getChatWidth() + (event.key === 'ArrowLeft' ? step : -step));
    }
  });
}

window.addEventListener('formulaXL:navigate', (event) => {
  const section = event.detail && event.detail.section;
  if (section) setVisibleSection(section);
});

window.addEventListener('formulaXL:backup-restored', () => {
  loadChats();
  activeChatId = state.chats[0]?.id || null;
  selectedChatIds.clear();
  render();
});

navItems.forEach((item) => {
  item.addEventListener('click', () => {
    setVisibleSection(item.dataset.label);
    setMobileMenuOpen(false);
  });
});

if (startStudyingButton) {
  startStudyingButton.addEventListener('click', () => setVisibleSection('Exercícios'));
}

setVisibleSection('Início');
