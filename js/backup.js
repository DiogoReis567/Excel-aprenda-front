(() => {
  const CHAT_STORAGE_KEY = 'formulaX-chat-history';
  const LEARNING_STORAGE_KEY = 'formulaXL-learning-history-v1';
  const BACKUP_FORMAT = 'formulaXL-backup';
  const BACKUP_VERSION = 1;

  const elements = {
    exportButton: document.getElementById('backup-export'),
    fileInput: document.getElementById('backup-import-file'),
    status: document.getElementById('backup-status'),
    modal: document.getElementById('backup-import-modal'),
    fileSummary: document.getElementById('backup-file-summary'),
    importError: document.getElementById('backup-import-error'),
    confirmButton: document.getElementById('backup-import-confirm'),
    cancelButton: document.getElementById('backup-import-cancel'),
    replaceWarning: document.getElementById('backup-replace-warning')
  };

  if (Object.values(elements).some((element) => !element) || !window.FormulaXLearning) return;

  let pendingBackup = null;

  function setStatus(message, isError = false) {
    elements.status.textContent = message;
    elements.status.dataset.error = String(isError);
  }

  function validateChatHistory(history) {
    if (!history || typeof history !== 'object' || !Array.isArray(history.chats)) {
      return { ok: false, error: 'O arquivo não contém um histórico de conversas compatível.' };
    }

    const chatIds = new Set();
    const chats = [];
    for (const chat of history.chats) {
      if (
        !chat
        || typeof chat.id !== 'string'
        || !chat.id.trim()
        || typeof chat.title !== 'string'
        || typeof chat.archived !== 'boolean'
        || !Number.isFinite(chat.createdAt)
        || !Number.isFinite(chat.updatedAt)
        || !Array.isArray(chat.messages)
      ) {
        return { ok: false, error: 'O backup contém uma conversa com formato inválido.' };
      }
      if (chatIds.has(chat.id)) return { ok: false, error: 'O backup contém identificadores de conversa duplicados.' };
      chatIds.add(chat.id);

      const messageIds = new Set();
      const messages = [];
      for (const message of chat.messages) {
        if (
          !message
          || typeof message.id !== 'string'
          || !message.id.trim()
          || !['assistant', 'user'].includes(message.role)
          || typeof message.content !== 'string'
          || !Number.isFinite(message.createdAt)
        ) {
          return { ok: false, error: 'O backup contém uma mensagem com formato inválido.' };
        }
        if (messageIds.has(message.id)) return { ok: false, error: 'Uma conversa contém mensagens duplicadas.' };
        messageIds.add(message.id);
        messages.push({
          id: message.id,
          role: message.role,
          content: message.content,
          createdAt: message.createdAt
        });
      }

      chats.push({
        id: chat.id,
        title: chat.title,
        archived: chat.archived,
        createdAt: chat.createdAt,
        updatedAt: chat.updatedAt,
        messages
      });
    }

    return { ok: true, history: { chats } };
  }

  function normalizeBackup(value) {
    if (
      !value
      || typeof value !== 'object'
      || value.format !== BACKUP_FORMAT
      || value.version !== BACKUP_VERSION
      || !value.data
      || typeof value.data !== 'object'
    ) {
      return { ok: false, error: 'Este arquivo não é um backup FórmulaXL compatível.' };
    }

    const chatResult = validateChatHistory(value.data.chatHistory);
    if (!chatResult.ok) return chatResult;
    if (!Object.prototype.hasOwnProperty.call(value.data, 'learningHistory')) {
      return { ok: false, error: 'O backup não contém os dados de aprendizado.' };
    }

    const learningError = window.FormulaXLearning.validateHistory(value.data.learningHistory);
    if (learningError) return { ok: false, error: learningError };

    return {
      ok: true,
      backup: {
        chatHistory: chatResult.history,
        learningHistory: {
          version: 1,
          corrections: value.data.learningHistory.corrections
        }
      }
    };
  }

  function readStoredData() {
    const rawChats = localStorage.getItem(CHAT_STORAGE_KEY);
    const rawLearning = localStorage.getItem(LEARNING_STORAGE_KEY);
    const chats = rawChats === null ? { chats: [] } : JSON.parse(rawChats);
    const learning = rawLearning === null ? { version: 1, corrections: [] } : JSON.parse(rawLearning);
    const chatResult = validateChatHistory(chats);
    if (!chatResult.ok) throw new Error('Os dados locais de conversas estão inválidos e não puderam ser incluídos no backup.');
    const learningError = window.FormulaXLearning.validateHistory(learning);
    if (learningError) throw new Error('Os dados locais de aprendizado estão inválidos e não puderam ser incluídos no backup.');
    return { chatHistory: chatResult.history, learningHistory: learning };
  }

  function mergeChatHistories(currentHistory, importedHistory) {
    const chatsById = new Map(currentHistory.chats.map((chat) => [chat.id, { ...chat, messages: [...chat.messages] }]));

    importedHistory.chats.forEach((importedChat) => {
      const currentChat = chatsById.get(importedChat.id);
      if (!currentChat) {
        chatsById.set(importedChat.id, { ...importedChat, messages: [...importedChat.messages] });
        return;
      }

      const messagesById = new Map(currentChat.messages.map((message) => [message.id, message]));
      importedChat.messages.forEach((message) => {
        const existing = messagesById.get(message.id);
        if (!existing || message.createdAt > existing.createdAt) messagesById.set(message.id, message);
      });
      currentChat.messages = [...messagesById.values()].sort((left, right) => left.createdAt - right.createdAt);

      if (importedChat.updatedAt >= currentChat.updatedAt) {
        currentChat.title = importedChat.title;
        currentChat.archived = importedChat.archived;
      }
      currentChat.createdAt = Math.min(currentChat.createdAt, importedChat.createdAt);
      currentChat.updatedAt = Math.max(currentChat.updatedAt, importedChat.updatedAt);
    });

    return { chats: [...chatsById.values()] };
  }

  function mergeLearningHistories(currentHistory, importedHistory) {
    const correctionsById = new Map(currentHistory.corrections.map((correction) => [correction.correctionId, correction]));
    importedHistory.corrections.forEach((correction) => {
      if (!correctionsById.has(correction.correctionId)) correctionsById.set(correction.correctionId, correction);
    });
    return { version: 1, corrections: [...correctionsById.values()] };
  }

  function writeStoredData(data) {
    const oldValues = new Map([
      [CHAT_STORAGE_KEY, localStorage.getItem(CHAT_STORAGE_KEY)],
      [LEARNING_STORAGE_KEY, localStorage.getItem(LEARNING_STORAGE_KEY)]
    ]);

    try {
      localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(data.chatHistory));
      localStorage.setItem(LEARNING_STORAGE_KEY, JSON.stringify(data.learningHistory));
    } catch (error) {
      let rollbackError = null;
      oldValues.forEach((value, key) => {
        try {
          if (value === null) localStorage.removeItem(key);
          else localStorage.setItem(key, value);
        } catch (restoreError) {
          rollbackError = rollbackError || restoreError;
        }
      });
      if (rollbackError) {
        console.error('Não foi possível restaurar completamente os dados locais após a falha na importação:', rollbackError);
        throw new Error('A importação falhou e o navegador também não conseguiu restaurar todos os dados anteriores. Verifique os históricos antes de continuar.');
      }
      throw error;
    }
  }

  function openImportModal(backup, filename) {
    pendingBackup = backup;
    elements.fileSummary.textContent = `Arquivo válido: ${filename}. Escolha como deseja aplicar os dados.`;
    document.querySelector('input[name="backup-import-mode"][value="merge"]').checked = true;
    elements.replaceWarning.hidden = true;
    elements.importError.hidden = true;
    elements.importError.textContent = '';
    elements.modal.classList.remove('hidden');
    elements.modal.setAttribute('aria-hidden', 'false');
    elements.confirmButton.focus();
  }

  function closeImportModal() {
    pendingBackup = null;
    elements.modal.classList.add('hidden');
    elements.modal.setAttribute('aria-hidden', 'true');
  }

  function exportBackup() {
    try {
      const data = readStoredData();
      const backup = {
        format: BACKUP_FORMAT,
        version: BACKUP_VERSION,
        exportedAt: new Date().toISOString(),
        data
      };
      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `formulaXL-backup-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
      setStatus('Backup exportado com o histórico de conversas e os dados de aprendizado.');
    } catch (error) {
      console.error('Não foi possível exportar o backup do FórmulaXL:', error);
      setStatus(error instanceof Error ? error.message : 'Não foi possível exportar o backup.', true);
    }
  }

  async function handleFileSelection() {
    const file = elements.fileInput.files && elements.fileInput.files[0];
    elements.fileInput.value = '';
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.json')) {
      setStatus('Selecione um arquivo de backup .json.', true);
      return;
    }

    try {
      const parsed = JSON.parse(await file.text());
      const result = normalizeBackup(parsed);
      if (!result.ok) {
        setStatus(result.error, true);
        return;
      }
      setStatus('');
      openImportModal(result.backup, file.name);
    } catch (error) {
      console.error('Não foi possível ler o arquivo de backup:', error);
      setStatus('O arquivo não pôde ser lido. Verifique se é um JSON válido.', true);
    }
  }

  function importBackup() {
    if (!pendingBackup) return;
    const mode = document.querySelector('input[name="backup-import-mode"]:checked')?.value;

    try {
      let data = pendingBackup;
      if (mode === 'merge') {
        const current = readStoredData();
        data = {
          chatHistory: mergeChatHistories(current.chatHistory, pendingBackup.chatHistory),
          learningHistory: mergeLearningHistories(current.learningHistory, pendingBackup.learningHistory)
        };
      } else if (mode !== 'replace') {
        throw new Error('Escolha como deseja importar os dados.');
      }

      writeStoredData(data);
      closeImportModal();
      window.FormulaXLearning.refresh();
      window.dispatchEvent(new CustomEvent('formulaXL:backup-restored'));
      setStatus(mode === 'merge'
        ? 'Dados mesclados. Conversas e correções repetidas foram evitadas.'
        : 'Dados atuais substituídos pelo conteúdo do backup.');
    } catch (error) {
      console.error('Não foi possível importar o backup do FórmulaXL:', error);
      const message = error instanceof Error ? error.message : 'Não foi possível importar o backup.';
      elements.importError.textContent = message;
      elements.importError.hidden = false;
      setStatus(message, true);
    }
  }

  elements.exportButton.addEventListener('click', exportBackup);
  elements.fileInput.addEventListener('change', handleFileSelection);
  elements.confirmButton.addEventListener('click', importBackup);
  elements.cancelButton.addEventListener('click', closeImportModal);
  elements.modal.addEventListener('click', (event) => {
    if (event.target === elements.modal) closeImportModal();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !elements.modal.classList.contains('hidden')) closeImportModal();
  });
  document.querySelectorAll('input[name="backup-import-mode"]').forEach((input) => {
    input.addEventListener('change', () => {
      elements.replaceWarning.hidden = input.value !== 'replace' || !input.checked;
    });
  });
})();
