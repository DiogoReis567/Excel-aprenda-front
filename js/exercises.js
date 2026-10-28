(() => {
  const catalog = window.EXCEL_FUNCTIONS;
  const categoryContainer = document.getElementById('function-categories');
  const functionList = document.getElementById('function-list');
  const emptyResults = document.getElementById('function-empty-results');
  const searchInput = document.getElementById('function-search');
  const selectedList = document.getElementById('selected-function-list');
  const selectedCount = document.getElementById('selected-function-count');
  const clearSelectionButton = document.getElementById('clear-function-selection');
  const generateChallengeButton = document.getElementById('generate-exercise');
  const challengeDifficulty = document.getElementById('exercise-difficulty');
  const challengeFormat = document.getElementById('exercise-format');
  const challengeStatus = document.getElementById('exercise-generation-status');
  const challengeEmpty = document.getElementById('exercise-challenge-empty');
  const challengeContent = document.getElementById('exercise-challenge-content');
  const downloadChallengeButton = document.querySelector('.exercise-download-button');
  const exerciseApp = document.getElementById('exercise-app');
  const randomQuantityInput = document.getElementById('random-function-quantity');
  const randomQuantityDecrease = document.getElementById('random-function-decrease');
  const randomQuantityIncrease = document.getElementById('random-function-increase');
  const randomKeepSelected = document.getElementById('random-keep-selected');
  const randomFunctionButton = document.getElementById('random-function-button');
  const randomFunctionStatus = document.getElementById('random-function-status');
  const workbookInput = document.getElementById('workbook-file');
  const workbookDropArea = document.getElementById('workbook-drop-area');
  const workbookName = document.getElementById('selected-workbook-name');
  const workbookCard = document.getElementById('selected-workbook-card');
  const noWorkbookSelected = document.getElementById('no-workbook-selected');
  const removeWorkbookButton = document.getElementById('remove-workbook-file');
  const workbookStatus = document.getElementById('workbook-upload-status');
  const correctExerciseButton = document.getElementById('correct-exercise');

  if (!categoryContainer || !functionList || !emptyResults || !searchInput || !selectedList || !selectedCount || !clearSelectionButton || !generateChallengeButton || !challengeDifficulty || !challengeFormat || !challengeStatus || !challengeEmpty || !challengeContent || !downloadChallengeButton || !randomQuantityInput || !randomQuantityDecrease || !randomQuantityIncrease || !randomKeepSelected || !randomFunctionButton || !randomFunctionStatus || !workbookInput || !workbookDropArea || !workbookName || !workbookCard || !noWorkbookSelected || !removeWorkbookButton || !workbookStatus || !correctExerciseButton) {
    return;
  }

  if (!Array.isArray(catalog) || catalog.length !== 100) {
    console.error('O catálogo de exercícios deve conter exatamente 100 funções.');
    functionList.textContent = 'Não foi possível carregar o catálogo de funções.';
    return;
  }

  const selectedFunctions = new Set();
  const categories = ['Todas', ...new Set(catalog.map((item) => item.category))];
  let activeCategory = 'Todas';
  let selectedWorkbook = null;
  let workbookValidationRequest = 0;
  let challengeRequestInProgress = false;
  let challengeText = '';
  let activeChallenge = null;
  let correctionRequestInProgress = false;
  let workbookDownloadInProgress = false;

  function updateCorrectionButton() {
    correctExerciseButton.disabled = !selectedWorkbook || !challengeText || correctionRequestInProgress;
  }

  function clearWorkbookSelection() {
    workbookValidationRequest += 1;
    selectedWorkbook = null;
    workbookInput.value = '';
    workbookName.textContent = '';
    workbookCard.hidden = true;
    noWorkbookSelected.hidden = false;
    workbookDropArea.classList.remove('is-valid', 'is-invalid', 'is-dragging');
  }

  function showWorkbookError(message) {
    workbookStatus.textContent = message;
    workbookStatus.classList.add('is-error');
    workbookDropArea.classList.toggle('is-invalid', !selectedWorkbook);
  }

  async function isXlsxPackage(file) {
    const tailStart = Math.max(0, file.size - 65557);
    const tailBuffer = await file.slice(tailStart).arrayBuffer();
    const tailBytes = new Uint8Array(tailBuffer);
    const tailView = new DataView(tailBuffer);
    let endRecordOffset = -1;

    for (let offset = tailBytes.length - 22; offset >= 0; offset -= 1) {
      if (tailView.getUint32(offset, true) === 0x06054b50) {
        endRecordOffset = offset;
        break;
      }
    }

    if (endRecordOffset < 0) return false;

    const entryCount = tailView.getUint16(endRecordOffset + 10, true);
    const directorySize = tailView.getUint32(endRecordOffset + 12, true);
    const directoryOffset = tailView.getUint32(endRecordOffset + 16, true);
    if (directoryOffset + directorySize > file.size) return false;

    const directoryBuffer = await file.slice(directoryOffset, directoryOffset + directorySize).arrayBuffer();
    const directoryBytes = new Uint8Array(directoryBuffer);
    const directoryView = new DataView(directoryBuffer);
    let hasContentTypes = false;
    let hasWorkbook = false;
    let hasWorksheet = false;
    let offset = 0;
    const decoder = new TextDecoder();

    for (let index = 0; index < entryCount; index += 1) {
      if (offset + 46 > directoryBytes.length || directoryView.getUint32(offset, true) !== 0x02014b50) return false;

      const filenameLength = directoryView.getUint16(offset + 28, true);
      const extraLength = directoryView.getUint16(offset + 30, true);
      const commentLength = directoryView.getUint16(offset + 32, true);
      const uncompressedSize = directoryView.getUint32(offset + 24, true);
      const filenameStart = offset + 46;
      const filenameEnd = filenameStart + filenameLength;
      if (filenameEnd > directoryBytes.length) return false;

      const filename = decoder.decode(directoryBytes.subarray(filenameStart, filenameEnd));
      if (filename === '[Content_Types].xml' && uncompressedSize > 0) hasContentTypes = true;
      if (filename === 'xl/workbook.xml' && uncompressedSize > 0) hasWorkbook = true;
      if (/^xl\/worksheets\/sheet\d+\.xml$/i.test(filename) && uncompressedSize > 0) hasWorksheet = true;
      offset = filenameEnd + extraLength + commentLength;
    }

    return offset === directoryBytes.length && hasContentTypes && hasWorkbook && hasWorksheet;
  }

  async function handleWorkbookSelection(file) {
    if (!file) {
      return;
    }

    const requestId = ++workbookValidationRequest;
    if (!/\.xlsx$/i.test(file.name)) {
      showWorkbookError('Formato inválido. Selecione somente um arquivo com extensão .xlsx. O arquivo anterior foi mantido.');
      workbookInput.value = '';
      return;
    }

    workbookStatus.textContent = 'Validando a planilha .xlsx localmente...';
    workbookStatus.classList.remove('is-error');
    workbookDropArea.classList.remove('is-invalid');

    try {
      if (!(await isXlsxPackage(file))) {
        if (requestId !== workbookValidationRequest) return;
        showWorkbookError('O arquivo não parece ser uma planilha .xlsx válida. O arquivo anterior foi mantido.');
        workbookInput.value = '';
        return;
      }
    } catch (error) {
      console.error('Falha ao validar a planilha selecionada:', error);
      if (requestId !== workbookValidationRequest) return;
      showWorkbookError('Não foi possível ler o arquivo. Selecione uma planilha .xlsx válida. O arquivo anterior foi mantido.');
      workbookInput.value = '';
      return;
    }

    if (requestId !== workbookValidationRequest) return;

    selectedWorkbook = file;
    workbookName.textContent = file.name;
    workbookCard.hidden = false;
    noWorkbookSelected.hidden = true;
    workbookStatus.textContent = 'Arquivo .xlsx válido. Ele permanece neste dispositivo e ainda não será enviado.';
    workbookStatus.classList.remove('is-error');
    workbookDropArea.classList.add('is-valid');
    updateCorrectionButton();
  }

  function normalize(value) {
    return value.normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLocaleLowerCase('pt-BR')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  function renderCategories() {
    categoryContainer.replaceChildren();

    categories.forEach((category) => {
      const button = document.createElement('button');
      button.className = 'exercise-category-button';
      button.type = 'button';
      button.textContent = category;
      button.setAttribute('aria-pressed', String(category === activeCategory));
      button.addEventListener('click', () => {
        activeCategory = category;
        renderCategories();
        renderFunctions();
      });
      categoryContainer.append(button);
    });
  }

  function renderFunctions() {
    const query = normalize(searchInput.value.trim());
    const matchingFunctions = catalog.filter((item) => {
      const matchesCategory = activeCategory === 'Todas' || item.category === activeCategory;
      const matchesSearch = !query || normalize(`${item.name} ${item.description}`).includes(query);
      return matchesCategory && matchesSearch;
    });

    functionList.replaceChildren();
    emptyResults.hidden = matchingFunctions.length > 0;

    matchingFunctions.forEach((item) => {
      const button = document.createElement('button');
      button.className = 'exercise-function-option';
      button.type = 'button';
      button.setAttribute('aria-pressed', String(selectedFunctions.has(item.name)));
      button.addEventListener('click', () => {
        if (selectedFunctions.has(item.name)) {
          selectedFunctions.delete(item.name);
        } else {
          selectedFunctions.add(item.name);
        }
        renderFunctions();
        renderSelectedFunctions();
      });

      const heading = document.createElement('span');
      heading.className = 'exercise-function-option-heading';
      const name = document.createElement('strong');
      name.textContent = item.name;
      const category = document.createElement('span');
      category.textContent = item.category;
      heading.append(name, category);

      const description = document.createElement('span');
      description.className = 'exercise-function-description';
      description.textContent = item.description;

      const example = document.createElement('code');
      example.className = 'exercise-function-example';
      example.textContent = item.example;

      const compatibility = document.createElement('span');
      compatibility.className = 'exercise-function-compatibility';
      compatibility.textContent = item.compatibility;

      button.append(heading, description, example, compatibility);
      functionList.append(button);
    });
  }

  function renderSelectedFunctions() {
    selectedCount.textContent = String(selectedFunctions.size);
    clearSelectionButton.disabled = selectedFunctions.size === 0;
    generateChallengeButton.disabled = selectedFunctions.size === 0 || challengeRequestInProgress;
    selectedList.replaceChildren();

    if (!selectedFunctions.size) {
      const placeholder = document.createElement('span');
      placeholder.className = 'exercise-selection-placeholder';
      placeholder.textContent = 'Nenhuma função selecionada.';
      selectedList.append(placeholder);
      return;
    }

    selectedFunctions.forEach((name) => {
      const button = document.createElement('button');
      button.className = 'exercise-selected-chip';
      button.type = 'button';
      button.setAttribute('aria-label', `Remover função ${name}`);
      button.textContent = `${name} ×`;
      button.addEventListener('click', () => {
        selectedFunctions.delete(name);
        renderFunctions();
        renderSelectedFunctions();
      });
      selectedList.append(button);
    });
  }

  function updateExerciseHeaderTransparency() {
    exerciseApp.classList.toggle('is-scrolled', exerciseApp.scrollTop > 8 || window.scrollY > 8);
  }

  function adjustRandomQuantity(amount) {
    const current = Number(randomQuantityInput.value);
    const value = Number.isInteger(current) ? current : 3;
    randomQuantityInput.value = String(Math.max(1, Math.min(catalog.length, value + amount)));
    randomFunctionStatus.textContent = '';
  }

  randomQuantityDecrease.addEventListener('click', () => adjustRandomQuantity(-1));
  randomQuantityIncrease.addEventListener('click', () => adjustRandomQuantity(1));
  randomQuantityInput.addEventListener('input', () => {
    randomFunctionStatus.textContent = '';
  });

  searchInput.addEventListener('input', renderFunctions);
  window.addEventListener('formulaXL:search-function', (event) => {
    const functionName = event.detail && event.detail.functionName;
    if (typeof functionName !== 'string' || !catalog.some((item) => item.name === functionName)) return;
    activeCategory = 'Todas';
    searchInput.value = functionName;
    renderCategories();
    renderFunctions();
    functionList.scrollTop = 0;
    searchInput.focus();
  });
  exerciseApp.addEventListener('scroll', updateExerciseHeaderTransparency, { passive: true });
  window.addEventListener('scroll', updateExerciseHeaderTransparency, { passive: true });
  clearSelectionButton.addEventListener('click', () => {
    selectedFunctions.clear();
    renderFunctions();
    renderSelectedFunctions();
  });
  generateChallengeButton.addEventListener('click', async () => {
    if (challengeRequestInProgress || !selectedFunctions.size) return;
    if (!window.FormulaXLChat) {
      challengeStatus.textContent = 'O chat do Axel AI não está disponível nesta página.';
      return;
    }

    const functions = [...selectedFunctions];
    const difficulty = challengeDifficulty.options[challengeDifficulty.selectedIndex].text;
    const format = challengeFormat.value === 'xlsx'
      ? 'Inclua os dados em uma tabela Markdown com cabeçalhos; ela também será usada para montar a planilha Excel para download.'
      : 'Apresente o desafio em texto, com uma tabela Markdown para os dados.';
    const prompt = [
      'Crie um novo desafio prático de Microsoft Excel para a seção Exercícios da plataforma FórmulaXL.',
      `Funções que devem ser praticadas: ${functions.join(', ')}.`,
      `Dificuldade: ${difficulty}.`,
      format,
      'Responda em português do Brasil usando Markdown, com título, contexto, objetivo, tabela com dados de exemplo e tarefas numeradas.',
      'Crie dados suficientes para que o aluno use as funções selecionadas. Não inclua gabarito, respostas finais ou fórmulas resolvidas.',
      'Mantenha o desafio autocontido e não peça para o aluno anexar um arquivo.'
    ].join('\n');

    challengeRequestInProgress = true;
    generateChallengeButton.disabled = true;
    generateChallengeButton.textContent = 'Gerando desafio...';
    challengeStatus.textContent = 'O Axel AI está criando o desafio. Ele também aparecerá nesta conversa compartilhada.';
    challengeStatus.classList.remove('is-error');

    if (document.querySelector('.app-shell').classList.contains('exercise-chat-collapsed')) {
      document.getElementById('exercise-chat-toggle').click();
    }

    try {
      const response = await window.FormulaXLChat.sendMessage(prompt);
      if (!response) {
        challengeStatus.textContent = 'Não foi possível gerar o desafio. Verifique o status do servidor e tente novamente.';
        challengeStatus.classList.add('is-error');
        return;
      }

      challengeEmpty.hidden = true;
      challengeText = response;
      const heading = response.match(/^\s{0,3}#{1,6}\s+(.+?)\s*#*\s*$/m);
      activeChallenge = {
        id: crypto.randomUUID(),
        title: heading ? heading[1].replace(/\*\*/g, '').trim() : 'Desafio de Excel',
        text: response,
        difficulty,
        functions
      };
      challengeContent.innerHTML = window.FormulaXLChat.renderMarkdown(response);
      challengeContent.hidden = false;
      downloadChallengeButton.disabled = false;
      challengeStatus.textContent = 'Desafio gerado e adicionado à conversa do Axel AI.';
      if (challengeFormat.value === 'xlsx') {
        challengeStatus.textContent += ' A planilha poderá ser baixada pelo botão abaixo.';
      }
      updateCorrectionButton();
    } catch (error) {
      console.error('Não foi possível gerar o desafio:', error);
      challengeStatus.textContent = 'Ocorreu um erro ao gerar o desafio. Tente novamente.';
      challengeStatus.classList.add('is-error');
    } finally {
      challengeRequestInProgress = false;
      generateChallengeButton.textContent = 'Gerar desafio';
      renderSelectedFunctions();
    }
  });
  downloadChallengeButton.addEventListener('click', async () => {
    if (!activeChallenge || workbookDownloadInProgress) return;
    if (!window.FormulaXLChat?.downloadExerciseWorkbook) {
      challengeStatus.textContent = 'O serviço de geração de planilhas não está disponível no frontend.';
      challengeStatus.classList.add('is-error');
      return;
    }

    workbookDownloadInProgress = true;
    downloadChallengeButton.disabled = true;
    downloadChallengeButton.textContent = 'Preparando planilha...';
    challengeStatus.textContent = 'Gerando a planilha para download...';
    challengeStatus.classList.remove('is-error');
    try {
      await window.FormulaXLChat.downloadExerciseWorkbook(activeChallenge);
      challengeStatus.textContent = 'Planilha do desafio baixada.';
    } catch (error) {
      console.error('Não foi possível baixar a planilha do desafio:', error);
      challengeStatus.textContent = error.message;
      challengeStatus.classList.add('is-error');
    } finally {
      workbookDownloadInProgress = false;
      downloadChallengeButton.disabled = !activeChallenge;
      downloadChallengeButton.textContent = 'Baixar planilha .xlsx';
    }
  });
  randomFunctionButton.addEventListener('click', () => {
    const quantity = Number(randomQuantityInput.value);
    if (!Number.isInteger(quantity) || quantity < 1 || quantity > catalog.length) {
      randomFunctionStatus.textContent = `Informe uma quantidade inteira entre 1 e ${catalog.length}.`;
      randomQuantityInput.focus();
      return;
    }

    const keepSelected = randomKeepSelected.checked;
    const candidates = catalog.filter((item) => !keepSelected || !selectedFunctions.has(item.name));
    if (quantity > candidates.length) {
      randomFunctionStatus.textContent = `Há apenas ${candidates.length} função(ões) disponível(is) para sortear sem repetir as já selecionadas.`;
      return;
    }

    if (!keepSelected) selectedFunctions.clear();
    for (let index = candidates.length - 1; index > 0; index -= 1) {
      const randomIndex = Math.floor(Math.random() * (index + 1));
      [candidates[index], candidates[randomIndex]] = [candidates[randomIndex], candidates[index]];
    }
    candidates.slice(0, quantity).forEach((item) => selectedFunctions.add(item.name));

    randomFunctionStatus.textContent = `${quantity} função(ões) sorteada(s)${keepSelected ? ' e adicionada(s) à seleção atual' : ''}.`;
    renderFunctions();
    renderSelectedFunctions();
  });
  workbookInput.addEventListener('change', () => {
    handleWorkbookSelection(workbookInput.files[0] || null);
  });
  removeWorkbookButton.addEventListener('click', () => {
    clearWorkbookSelection();
    workbookStatus.textContent = 'Arquivo removido. Selecione outro arquivo .xlsx quando quiser.';
    workbookStatus.classList.remove('is-error');
    workbookInput.focus();
    updateCorrectionButton();
  });
  correctExerciseButton.addEventListener('click', async () => {
    if (!selectedWorkbook || !challengeText || correctionRequestInProgress) return;
    if (!window.FormulaXLChat?.correctExerciseWorkbook) {
      workbookStatus.textContent = 'O serviço de correção do Axel AI não está disponível.';
      workbookStatus.classList.add('is-error');
      return;
    }

    correctionRequestInProgress = true;
    correctExerciseButton.textContent = 'Corrigindo...';
    workbookStatus.textContent = 'Enviando a planilha para correção pelo Axel AI...';
    workbookStatus.classList.remove('is-error');
    updateCorrectionButton();

    if (document.querySelector('.app-shell').classList.contains('exercise-chat-collapsed')) {
      document.getElementById('exercise-chat-toggle').click();
    }

    try {
      const result = await window.FormulaXLChat.correctExerciseWorkbook(selectedWorkbook, activeChallenge);
      if (!result) {
        workbookStatus.textContent = 'Não foi possível corrigir a planilha. Consulte a resposta do Axel AI no chat e tente novamente.';
        workbookStatus.classList.add('is-error');
        return;
      }

      workbookStatus.textContent = 'Correção concluída e registrada em Meu aprendizado. O desafio e o feedback estão na conversa compartilhada do Axel AI.';
    } catch (error) {
      console.error('Erro ao corrigir exercício:', error);
      workbookStatus.textContent = 'Ocorreu um erro ao corrigir a planilha. Tente novamente.';
      workbookStatus.classList.add('is-error');
    } finally {
      correctionRequestInProgress = false;
      correctExerciseButton.textContent = 'Corrigir exercício';
      updateCorrectionButton();
    }
  });
  workbookDropArea.addEventListener('dragover', (event) => {
    event.preventDefault();
    workbookDropArea.classList.add('is-dragging');
  });
  workbookDropArea.addEventListener('dragleave', (event) => {
    if (!workbookDropArea.contains(event.relatedTarget)) {
      workbookDropArea.classList.remove('is-dragging');
    }
  });
  workbookDropArea.addEventListener('drop', (event) => {
    event.preventDefault();
    workbookDropArea.classList.remove('is-dragging');
    if (event.dataTransfer.files.length > 1) {
      workbookValidationRequest += 1;
      showWorkbookError('Solte apenas um arquivo .xlsx por vez. O arquivo anterior foi mantido.');
      return;
    }
    const file = event.dataTransfer.files[0] || null;
    handleWorkbookSelection(file);
  });

  renderCategories();
  renderFunctions();
  renderSelectedFunctions();
})();
