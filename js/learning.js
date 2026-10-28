(() => {
  const STORAGE_KEY = 'formulaXL-learning-history-v1';
  const STORAGE_VERSION = 1;
  const allowedStatuses = new Set(['correct', 'incorrect', 'unanswered', 'ungradable']);
  const catalogNames = new Set((window.EXCEL_FUNCTIONS || []).map((item) => item.name));

  const elements = {
    app: document.getElementById('learning-app'),
    storageError: document.getElementById('learning-storage-error'),
    progressRing: document.getElementById('learning-progress-ring'),
    progressValue: document.getElementById('learning-progress-value'),
    progressCaption: document.getElementById('learning-progress-caption'),
    completedCount: document.getElementById('learning-completed-count'),
    evaluatedCount: document.getElementById('learning-evaluated-count'),
    correctCount: document.getElementById('learning-correct-count'),
    incorrectCount: document.getElementById('learning-incorrect-count'),
    historyCount: document.getElementById('learning-history-count'),
    historyEmpty: document.getElementById('learning-history-empty'),
    historyList: document.getElementById('learning-history-list'),
    recommendationsEmpty: document.getElementById('learning-recommendations-empty'),
    recommendationsList: document.getElementById('learning-recommendations-list')
  };

  if (Object.values(elements).some((element) => !element)) return;

  function displayStorageError(message) {
    elements.storageError.textContent = message;
    elements.storageError.hidden = !message;
  }

  function parseTimestamp(value) {
    if (value === undefined || value === null) return Date.now();
    if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
    return Date.parse(value);
  }

  function readHistory() {
    let serialized;
    try {
      serialized = localStorage.getItem(STORAGE_KEY);
    } catch (error) {
      console.error('Não foi possível ler o histórico local de aprendizado:', error);
      return { ok: false, error: 'Não foi possível acessar o armazenamento local deste navegador.' };
    }

    if (serialized === null) {
      return { ok: true, corrections: [] };
    }

    try {
      const data = JSON.parse(serialized);
      if (data.version !== STORAGE_VERSION || !Array.isArray(data.corrections)) {
        return { ok: false, error: 'O histórico local de aprendizado tem um formato incompatível e não foi alterado.' };
      }
      if (data.corrections.some((correction) => validateCorrection(correction))) {
        return { ok: false, error: 'Um ou mais registros do histórico local estão inválidos; os dados existentes foram preservados.' };
      }
      return { ok: true, corrections: data.corrections };
    } catch (error) {
      console.error('Não foi possível interpretar o histórico local de aprendizado:', error);
      return { ok: false, error: 'O histórico local de aprendizado não pôde ser lido; os dados existentes foram preservados.' };
    }
  }

  function countEvaluatedItems(items) {
    return items.reduce((totals, item) => {
      if (item.status === 'correct') totals.correct += 1;
      if (item.status === 'incorrect') totals.incorrect += 1;
      return totals;
    }, { correct: 0, incorrect: 0 });
  }

  function getLatestCorrections(corrections) {
    const latestByChallenge = new Map();
    corrections.forEach((correction, index) => {
      const current = latestByChallenge.get(correction.challengeId);
      if (!current || correction.correctedAt > current.correctedAt || (correction.correctedAt === current.correctedAt && index > current.index)) {
        latestByChallenge.set(correction.challengeId, { ...correction, index });
      }
    });
    return [...latestByChallenge.values()];
  }

  function getCorrectedErrors(corrections) {
    const sorted = corrections
      .map((correction, index) => ({ ...correction, index }))
      .sort((left, right) => left.correctedAt - right.correctedAt || left.index - right.index);
    const previousStatus = new Map();
    const correctedByCorrection = new Map();

    sorted.forEach((correction) => {
      const correctedTasks = [];
      correction.items.forEach((item) => {
        if (!item.taskId) return;
        const key = `${correction.challengeId}\u0000${item.taskId}`;
        if (previousStatus.get(key) === 'incorrect' && item.status === 'correct') {
          correctedTasks.push(item);
        }
        previousStatus.set(key, item.status);
      });
      correctedByCorrection.set(correction.correctionId, correctedTasks);
    });

    return correctedByCorrection;
  }

  function createTextElement(tagName, className, text) {
    const element = document.createElement(tagName);
    if (className) element.className = className;
    element.textContent = text;
    return element;
  }

  function formatDate(value) {
    return new Date(value).toLocaleString('pt-BR', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }

  function renderSummary(corrections) {
    const latest = getLatestCorrections(corrections);
    const totals = latest.reduce((summary, correction) => {
      const counts = countEvaluatedItems(correction.items);
      summary.correct += counts.correct;
      summary.incorrect += counts.incorrect;
      return summary;
    }, { correct: 0, incorrect: 0 });
    const evaluated = totals.correct + totals.incorrect;
    const percentage = evaluated ? Math.round((totals.correct / evaluated) * 100) : null;

    elements.progressValue.textContent = percentage === null ? '--' : `${percentage}%`;
    elements.progressRing.style.setProperty('--learning-progress', `${percentage || 0}%`);
    elements.progressRing.setAttribute(
      'aria-label',
      percentage === null ? 'Ainda não há questões corrigidas para calcular o progresso' : `Taxa de acertos de ${percentage} por cento`
    );
    elements.progressCaption.textContent = percentage === null
      ? 'Será calculado após uma correção real'
      : 'Taxa de acertos nas correções mais recentes';
    elements.completedCount.textContent = String(latest.length);
    elements.evaluatedCount.textContent = String(evaluated);
    elements.correctCount.textContent = String(totals.correct);
    elements.incorrectCount.textContent = String(totals.incorrect);
  }

  function renderHistory(corrections) {
    const ordered = [...corrections].sort((left, right) => right.correctedAt - left.correctedAt);
    const correctedByCorrection = getCorrectedErrors(corrections);
    elements.historyCount.textContent = `${ordered.length} ${ordered.length === 1 ? 'registro' : 'registros'}`;
    elements.historyEmpty.hidden = ordered.length > 0;
    elements.historyList.replaceChildren();

    ordered.slice(0, 20).forEach((correction) => {
      const counts = countEvaluatedItems(correction.items);
      const evaluated = counts.correct + counts.incorrect;
      const accuracy = evaluated ? Math.round((counts.correct / evaluated) * 100) : null;
      const item = document.createElement('article');
      item.className = 'learning-history-item';

      const content = document.createElement('div');
      content.className = 'learning-history-content';
      const heading = document.createElement('div');
      heading.className = 'learning-history-heading';
      heading.append(
        createTextElement('h3', '', correction.challengeTitle),
        createTextElement('time', '', formatDate(correction.correctedAt))
      );

      const metadata = document.createElement('div');
      metadata.className = 'learning-history-metadata';
      metadata.append(
        createTextElement('span', '', correction.difficulty || 'Dificuldade não informada'),
        createTextElement('span', 'learning-history-score', accuracy === null ? 'Sem questões avaliáveis' : `${accuracy}% de acertos`),
        createTextElement('span', '', `${counts.correct}/${evaluated} corretas`)
      );
      content.append(heading, metadata);

      const functions = document.createElement('div');
      functions.className = 'learning-history-functions';
      correction.functions.forEach((name) => functions.append(createTextElement('span', 'learning-function-tag', name)));
      if (correction.functions.length) content.append(functions);

      const incorrectItems = correction.items.filter((result) => result.status === 'incorrect');
      const correctedItems = correctedByCorrection.get(correction.correctionId) || [];
      if (incorrectItems.length || correctedItems.length) {
        const outcome = document.createElement('p');
        outcome.className = 'learning-history-outcome';
        const parts = [];
        if (incorrectItems.length) parts.push(`${incorrectItems.length} ${incorrectItems.length === 1 ? 'erro identificado' : 'erros identificados'}`);
        if (correctedItems.length) parts.push(`${correctedItems.length} ${correctedItems.length === 1 ? 'erro corrigido em nova tentativa' : 'erros corrigidos em nova tentativa'}`);
        outcome.textContent = parts.join(' · ');
        content.append(outcome);
      }

      if (incorrectItems.length) {
        const errorList = document.createElement('ul');
        errorList.className = 'learning-error-list';
        incorrectItems.forEach((result) => {
          const errorItem = document.createElement('li');
          const label = result.functionName ? `${result.taskId} · ${result.functionName}` : result.taskId;
          errorItem.append(createTextElement('strong', '', label));
          if (result.feedback) errorItem.append(createTextElement('span', '', result.feedback));
          if (result.expected !== undefined || result.actual !== undefined) {
            const comparison = [];
            if (result.expected !== undefined) comparison.push(`Esperado: ${result.expected}`);
            if (result.actual !== undefined) comparison.push(`Obtido: ${result.actual}`);
            errorItem.append(createTextElement('span', 'learning-error-comparison', comparison.join(' · ')));
          }
          errorList.append(errorItem);
        });
        content.append(errorList);
      }

      item.append(content);
      elements.historyList.append(item);
    });
  }

  function renderRecommendations(corrections) {
    const latest = getLatestCorrections(corrections);
    const challengeCountByFunction = new Map();

    latest.forEach((correction) => {
      const incorrectFunctions = new Set(
        correction.items
          .filter((item) => item.status === 'incorrect' && item.functionName)
          .map((item) => item.functionName)
      );
      incorrectFunctions.forEach((functionName) => {
        challengeCountByFunction.set(functionName, (challengeCountByFunction.get(functionName) || 0) + 1);
      });
    });

    const recommendations = [...challengeCountByFunction.entries()]
      .filter(([, challengeCount]) => challengeCount >= 2)
      .sort((left, right) => right[1] - left[1])
      .slice(0, 5);

    elements.recommendationsEmpty.hidden = recommendations.length > 0;
    elements.recommendationsList.replaceChildren();
    recommendations.forEach(([functionName, challengeCount]) => {
      const card = document.createElement('article');
      card.className = 'learning-recommendation-card';
      const content = document.createElement('div');
      content.append(
        createTextElement('span', 'learning-recommendation-label', 'Prática recomendada'),
        createTextElement('h3', '', `Função ${functionName}`),
        createTextElement('p', '', `Apareceu em resultados incorretos de ${challengeCount} desafios diferentes. Vale praticar essa função novamente.`)
      );
      const button = document.createElement('button');
      button.className = 'learning-recommendation-action';
      button.type = 'button';
      button.textContent = 'Ver no catálogo';
      button.addEventListener('click', () => {
        window.dispatchEvent(new CustomEvent('formulaXL:navigate', { detail: { section: 'Exercícios' } }));
        window.dispatchEvent(new CustomEvent('formulaXL:search-function', { detail: { functionName } }));
      });
      card.append(content, button);
      elements.recommendationsList.append(card);
    });
  }

  function render() {
    const result = readHistory();
    if (!result.ok) {
      displayStorageError(result.error);
      return;
    }

    displayStorageError('');
    renderSummary(result.corrections);
    renderHistory(result.corrections);
    renderRecommendations(result.corrections);
  }

  function validateCorrection(correction) {
    if (!correction || typeof correction !== 'object') return 'O resultado da correção está vazio ou inválido.';
    if (typeof correction.correctionId !== 'string' || !correction.correctionId.trim()) return 'A correção precisa de um identificador único.';
    if (typeof correction.challengeId !== 'string' || !correction.challengeId.trim()) return 'A correção precisa estar vinculada a um desafio.';
    if (typeof correction.challengeTitle !== 'string' || !correction.challengeTitle.trim()) return 'O desafio precisa ter um título.';
    if (!Array.isArray(correction.functions) || !correction.functions.every((name) => catalogNames.has(name))) return 'A correção deve informar funções válidas do catálogo.';
    if (!Array.isArray(correction.items) || correction.items.length === 0) return 'A correção precisa conter os resultados reais das questões avaliadas.';
    if (!correction.items.every((item) => (
      item
      && typeof item.taskId === 'string'
      && item.taskId.trim()
      && allowedStatuses.has(item.status)
      && (item.functionName === null || item.functionName === undefined || (catalogNames.has(item.functionName) && correction.functions.includes(item.functionName)))
      && ['feedback', 'expected', 'actual'].every((field) => item[field] === undefined || typeof item[field] === 'string' || typeof item[field] === 'number')
    ))) {
      return 'Um ou mais resultados da correção têm formato inválido.';
    }
    if (correction.correctedAt !== undefined && !Number.isFinite(parseTimestamp(correction.correctedAt))) {
      return 'A data da correção está inválida.';
    }
    return null;
  }

  function validateHistory(history) {
    if (!history || typeof history !== 'object' || history.version !== STORAGE_VERSION || !Array.isArray(history.corrections)) {
      return 'O arquivo não contém um histórico de aprendizado compatível.';
    }

    const correctionIds = new Set();
    for (const correction of history.corrections) {
      const validationError = validateCorrection(correction);
      if (validationError) return validationError;
      if (correctionIds.has(correction.correctionId)) return 'O histórico contém identificadores de correção duplicados.';
      if (correction.correctedAt === undefined) return 'O histórico contém uma correção sem data válida.';
      if (correction.difficulty !== undefined && typeof correction.difficulty !== 'string') {
        return 'O histórico contém uma dificuldade inválida.';
      }
      correctionIds.add(correction.correctionId);
    }

    return null;
  }

  function recordCorrection(correction) {
    const validationError = validateCorrection(correction);
    if (validationError) return { ok: false, error: validationError };

    const current = readHistory();
    if (!current.ok) {
      displayStorageError(current.error);
      return { ok: false, error: current.error };
    }

    if (current.corrections.some((entry) => entry.correctionId === correction.correctionId)) {
      return { ok: true, duplicate: true };
    }

    const normalized = {
      correctionId: correction.correctionId.trim(),
      challengeId: correction.challengeId.trim(),
      challengeTitle: correction.challengeTitle.trim(),
      correctedAt: parseTimestamp(correction.correctedAt),
      difficulty: typeof correction.difficulty === 'string' ? correction.difficulty : '',
      functions: [...new Set(correction.functions)],
      items: correction.items.map((item) => ({
        taskId: item.taskId.trim(),
        functionName: item.functionName || null,
        status: item.status,
        ...(item.feedback !== undefined ? { feedback: String(item.feedback) } : {}),
        ...(item.expected !== undefined ? { expected: String(item.expected) } : {}),
        ...(item.actual !== undefined ? { actual: String(item.actual) } : {})
      }))
    };

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        version: STORAGE_VERSION,
        corrections: [...current.corrections, normalized]
      }));
    } catch (error) {
      console.error('Não foi possível salvar a correção no histórico local:', error);
      const message = 'A correção foi recebida, mas não pôde ser salva no armazenamento deste navegador.';
      displayStorageError(message);
      return { ok: false, error: message };
    }

    render();
    return { ok: true, duplicate: false };
  }

  window.FormulaXLearning = Object.freeze({
    recordCorrection,
    getHistory: readHistory,
    validateHistory,
    refresh: render
  });

  window.addEventListener('storage', (event) => {
    if (event.key === STORAGE_KEY) render();
  });

  window.addEventListener('formulaXL:learning-updated', render);
  render();
})();
