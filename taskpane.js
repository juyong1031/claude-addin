/* ============================================================
   Claude AI Office Add-in — taskpane.js
   Excel + PowerPoint 공용
   ============================================================ */

let officeApp = null; // 'excel' | 'powerpoint'
let currentContext = '';
let lastResponse = '';
let currentMode = 'generate';

// ── 초기화 ──────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    officeApp = 'excel';
    document.getElementById('app-type-badge').textContent = 'Excel';
  } else if (info.host === Office.HostType.PowerPoint) {
    officeApp = 'powerpoint';
    document.getElementById('app-type-badge').textContent = 'PowerPoint';
  }

  // 저장된 API 키 복원
  const savedKey = localStorage.getItem('claude_api_key');
  if (savedKey) {
    document.getElementById('api-key-input').value = savedKey;
  }
});

// ── API 키 저장 ──────────────────────────────────────────────
function saveApiKey() {
  const key = document.getElementById('api-key-input').value.trim();
  if (!key.startsWith('sk-ant-')) {
    showStatus('API 키 형식이 올바르지 않습니다 (sk-ant-로 시작해야 함)', 'err');
    return;
  }
  localStorage.setItem('claude_api_key', key);
  showStatus('API 키가 저장되었습니다 ✓', 'ok');
}

// ── 모드 전환 ────────────────────────────────────────────────
function switchMode(mode) {
  currentMode = mode;
  document.querySelectorAll('.mode-tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + mode).classList.add('active');

  const placeholders = {
    generate: 'Claude에게 질문하거나 요청하세요...',
    analyze:  '분석할 내용을 입력하거나 컨텍스트를 가져온 뒤 분석을 요청하세요...',
    insert:   '문서에 삽입할 내용을 생성해 달라고 요청하세요...',
  };
  document.getElementById('prompt-input').placeholder = placeholders[mode];
}

// ── 빠른 프롬프트 ────────────────────────────────────────────
function setPrompt(text) {
  document.getElementById('prompt-input').value = text;
  document.getElementById('prompt-input').focus();
}

// ── 컨텍스트 가져오기 ────────────────────────────────────────
async function getContext() {
  try {
    if (officeApp === 'excel') {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(['values', 'address', 'rowCount', 'columnCount']);
        await ctx.sync();

        const addr = range.address;
        const rows = range.rowCount;
        const cols = range.columnCount;
        const values = range.values;

        // 값을 텍스트로 변환 (최대 50행)
        const preview = values.slice(0, 50)
          .map(row => row.join('\t'))
          .join('\n');

        currentContext = `[Excel 선택 범위: ${addr} (${rows}행 × ${cols}열)]\n${preview}`;
        document.getElementById('context-desc').textContent =
          `${addr} (${rows}행 × ${cols}열) 로드됨`;
      });
    } else if (officeApp === 'powerpoint') {
      // PowerPoint: 현재 슬라이드 텍스트 수집
      await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
        slide.load('id');
        await ctx.sync();

        // 슬라이드의 shapes에서 텍스트 추출
        const shapes = slide.shapes;
        shapes.load('items');
        await ctx.sync();

        let texts = [];
        for (const shape of shapes.items) {
          try {
            shape.textFrame.load('text');
            await ctx.sync();
            if (shape.textFrame.text) {
              texts.push(shape.textFrame.text);
            }
          } catch (_) {}
        }

        currentContext = `[PowerPoint 현재 슬라이드 텍스트]\n${texts.join('\n---\n')}`;
        document.getElementById('context-desc').textContent =
          `슬라이드 ${slide.id} 텍스트 로드됨 (${texts.length}개 텍스트 박스)`;
      });
    }

    showStatus('컨텍스트 로드 완료 ✓', 'ok');
  } catch (e) {
    showStatus('컨텍스트 가져오기 실패: ' + e.message, 'err');
  }
}

// ── Claude API 호출 ──────────────────────────────────────────
async function sendToClaude() {
  const apiKey = localStorage.getItem('claude_api_key') || document.getElementById('api-key-input').value.trim();
  if (!apiKey) {
    showStatus('API 키를 먼저 입력해주세요', 'err');
    return;
  }

  const userPrompt = document.getElementById('prompt-input').value.trim();
  if (!userPrompt) {
    showStatus('프롬프트를 입력해주세요', 'err');
    return;
  }

  // UI 상태 전환
  setLoading(true);
  hideElement('response-area');
  hideElement('response-actions');

  // 시스템 프롬프트 구성
  const systemPrompt = officeApp === 'excel'
    ? 'You are an Excel expert and data analyst assistant. When given spreadsheet data, provide clear analysis, insights, or generate content as requested. Respond in Korean unless asked otherwise.'
    : 'You are a PowerPoint presentation expert. Help improve slides, generate scripts, or create compelling content. Respond in Korean unless asked otherwise.';

  // 메시지 구성 (컨텍스트 포함)
  const fullMessage = currentContext
    ? `${currentContext}\n\n사용자 요청: ${userPrompt}`
    : userPrompt;

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-calls': 'true',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1500,
        system: systemPrompt,
        messages: [
          { role: 'user', content: fullMessage }
        ],
      }),
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    lastResponse = data.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n');

    showResponse(lastResponse);

  } catch (e) {
    showStatus('오류: ' + e.message, 'err');
  } finally {
    setLoading(false);
  }
}

// ── 문서에 삽입 ──────────────────────────────────────────────
async function insertResponse() {
  if (!lastResponse) return;

  try {
    if (officeApp === 'excel') {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load('address');
        await ctx.sync();

        // 줄바꿈 기준으로 각 셀에 삽입
        const lines = lastResponse.split('\n').filter(l => l.trim());
        const cell = range.getCell(0, 0);

        if (lines.length === 1) {
          cell.values = [[lines[0]]];
        } else {
          // 여러 줄 → 여러 행에 삽입
          const insertRange = cell.getResizedRange(lines.length - 1, 0);
          insertRange.values = lines.map(l => [l]);
        }
        await ctx.sync();
        showStatus('Excel에 삽입 완료 ✓', 'ok');
      });
    } else if (officeApp === 'powerpoint') {
      await PowerPoint.run(async (ctx) => {
        ctx.presentation.insertSlidesFromBase64; // ensure API available
        // 현재 선택 슬라이드에 텍스트 박스 추가
        const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
        slide.shapes.addTextBox(lastResponse, {
          left: 50, top: 300, width: 600, height: 200
        });
        await ctx.sync();
        showStatus('슬라이드에 텍스트 박스 삽입 완료 ✓', 'ok');
      });
    }
  } catch (e) {
    showStatus('삽입 실패: ' + e.message, 'err');
  }
}

// ── 복사 ─────────────────────────────────────────────────────
function copyResponse() {
  if (!lastResponse) return;
  navigator.clipboard.writeText(lastResponse)
    .then(() => showStatus('클립보드에 복사됨 ✓', 'ok'))
    .catch(() => showStatus('복사 실패', 'err'));
}

// ── 초기화 ───────────────────────────────────────────────────
function clearAll() {
  lastResponse = '';
  currentContext = '';
  document.getElementById('prompt-input').value = '';
  document.getElementById('context-desc').textContent = '버튼을 눌러 현재 선택 영역을 가져오세요';
  hideElement('response-area');
  hideElement('response-actions');
  hideElement('status-msg');
}

// ── 유틸 ─────────────────────────────────────────────────────
function showResponse(text) {
  const el = document.getElementById('response-area');
  el.textContent = text;
  el.classList.remove('hidden');
  document.getElementById('response-actions').classList.remove('hidden');
}

function setLoading(on) {
  document.getElementById('loading').classList.toggle('hidden', !on);
  document.getElementById('send-btn').disabled = on;
}

function hideElement(id) {
  document.getElementById(id).classList.add('hidden');
}

function showStatus(msg, type) {
  const el = document.getElementById('status-msg');
  el.textContent = msg;
  el.className = type === 'ok' ? 'status-ok' : 'status-err';
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 4000);
}
