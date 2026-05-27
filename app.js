// ===== データ管理（localStorage、書き込みは数ミリ秒） =====
const DB_KEY = 'minibasApp_v1';
const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
const PDF_LIB_URL = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';
const FONTKIT_URL = 'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js';
const OFFICIAL_PDF_URL = 'MINI_scoresheet_20190401.pdf';
const PDF_FONT_URL = 'NotoSansJP-VF.ttf';
const DEFAULT_DB = { games: [], currentGameId: null, ownTeam: emptyTeam('') };
let db = JSON.parse(localStorage.getItem(DB_KEY) || 'null') || structuredClone(DEFAULT_DB);
let xlsxLoadPromise = null;
let pdfLibLoadPromise = null;
function save() { localStorage.setItem(DB_KEY, JSON.stringify(db)); }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
normalizeDb();

// ===== 新規試合作成 =====
function newGame() {
  const ownTeam = hasOwnTeam() ? cloneTeam(db.ownTeam) : emptyTeam('Aチーム');
  return {
    id: uid(),
    createdAt: new Date().toISOString(),
    info: { tournament: '', gameNo: '', date: new Date().toISOString().slice(0,10),
            venue: '', referee1: '', referee2: '', commissioner: '' },
    teams: {
      A: ownTeam,
      B: emptyTeam('Bチーム')
    },
    quarter: 1,                   // 1-4, 5+はOT
    selected: { team: null, no: null },
    events: [],                   // {id,t,team,no,type:'score'|'foul'|'play'|'to',pts,q}
    timeouts: { A: [0,0,0,0], B: [0,0,0,0] }, // Qごと
  };
}

function currentGame() { return db.games.find(g => g.id === db.currentGameId); }

function emptyTeam(name) {
  return { name, coach: '', assistant: '', players: [] };
}

function normalizeDb() {
  if (!Array.isArray(db.games)) db.games = [];
  if (!('currentGameId' in db)) db.currentGameId = null;
  db.ownTeam = normalizeTeam(db.ownTeam, '');
}

function normalizeTeam(team, fallbackName) {
  const players = Array.isArray(team?.players) ? team.players : [];
  return {
    name: team?.name || fallbackName,
    coach: team?.coach || '',
    assistant: team?.assistant || '',
    players: players.map(p => ({
      no: p?.no || '',
      name: p?.name || '',
      license: p?.license || ''
    }))
  };
}

function cloneTeam(team) {
  return JSON.parse(JSON.stringify(normalizeTeam(team, '')));
}

function hasOwnTeam() {
  const t = db.ownTeam;
  return !!(t.name || t.coach || t.assistant ||
    t.players.some(p => p.no || p.name || p.license));
}

// ===== ルーティング =====
let view = 'home';
function render() {
  const root = document.getElementById('app');
  if (view === 'home')   root.innerHTML = renderHome();
  if (view === 'own')    root.innerHTML = renderOwnTeam();
  if (view === 'setup')  root.innerHTML = renderSetup();
  if (view === 'play')   root.innerHTML = renderPlay();
  if (view === 'export') root.innerHTML = renderExport();
  bindEvents();
}

// ===== ホーム画面 =====
function renderHome() {
  const own = db.ownTeam;
  const ownName = own.name || '未登録';
  const ownPlayers = own.players.filter(p => p.no || p.name || p.license).length;
  const rows = db.games.slice().reverse().map(g => `
    <div class="game-row">
      <div>
        <div><b>${esc(g.teams.A.name)} vs ${esc(g.teams.B.name)}</b></div>
        <div class="meta">${g.info.date} ${g.info.tournament || ''}　
          ${getScore(g,'A')} - ${getScore(g,'B')}</div>
      </div>
      <div>
        <button class="btn small" data-act="open" data-id="${g.id}">開く</button>
        <button class="btn small red" data-act="del" data-id="${g.id}">削除</button>
      </div>
    </div>`).join('');
  return `
    <div class="header"><h1>🏀 ミニバス得点表</h1>
      <div class="header-actions">
        <button class="btn gray small" data-act="own">自チーム</button>
        <button class="btn green" data-act="new">＋新規試合</button>
      </div></div>
    <div class="container">
      <div class="card">
        <div class="card-head">
          <div>
            <h3>自チーム</h3>
            <div class="meta">${esc(ownName)}　選手${ownPlayers}人</div>
          </div>
          <button class="btn small" data-act="own">登録・編集</button>
        </div>
      </div>
      <div class="card">
        <h3 style="margin-top:0">試合一覧</h3>
        ${rows || '<div style="color:#999;padding:20px;text-align:center">試合がありません</div>'}
      </div>
    </div>`;
}

// ===== 自チーム登録 =====
function renderOwnTeam() {
  const t = db.ownTeam;
  const playerRows = t.players.map((p,i) => `
    <div class="player-row">
      <input data-act="ownpl" data-i="${i}" data-f="no" value="${esc(p.no)}" placeholder="背番">
      <input data-act="ownpl" data-i="${i}" data-f="name" value="${esc(p.name)}" placeholder="氏名">
      <input data-act="ownpl" data-i="${i}" data-f="license" value="${esc(p.license||'')}" placeholder="ID下3桁">
      <button class="btn small red" data-act="owndelpl" data-i="${i}">×</button>
    </div>`).join('');

  return `
    <div class="header">
      <button class="btn gray small" data-act="home">← 戻る</button>
      <h1>🏀 自チーム登録</h1>
      <button class="btn green small" data-act="new">新規試合</button>
    </div>
    <div class="container">
      <div class="card">
        <h3 style="margin-top:0">チーム情報</h3>
        <label>チーム名</label><input data-act="ownfield" data-f="name" value="${esc(t.name)}">
        <label>ヘッドコーチ</label><input data-act="ownfield" data-f="coach" value="${esc(t.coach)}">
        <label>アシスタントコーチ</label><input data-act="ownfield" data-f="assistant" value="${esc(t.assistant)}">
      </div>
      <div class="card">
        <div class="card-head">
          <h3>選手</h3>
          <button class="btn small" data-act="ownaddpl">＋選手追加</button>
        </div>
        ${playerRows || '<div style="color:#999;padding:12px 0">選手が未登録です</div>'}
      </div>
    </div>`;
}

// ===== 試合セットアップ =====
function renderSetup() {
  const g = currentGame();
  const playerRows = (team) => g.teams[team].players.map((p,i) => `
    <div class="player-row">
      <input data-act="pl" data-t="${team}" data-i="${i}" data-f="no" value="${esc(p.no)}" placeholder="背番">
      <input data-act="pl" data-t="${team}" data-i="${i}" data-f="name" value="${esc(p.name)}" placeholder="氏名">
      <input data-act="pl" data-t="${team}" data-i="${i}" data-f="license" value="${esc(p.license||'')}" placeholder="ID下3桁">
      <button class="btn small red" data-act="delpl" data-t="${team}" data-i="${i}">×</button>
    </div>`).join('');

  return `
    <div class="header">
      <button class="btn gray small" data-act="home">← 戻る</button>
      <h1>🏀 試合設定</h1>
      <button class="btn green small" data-act="play">試合開始 →</button>
    </div>
    <div class="container">
      <div class="card">
        <h3 style="margin-top:0">試合情報</h3>
        <label>大会名</label><input data-act="info" data-f="tournament" value="${esc(g.info.tournament)}">
        <label>試合No</label><input data-act="info" data-f="gameNo" value="${esc(g.info.gameNo)}">
        <label>日付</label><input type="date" data-act="info" data-f="date" value="${g.info.date}">
        <label>会場</label><input data-act="info" data-f="venue" value="${esc(g.info.venue)}">
        <label>レフリー1</label><input data-act="info" data-f="referee1" value="${esc(g.info.referee1)}">
        <label>レフリー2</label><input data-act="info" data-f="referee2" value="${esc(g.info.referee2)}">
        <label>コミッショナー</label><input data-act="info" data-f="commissioner" value="${esc(g.info.commissioner)}">
      </div>
      ${['A','B'].map(t => `
        <div class="card">
          <div class="card-head">
            <h3>${t}チーム</h3>
            <button class="btn small gray" data-act="applyown" data-t="${t}">自チームを反映</button>
          </div>
          <label>チーム名</label><input data-act="team" data-t="${t}" data-f="name" value="${esc(g.teams[t].name)}">
          <label>ヘッドコーチ</label><input data-act="team" data-t="${t}" data-f="coach" value="${esc(g.teams[t].coach)}">
          <label>アシスタントコーチ</label><input data-act="team" data-t="${t}" data-f="assistant" value="${esc(g.teams[t].assistant)}">
          <label style="margin-top:10px">選手（背番号／氏名／JBA-ID下3桁）</label>
          ${playerRows(t)}
          <button class="btn small" data-act="addpl" data-t="${t}">＋選手追加</button>
        </div>`).join('')}
    </div>`;
}

// ===== 試合中（得点入力） =====
function renderPlay() {
  const g = currentGame();
  const qlabel = g.quarter <= 4 ? `第${g.quarter}Q` : `OT${g.quarter-4}`;

  const playerBtns = (team) => g.teams[team].players.map(p => {
    const sel = g.selected.team===team && g.selected.no===p.no ? 'selected' : '';
    const pts = getPlayerPoints(g, team, p.no);
    const fouls = getPlayerFouls(g, team, p.no);
    return `<div class="player-btn ${sel}" data-act="sel" data-t="${team}" data-no="${esc(p.no)}">
      <div class="num">${esc(p.no)}</div>
      <div class="name">${esc(p.name)}</div>
      <div style="font-size:11px;color:${sel?'#fff':'#888'}">${pts}P${fouls?`<span class="foul-badge">F${fouls}</span>`:''}</div>
    </div>`;
  }).join('');

  const history = g.events.slice(-15).reverse().map(e => {
    const t = g.teams[e.team];
    const pl = e.no ? t.players.find(p => p.no===e.no) : null;
    const name = e.no ? (pl ? `#${pl.no} ${pl.name}` : `#${e.no}`) : '';
    const label = e.type==='score'? `${e.pts}点` : e.type==='foul'? 'ファウル' : e.type==='play'? '出場' : 'T.O.';
    return `<div><span>Q${e.q} ${t.name} ${name} ${label}</span>
      <span class="undo" data-act="undo" data-id="${e.id}">取消</span></div>`;
  }).join('');

  return `
    <div class="header">
      <button class="btn gray small" data-act="setup">設定</button>
      <h1>🏀 ${esc(g.teams.A.name)} vs ${esc(g.teams.B.name)}</h1>
      <button class="btn small" data-act="export">出力</button>
    </div>
    <div class="container">
      <div class="scoreboard">
        <div><div class="team-name">${esc(g.teams.A.name)}</div>
          <div class="team-score">${getScore(g,'A')}</div></div>
        <div class="vs">VS</div>
        <div><div class="team-name">${esc(g.teams.B.name)}</div>
          <div class="team-score">${getScore(g,'B')}</div></div>
      </div>
      <div class="qinfo">
        <button class="btn small gray" data-act="qprev">◀</button>
        　${qlabel}　
        <button class="btn small gray" data-act="qnext">▶</button>
      </div>

      ${['A','B'].map(t => `
        <div class="card team-block">
          <h3>${esc(g.teams[t].name)}（${getScore(g,t)}点）</h3>
          <div class="player-grid">${playerBtns(t)}</div>
        </div>`).join('')}

      <div class="card">
        <div style="font-size:13px;margin-bottom:6px;color:#555">
          ${g.selected.no ? `選択中：${g.selected.team}チーム #${g.selected.no}` : '↑ 先に選手を選んでください'}
        </div>
        <div class="point-bar">
          <button class="btn green" data-act="pt" data-p="1">+1</button>
          <button class="btn green" data-act="pt" data-p="2">+2</button>
          <button class="btn green" data-act="pt" data-p="3">+3</button>
          <button class="btn red"   data-act="foul">ファウル</button>
        </div>
        <div style="margin-top:6px;display:grid;grid-template-columns:1fr 1fr;gap:6px">
          <button class="btn" data-act="playin">出場記録</button>
          <button class="btn gray" data-act="clearSel">選択解除</button>
        </div>
        <div style="margin-top:6px;display:grid;grid-template-columns:1fr 1fr;gap:6px">
          <button class="btn gray" data-act="toA">A タイムアウト (${g.timeouts.A[Math.min(g.quarter-1,3)]||0})</button>
          <button class="btn gray" data-act="toB">B タイムアウト (${g.timeouts.B[Math.min(g.quarter-1,3)]||0})</button>
        </div>
      </div>

      <div class="card">
        <h3 style="margin-top:0">履歴（最新15件）</h3>
        <div class="history">${history || '<div style="color:#999">まだ記録はありません</div>'}</div>
      </div>
    </div>`;
}

// ===== 出力画面 =====
function renderExport() {
  const g = currentGame();
  return `
    <div class="header">
      <button class="btn gray small" data-act="play">← 戻る</button>
      <h1>🏀 出力</h1><span></span>
    </div>
    <div class="container">
      <div class="card">
        <h3 style="margin-top:0">最終スコア</h3>
        <p>${esc(g.teams.A.name)} <b>${getScore(g,'A')}</b> - <b>${getScore(g,'B')}</b> ${esc(g.teams.B.name)}</p>
        <p>Q別：${[1,2,3,4].map(q=>`${getScoreByQ(g,'A',q)}-${getScoreByQ(g,'B',q)}`).join(' / ')}</p>
      </div>
      <div class="card">
        <h3 style="margin-top:0">ファイル出力</h3>
        <p style="font-size:13px;color:#555">
          公式ミニバススコアシートPDF、スコアシート用Excel、<br>
          ランニングスコアCSVを出力できます。
        </p>
        <button class="btn green" data-act="officialpdf" style="width:100%;margin-bottom:8px">
          🏀 公式PDFに書き込んで保存
        </button>
        <button class="btn green" data-act="pdf" style="width:100%;margin-bottom:8px">
          🏀 PDF用スコアシートを開く
        </button>
        <button class="btn green" data-act="xlsx" style="width:100%;margin-bottom:8px">
          📊 スコアシートExcel をダウンロード
        </button>
        <button class="btn" data-act="csv" style="width:100%;margin-bottom:8px">
          📄 ランニングスコアCSV をダウンロード
        </button>
        <button class="btn gray" data-act="json" style="width:100%">
          💾 試合データJSON（バックアップ）
        </button>
      </div>
    </div>`;
}

// ===== 集計 =====
function getScore(g, team) {
  return g.events.filter(e => e.team===team && e.type==='score').reduce((s,e)=>s+e.pts, 0);
}
function getScoreByQ(g, team, q) {
  return g.events.filter(e => e.team===team && e.type==='score' && e.q===q).reduce((s,e)=>s+e.pts, 0);
}
function getPlayerPoints(g, team, no) {
  return g.events.filter(e => e.team===team && e.no===no && e.type==='score').reduce((s,e)=>s+e.pts, 0);
}
function getPlayerFouls(g, team, no) {
  return g.events.filter(e => e.team===team && e.no===no && e.type==='foul').length;
}
function getPlayerQuarterSet(g, team, no) {
  return new Set(g.events
    .filter(e => e.team===team && e.no===no && e.q<=4 && ['play','score','foul'].includes(e.type))
    .map(e => e.q));
}
function getPlayerFoulEvents(g, team, no) {
  return g.events.filter(e => e.team===team && e.no===no && e.type==='foul').sort((a,b)=>a.t-b.t);
}
function getTeamFoulsByQ(g, team, q) {
  return g.events.filter(e => e.team===team && e.type==='foul' && e.q===q).length;
}

// ===== イベントバインド =====
function bindEvents() {
  document.querySelectorAll('[data-act]').forEach(el => {
    const evt = el.tagName==='INPUT' || el.tagName==='SELECT' ? 'change' : 'click';
    el.addEventListener(evt, handle);
    if (el.tagName==='INPUT' && el.type==='text') el.addEventListener('blur', handle);
  });
}
function handle(e) {
  const el = e.currentTarget;
  const act = el.dataset.act;
  const g = currentGame();

  switch(act) {
    case 'new': db.currentGameId = (db.games.push(newGame()), db.games.at(-1).id); save(); view='setup'; break;
    case 'open': db.currentGameId = el.dataset.id; save(); view='play'; break;
    case 'del':
      if (confirm('この試合を削除しますか？')) {
        db.games = db.games.filter(x => x.id !== el.dataset.id); save();
      } break;
    case 'home': view='home'; break;
    case 'own': view='own'; break;
    case 'setup': view='setup'; break;
    case 'play': view='play'; break;
    case 'export': view='export'; break;

    case 'info': g.info[el.dataset.f] = el.value; save(); return;
    case 'team': g.teams[el.dataset.t][el.dataset.f] = el.value; save(); return;
    case 'applyown': {
      if (!hasOwnTeam()) { alert('先に自チームを登録してください'); return; }
      const team = el.dataset.t;
      const hasExisting = g.teams[team].name || g.teams[team].coach || g.teams[team].assistant ||
        g.teams[team].players.some(p => p.no || p.name || p.license);
      if (hasExisting && !confirm(`${team}チームを自チーム情報で上書きしますか？`)) return;
      g.teams[team] = cloneTeam(db.ownTeam);
      if (g.selected.team === team) g.selected = { team: null, no: null };
      save();
      break;
    }
    case 'addpl': g.teams[el.dataset.t].players.push({no:'', name:'', license:''}); save(); break;
    case 'delpl': g.teams[el.dataset.t].players.splice(+el.dataset.i, 1); save(); break;
    case 'pl': g.teams[el.dataset.t].players[+el.dataset.i][el.dataset.f] = el.value; save(); return;
    case 'ownfield': db.ownTeam[el.dataset.f] = el.value; save(); return;
    case 'ownaddpl': db.ownTeam.players.push({no:'', name:'', license:''}); save(); break;
    case 'owndelpl': db.ownTeam.players.splice(+el.dataset.i, 1); save(); break;
    case 'ownpl': db.ownTeam.players[+el.dataset.i][el.dataset.f] = el.value; save(); return;

    case 'sel': g.selected = { team: el.dataset.t, no: el.dataset.no }; save(); break;
    case 'clearSel': g.selected = { team: null, no: null }; save(); break;
    case 'playin':
      if (!g.selected.no) { alert('先に選手を選んでください'); return; }
      if (g.quarter > 4) { alert('出場時限は1Q〜4Qに記録します'); return; }
      if (!g.events.some(x => x.team===g.selected.team && x.no===g.selected.no && x.type==='play' && x.q===g.quarter)) {
        g.events.push({id:uid(), t:Date.now(), team:g.selected.team, no:g.selected.no, type:'play', q:g.quarter});
        save();
      }
      break;
    case 'pt':
      if (!g.selected.no) { alert('先に選手を選んでください'); return; }
      g.events.push({id:uid(), t:Date.now(), team:g.selected.team, no:g.selected.no,
                     type:'score', pts:+el.dataset.p, q:g.quarter}); save(); break;
    case 'foul':
      if (!g.selected.no) { alert('先に選手を選んでください'); return; }
      g.events.push({id:uid(), t:Date.now(), team:g.selected.team, no:g.selected.no,
                     type:'foul', q:g.quarter}); save(); break;
    case 'toA': case 'toB': {
      const team = act==='toA'?'A':'B';
      const idx = Math.min(g.quarter-1, 3);
      g.timeouts[team][idx] = (g.timeouts[team][idx]||0) + 1;
      g.events.push({id:uid(), t:Date.now(), team, type:'to', q:g.quarter}); save(); break;
    }
    case 'undo': g.events = g.events.filter(x => x.id !== el.dataset.id); save(); break;
    case 'qprev': if (g.quarter>1){g.quarter--; save();} break;
    case 'qnext': if (g.quarter<8){g.quarter++; save();} break;

    case 'xlsx': exportXlsx(g); return;
    case 'officialpdf': exportOfficialPdf(g); return;
    case 'pdf':  printScoreSheet(g); return;
    case 'csv':  exportCsv(g);  return;
    case 'json': exportJson(g); return;
  }
  render();
}

function esc(s){ return String(s??'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

// ===== JBAミニバス公式PDFへ書き込み =====
async function exportOfficialPdf(g) {
  if (!(await ensurePdfLib())) {
    alert('PDF出力ライブラリを読み込めませんでした。ネット接続を確認して、もう一度お試しください。');
    return;
  }

  try {
    const { PDFDocument, rgb } = window.PDFLib;
    const [templateBytes, fontBytes] = await Promise.all([
      fetch(OFFICIAL_PDF_URL).then(r => {
        if (!r.ok) throw new Error('official PDF not found');
        return r.arrayBuffer();
      }),
      fetch(PDF_FONT_URL).then(r => {
        if (!r.ok) throw new Error('Japanese font not found');
        return r.arrayBuffer();
      })
    ]);

    const pdfDoc = await PDFDocument.load(templateBytes, { ignoreEncryption: true });
    pdfDoc.registerFontkit(window.fontkit);
    const font = await pdfDoc.embedFont(fontBytes, { subset: true });
    const page = pdfDoc.getPage(0);
    const ink = rgb(0.02, 0.02, 0.02);
    const red = rgb(0.78, 0.06, 0.18);

    const draw = (text, x, y, size=8, maxWidth=90, color=ink) => {
      const value = fitPdfText(font, String(text ?? ''), size, maxWidth);
      if (!value) return;
      page.drawText(value, { x, y, size, font, color });
    };
    const center = (text, x, y, width, size=8, color=ink) => {
      const value = fitPdfText(font, String(text ?? ''), size, width);
      if (!value) return;
      const textWidth = font.widthOfTextAtSize(value, size);
      page.drawText(value, { x: x + Math.max(0, (width - textWidth) / 2), y, size, font, color });
    };

    // Header
    draw(g.info.tournament, 58, 770, 8, 230);
    draw(formatPdfDate(g.info.date), 335, 778, 8, 95);
    draw(g.info.venue, 335, 762, 8, 110);
    draw(g.info.gameNo, 500, 758, 8, 58);
    draw(g.info.referee1, 355, 752, 7, 75);
    draw(g.info.referee2, 471, 752, 7, 75);
    draw(g.info.commissioner, 355, 702, 7, 75);

    // Top score strip
    center(g.teams.A.name, 65, 736, 120, 8);
    center(g.teams.B.name, 242, 736, 120, 8);
    center(getScore(g,'A'), 145, 718, 38, 16, red);
    center(getScore(g,'B'), 225, 718, 38, 16, red);
    draw(getScore(g,'A') === getScore(g,'B') ? '' :
      (getScore(g,'A') > getScore(g,'B') ? g.teams.A.name : g.teams.B.name), 416, 49, 8, 150);

    // Team blocks
    fillOfficialTeam(page, font, draw, center, g, 'A', {
      name: [55, 654], coach: [58, 374], assistant: [58, 360],
      rowStart: 586.2, timeoutY: 632,
      teamFoulY: { 1: 558.1, 2: 558.1, 3: 473.4, 4: 473.4 }
    });
    fillOfficialTeam(page, font, draw, center, g, 'B', {
      name: [55, 329], coach: [58, 49], assistant: [58, 35],
      rowStart: 261.2, timeoutY: 307,
      teamFoulY: { 1: 232.9, 2: 232.9, 3: 148.2, 4: 148.2 }
    });

    fillOfficialRunningScore(page, font, g, ink);

    const pdfBytes = await pdfDoc.save();
    downloadBlob(pdfBytes, 'application/pdf',
      `${safeFileName(g.info.date)}_${safeFileName(g.teams.A.name)}_vs_${safeFileName(g.teams.B.name)}_official.pdf`);
  } catch (err) {
    console.error(err);
    alert('公式PDFへの書き込みに失敗しました。PDFテンプレートやネット接続を確認してください。');
  }
}

function fillOfficialTeam(page, font, draw, center, g, team, pos) {
  const t = g.teams[team];
  const rowGap = 14.17;
  const ink = window.PDFLib.rgb(0.02, 0.02, 0.02);
  const red = window.PDFLib.rgb(0.78, 0.06, 0.18);
  const playXs = [183.5, 197.2, 210.9, 224.6];
  const foulXs = [238.5, 252.2, 265.9, 279.6, 293.3];

  draw(t.name, pos.name[0], pos.name[1], 8, 155);
  draw(t.coach, pos.coach[0], pos.coach[1], 7, 160);
  draw(t.assistant, pos.assistant[0], pos.assistant[1], 7, 150);

  t.players.slice(0, 15).forEach((p, i) => {
    const y = pos.rowStart - i * rowGap;
    const quarters = getPlayerQuarterSet(g, team, p.no);
    center(p.license || '', 29, y, 42, 6.4);
    draw(p.name || '', 72, y, 6.8, 92);
    center(p.no || '', 168, y, 16, 7.2, red);

    [1,2,3,4].forEach((q, qIndex) => {
      if (quarters.has(q)) drawDiagonal(page, playXs[qIndex], y + 0.6, 8, 8, quarterInk(q));
    });

    getPlayerFoulEvents(g, team, p.no).slice(0, 5).forEach((foul, fIndex) => {
      const color = quarterInk(foul.q);
      page.drawText('P', { x: foulXs[fIndex] + 2.4, y, size: 6.8, font, color });
      if (foul.q) page.drawText(String(foul.q), { x: foulXs[fIndex] + 7.4, y: y - 0.2, size: 4.2, font, color });
    });
  });

  [1,2,3,4].forEach(q => {
    const count = Math.min(getTeamFoulsByQ(g, team, q), 4);
    const x = q === 1 || q === 3 ? 323.5 : 338.8;
    const startY = pos.teamFoulY[q] ?? 0;
    for (let i = 0; i < count; i++) {
      drawDiagonal(page, x, startY - i * rowGap, 8, 8, quarterInk(q));
    }
  });

  [0,1,2,3].forEach(i => {
    const count = g.timeouts[team]?.[i] || 0;
    if (count) center(String(count), 234 + i * 13.7, pos.timeoutY, 11, 7, ink);
  });
}

function fillOfficialRunningScore(page, font, g, color) {
  const marks = [];
  const coords = (team, score) => {
    if (score < 1 || score > 120) return null;
    const group = Math.floor((score - 1) / 40);
    const row = (score - 1) % 40;
    const y = 628.5 - row * 14.15;
    const scoreXs = team === 'A' ? [385.4, 460.6, 537.4] : [400.8, 475.9, 552.7];
    const scoreX = scoreXs[group];
    return {
      score,
      scoreX,
      playerX: scoreX + (team === 'A' ? -13.8 : 15.1),
      y
    };
  };
  const centerSmall = (text, x, y, width, drawColor) => {
    const value = String(text ?? '');
    if (!value) return;
    const size = 5.8;
    const textWidth = font.widthOfTextAtSize(value, size);
    page.drawText(value, { x: x + Math.max(0, (width - textWidth) / 2), y, size, font, color: drawColor });
  };

  let sums = { A: 0, B: 0 };
  g.events.filter(e => e.type === 'score').sort((a,b)=>a.t-b.t).forEach(e => {
    sums[e.team] += e.pts;
    const p = coords(e.team, sums[e.team]);
    if (!p) return;
    const drawColor = quarterInk(e.q);

    centerSmall(e.no || '', p.playerX, p.y, 12, drawColor);
    if (e.pts === 1) {
      page.drawEllipse({ x: p.scoreX + 3.7, y: p.y + 5.4, xScale: 2.4, yScale: 2.4, color: drawColor });
    } else {
      drawDiagonal(page, p.scoreX - 1.2, p.y - 0.8, 10.4, 9.2, drawColor);
      if (e.pts === 3) {
        page.drawEllipse({ x: p.playerX + 6.0, y: p.y + 3.1, xScale: 6.0, yScale: 4.5, borderColor: drawColor, borderWidth: 0.8 });
      }
    }
    marks.push({...p, q:e.q, color:drawColor});
  });

  const lastByQ = new Map();
  marks.forEach(mark => lastByQ.set(mark.q, mark));
  lastByQ.forEach(mark => markOfficialPeriodEnd(page, mark, false));
  if (marks.length) markOfficialPeriodEnd(page, marks[marks.length - 1], true);
}

function quarterInk(q) {
  return (q === 1 || q === 3) ? window.PDFLib.rgb(0.78, 0.06, 0.18) : window.PDFLib.rgb(0.02, 0.02, 0.02);
}

function drawDiagonal(page, x, y, width, height, color, thickness=0.8) {
  page.drawLine({
    start: { x, y },
    end: { x: x + width, y: y + height },
    thickness,
    color
  });
}

function markOfficialPeriodEnd(page, mark, isFinal) {
  page.drawEllipse({
    x: mark.scoreX + 3.8,
    y: mark.y + 3.2,
    xScale: 6.0,
    yScale: 4.8,
    borderColor: mark.color,
    borderWidth: isFinal ? 1.4 : 0.9
  });
  const y1 = mark.y - 2.3;
  page.drawLine({
    start: { x: mark.playerX - 1.0, y: y1 },
    end: { x: mark.scoreX + 11.0, y: y1 },
    thickness: isFinal ? 1.2 : 0.8,
    color: mark.color
  });
  if (isFinal) {
    page.drawLine({
      start: { x: mark.playerX - 1.0, y: y1 - 2.0 },
      end: { x: mark.scoreX + 11.0, y: y1 - 2.0 },
      thickness: 1.2,
      color: mark.color
    });
  }
}

function fitPdfText(font, text, size, maxWidth) {
  let value = String(text || '').trim();
  if (!value) return '';
  while (font.widthOfTextAtSize(value, size) > maxWidth && value.length > 1) {
    value = value.slice(0, -1);
  }
  return value;
}

function formatPdfDate(dateText) {
  const m = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return dateText || '';
  return `${m[1]}年${m[2]}月${m[3]}日`;
}

function safeFileName(text) {
  return String(text || 'score').replace(/[\\/:*?"<>|]/g, '_').trim() || 'score';
}

// ===== PDF用スコアシート（印刷画面） =====
function printScoreSheet(g) {
  const w = window.open('', '_blank');
  if (!w) {
    alert('PDF用画面を開けませんでした。ポップアップ許可を確認してください。');
    return;
  }

  w.document.open();
  w.document.write(buildScoreSheetHtml(g));
  w.document.close();
  w.focus();
  setTimeout(() => w.print(), 500);
}

function buildScoreSheetHtml(g) {
  const qScores = [1,2,3,4].map(q => ({
    q,
    a: getScoreByQ(g,'A',q),
    b: getScoreByQ(g,'B',q)
  }));
  const otA = g.events.filter(e => e.team==='A' && e.type==='score' && e.q > 4).reduce((s,e)=>s+e.pts, 0);
  const otB = g.events.filter(e => e.team==='B' && e.type==='score' && e.q > 4).reduce((s,e)=>s+e.pts, 0);

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(g.info.date)}_${esc(g.teams.A.name)}_vs_${esc(g.teams.B.name)}_scoresheet</title>
<style>
@page { size: A4 landscape; margin: 8mm; }
* { box-sizing: border-box; }
body { margin: 0; color: #111; font-family: -apple-system, "Hiragino Sans", "Yu Gothic UI", sans-serif; background: #fff; }
.sheet { width: 281mm; min-height: 194mm; margin: 0 auto; padding: 5mm; border: 2px solid #111; background: #fff; }
.top { display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 4mm; border-bottom: 2px solid #111; padding-bottom: 3mm; }
.title { text-align: center; font-size: 18px; font-weight: 900; letter-spacing: 0.02em; }
.brand { color: #c8102e; font-size: 12px; font-weight: 800; }
.final { text-align: center; border: 2px solid #111; padding: 2mm 4mm; min-width: 54mm; }
.final .score { font-size: 30px; line-height: 1; font-weight: 900; }
.final .teams { font-size: 10px; margin-top: 2mm; }
.info { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.5mm; margin: 3mm 0; font-size: 9px; }
.field { border: 1px solid #111; min-height: 7mm; display: grid; grid-template-columns: 20mm 1fr; }
.field span:first-child { background: #f4f4f4; border-right: 1px solid #111; padding: 1.5mm; font-weight: 700; }
.field span:last-child { padding: 1.5mm; }
.teams-grid { display: grid; grid-template-columns: 1fr 52mm 1fr; gap: 2.5mm; align-items: start; }
.team-box { border: 2px solid #111; }
.team-title { background: #111; color: #fff; padding: 1.5mm 2mm; display: flex; justify-content: space-between; font-weight: 900; }
.team-title strong { color: #ffb4c4; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
th, td { border: 1px solid #111; padding: 1mm; font-size: 8.5px; text-align: center; height: 6.5mm; overflow: hidden; }
th { background: #f3f3f3; font-weight: 800; }
td.name { text-align: left; font-size: 8px; }
.summary { border: 2px solid #111; }
.summary h2 { margin: 0; padding: 1.5mm; background: #c8102e; color: #fff; font-size: 11px; text-align: center; }
.summary td, .summary th { height: 7.5mm; font-size: 9px; }
.win { font-size: 9px; margin-top: 2mm; border: 1px solid #111; padding: 1.5mm; min-height: 10mm; }
.running { margin-top: 3mm; border: 2px solid #111; }
.running h2 { margin: 0; background: #111; color: #fff; padding: 1.5mm 2mm; font-size: 11px; }
.running-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 0; }
.running-grid table:first-child { border-right: 2px solid #111; }
.running td, .running th { height: 5.4mm; font-size: 7.5px; padding: 0.6mm; }
.sign { display: grid; grid-template-columns: repeat(4, 1fr); gap: 2mm; margin-top: 3mm; font-size: 8px; }
.sign div { border: 1px solid #111; min-height: 9mm; padding: 1mm; }
.no-print { position: fixed; right: 12px; top: 12px; display: flex; gap: 8px; }
.no-print button { border: 0; border-radius: 999px; padding: 10px 14px; color: #fff; background: #c8102e; font-weight: 800; }
@media print {
  .no-print { display: none; }
  .sheet { border-width: 2px; margin: 0; }
}
</style>
</head>
<body>
<div class="no-print">
  <button onclick="window.print()">PDFに保存 / 印刷</button>
  <button onclick="window.close()" style="background:#111">閉じる</button>
</div>
<main class="sheet">
  <section class="top">
    <div>
      <div class="brand">HARUTA HIGASHI MINI BASKETBALL</div>
      <div style="font-size:9px">Official Score Sheet</div>
    </div>
    <div class="title">ミニバスケットボール スコアシート</div>
    <div class="final">
      <div class="score">${getScore(g,'A')} - ${getScore(g,'B')}</div>
      <div class="teams">${esc(g.teams.A.name)}　vs　${esc(g.teams.B.name)}</div>
    </div>
  </section>

  <section class="info">
    ${sheetField('大会名', g.info.tournament)}
    ${sheetField('試合No', g.info.gameNo)}
    ${sheetField('日付', g.info.date)}
    ${sheetField('会場', g.info.venue)}
    ${sheetField('レフリー1', g.info.referee1)}
    ${sheetField('レフリー2', g.info.referee2)}
    ${sheetField('コミッショナー', g.info.commissioner)}
    ${sheetField('作成時刻', new Date().toLocaleString('ja-JP'))}
  </section>

  <section class="teams-grid">
    ${teamSheet(g, 'A')}
    <aside class="summary">
      <h2>得点集計</h2>
      <table>
        <tr><th></th><th>1Q</th><th>2Q</th><th>3Q</th><th>4Q</th><th>OT</th><th>合計</th></tr>
        <tr><th>A</th>${qScores.map(x=>`<td>${x.a}</td>`).join('')}<td>${otA}</td><td><b>${getScore(g,'A')}</b></td></tr>
        <tr><th>B</th>${qScores.map(x=>`<td>${x.b}</td>`).join('')}<td>${otB}</td><td><b>${getScore(g,'B')}</b></td></tr>
      </table>
      <h2 style="background:#111">タイムアウト</h2>
      <table>
        <tr><th></th><th>1Q</th><th>2Q</th><th>3Q</th><th>4Q/OT</th></tr>
        <tr><th>A</th>${[0,1,2,3].map(i=>`<td>${g.timeouts.A[i]||0}</td>`).join('')}</tr>
        <tr><th>B</th>${[0,1,2,3].map(i=>`<td>${g.timeouts.B[i]||0}</td>`).join('')}</tr>
      </table>
      <div class="win">勝者<br><b>${getScore(g,'A')===getScore(g,'B') ? '' : esc(getScore(g,'A') > getScore(g,'B') ? g.teams.A.name : g.teams.B.name)}</b></div>
    </aside>
    ${teamSheet(g, 'B')}
  </section>

  ${runningScoreSheet(g)}

  <section class="sign">
    <div>スコアラー</div>
    <div>アシスタントスコアラー</div>
    <div>タイマー</div>
    <div>ショットクロック</div>
  </section>
</main>
</body>
</html>`;
}

function sheetField(label, value) {
  return `<div class="field"><span>${esc(label)}</span><span>${esc(value)}</span></div>`;
}

function teamSheet(g, team) {
  const t = g.teams[team];
  const rows = Array.from({length: 15}, (_, i) => t.players[i] || { no:'', name:'', license:'' }).map(p => `
    <tr>
      <td>${esc(p.no)}</td>
      <td class="name">${esc(p.name)}</td>
      <td>${esc(p.license || '')}</td>
      <td>${p.no ? getPlayerPoints(g,team,p.no) : ''}</td>
      <td>${p.no ? getPlayerFouls(g,team,p.no) : ''}</td>
    </tr>`).join('');

  return `<div class="team-box">
    <div class="team-title"><span>${team}チーム</span><strong>${esc(t.name)}</strong></div>
    <table>
      <tr><th style="width:12mm">No</th><th>氏名</th><th style="width:16mm">ID下3</th><th style="width:14mm">得点</th><th style="width:12mm">F</th></tr>
      ${rows}
    </table>
    <table>
      <tr><th style="width:23mm">HC</th><td class="name">${esc(t.coach)}</td></tr>
      <tr><th>AC</th><td class="name">${esc(t.assistant)}</td></tr>
    </table>
  </div>`;
}

function runningScoreSheet(g) {
  let aSum = 0, bSum = 0;
  const events = g.events.filter(e => e.type==='score').sort((a,b)=>a.t-b.t);
  const rows = events.map((e, i) => {
    if (e.team === 'A') aSum += e.pts; else bSum += e.pts;
    const time = new Date(e.t).toLocaleTimeString('ja-JP',{hour:'2-digit',minute:'2-digit'});
    return {
      n: i + 1,
      q: e.q <= 4 ? e.q : `OT${e.q - 4}`,
      time,
      aNo: e.team === 'A' ? e.no : '',
      aPts: e.team === 'A' ? e.pts : '',
      aSum: e.team === 'A' ? aSum : '',
      bNo: e.team === 'B' ? e.no : '',
      bPts: e.team === 'B' ? e.pts : '',
      bSum: e.team === 'B' ? bSum : ''
    };
  });

  while (rows.length < 28) {
    rows.push({ n: rows.length + 1, q:'', time:'', aNo:'', aPts:'', aSum:'', bNo:'', bPts:'', bSum:'' });
  }

  const tableRows = (items) => items.map(r => `<tr>
    <td>${r.n}</td><td>${r.q}</td><td>${r.time}</td>
    <td>${esc(r.aNo)}</td><td>${r.aPts}</td><td>${r.aSum}</td>
    <td>${esc(r.bNo)}</td><td>${r.bPts}</td><td>${r.bSum}</td>
  </tr>`).join('');

  const mid = Math.ceil(rows.length / 2);
  const head = '<tr><th>No</th><th>Q</th><th>時刻</th><th>A No</th><th>A点</th><th>A計</th><th>B No</th><th>B点</th><th>B計</th></tr>';
  return `<section class="running">
    <h2>ランニングスコア</h2>
    <div class="running-grid">
      <table>${head}${tableRows(rows.slice(0, mid))}</table>
      <table>${head}${tableRows(rows.slice(mid))}</table>
    </div>
  </section>`;
}

// ===== Excel 出力（JBA公式ミニバススコアシートのレイアウトに準拠） =====
async function exportXlsx(g) {
  if (!(await ensureXlsx())) {
    alert('Excel出力ライブラリを読み込めませんでした。ネット接続を確認して、もう一度お試しください。');
    return;
  }

  const wb = XLSX.utils.book_new();
  const ws = {};
  const set = (addr, v, opts={}) => {
    ws[addr] = { v: v, t: typeof v === 'number' ? 'n' : 's',
                 s: { alignment: { horizontal:'center', vertical:'center', wrapText:true },
                      font:{ sz: opts.sz||10, bold: opts.bold||false },
                      border: opts.border ? {
                        top:{style:'thin'}, bottom:{style:'thin'},
                        left:{style:'thin'}, right:{style:'thin'}} : undefined }};
  };

  // --- ヘッダー ---
  set('A1','ミニバスケットボール オフィシャルスコアシート',{sz:14,bold:true});
  set('A3','大会名'); set('B3', g.info.tournament);
  set('E3','試合No'); set('F3', g.info.gameNo);
  set('A4','日付');   set('B4', g.info.date);
  set('E4','会場');   set('F4', g.info.venue);
  set('A5','Aチーム'); set('B5', g.teams.A.name, {bold:true});
  set('E5','Bチーム'); set('F5', g.teams.B.name, {bold:true});

  // --- 両チーム選手一覧 ---
  set('A7','Aチーム',{bold:true});
  set('A8','No'); set('B8','氏名'); set('C8','ID下3'); set('D8','得点'); set('E8','F');
  g.teams.A.players.forEach((p,i)=>{
    const r = 9+i;
    set(`A${r}`,p.no,{border:true}); set(`B${r}`,p.name,{border:true});
    set(`C${r}`,p.license,{border:true});
    set(`D${r}`,getPlayerPoints(g,'A',p.no),{border:true});
    set(`E${r}`,getPlayerFouls(g,'A',p.no),{border:true});
  });

  const bStart = 9 + Math.max(g.teams.A.players.length, 12) + 3;
  set(`A${bStart-1}`,'Bチーム',{bold:true});
  set(`A${bStart}`,'No'); set(`B${bStart}`,'氏名'); set(`C${bStart}`,'ID下3');
  set(`D${bStart}`,'得点'); set(`E${bStart}`,'F');
  g.teams.B.players.forEach((p,i)=>{
    const r = bStart+1+i;
    set(`A${r}`,p.no,{border:true}); set(`B${r}`,p.name,{border:true});
    set(`C${r}`,p.license,{border:true});
    set(`D${r}`,getPlayerPoints(g,'B',p.no),{border:true});
    set(`E${r}`,getPlayerFouls(g,'B',p.no),{border:true});
  });

  // --- Q別スコア ---
  const qRow = bStart + Math.max(g.teams.B.players.length, 12) + 3;
  set(`A${qRow}`,'Q別',{bold:true});
  set(`A${qRow+1}`,'チーム'); set(`B${qRow+1}`,'1Q'); set(`C${qRow+1}`,'2Q');
  set(`D${qRow+1}`,'3Q'); set(`E${qRow+1}`,'4Q'); set(`F${qRow+1}`,'合計');
  set(`A${qRow+2}`,g.teams.A.name,{border:true});
  [1,2,3,4].forEach((q,i)=> set(String.fromCharCode(66+i)+(qRow+2), getScoreByQ(g,'A',q), {border:true}));
  set(`F${qRow+2}`, getScore(g,'A'), {border:true,bold:true});
  set(`A${qRow+3}`,g.teams.B.name,{border:true});
  [1,2,3,4].forEach((q,i)=> set(String.fromCharCode(66+i)+(qRow+3), getScoreByQ(g,'B',q), {border:true}));
  set(`F${qRow+3}`, getScore(g,'B'), {border:true,bold:true});

  // --- ランニングスコア（JBA公式仕様：得点ごとに累計と背番号） ---
  const runRow = qRow + 6;
  set(`A${runRow}`,'ランニングスコア',{bold:true});
  set(`A${runRow+1}`,'時刻'); set(`B${runRow+1}`,'Q');
  set(`C${runRow+1}`,'Aチーム累計'); set(`D${runRow+1}`,'A背番');
  set(`E${runRow+1}`,'B背番'); set(`F${runRow+1}`,'Bチーム累計');
  let aSum=0, bSum=0;
  const scoreEvents = g.events.filter(e => e.type==='score').sort((a,b)=>a.t-b.t);
  scoreEvents.forEach((e,i)=>{
    const r = runRow+2+i;
    if (e.team==='A') aSum+=e.pts; else bSum+=e.pts;
    const time = new Date(e.t).toLocaleTimeString('ja-JP',{hour:'2-digit',minute:'2-digit'});
    set(`A${r}`, time, {border:true});
    set(`B${r}`, e.q, {border:true});
    if (e.team==='A') { set(`C${r}`,aSum,{border:true,bold:true}); set(`D${r}`,e.no,{border:true}); set(`E${r}`,'',{border:true}); set(`F${r}`,'',{border:true});}
    else { set(`C${r}`,'',{border:true}); set(`D${r}`,'',{border:true}); set(`E${r}`,e.no,{border:true}); set(`F${r}`,bSum,{border:true,bold:true});}
  });

  ws['!ref'] = 'A1:F' + (runRow + 2 + scoreEvents.length + 5);
  ws['!cols'] = [{wch:10},{wch:18},{wch:10},{wch:10},{wch:10},{wch:12}];

  XLSX.utils.book_append_sheet(wb, ws, 'スコアシート');
  const fname = `${g.info.date}_${g.teams.A.name}_vs_${g.teams.B.name}.xlsx`;
  XLSX.writeFile(wb, fname);
}

function ensureXlsx(timeoutMs = 15000) {
  if (window.XLSX) return Promise.resolve(true);
  if (!xlsxLoadPromise) {
    xlsxLoadPromise = new Promise(resolve => {
      const s = document.createElement('script');
      s.src = XLSX_URL;
      s.async = true;
      s.onload = () => resolve(true);
      s.onerror = () => resolve(false);
      document.head.appendChild(s);
    });
  }

  return Promise.race([
    xlsxLoadPromise,
    new Promise(resolve => setTimeout(() => resolve(false), timeoutMs))
  ]);
}

function ensurePdfLib(timeoutMs = 20000) {
  if (window.PDFLib && window.fontkit) return Promise.resolve(true);
  if (!pdfLibLoadPromise) {
    pdfLibLoadPromise = loadScript(PDF_LIB_URL)
      .then(() => loadScript(FONTKIT_URL))
      .then(() => !!(window.PDFLib && window.fontkit))
      .catch(() => false);
  }

  return Promise.race([
    pdfLibLoadPromise,
    new Promise(resolve => setTimeout(() => resolve(false), timeoutMs))
  ]);
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = [...document.scripts].find(s => s.src === src);
    if (existing) {
      existing.addEventListener('load', resolve, { once: true });
      existing.addEventListener('error', reject, { once: true });
      return;
    }
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// ===== CSV =====
function exportCsv(g) {
  const rows = [['時刻','Q','チーム','種別','背番','得点']];
  g.events.slice().sort((a,b)=>a.t-b.t).forEach(e=>{
    rows.push([new Date(e.t).toLocaleString('ja-JP'), e.q,
               g.teams[e.team]?.name||e.team, e.type, e.no||'', e.pts||'']);
  });
  const csv = '\uFEFF' + rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
  downloadBlob(csv, 'text/csv', `${g.info.date}_running_score.csv`);
}

// ===== JSON =====
function exportJson(g) {
  downloadBlob(JSON.stringify(g,null,2), 'application/json', `${g.info.date}_game.json`);
}

function downloadBlob(content, mime, fname) {
  const blob = new Blob([content], {type: mime});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = fname; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
}

// ===== 初期化 =====
render();
ensureXlsx(1);
