// ===== データ管理（localStorage、書き込みは数ミリ秒） =====
const DB_KEY = 'minibasApp_v1';
const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
const DEFAULT_DB = { games: [], currentGameId: null, ownTeam: emptyTeam('') };
let db = JSON.parse(localStorage.getItem(DB_KEY) || 'null') || structuredClone(DEFAULT_DB);
let xlsxLoadPromise = null;
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
    events: [],                   // {id,t,team,no,type:'score'|'foul'|'to',pts,q}
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
    const pl = t.players.find(p => p.no===e.no);
    const name = pl ? `#${pl.no} ${pl.name}` : `#${e.no}`;
    const label = e.type==='score'? `${e.pts}点` : e.type==='foul'? 'ファウル' : 'T.O.';
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
          スコアシート用Excel（JBA公式テンプレートに準拠したレイアウト）と、<br>
          ランニングスコアCSVを出力できます。
        </p>
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
    case 'csv':  exportCsv(g);  return;
    case 'json': exportJson(g); return;
  }
  render();
}

function esc(s){ return String(s??'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

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
