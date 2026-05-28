[index.html.html](https://github.com/user-attachments/files/28366822/index.html.html)
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>FinTrack</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f1117;--bg2:#171b26;--bg3:#1e2333;--bg4:#252b3d;
  --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
  --text:#f0f2f8;--text2:#8b93a8;--text3:#5a6280;
  --green:#22d3a0;--green-bg:rgba(34,211,160,0.1);
  --red:#f56565;--red-bg:rgba(245,101,101,0.1);
  --blue:#60a5fa;--amber:#fbbf24;--purple:#a78bfa;
  --radius:12px;--radius-sm:8px;
  --sidebar:220px;
}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;line-height:1.6;-webkit-tap-highlight-color:transparent}
.app{display:flex;min-height:100vh}

/* ── SIDEBAR ── */
.sidebar{width:var(--sidebar);background:var(--bg2);border-right:1px solid var(--border);padding:24px 0;display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;transition:transform .25s}
.logo{padding:0 20px 24px;font-size:16px;font-weight:600;letter-spacing:-.3px;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border);margin-bottom:16px}
.logo-icon{width:30px;height:30px;background:linear-gradient(135deg,var(--green),#0ea5e9);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 20px;cursor:pointer;transition:background .15s;color:var(--text2);font-size:13.5px;font-weight:400;border-left:2px solid transparent;user-select:none}
.nav-item:hover{background:var(--bg3);color:var(--text)}
.nav-item.active{background:var(--bg3);color:var(--green);border-left-color:var(--green)}
.nav-icon{font-size:16px;width:20px;text-align:center;flex-shrink:0}

/* ── MAIN ── */
.main{margin-left:var(--sidebar);flex:1;padding:28px 32px;max-width:1280px}
.page-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:24px;flex-wrap:wrap;gap:12px}
.page-title{font-size:21px;font-weight:600;letter-spacing:-.4px}
.page-sub{font-size:13px;color:var(--text2);margin-top:2px}
.hdr-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}

/* ── CARDS ── */
.cards-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px}
.card-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px}
.card-value{font-size:22px;font-weight:600;font-family:'DM Mono',monospace;letter-spacing:-.5px}
.card-sub{font-size:12px;color:var(--text2);margin-top:3px}

/* ── CHARTS ── */
.content-grid{display:grid;grid-template-columns:1fr 360px;gap:18px;align-items:start}
.chart-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:22px}
.chart-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:8px}
.chart-title{font-size:14px;font-weight:500}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--text2)}
.legend-item{display:flex;align-items:center;gap:5px}
.legend-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0}

/* ── TRANSACTIONS ── */
.transactions-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.tx-header{display:flex;align-items:center;justify-content:space-between;padding:18px 20px 14px;border-bottom:1px solid var(--border);flex-wrap:wrap;gap:8px}
.tx-title{font-size:14px;font-weight:500}
.tx-list{overflow-y:auto;max-height:420px}
.tx-list::-webkit-scrollbar{width:4px}
.tx-list::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}
.tx-item{display:flex;align-items:center;gap:11px;padding:11px 18px;border-bottom:1px solid var(--border);transition:background .1s;cursor:default}
.tx-item:hover{background:var(--bg3)}
.tx-item:last-child{border-bottom:none}
.tx-icon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.tx-info{flex:1;min-width:0}
.tx-name{font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tx-cat{font-size:11px;color:var(--text2)}
.tx-date{font-size:11px;color:var(--text3);white-space:nowrap;margin-right:4px}
.tx-amount{font-family:'DM Mono',monospace;font-size:13px;font-weight:500;white-space:nowrap}
.tx-actions{display:flex;gap:2px;opacity:0;transition:opacity .15s}
.tx-item:hover .tx-actions{opacity:1}
.tx-btn{background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;padding:3px 6px;border-radius:5px;transition:all .15s}
.tx-btn:hover.edit{background:var(--blue);color:#fff}
.tx-btn:hover.del{background:var(--red);color:#fff}

/* ── BUTTONS ── */
.btn{padding:7px 14px;border-radius:var(--radius-sm);border:1px solid var(--border2);background:var(--bg3);color:var(--text);font-size:12px;font-weight:500;cursor:pointer;font-family:'DM Sans',sans-serif;transition:background .15s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:hover{background:var(--bg4)}
.btn-primary{background:var(--green);color:#0f1117;border-color:var(--green)}
.btn-primary:hover{background:#1ab88c}
.btn-danger{background:var(--red-bg);color:var(--red);border-color:rgba(245,101,101,.3)}

/* ── MODAL ── */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:200;align-items:center;justify-content:center;padding:16px}
.modal-bg.open{display:flex}
.modal{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--radius);padding:26px;width:460px;max-width:100%;max-height:90vh;overflow-y:auto}
.modal-title{font-size:15px;font-weight:600;margin-bottom:18px}
.form-row{margin-bottom:13px}
.form-label{font-size:11.5px;color:var(--text2);margin-bottom:5px;display:block}
.form-input,.form-select{width:100%;padding:9px 11px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;outline:none;transition:border .15s}
.form-input:focus,.form-select:focus{border-color:var(--green)}
.form-select option{background:var(--bg3)}
.form-row-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.radio-group{display:flex;gap:8px}
.radio-btn{flex:1;padding:8px 10px;border-radius:var(--radius-sm);border:1px solid var(--border2);cursor:pointer;text-align:center;font-size:13px;font-weight:500;transition:all .15s;color:var(--text2);background:var(--bg3)}
.radio-btn.sel-income{background:var(--green-bg);border-color:var(--green);color:var(--green)}
.radio-btn.sel-expense{background:var(--red-bg);border-color:var(--red);color:var(--red)}
.modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:18px}

/* ── FORM CONTROLS ── */
.period-select{padding:7px 11px;background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:13px;font-family:'DM Sans',sans-serif;cursor:pointer;outline:none}
.filter-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;align-items:center}
.filter-chip{padding:5px 12px;border-radius:20px;border:1px solid var(--border);font-size:12px;cursor:pointer;background:var(--bg2);color:var(--text2);transition:all .15s;user-select:none}
.filter-chip.active{background:var(--bg4);color:var(--text);border-color:var(--border2)}

/* ── SECTIONS ── */
.section{display:none}
.section.active{display:block}
.empty{text-align:center;padding:36px 20px;color:var(--text3);font-size:13px}
.empty-icon{font-size:28px;margin-bottom:8px}

/* ── CAT BARS ── */
.cat-list{padding:6px 0}
.cat-row{display:flex;align-items:center;gap:11px;padding:9px 20px}
.cat-bar-wrap{flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden}
.cat-bar{height:100%;border-radius:3px;transition:width .4s ease}
.cat-name-col{font-size:13px;width:130px;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cat-amt{font-family:'DM Mono',monospace;font-size:12px;color:var(--text);min-width:90px;text-align:right}
.cat-pct{font-size:11px;color:var(--text3);min-width:34px;text-align:right}
.cat-manager{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.cat-section-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px}
.cat-section-title{font-size:14px;font-weight:500;margin-bottom:14px}
.cat-item-row{display:flex;align-items:center;gap:9px;padding:7px 0;border-bottom:1px solid var(--border)}
.cat-item-row:last-child{border-bottom:none}
.cat-color-dot{width:11px;height:11px;border-radius:50%;flex-shrink:0}
.cat-item-name{flex:1;font-size:13px}
.cat-item-badge{font-size:10px;padding:2px 7px;border-radius:10px;background:var(--bg3);color:var(--text3)}
.cat-item-del{background:none;border:none;color:var(--text3);cursor:pointer;font-size:13px;opacity:.5;padding:2px 6px}
.cat-item-del:hover{color:var(--red);opacity:1}
.new-cat-form{display:flex;gap:7px;margin-top:14px;flex-wrap:wrap}
.color-picker{width:38px;height:34px;border:1px solid var(--border2);border-radius:var(--radius-sm);cursor:pointer;background:none;padding:2px}

/* ── IMPORT ── */
.drop-zone{border:2px dashed var(--border2);border-radius:var(--radius);padding:44px 20px;text-align:center;cursor:pointer;transition:all .2s;background:var(--bg2)}
.drop-zone:hover,.drop-zone.dragover{border-color:var(--green);background:var(--green-bg)}
.drop-icon{font-size:36px;margin-bottom:10px}
.drop-title{font-size:15px;font-weight:500;margin-bottom:5px}
.drop-sub{font-size:13px;color:var(--text2)}
.import-preview{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-top:18px}
.import-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0}
.istat{background:var(--bg3);border-radius:var(--radius-sm);padding:12px;text-align:center}
.istat-val{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.istat-label{font-size:11px;color:var(--text2);margin-top:3px}
.import-table-wrap{max-height:260px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius-sm);margin-top:14px}
.import-table{width:100%;border-collapse:collapse;font-size:12px}
.import-table th{padding:8px 12px;background:var(--bg3);color:var(--text2);font-weight:500;text-align:left;position:sticky;top:0}
.import-table td{padding:8px 12px;border-bottom:1px solid var(--border)}
.import-table tr:last-child td{border-bottom:none}
.import-warn{background:rgba(251,191,36,.1);border:1px solid rgba(251,191,36,.3);border-radius:var(--radius-sm);padding:11px 14px;font-size:13px;color:var(--amber);margin-top:10px}

/* ── CLOUD SETUP ── */
.cloud-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:24px;margin-bottom:20px}
.cloud-steps{counter-reset:step}
.cloud-step{display:flex;gap:14px;margin-bottom:18px;align-items:flex-start}
.cloud-step-num{width:26px;height:26px;border-radius:50%;background:var(--green-bg);border:1px solid var(--green);color:var(--green);font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.cloud-step-body{flex:1}
.cloud-step-title{font-size:13.5px;font-weight:500;margin-bottom:4px}
.cloud-step-desc{font-size:12.5px;color:var(--text2);line-height:1.5}
.cloud-step-desc a{color:var(--blue);text-decoration:none}
.cloud-step-desc a:hover{text-decoration:underline}
.sync-status{display:inline-flex;align-items:center;gap:6px;font-size:12px;padding:4px 10px;border-radius:20px}
.sync-dot{width:7px;height:7px;border-radius:50%}
.sync-ok{background:var(--green-bg);color:var(--green)}
.sync-dot.ok{background:var(--green)}
.sync-err{background:var(--red-bg);color:var(--red)}
.sync-dot.err{background:var(--red)}
.sync-local{background:var(--bg3);color:var(--text2)}
.sync-dot.local{background:var(--text3)}

/* ── MOBILE ── */
.mobile-bar{display:none;position:fixed;bottom:0;left:0;right:0;background:var(--bg2);border-top:1px solid var(--border);z-index:100;padding:6px 0 calc(6px + env(safe-area-inset-bottom))}
.mobile-bar-inner{display:flex;justify-content:space-around}
.mob-btn{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 12px;cursor:pointer;color:var(--text3);font-size:10px;border-radius:8px;transition:color .15s;border:none;background:none}
.mob-btn.active{color:var(--green)}
.mob-btn .mob-icon{font-size:18px}

@media(max-width:768px){
  :root{--sidebar:0px}
  .sidebar{transform:translateX(-220px);width:220px}
  .sidebar.open{transform:translateX(0)}
  .main{margin-left:0;padding:16px 16px 80px}
  .cards-grid{grid-template-columns:repeat(2,1fr);gap:10px}
  .content-grid{grid-template-columns:1fr}
  .cat-manager{grid-template-columns:1fr}
  .import-stats{grid-template-columns:repeat(2,1fr)}
  .form-row-2{grid-template-columns:1fr}
  .mobile-bar{display:block}
  .page-header{margin-bottom:16px}
  .hdr-actions .btn span{display:none}
}
@media(min-width:769px) and (max-width:1024px){
  :root{--sidebar:60px}
  .sidebar .nav-item span,.sidebar .logo span{display:none}
  .cards-grid{grid-template-columns:repeat(2,1fr)}
}
</style>
</head>
<body>
<div class="app">

<!-- Sidebar desktop -->
<aside class="sidebar" id="sidebar">
  <div class="logo">
    <div class="logo-icon">💰</div>
    <span>FinTrack</span>
  </div>
  <nav>
    <div class="nav-item active" onclick="showSection('dashboard')"><span class="nav-icon">📊</span><span>Dashboard</span></div>
    <div class="nav-item" onclick="showSection('transacoes')"><span class="nav-icon">↕️</span><span>Transações</span></div>
    <div class="nav-item" onclick="showSection('categorias')"><span class="nav-icon">🏷️</span><span>Categorias</span></div>
    <div class="nav-item" onclick="showSection('anual')"><span class="nav-icon">📅</span><span>Visão Anual</span></div>
    <div class="nav-item" onclick="showSection('importar')"><span class="nav-icon">📥</span><span>Importar Excel</span></div>
    <div class="nav-item" onclick="showSection('gcat')"><span class="nav-icon">⚙️</span><span>Categorias</span></div>
    <div class="nav-item" onclick="showSection('cloud')"><span class="nav-icon">☁️</span><span>Nuvem</span></div>
  </nav>
</aside>

<main class="main">

<!-- ===== DASHBOARD ===== -->
<section id="sec-dashboard" class="section active">
  <div class="page-header">
    <div><div class="page-title">Dashboard</div><div class="page-sub" id="dash-period-label">—</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="dash-month" onchange="updateDashboard()"></select>
      <select class="period-select" id="dash-year" onchange="updateDashboard()"></select>
      <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
    </div>
  </div>
  <div class="cards-grid">
    <div class="card"><div class="card-label">Receitas</div><div class="card-value" id="c-income" style="color:var(--green)">R$ 0</div><div class="card-sub" id="c-income-count">0 entradas</div></div>
    <div class="card"><div class="card-label">Despesas</div><div class="card-value" id="c-expense" style="color:var(--red)">R$ 0</div><div class="card-sub" id="c-expense-count">0 saídas</div></div>
    <div class="card"><div class="card-label">Saldo do mês</div><div class="card-value" id="c-balance">R$ 0</div><div class="card-sub" id="c-balance-pct">—</div></div>
    <div class="card"><div class="card-label">Maior gasto</div><div class="card-value" id="c-topcat" style="font-size:14px;padding-top:4px;color:var(--amber)">—</div><div class="card-sub" id="c-topcat-amt">—</div></div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header">
        <div class="chart-title">Receitas vs Despesas — últimos 6 meses</div>
        <div class="legend">
          <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span>
          <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span>
        </div>
      </div>
      <div style="position:relative;height:220px"><canvas id="chart-monthly"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Gastos por categoria</div></div>
      <div style="position:relative;height:170px"><canvas id="chart-pie"></canvas></div>
      <div id="pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
  <div class="transactions-card" style="margin-top:18px">
    <div class="tx-header"><div class="tx-title">Últimas transações</div><button class="btn" onclick="showSection('transacoes')">Ver todas →</button></div>
    <div class="tx-list" id="recent-list"></div>
  </div>
</section>

<!-- ===== TRANSAÇÕES ===== -->
<section id="sec-transacoes" class="section">
  <div class="page-header">
    <div><div class="page-title">Transações</div><div class="page-sub">Histórico completo</div></div>
    <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
  </div>
  <div class="filter-bar">
    <select class="period-select" id="tx-month" onchange="renderTxList()"></select>
    <select class="period-select" id="tx-year" onchange="renderTxList()"></select>
    <div id="filter-chips" style="display:flex;gap:6px;flex-wrap:wrap"></div>
  </div>
  <div class="transactions-card">
    <div class="tx-list" id="tx-list" style="max-height:none"></div>
  </div>
</section>

<!-- ===== CATEGORIAS (análise) ===== -->
<section id="sec-categorias" class="section">
  <div class="page-header">
    <div><div class="page-title">Categorias</div><div class="page-sub">Análise por categoria</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="cat-month" onchange="renderCategorias()"></select>
      <select class="period-select" id="cat-year" onchange="renderCategorias()"></select>
    </div>
  </div>
  <div class="content-grid">
    <div>
      <div class="chart-card" style="margin-bottom:18px">
        <div class="chart-header"><div class="chart-title">Despesas por categoria</div></div>
        <div class="cat-list" id="cat-expense-bars"></div>
      </div>
      <div class="chart-card">
        <div class="chart-header"><div class="chart-title">Receitas por categoria</div></div>
        <div class="cat-list" id="cat-income-bars"></div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Distribuição de despesas</div></div>
      <div style="position:relative;height:260px"><canvas id="chart-cat-pie"></canvas></div>
      <div id="cat-pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
</section>

<!-- ===== VISÃO ANUAL ===== -->
<section id="sec-anual" class="section">
  <div class="page-header">
    <div><div class="page-title">Visão Anual</div><div class="page-sub">Resumo e análise por categoria</div></div>
    <select class="period-select" id="anual-year" onchange="renderAnual()"></select>
  </div>
  <div class="cards-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="card"><div class="card-label">Total Receitas</div><div class="card-value" id="anual-income" style="color:var(--green)">R$ 0</div></div>
    <div class="card"><div class="card-label">Total Despesas</div><div class="card-value" id="anual-expense" style="color:var(--red)">R$ 0</div></div>
    <div class="card"><div class="card-label">Saldo Anual</div><div class="card-value" id="anual-balance">R$ 0</div></div>
  </div>

  <!-- Gráfico geral -->
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Mensal — Receitas vs Despesas</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Saldo</span>
      </div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-anual"></canvas></div>
  </div>

  <!-- Gráfico por categoria anual -->
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Evolução anual por categoria</div>
      <div class="hdr-actions" style="gap:8px">
        <select class="period-select" id="cat-anual-type" onchange="renderCatAnual()" style="font-size:12px">
          <option value="expense">Despesas</option>
          <option value="income">Receitas</option>
        </select>
        <select class="period-select" id="cat-anual-sel" onchange="renderCatAnual()" style="font-size:12px"></select>
      </div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-cat-anual"></canvas></div>
    <div id="cat-anual-summary" style="display:flex;gap:12px;flex-wrap:wrap;margin-top:14px;font-size:12px"></div>
  </div>

  <!-- Tabela mensal -->
  <div class="transactions-card">
    <div class="tx-header"><div class="tx-title">Resumo mensal</div></div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="border-bottom:1px solid var(--border)">
            <th style="padding:11px 18px;text-align:left;color:var(--text2);font-weight:500">Mês</th>
            <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Receitas</th>
            <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Despesas</th>
            <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Saldo</th>
            <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Qtd.</th>
          </tr>
        </thead>
        <tbody id="anual-tbody"></tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== IMPORTAR EXCEL ===== -->
<section id="sec-importar" class="section">
  <div class="page-header">
    <div><div class="page-title">Importar Excel</div><div class="page-sub">Compatível com a sua planilha</div></div>
  </div>
  <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="dropFile(event)">
    <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none" onchange="handleFile(this.files[0])">
    <div class="drop-icon">📂</div>
    <div class="drop-title">Arraste o arquivo Excel aqui</div>
    <div class="drop-sub">ou clique para selecionar • .xlsx / .xls</div>
    <div class="drop-sub" style="margin-top:6px;font-size:11px;opacity:.6">Formato: Dia | Gasto/Recebido | Departamento | Valor | Descrição | Mês | Ano</div>
  </div>
  <div id="import-preview" style="display:none">
    <div class="import-preview">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
        <div>
          <div style="font-size:14px;font-weight:500" id="import-filename">—</div>
          <div style="font-size:12px;color:var(--text2);margin-top:2px" id="import-fileinfo">—</div>
        </div>
        <div style="display:flex;gap:8px">
          <button class="btn" onclick="resetImport()">✕ Cancelar</button>
          <button class="btn btn-primary" id="btn-confirm-import" onclick="confirmImport()">✓ Importar</button>
        </div>
      </div>
      <div class="import-stats">
        <div class="istat"><div class="istat-val" id="imp-total">0</div><div class="istat-label">Lidas</div></div>
        <div class="istat"><div class="istat-val" id="imp-income" style="color:var(--green)">0</div><div class="istat-label">Receitas</div></div>
        <div class="istat"><div class="istat-val" id="imp-expense" style="color:var(--red)">0</div><div class="istat-label">Despesas</div></div>
        <div class="istat"><div class="istat-val" id="imp-skip" style="color:var(--amber)">0</div><div class="istat-label">Ignoradas</div></div>
      </div>
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:8px">
        <div style="font-size:13px;font-weight:500">Prévia (20 primeiras)</div>
        <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer">
          <input type="checkbox" id="imp-skip-dup" checked style="accent-color:var(--green)"> Ignorar duplicatas
        </label>
      </div>
      <div class="import-table-wrap">
        <table class="import-table">
          <thead><tr><th>Data</th><th>Tipo</th><th>Categoria</th><th>Valor</th><th>Descrição</th></tr></thead>
          <tbody id="import-tbody"></tbody>
        </table>
      </div>
      <div id="import-warn" class="import-warn" style="display:none"></div>
      <div style="margin-top:14px">
        <div style="font-size:12px;color:var(--text2);margin-bottom:6px">Novas categorias que serão criadas:</div>
        <div id="import-new-cats" style="display:flex;gap:6px;flex-wrap:wrap"></div>
      </div>
    </div>
  </div>
  <div class="chart-card" style="margin-top:18px">
    <div class="chart-header"><div class="chart-title">Histórico de importações</div></div>
    <div id="import-history-list"><div class="empty"><div class="empty-icon">📋</div>Nenhuma importação ainda</div></div>
  </div>
</section>

<!-- ===== GERENCIAR CATEGORIAS ===== -->
<section id="sec-gcat" class="section">
  <div class="page-header"><div><div class="page-title">Gerenciar Categorias</div><div class="page-sub">Crie e edite categorias personalizadas</div></div></div>
  <div class="cat-manager">
    <div class="cat-section-card">
      <div class="cat-section-title">⬇ Despesas</div>
      <div id="gcat-expense-list"></div>
      <div class="new-cat-form">
        <input class="form-input" id="new-exp-name" placeholder="Nome" style="flex:1;min-width:100px">
        <input type="color" class="color-picker" id="new-exp-color" value="#f56565">
        <input class="form-input" id="new-exp-icon" placeholder="🏷️" style="width:46px;text-align:center">
        <button class="btn btn-primary" onclick="addCat('expense')">+</button>
      </div>
    </div>
    <div class="cat-section-card">
      <div class="cat-section-title">⬆ Receitas</div>
      <div id="gcat-income-list"></div>
      <div class="new-cat-form">
        <input class="form-input" id="new-inc-name" placeholder="Nome" style="flex:1;min-width:100px">
        <input type="color" class="color-picker" id="new-inc-color" value="#22d3a0">
        <input class="form-input" id="new-inc-icon" placeholder="💰" style="width:46px;text-align:center">
        <button class="btn btn-primary" onclick="addCat('income')">+</button>
      </div>
    </div>
  </div>
</section>

<!-- ===== NUVEM ===== -->
<section id="sec-cloud" class="section">
  <div class="page-header">
    <div><div class="page-title">Sincronização na Nuvem</div><div class="page-sub">Acesse seus dados em qualquer dispositivo</div></div>
    <div id="sync-status-badge" class="sync-status sync-local"><span class="sync-dot local"></span>Armazenamento local</div>
  </div>

  <div class="cloud-card">
    <div style="font-size:15px;font-weight:500;margin-bottom:6px">Como funciona</div>
    <div style="font-size:13px;color:var(--text2);margin-bottom:20px;line-height:1.6">
      O FinTrack usa o <strong style="color:var(--text)">Firebase Realtime Database</strong> (gratuito até 1 GB) para sincronizar seus dados entre computador, celular e tablet em tempo real. Configure uma vez e pronto.
    </div>
    <div class="cloud-steps">
      <div class="cloud-step">
        <div class="cloud-step-num">1</div>
        <div class="cloud-step-body">
          <div class="cloud-step-title">Crie um projeto Firebase (gratuito)</div>
          <div class="cloud-step-desc">Acesse <a href="https://console.firebase.google.com" target="_blank">console.firebase.google.com</a> → "Adicionar projeto" → dê um nome → desative Google Analytics → Criar projeto.</div>
        </div>
      </div>
      <div class="cloud-step">
        <div class="cloud-step-num">2</div>
        <div class="cloud-step-body">
          <div class="cloud-step-title">Crie o Realtime Database</div>
          <div class="cloud-step-desc">No menu lateral → "Realtime Database" → "Criar banco de dados" → escolha localização → modo de teste (por 30 dias) → Ativar.</div>
        </div>
      </div>
      <div class="cloud-step">
        <div class="cloud-step-num">3</div>
        <div class="cloud-step-body">
          <div class="cloud-step-title">Copie a URL do banco</div>
          <div class="cloud-step-desc">Na página do Realtime Database, copie a URL que aparece no topo — parecida com <code style="background:var(--bg3);padding:1px 6px;border-radius:4px;font-size:11px">https://seu-projeto-default-rtdb.firebaseio.com</code></div>
        </div>
      </div>
      <div class="cloud-step">
        <div class="cloud-step-num">4</div>
        <div class="cloud-step-body">
          <div class="cloud-step-title">Cole a URL abaixo e salve</div>
          <div class="cloud-step-desc">Use o mesmo arquivo HTML em todos os dispositivos (salve localmente ou hospede no <a href="https://pages.github.com" target="_blank">GitHub Pages</a>).</div>
        </div>
      </div>
    </div>

    <div class="form-row" style="margin-bottom:0">
      <label class="form-label">URL do Firebase Realtime Database</label>
      <div style="display:flex;gap:8px">
        <input class="form-input" id="firebase-url" placeholder="https://seu-projeto-default-rtdb.firebaseio.com" style="flex:1">
        <button class="btn btn-primary" onclick="connectFirebase()">Conectar</button>
      </div>
      <div style="font-size:11px;color:var(--text3);margin-top:5px">A URL é salva localmente neste dispositivo. Insira a mesma URL em todos os dispositivos que quiser sincronizar.</div>
    </div>
  </div>

  <div class="chart-card" id="cloud-actions" style="display:none">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:16px">
      <div>
        <div style="font-size:14px;font-weight:500">Firebase conectado</div>
        <div style="font-size:12px;color:var(--text2);margin-top:2px" id="cloud-url-display">—</div>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn" onclick="syncFromCloud()">⬇ Baixar da nuvem</button>
        <button class="btn btn-primary" onclick="syncToCloud()">⬆ Enviar para nuvem</button>
        <button class="btn btn-danger" onclick="disconnectFirebase()">Desconectar</button>
      </div>
    </div>
    <div id="cloud-log" style="font-size:12px;color:var(--text2);background:var(--bg3);border-radius:8px;padding:12px;min-height:48px;line-height:1.7"></div>
  </div>
</section>

</main>
</div>

<!-- Mobile bottom bar -->
<div class="mobile-bar">
  <div class="mobile-bar-inner">
    <button class="mob-btn active" id="mob-dashboard" onclick="showSection('dashboard')"><span class="mob-icon">📊</span>Início</button>
    <button class="mob-btn" id="mob-transacoes" onclick="showSection('transacoes')"><span class="mob-icon">↕️</span>Transações</button>
    <button class="mob-btn" id="mob-add" onclick="openModal()"><span class="mob-icon">➕</span>Novo</button>
    <button class="mob-btn" id="mob-anual" onclick="showSection('anual')"><span class="mob-icon">📅</span>Anual</button>
    <button class="mob-btn" id="mob-menu" onclick="toggleMobileMenu()"><span class="mob-icon">☰</span>Menu</button>
  </div>
</div>

<!-- Modal nova/editar transação -->
<div class="modal-bg" id="modal-bg" onclick="closeModalBg(event)">
  <div class="modal">
    <div class="modal-title" id="modal-title">Nova transação</div>
    <div class="form-row">
      <label class="form-label">Tipo</label>
      <div class="radio-group">
        <div class="radio-btn sel-income" id="rb-income" onclick="setType('income')">⬆ Receita</div>
        <div class="radio-btn" id="rb-expense" onclick="setType('expense')">⬇ Despesa</div>
      </div>
    </div>
    <div class="form-row">
      <label class="form-label">Descrição</label>
      <input class="form-input" id="f-desc" placeholder="Ex: Salário, Aluguel, Mercado...">
    </div>
    <div class="form-row-2">
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Valor (R$)</label>
        <input class="form-input" id="f-amount" type="number" min="0.01" step="0.01" placeholder="0,00">
      </div>
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Data</label>
        <input class="form-input" id="f-date" type="date">
      </div>
    </div>
    <div class="form-row" style="margin-top:13px">
      <label class="form-label">Categoria</label>
      <select class="form-select" id="f-cat"></select>
    </div>
    <input type="hidden" id="f-editing-id">
    <div class="modal-actions">
      <button class="btn" onclick="closeModal()">Cancelar</button>
      <button class="btn btn-primary" onclick="saveTransaction()">Salvar</button>
    </div>
  </div>
</div>

<script>
// ===================== STORAGE =====================
const DB_KEY='fintrack_v2', CATS_KEY='fintrack_cats_v2', IMP_KEY='fintrack_imports', FB_KEY='fintrack_firebase_url';

const store = {
  async get(k){ if(window.storage){try{const r=await window.storage.get(k);return r?r.value:null}catch{}}return localStorage.getItem(k); },
  async set(k,v){ if(window.storage){try{await window.storage.set(k,v);return}catch{}}localStorage.setItem(k,v); }
};

const DEFAULT_CATS = {
  expense:[
    {id:'moradia',name:'Moradia',icon:'🏠',color:'#60a5fa',builtin:true},
    {id:'alimentacao',name:'Alimentação',icon:'🍽️',color:'#fb923c',builtin:true},
    {id:'transporte',name:'Transporte',icon:'🚗',color:'#a78bfa',builtin:true},
    {id:'saude',name:'Saúde',icon:'❤️',color:'#f87171',builtin:true},
    {id:'educacao',name:'Educação',icon:'📚',color:'#34d399',builtin:true},
    {id:'lazer',name:'Lazer',icon:'🎮',color:'#fbbf24',builtin:true},
    {id:'vestuario',name:'Vestuário',icon:'👕',color:'#c084fc',builtin:true},
    {id:'servicos',name:'Serviços',icon:'📱',color:'#38bdf8',builtin:true},
    {id:'outros_d',name:'Outros',icon:'📦',color:'#94a3b8',builtin:true},
  ],
  income:[
    {id:'salario',name:'Salário',icon:'💼',color:'#22d3a0',builtin:true},
    {id:'freelance',name:'Freelance',icon:'💻',color:'#60a5fa',builtin:true},
    {id:'investimento',name:'Investimento',icon:'📈',color:'#fbbf24',builtin:true},
    {id:'bonus',name:'Bônus',icon:'🎁',color:'#f472b6',builtin:true},
    {id:'outros_r',name:'Outros',icon:'💰',color:'#94a3b8',builtin:true},
  ]
};

let _db={transactions:[]}, _cats=JSON.parse(JSON.stringify(DEFAULT_CATS)), _imports=[], _fbUrl='';

async function loadAll(){
  try{
    const [dbR,catsR,impsR,fbR]=await Promise.all([store.get(DB_KEY),store.get(CATS_KEY),store.get(IMP_KEY),store.get(FB_KEY)]);
    _db   = dbR   ? JSON.parse(dbR)   : {transactions:[]};
    _cats = catsR ? JSON.parse(catsR) : JSON.parse(JSON.stringify(DEFAULT_CATS));
    _imports = impsR ? JSON.parse(impsR) : [];
    _fbUrl = fbR || '';
  }catch{}
}
function loadCats(){return _cats}
function loadDB(){return _db}
function loadImports(){return _imports}
async function saveCats(c){_cats=c;await store.set(CATS_KEY,JSON.stringify(c))}
async function saveDB(db){_db=db;await store.set(DB_KEY,JSON.stringify(db))}
async function saveImports(a){_imports=a;await store.set(IMP_KEY,JSON.stringify(a))}

function allCats(){const c=loadCats();return[...c.expense,...c.income]}
function getCat(id){return allCats().find(c=>c.id===id)||{name:id||'—',icon:'📦',color:'#888'}}
function slugify(s){return s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').substring(0,30)}
function randomColor(){const p=['#f87171','#fb923c','#fbbf24','#34d399','#60a5fa','#a78bfa','#f472b6','#38bdf8','#22d3a0','#c084fc'];return p[Math.floor(Math.random()*p.length)]}

const MONTHS=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const MONTHS_SHORT=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function fmt(v){return 'R$ '+Math.abs(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}

function getYears(){
  const db=loadDB(), now=new Date().getFullYear(), ys=new Set([now]);
  db.transactions.forEach(t=>ys.add(new Date(t.date+'T12:00').getFullYear()));
  return[...ys].sort((a,b)=>b-a);
}
function populateMonthSelect(id,cur){
  const sel=document.getElementById(id); if(!sel)return;
  sel.innerHTML=(id.startsWith('tx-')?'<option value="-1">Todos os meses</option>':'')+MONTHS.map((m,i)=>`<option value="${i}"${i===cur?' selected':''}>${m}</option>`).join('');
}
function populateYearSelect(id,cur){
  const sel=document.getElementById(id); if(!sel)return;
  sel.innerHTML=getYears().map(y=>`<option value="${y}"${y===cur?' selected':''}>${y}</option>`).join('');
}
function filterTx(txs,month,year){
  return txs.filter(t=>{const d=new Date(t.date+'T12:00');return d.getFullYear()===year&&(month===-1||d.getMonth()===month)});
}

// ===================== CHARTS STATE =====================
let chartMonthly=null,chartPie=null,chartCatPie=null,chartAnual=null,chartCatAnual=null;
function destroyChart(c){if(c){try{c.destroy()}catch{}}return null}

// ===================== SECTIONS =====================
const SEC_IDX={dashboard:0,transacoes:1,categorias:2,anual:3,importar:4,gcat:5,cloud:6};
function showSection(name){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('sec-'+name).classList.add('active');
  document.querySelectorAll('.nav-item')[SEC_IDX[name]]?.classList.add('active');
  // mobile bar
  document.querySelectorAll('.mob-btn').forEach(b=>b.classList.remove('active'));
  const mobMap={dashboard:'mob-dashboard',transacoes:'mob-transacoes',anual:'mob-anual'};
  if(mobMap[name]) document.getElementById(mobMap[name])?.classList.add('active');
  // close sidebar on mobile
  document.getElementById('sidebar').classList.remove('open');
  if(name==='dashboard') updateDashboard();
  if(name==='transacoes') renderTxList();
  if(name==='categorias') renderCategorias();
  if(name==='anual') renderAnual();
  if(name==='importar') renderImportHistory();
  if(name==='gcat') renderCatManager();
  if(name==='cloud') renderCloudSection();
}
function toggleMobileMenu(){document.getElementById('sidebar').classList.toggle('open')}

// ===================== MODAL NOVA/EDITAR =====================
let currentType='expense', editingId=null;

function openModal(id){
  editingId = id||null;
  const modal = document.getElementById('modal-bg');
  modal.classList.add('open');
  document.getElementById('modal-title').textContent = id ? 'Editar transação' : 'Nova transação';
  document.getElementById('f-editing-id').value = id||'';

  if(id){
    const t = loadDB().transactions.find(x=>x.id===id);
    if(!t) return;
    setType(t.type);
    document.getElementById('f-desc').value = t.description;
    document.getElementById('f-amount').value = t.amount;
    document.getElementById('f-date').value = t.date;
    // set category after setType populates the select
    document.getElementById('f-cat').value = t.category;
  } else {
    document.getElementById('f-date').value = new Date().toISOString().split('T')[0];
    document.getElementById('f-desc').value='';
    document.getElementById('f-amount').value='';
    setType('expense');
  }
}
function closeModal(){document.getElementById('modal-bg').classList.remove('open');editingId=null}
function closeModalBg(e){if(e.target.id==='modal-bg')closeModal()}

function setType(type){
  currentType=type;
  const cats=loadCats()[type];
  document.getElementById('f-cat').innerHTML=cats.map(c=>`<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
  document.getElementById('rb-income').className='radio-btn'+(type==='income'?' sel-income':'');
  document.getElementById('rb-expense').className='radio-btn'+(type==='expense'?' sel-expense':'');
}

async function saveTransaction(){
  const desc=document.getElementById('f-desc').value.trim();
  const amount=parseFloat(document.getElementById('f-amount').value);
  const date=document.getElementById('f-date').value;
  const cat=document.getElementById('f-cat').value;
  if(!desc||!amount||!date||amount<=0){alert('Preencha todos os campos.');return}
  const db=loadDB();
  const id=document.getElementById('f-editing-id').value;
  if(id){
    const idx=db.transactions.findIndex(t=>t.id===id);
    if(idx>=0) db.transactions[idx]={...db.transactions[idx],type:currentType,description:desc,amount,date,category:cat};
  } else {
    db.transactions.unshift({id:Date.now().toString(),type:currentType,description:desc,amount,date,category:cat});
  }
  db.transactions.sort((a,b)=>new Date(b.date)-new Date(a.date));
  await saveDB(db);
  closeModal();
  refreshAll();
}

async function delTx(id){
  if(!confirm('Excluir esta transação?'))return;
  const db=loadDB();
  db.transactions=db.transactions.filter(t=>t.id!==id);
  await saveDB(db);
  refreshAll();
}

// ===================== DASHBOARD =====================
function updateDashboard(){
  const now=new Date();
  const month=parseInt(document.getElementById('dash-month')?.value??now.getMonth());
  const year=parseInt(document.getElementById('dash-year')?.value??now.getFullYear());
  document.getElementById('dash-period-label').textContent=MONTHS[month]+' de '+year;
  const db=loadDB(), txs=filterTx(db.transactions,month,year);
  const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const balance=income-expense;
  document.getElementById('c-income').textContent=fmt(income);
  document.getElementById('c-income-count').textContent=txs.filter(t=>t.type==='income').length+' entrada(s)';
  document.getElementById('c-expense').textContent=fmt(expense);
  document.getElementById('c-expense-count').textContent=txs.filter(t=>t.type==='expense').length+' saída(s)';
  document.getElementById('c-balance').textContent=fmt(balance);
  document.getElementById('c-balance').style.color=balance>=0?'var(--green)':'var(--red)';
  document.getElementById('c-balance-pct').textContent=income>0?'Poupança: '+Math.round((balance/income)*100)+'%':'—';
  const catMap={};
  txs.filter(t=>t.type==='expense').forEach(t=>{catMap[t.category]=(catMap[t.category]||0)+t.amount});
  const topCat=Object.entries(catMap).sort((a,b)=>b[1]-a[1])[0];
  if(topCat){const c=getCat(topCat[0]);document.getElementById('c-topcat').textContent=c.icon+' '+c.name;document.getElementById('c-topcat-amt').textContent=fmt(topCat[1])}
  else{document.getElementById('c-topcat').textContent='—';document.getElementById('c-topcat-amt').textContent='Sem despesas'}
  // últimos 6 meses
  const last6=[];
  for(let i=5;i>=0;i--){let m=month-i,y=year;if(m<0){m+=12;y--}const t=filterTx(db.transactions,m,y);last6.push({label:MONTHS_SHORT[m],income:t.filter(x=>x.type==='income').reduce((s,x)=>s+x.amount,0),expense:t.filter(x=>x.type==='expense').reduce((s,x)=>s+x.amount,0)})}
  chartMonthly=destroyChart(chartMonthly);
  chartMonthly=new Chart(document.getElementById('chart-monthly'),{type:'bar',data:{labels:last6.map(x=>x.label),datasets:[{label:'Receitas',data:last6.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},{label:'Despesas',data:last6.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  // pizza
  const expCats=allCats().filter(c=>loadCats().expense.find(e=>e.id===c.id)).map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartPie=destroyChart(chartPie);
  const pieLeg=document.getElementById('pie-legend');
  if(!expCats.length){pieLeg.innerHTML='<div style="color:var(--text3);font-size:12px">Sem despesas no período</div>';return}
  chartPie=new Chart(document.getElementById('chart-pie'),{type:'doughnut',data:{labels:expCats.map(x=>x.cat.name),datasets:[{data:expCats.map(x=>x.total),backgroundColor:expCats.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  pieLeg.innerHTML=expCats.map(x=>`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-family:DM Mono;font-size:12px;color:var(--text)">${fmt(x.total)}</span></span>`).join('');
  pieLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
  renderTxItems(document.getElementById('recent-list'),db.transactions.slice(0,8));
}

// ===================== TX LIST =====================
let filterCat=null;
function renderTxList(){
  const month=parseInt(document.getElementById('tx-month')?.value??-1);
  const year=parseInt(document.getElementById('tx-year')?.value??new Date().getFullYear());
  const db=loadDB();
  let txs=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(filterCat)txs=txs.filter(t=>t.category===filterCat);
  const usedCats=[...new Set(db.transactions.map(t=>t.category))];
  const chips=document.getElementById('filter-chips');
  chips.innerHTML=`<div class="filter-chip ${filterCat===null?'active':''}" onclick="setCatFilter(null)">Todas</div>`+usedCats.map(cid=>{const c=getCat(cid);return`<div class="filter-chip ${filterCat===cid?'active':''}" onclick="setCatFilter('${cid}')">${c.icon} ${c.name}</div>`}).join('');
  renderTxItems(document.getElementById('tx-list'),txs);
}
function setCatFilter(cat){filterCat=cat;renderTxList()}

function renderTxItems(container,txs){
  if(!container)return;
  if(!txs.length){container.innerHTML='<div class="empty"><div class="empty-icon">🔍</div>Nenhuma transação encontrada</div>';return}
  container.innerHTML=txs.map(t=>{
    const c=getCat(t.category);
    const d=new Date(t.date+'T12:00');
    const amtColor=t.type==='income'?'var(--green)':'var(--red)';
    const sign=t.type==='income'?'+':'-';
    return`<div class="tx-item">
      <div class="tx-icon" style="background:${c.color}22">${c.icon}</div>
      <div class="tx-info"><div class="tx-name">${t.description}</div><div class="tx-cat">${c.name}</div></div>
      <div class="tx-date">${d.toLocaleDateString('pt-BR')}</div>
      <div class="tx-amount" style="color:${amtColor}">${sign} ${fmt(t.amount)}</div>
      <div class="tx-actions">
        <button class="tx-btn edit" onclick="openModal('${t.id}')" title="Editar">✏️</button>
        <button class="tx-btn del" onclick="delTx('${t.id}')" title="Excluir">🗑</button>
      </div>
    </div>`;
  }).join('');
}

// ===================== CATEGORIAS =====================
function renderCategorias(){
  const month=parseInt(document.getElementById('cat-month')?.value??new Date().getMonth());
  const year=parseInt(document.getElementById('cat-year')?.value??new Date().getFullYear());
  const db=loadDB(), txs=filterTx(db.transactions,month,year), cats=loadCats();
  function makeBars(type,cid){
    const total=txs.filter(t=>t.type===type).reduce((s,t)=>s+t.amount,0);
    const data=cats[type].map(c=>({cat:c,total:txs.filter(t=>t.type===type&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
    document.getElementById(cid).innerHTML=!data.length?'<div class="empty" style="padding:16px">Sem dados no período</div>':data.map(x=>{const pct=total>0?Math.round((x.total/total)*100):0;return`<div class="cat-row"><div class="cat-name-col">${x.cat.icon} ${x.cat.name}</div><div class="cat-bar-wrap"><div class="cat-bar" style="width:${pct}%;background:${x.cat.color}"></div></div><div class="cat-amt">${fmt(x.total)}</div><div class="cat-pct">${pct}%</div></div>`}).join('');
  }
  makeBars('expense','cat-expense-bars');
  makeBars('income','cat-income-bars');
  const expTotal=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const expData=cats.expense.map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartCatPie=destroyChart(chartCatPie);
  const catLeg=document.getElementById('cat-pie-legend');
  if(!expData.length){catLeg.innerHTML='';return}
  chartCatPie=new Chart(document.getElementById('chart-cat-pie'),{type:'doughnut',data:{labels:expData.map(x=>x.cat.name),datasets:[{data:expData.map(x=>x.total),backgroundColor:expData.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  catLeg.innerHTML=expData.map(x=>{const pct=expTotal>0?Math.round((x.total/expTotal)*100):0;return`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-size:11px;color:var(--text3)">${pct}%</span></span>`}).join('');
  catLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
}

// ===================== ANUAL =====================
function renderAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const db=loadDB();
  const months=MONTHS.map((name,m)=>{const txs=filterTx(db.transactions,m,year);const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);return{name,income,expense,balance:income-expense,count:txs.length}});
  const totI=months.reduce((s,m)=>s+m.income,0),totE=months.reduce((s,m)=>s+m.expense,0),totB=totI-totE;
  document.getElementById('anual-income').textContent=fmt(totI);
  document.getElementById('anual-expense').textContent=fmt(totE);
  document.getElementById('anual-balance').textContent=fmt(totB);
  document.getElementById('anual-balance').style.color=totB>=0?'var(--green)':'var(--red)';
  chartAnual=destroyChart(chartAnual);
  chartAnual=new Chart(document.getElementById('chart-anual'),{type:'bar',data:{labels:MONTHS_SHORT,datasets:[{label:'Receitas',data:months.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},{label:'Despesas',data:months.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4},{type:'line',label:'Saldo',data:months.map(x=>x.balance),borderColor:'rgba(96,165,250,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  document.getElementById('anual-tbody').innerHTML=months.map(m=>{const bc=m.balance>=0?'var(--green)':'var(--red)';return`<tr style="border-bottom:1px solid var(--border);${!m.count?'opacity:.35':''}"><td style="padding:11px 18px;font-weight:500">${m.name}</td><td style="padding:11px 18px;text-align:right;color:var(--green);font-family:DM Mono;font-size:12px">${m.income>0?fmt(m.income):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--red);font-family:DM Mono;font-size:12px">${m.expense>0?fmt(m.expense):'—'}</td><td style="padding:11px 18px;text-align:right;color:${bc};font-family:DM Mono;font-size:12px">${m.count?fmt(m.balance):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--text2)">${m.count||'—'}</td></tr>`}).join('');
  // popular select de categorias do gráfico por categoria
  populateCatAnualSelect();
  renderCatAnual();
}

function populateCatAnualSelect(){
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const sel=document.getElementById('cat-anual-sel'); if(!sel)return;
  const cats=loadCats()[type];
  sel.innerHTML=cats.map(c=>`<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
}

function renderCatAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const catId=document.getElementById('cat-anual-sel')?.value;
  if(!catId)return;
  const db=loadDB();
  const cat=getCat(catId);
  const monthData=MONTHS.map((name,m)=>{
    const txs=filterTx(db.transactions,m,year).filter(t=>t.type===type&&t.category===catId);
    return{name,short:MONTHS_SHORT[m],total:txs.reduce((s,t)=>s+t.amount,0),count:txs.length};
  });
  const annual=monthData.reduce((s,m)=>s+m.total,0);
  const avg=monthData.filter(m=>m.total>0).reduce((s,m,_,a)=>s+m.total/a.length,0);
  chartCatAnual=destroyChart(chartCatAnual);
  chartCatAnual=new Chart(document.getElementById('chart-cat-anual'),{
    type:'bar',
    data:{
      labels:monthData.map(x=>x.short),
      datasets:[{
        label:cat.name,
        data:monthData.map(x=>x.total),
        backgroundColor:monthData.map(x=>x.total>0?cat.color+'bb':'rgba(255,255,255,0.05)'),
        borderColor:cat.color,
        borderWidth:1,
        borderRadius:5,
      },{
        type:'line',
        label:'Média',
        data:monthData.map(()=>avg),
        borderColor:'rgba(255,255,255,0.25)',
        borderDash:[4,4],
        borderWidth:1.5,
        pointRadius:0,
        backgroundColor:'transparent',
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${fmt(ctx.raw)}`}}},
      scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(1)+'k'}}}
    }
  });
  const summaryEl=document.getElementById('cat-anual-summary');
  const bestMonth=monthData.reduce((a,b)=>b.total>a.total?b:a,monthData[0]);
  summaryEl.innerHTML=[
    `<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📦 Total anual: <strong style="color:${cat.color}">${fmt(annual)}</strong></span>`,
    `<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📊 Média mensal: <strong>${fmt(avg)}</strong></span>`,
    bestMonth.total>0?`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🏆 Maior mês: <strong>${bestMonth.name} (${fmt(bestMonth.total)})</strong></span>`:'',
    `<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🔢 Lançamentos: <strong>${monthData.reduce((s,m)=>s+m.count,0)}</strong></span>`,
  ].join('');
}

// ===================== IMPORT =====================
let pendingImport=[];
function dragOver(e){e.preventDefault();document.getElementById('drop-zone').classList.add('dragover')}
function dragLeave(){document.getElementById('drop-zone').classList.remove('dragover')}
function dropFile(e){e.preventDefault();dragLeave();const f=e.dataTransfer.files[0];if(f)handleFile(f)}

function handleFile(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=(e)=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});
      const sheetName=wb.SheetNames.find(s=>/gastos?/i.test(s))||wb.SheetNames[0];
      const ws=wb.Sheets[sheetName];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
      let headerIdx=0;
      for(let i=0;i<Math.min(10,rows.length);i++){
        const r=rows[i]; const hasDate=r.some(c=>c&&/^dia$|^data$/i.test(String(c).trim())); const hasValor=r.some(c=>c&&/^valor$/i.test(String(c).trim()));
        if(hasDate&&hasValor){headerIdx=i;break}
      }
      const hdr=rows[headerIdx].map(c=>String(c||'').toLowerCase().trim());
      let cDia=-1,cTipo=-1,cDept=-1,cValor=-1,cDesc=-1;
      hdr.forEach((h,i)=>{
        if(/^dia$|^data$/.test(h)) cDia=i;
        else if(/^valor$/.test(h)) cValor=i;
        else if(/departamento|categoria/.test(h)) cDept=i;
        else if(/descri/.test(h)) cDesc=i;
        else if((h===''||h==='unnamed')&&cTipo===-1) cTipo=i;
      });
      if(cDia<0)cDia=0; if(cTipo<0)cTipo=1; if(cDept<0)cDept=2; if(cValor<0)cValor=3; if(cDesc<0)cDesc=4;
      const dataRows=rows.slice(headerIdx+1);
      const parsed=[],skipped=[],newCatMap={};
      const cats=loadCats(), allExisting=[...cats.expense,...cats.income];
      const norm=s=>String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
      dataRows.forEach((row,idx)=>{
        if(!row){skipped.push(idx);return}
        const rawDate=row[cDia],rawTipo=String(row[cTipo]||'').trim(),rawDept=String(row[cDept]||'').trim(),rawValor=row[cValor],rawDesc=String(row[cDesc]||'').trim();
        if(!rawDate||rawValor==null||isNaN(Number(rawValor))){skipped.push(idx);return}
        const amount=Math.abs(Number(rawValor)); if(amount===0){skipped.push(idx);return}
        let dateStr='';
        if(rawDate instanceof Date&&!isNaN(rawDate)){const y=rawDate.getFullYear(),m=String(rawDate.getMonth()+1).padStart(2,'0'),d=String(rawDate.getDate()).padStart(2,'0');dateStr=`${y}-${m}-${d}`}
        else if(typeof rawDate==='number'){const d=new Date(Math.round((rawDate-25569)*86400*1000));if(!isNaN(d))dateStr=d.toISOString().split('T')[0]}
        else{const d=new Date(String(rawDate));if(!isNaN(d))dateStr=d.toISOString().split('T')[0]}
        if(!dateStr){skipped.push(idx);return}
        const tipoN=norm(rawTipo); let type='expense';
        if(/recebid|entrada|receita|income/.test(tipoN))type='income';
        const deptN=norm(rawDept); let catId=null;
        for(const c of allExisting){const cn=norm(c.name);if(deptN===cn){catId=c.id;break}if(deptN.length>=4&&cn.length>=4&&(deptN.startsWith(cn.slice(0,5))||cn.startsWith(deptN.slice(0,5)))){catId=c.id;break}}
        if(!catId){catId='imp_'+slugify(rawDept);if(!newCatMap[catId])newCatMap[catId]={id:catId,name:rawDept.trim(),type,icon:'🏷️',color:randomColor()}}
        const uid=`${dateStr}_${amount}_${catId}_${(rawDesc||rawDept).slice(0,12).replace(/\s/g,'_')}`;
        parsed.push({id:uid,type,description:rawDesc||rawDept,amount,date:dateStr,category:catId});
      });
      pendingImport=parsed;
      document.getElementById('import-filename').textContent=file.name;
      document.getElementById('import-fileinfo').textContent=`Aba: "${sheetName}" • ${dataRows.length} linhas • cols detectadas: Dia=${cDia} Tipo=${cTipo} Dept=${cDept} Valor=${cValor}`;
      document.getElementById('imp-total').textContent=parsed.length+skipped.length;
      document.getElementById('imp-income').textContent=parsed.filter(t=>t.type==='income').length;
      document.getElementById('imp-expense').textContent=parsed.filter(t=>t.type==='expense').length;
      document.getElementById('imp-skip').textContent=skipped.length;
      document.getElementById('import-tbody').innerHTML=parsed.slice(0,20).map(t=>{const d=new Date(t.date+'T12:00'),color=t.type==='income'?'var(--green)':'var(--red)',sign=t.type==='income'?'+':'-';return`<tr><td>${d.toLocaleDateString('pt-BR')}</td><td style="color:${color}">${t.type==='income'?'Receita':'Despesa'}</td><td>${getCat(t.category).name||t.category}</td><td style="color:${color};font-family:DM Mono">${sign} ${fmt(t.amount)}</td><td>${t.description}</td></tr>`}).join('');
      const newCats=Object.values(newCatMap);
      document.getElementById('import-new-cats').innerHTML=newCats.length?newCats.map(c=>`<span style="display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:12px"><span style="width:7px;height:7px;border-radius:50%;background:${c.color}"></span>${c.icon} ${c.name}</span>`).join(''):'<span style="font-size:12px;color:var(--text3)">Todas já existem</span>';
      const warn=document.getElementById('import-warn');
      if(!parsed.length){warn.style.display='block';warn.textContent='⚠ Nenhuma linha válida encontrada. Verifique o formato do arquivo.'}else{warn.style.display='none'}
      document.getElementById('import-preview').style.display='block';
    }catch(err){alert('Erro: '+err.message)}
  };
  reader.readAsBinaryString(file);
}

async function confirmImport(){
  if(!pendingImport.length){alert('Nenhum dado para importar.');return}
  const btn=document.getElementById('btn-confirm-import');
  btn.textContent='⏳ Salvando...';btn.disabled=true;
  const skipDup=document.getElementById('imp-skip-dup').checked;
  // novas cats
  const cats=loadCats(), allIds=new Set([...cats.expense,...cats.income].map(c=>c.id));
  pendingImport.forEach(t=>{
    if(!allIds.has(t.category)){
      const nc={id:t.category,name:t.category.replace('imp_','').replace(/_/g,' '),icon:'🏷️',color:randomColor(),builtin:false};
      cats[t.type]=cats[t.type]||[];cats[t.type].push(nc);allIds.add(t.category);
    }
  });
  await saveCats(cats);
  const db=loadDB(), existingIds=skipDup?new Set(db.transactions.map(t=>t.id)):new Set();
  let added=0,dupes=0;
  pendingImport.forEach(t=>{if(skipDup&&existingIds.has(t.id)){dupes++;return}db.transactions.push(t);existingIds.add(t.id);added++});
  db.transactions.sort((a,b)=>new Date(b.date)-new Date(a.date));
  await saveDB(db);
  const history=loadImports();
  history.unshift({date:new Date().toISOString(),added,dupes,total:pendingImport.length});
  await saveImports(history.slice(0,20));
  btn.textContent='✓ Importar';btn.disabled=false;
  alert(`✓ ${added} transações adicionadas${dupes?' ('+dupes+' duplicatas ignoradas)':''}.`);
  resetImport();populateAllSelects();refreshAll();renderImportHistory();
}

function resetImport(){pendingImport=[];document.getElementById('import-preview').style.display='none';document.getElementById('file-input').value=''}

function renderImportHistory(){
  const history=loadImports(), el=document.getElementById('import-history-list');
  if(!history.length){el.innerHTML='<div class="empty"><div class="empty-icon">📋</div>Nenhuma importação ainda</div>';return}
  el.innerHTML=history.map(h=>{const d=new Date(h.date);return`<div class="tx-item"><div class="tx-icon" style="background:var(--green-bg)">📥</div><div class="tx-info"><div class="tx-name">${h.added} transações importadas</div><div class="tx-cat">${h.dupes||0} duplicatas ignoradas</div></div><div class="tx-date">${d.toLocaleDateString('pt-BR')} ${d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})}</div></div>`}).join('');
}

// ===================== CAT MANAGER =====================
function renderCatManager(){
  const cats=loadCats();
  ['expense','income'].forEach(type=>{
    const el=document.getElementById('gcat-'+type+'-list');
    el.innerHTML=cats[type].map(c=>`<div class="cat-item-row"><div class="cat-color-dot" style="background:${c.color}"></div><div style="font-size:16px">${c.icon}</div><div class="cat-item-name">${c.name}</div>${c.builtin?'<span class="cat-item-badge">padrão</span>':''} ${!c.builtin?`<button class="cat-item-del" onclick="deleteCat('${type}','${c.id}')">✕</button>`:''}</div>`).join('')||'<div style="color:var(--text3);font-size:13px;padding:8px 0">Nenhuma</div>';
  });
}
async function addCat(type){
  const nE=document.getElementById(type==='expense'?'new-exp-name':'new-inc-name');
  const cE=document.getElementById(type==='expense'?'new-exp-color':'new-inc-color');
  const iE=document.getElementById(type==='expense'?'new-exp-icon':'new-inc-icon');
  const name=nE.value.trim(); if(!name){alert('Digite um nome.');return}
  const cats=loadCats();
  cats[type].push({id:slugify(name)+'_'+Date.now().toString(36),name,icon:iE.value.trim()||(type==='expense'?'📦':'💰'),color:cE.value,builtin:false});
  await saveCats(cats); nE.value='';iE.value=''; renderCatManager();
}
async function deleteCat(type,id){
  const cats=loadCats(), cat=cats[type].find(c=>c.id===id);
  if(cat?.builtin){alert('Categorias padrão não podem ser excluídas.');return}
  if(loadDB().transactions.some(t=>t.category===id)&&!confirm(`"${cat?.name}" está em uso. Excluir mesmo assim?`))return;
  cats[type]=cats[type].filter(c=>c.id!==id);
  await saveCats(cats); renderCatManager();
}

// ===================== NUVEM / FIREBASE =====================
function renderCloudSection(){
  const url=_fbUrl;
  document.getElementById('firebase-url').value=url||'';
  const badge=document.getElementById('sync-status-badge');
  if(url){
    badge.className='sync-status sync-ok';badge.innerHTML='<span class="sync-dot ok"></span>Firebase conectado';
    document.getElementById('cloud-actions').style.display='block';
    document.getElementById('cloud-url-display').textContent=url;
  } else {
    badge.className='sync-status sync-local';badge.innerHTML='<span class="sync-dot local"></span>Armazenamento local';
    document.getElementById('cloud-actions').style.display='none';
  }
}

async function connectFirebase(){
  let url=document.getElementById('firebase-url').value.trim();
  if(!url){alert('Cole a URL do Firebase.');return}
  if(!url.startsWith('https://'))url='https://'+url;
  if(url.endsWith('/'))url=url.slice(0,-1);
  // testa conexão
  try{
    const r=await fetch(url+'/.json?shallow=true');
    if(!r.ok)throw new Error('HTTP '+r.status);
    _fbUrl=url;
    await store.set(FB_KEY,url);
    logCloud('✓ Conectado a '+url);
    renderCloudSection();
  }catch(err){alert('Não foi possível conectar ao Firebase.\nVerifique a URL e as regras do banco.\n\nErro: '+err.message)}
}

function disconnectFirebase(){
  if(!confirm('Desconectar da nuvem? Os dados locais não serão apagados.'))return;
  _fbUrl='';store.set(FB_KEY,'');renderCloudSection();
}

async function syncToCloud(){
  if(!_fbUrl){alert('Configure o Firebase primeiro.');return}
  try{
    logCloud('⏳ Enviando dados...');
    const payload={db:JSON.stringify(_db),cats:JSON.stringify(_cats),imports:JSON.stringify(_imports),updated:new Date().toISOString()};
    const r=await fetch(_fbUrl+'/fintrack.json',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(!r.ok)throw new Error('HTTP '+r.status);
    logCloud('✓ Dados enviados com sucesso em '+new Date().toLocaleTimeString('pt-BR'));
  }catch(err){logCloud('✗ Erro ao enviar: '+err.message)}
}

async function syncFromCloud(){
  if(!_fbUrl){alert('Configure o Firebase primeiro.');return}
  try{
    logCloud('⏳ Baixando dados...');
    const r=await fetch(_fbUrl+'/fintrack.json');
    if(!r.ok)throw new Error('HTTP '+r.status);
    const data=await r.json();
    if(!data){logCloud('⚠ Nenhum dado encontrado na nuvem.');return}
    if(data.db) await saveDB(JSON.parse(data.db));
    if(data.cats) await saveCats(JSON.parse(data.cats));
    if(data.imports) await saveImports(JSON.parse(data.imports));
    populateAllSelects(); refreshAll();
    logCloud('✓ Dados baixados em '+new Date().toLocaleTimeString('pt-BR')+' (salvo em: '+new Date(data.updated).toLocaleString('pt-BR')+')');
  }catch(err){logCloud('✗ Erro ao baixar: '+err.message)}
}

function logCloud(msg){const el=document.getElementById('cloud-log');if(el)el.innerHTML=msg+'<br>'+el.innerHTML}

// ===================== REFRESH =====================
function refreshAll(){
  const active=document.querySelector('.section.active');
  const id=active?.id;
  if(id==='sec-dashboard')updateDashboard();
  if(id==='sec-transacoes')renderTxList();
  if(id==='sec-categorias')renderCategorias();
  if(id==='sec-anual')renderAnual();
}

function populateAllSelects(){
  const now=new Date(), cy=now.getFullYear(), cm=now.getMonth();
  populateMonthSelect('dash-month',cm); populateYearSelect('dash-year',cy);
  populateMonthSelect('tx-month',cm);   populateYearSelect('tx-year',cy);
  populateMonthSelect('cat-month',cm);  populateYearSelect('cat-year',cy);
  populateYearSelect('anual-year',cy);
}

// ===================== SEED =====================
async function seedDemo(){
  const now=new Date(),y=now.getFullYear(),m=now.getMonth();
  const p=(n)=>String(n).padStart(2,'0');
  const demos=[
    {type:'income',description:'Salário',amount:5500,date:`${y}-${p(m+1)}-05`,category:'salario'},
    {type:'expense',description:'Aluguel',amount:1500,date:`${y}-${p(m+1)}-07`,category:'moradia'},
    {type:'expense',description:'Supermercado',amount:480,date:`${y}-${p(m+1)}-10`,category:'alimentacao'},
    {type:'expense',description:'Uber / Gasolina',amount:220,date:`${y}-${p(m+1)}-12`,category:'transporte'},
    {type:'income',description:'Freelance Design',amount:1200,date:`${y}-${p(m+1)}-15`,category:'freelance'},
    {type:'expense',description:'Plano de Saúde',amount:350,date:`${y}-${p(m+1)}-15`,category:'saude'},
    {type:'expense',description:'Netflix + Spotify',amount:75,date:`${y}-${p(m+1)}-18`,category:'lazer'},
    {type:'income',description:'Rendimento CDB',amount:320,date:`${y}-${p(m+1)}-22`,category:'investimento'},
    {type:'expense',description:'Restaurante',amount:145,date:`${y}-${p(m+1)}-23`,category:'alimentacao'},
  ];
  const db=loadDB();
  demos.forEach(t=>{t.id=Date.now().toString()+Math.random();db.transactions.push(t)});
  await saveDB(db);
}

// ===================== INIT =====================
(async function init(){
  await loadAll();
  if(!_db.transactions.length) await seedDemo();
  populateAllSelects();
  updateDashboard();
  // Atualiza badge de sync
  if(_fbUrl){const b=document.getElementById('sync-status-badge');b.className='sync-status sync-ok';b.innerHTML='<span class="sync-dot ok"></span>Firebase conectado'}
})();
</script>
</body>
</html>
