/**
 * 东鹏整装渠道积分管理系统
 * 版本：2.0 | 数据驱动 | 支持Excel上传
 */

// ============= CONFIG =============
// 注册 Chart.js 数据标签插件
if (typeof Chart !== 'undefined' && Chart.register) {
  Chart.register(ChartDataLabels);
}

const DEFAULT_USERS = [
  { username: 'admin', password: 'dp2026', name: '系统管理员', role: 'admin', status: 'active' },
  { username: 'viewer', password: 'dp123', name: '只读用户', role: 'viewer', status: 'active' }
];

const DEFAULT_DATA = {
  title: "2026年东鹏整装渠道积分使用情况",
  lastUpdate: "2026-04-03",
  totalBudget: 6000000,
  note: "积分效能导向：用积分换业绩增量，ROI才是核心指标（万元销售消耗积分越低越优）",
  // 明细数据，用于展示区域积分使用详情
  detailItems: [],
  // 全国总实际业绩（万元）- 用于计算渠道ROI
  totalActualSales: 0,
  regions: [
    // 新增字段说明：
    // totalUsed: 预提报已用积分
    // totalUsedActual: 实际报销已用积分（从明细数据汇总）
    // actualSales: 实际达成销售业绩（万元）
    { name: "粤东", annualBudget: 593000, coMarketing: 32400, gamble: 33300, storeRenovation: 0, sceneFee: 0, totalUsed: 65700, totalUsedActual: 0, annualSales: 8900, actualSales: 0, salesPct: 9.89 },
    { name: "粤西", annualBudget: 447000, coMarketing: 0, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 0, totalUsedActual: 0, annualSales: 6700, actualSales: 0, salesPct: 7.44 },
    { name: "西南", annualBudget: 933000, coMarketing: 0, gamble: 0, storeRenovation: 0, sceneFee: 12500, totalUsed: 12500, totalUsedActual: 0, annualSales: 14000, actualSales: 0, salesPct: 15.56 },
    { name: "华东", annualBudget: 1163000, coMarketing: 44000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 44000, totalUsedActual: 0, annualSales: 17450, actualSales: 0, salesPct: 19.39 },
    { name: "华北", annualBudget: 473000, coMarketing: 18000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 18000, totalUsedActual: 0, annualSales: 7100, actualSales: 0, salesPct: 7.89 },
    { name: "鲁豫晋", annualBudget: 507000, coMarketing: 9000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 9000, totalUsedActual: 0, annualSales: 7600, actualSales: 0, salesPct: 8.44 },
    { name: "湘鄂", annualBudget: 697000, coMarketing: 0, gamble: 6000, storeRenovation: 0, sceneFee: 0, totalUsed: 6000, totalUsedActual: 0, annualSales: 10450, actualSales: 0, salesPct: 11.61 },
    { name: "东北", annualBudget: 193000, coMarketing: 21000, gamble: 21300, storeRenovation: 0, sceneFee: 0, totalUsed: 42300, totalUsedActual: 0, annualSales: 2900, actualSales: 0, salesPct: 3.22 },
    { name: "西北", annualBudget: 380000, coMarketing: 11000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 11000, totalUsedActual: 0, annualSales: 5700, actualSales: 0, salesPct: 6.33 },
    { name: "赣皖", annualBudget: 613000, coMarketing: 35000, gamble: 0, storeRenovation: 35000, sceneFee: 0, totalUsed: 70000, totalUsedActual: 0, annualSales: 9200, actualSales: 0, salesPct: 10.22 }
  ]
};

// ============= STATE =============
let appData = null;
let currentUser = null;
let users = [];
let charts = {};
let listData = [];
let filteredList = [];
let currentPage = 1;
const PAGE_SIZE = 8;

// ============= INIT =============
function init() {
  loadUsers();
  checkAutoLogin();
}

function loadUsers() {
  const saved = localStorage.getItem('dp_users');
  if (saved) {
    users = JSON.parse(saved);
  } else {
    users = JSON.parse(JSON.stringify(DEFAULT_USERS));
    saveUsers();
  }
}

function saveUsers() {
  localStorage.setItem('dp_users', JSON.stringify(users));
}

function loadData() {
  const saved = localStorage.getItem('dp_data');
  if (saved) {
    try {
      appData = JSON.parse(saved);
    } catch (e) {
      appData = JSON.parse(JSON.stringify(DEFAULT_DATA));
    }
  } else {
    appData = JSON.parse(JSON.stringify(DEFAULT_DATA));
  }
  // Compute derived fields
  appData.regions = appData.regions.map(r => computeRegion(r));
}

function saveData() {
  localStorage.setItem('dp_data', JSON.stringify(appData));
}

function computeRegion(r) {
  // 预提报已用积分
  const used = (r.coMarketing || 0) + (r.gamble || 0) + (r.storeRenovation || 0) + (r.sceneFee || 0);
  r.totalUsed = used;
  r.usedPct = r.annualBudget > 0 ? parseFloat(((used / r.annualBudget) * 100).toFixed(2)) : 0;
  
  // 实际报销已用积分（如果没有则使用预提报数据）
  const usedActual = r.totalUsedActual || 0;
  r.usedActualPct = r.annualBudget > 0 ? parseFloat(((usedActual / r.annualBudget) * 100).toFixed(2)) : 0;
  
  // 预提报ROI（基于年度销售目标）
  r.roi = r.annualSales > 0 ? parseFloat((used / r.annualSales).toFixed(2)) : 0;
  
  // 实际ROI（基于实际销售业绩）
  r.roiActual = r.actualSales > 0 ? parseFloat((usedActual / r.actualSales).toFixed(2)) : 0;
  
  // 业绩完成率
  r.salesCompletionRate = r.annualSales > 0 ? parseFloat(((r.actualSales || 0) / r.annualSales * 100).toFixed(2)) : 0;
  
  // 效能评级（基于预提报数据）
  r.rating = computeRating(r.usedPct, r.salesPct || 0);
  return r;
}

function computeRating(usedPct, salesPct) {
  if (usedPct === 0) return '待优化';
  const ratio = salesPct > 0 ? usedPct / salesPct : 999;
  if (ratio <= 0.7) return '高效投放';
  if (ratio <= 1.2) return '效能良好';
  if (ratio <= 2.0) return '需关注';
  return '待优化';
}

function getRatingBadgeClass(rating) {
  const map = {
    '高效投放': 'badge badge-excellent',
    '效能良好': 'badge badge-good',
    '需关注': 'badge badge-watch',
    '待优化': 'badge badge-pending'
  };
  return map[rating] || 'badge';
}

// 金额格式化 - 统一使用万元为单位
function formatMoney(val) {
  if (val === 0 || val === null || val === undefined) return '0万';
  // 转换为万元并保留1位小数
  const wan = (val / 10000).toFixed(1);
  // 去掉末尾的.0
  return parseFloat(wan) + '万';
}

// 完整金额格式化（用于表格等需要精确显示的场景）
function formatMoneyFull(val) {
  if (val === 0 || val === null || val === undefined) return '0万';
  const wan = (val / 10000).toFixed(2);
  return parseFloat(wan) + '万';
}

// 原始数值格式化（用于需要原始数字的场景，如计算）
function formatNumber(val) {
  return (val || 0).toLocaleString();
}

// ============= LOGIN =============
function checkAutoLogin() {
  const remembered = localStorage.getItem('dp_remembered');
  if (remembered) {
    try {
      const d = JSON.parse(remembered);
      const u = users.find(u => u.username === d.username);
      if (u && u.status === 'active') {
        currentUser = u;
        document.getElementById('username').value = u.username;
        document.getElementById('rememberMe').checked = true;
        startApp();
        return;
      }
    } catch (e) {}
  }
  showLogin();
}

function showLogin() {
  document.getElementById('loginPage').style.display = '';
  document.getElementById('loginPage').classList.add('active');
  document.getElementById('app').style.display = 'none';
}

function doLogin() {
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value;
  const remember = document.getElementById('rememberMe').checked;
  const errorDiv = document.getElementById('loginError');

  if (!username || !password) {
    errorDiv.textContent = '请输入账号和密码';
    errorDiv.style.display = 'block';
    return;
  }

  const user = users.find(u => u.username === username && u.password === password && u.status === 'active');
  if (!user) {
    errorDiv.textContent = '账号或密码错误，请重试';
    errorDiv.style.display = 'block';
    return;
  }

  errorDiv.style.display = 'none';
  currentUser = user;

  if (remember) {
    localStorage.setItem('dp_remembered', JSON.stringify({ username }));
  } else {
    localStorage.removeItem('dp_remembered');
  }

  startApp();
}

function doLogout() {
  localStorage.removeItem('dp_remembered');
  currentUser = null;
  destroyCharts();
  showLogin();
}

function togglePwd() {
  const input = document.getElementById('password');
  const icon = document.getElementById('eyeIcon');
  if (input.type === 'password') {
    input.type = 'text';
    icon.innerHTML = '<path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19m-6.72-1.07a3 3 0 11-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>';
  } else {
    input.type = 'password';
    icon.innerHTML = '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>';
  }
}

document.addEventListener('keydown', function(e) {
  if (e.key === 'Enter') {
    const loginPage = document.getElementById('loginPage');
    if (loginPage.classList.contains('active')) {
      doLogin();
    }
  }
});

// ============= APP START =============
function startApp() {
  loadData();
  document.getElementById('loginPage').style.display = 'none';
  document.getElementById('loginPage').classList.remove('active');
  document.getElementById('app').style.display = 'flex';
  
  // Set user info
  document.getElementById('sidebarUsername').textContent = currentUser.name || currentUser.username;
  document.getElementById('sidebarAvatar').textContent = (currentUser.name || currentUser.username).charAt(0).toUpperCase();
  document.getElementById('sidebarRole').textContent = currentUser.role === 'admin' ? '管理员' : '只读用户';
  
  renderDashboard();
  renderList();
  renderCategory();
  renderAccounts();
}

// ============= NAVIGATION =============
function showPage(page, el) {
  event.preventDefault();
  
  // Sidebar nav active
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  if (el) el.classList.add('active');
  
  // Inner pages
  document.querySelectorAll('.inner-page').forEach(p => p.classList.remove('active'));
  document.getElementById(page + 'Page').classList.add('active');
  
  const titles = {
    dashboard: '数据看板',
    list: '区域列表',
    category: '类目分析',
    detail: '积分明细',
    upload: '数据更新',
    accounts: '账号管理'
  };
  document.getElementById('pageTitle').textContent = titles[page] || '';
  
  // Close sidebar on mobile
  if (window.innerWidth <= 768) {
    document.getElementById('sidebar').classList.remove('open');
  }
  
  // Re-render charts when visible
  if (page === 'category') {
    setTimeout(renderCategoryCharts, 100);
  }
  if (page === 'dashboard') {
    setTimeout(renderDashboardCharts, 100);
  }
  if (page === 'detail') {
    setTimeout(initDetailPage, 100);
  }
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

function refreshAll() {
  loadData();
  renderDashboard();
  renderList();
  renderCategory();
  renderAccounts();
  showToast('数据已刷新');
}

// ============= DASHBOARD =============
function renderDashboard() {
  const data = appData;
  const regions = data.regions;
  
  // 预提报数据汇总
  const totalUsed = regions.reduce((s, r) => s + r.totalUsed, 0);
  // 实际报销数据汇总
  const totalUsedActual = regions.reduce((s, r) => s + (r.totalUsedActual || 0), 0);
  
  const totalBudget = data.totalBudget || regions.reduce((s, r) => s + r.annualBudget, 0);
  const totalSales = regions.reduce((s, r) => s + (r.annualSales || 0), 0);
  const totalActualSales = data.totalActualSales || regions.reduce((s, r) => s + (r.actualSales || 0), 0);
  
  // 渠道ROI预计 = 现使用积分 / 现渠道业绩
  const roi = totalActualSales > 0 ? (totalUsedActual / totalActualSales).toFixed(2) : 
              (totalSales > 0 ? (totalUsed / totalSales).toFixed(2) : 0);
  const usedPct = totalBudget > 0 ? ((totalUsed / totalBudget) * 100).toFixed(1) : 0;

  // 1. 积分使用情况：显示预算和已用积分（万元）
  document.getElementById('pointsUsage').textContent = `${formatMoney(totalUsed)} / ${formatMoney(totalBudget)}`;
  document.getElementById('pointsUsageSub').textContent = `消耗率 ${usedPct}%`;
  
  // 2. 渠道年度目标销售额（万元）
  document.getElementById('totalSales').textContent = formatNumber(totalSales);
  document.getElementById('salesNote').textContent = '万元';
  
  // 3. 渠道ROI预计
  document.getElementById('roiValue').textContent = roi;
  document.getElementById('dataTime').textContent = `数据截止：${data.lastUpdate || '--'}`;

  // Efficiency table - 更新后的表格结构（所有金额单位为万元）
  const tbody = document.getElementById('efficiencyTableBody');
  tbody.innerHTML = regions.map(r => `
    <tr>
      <td><strong>${r.name}</strong></td>
      <td>${formatMoneyFull(r.annualBudget)}</td>
      <!-- 预提报数据 -->
      <td>${formatMoneyFull(r.totalUsed)}</td>
      <td><span style="font-weight:700;color:${r.usedPct >= 15 ? 'var(--danger)' : r.usedPct >= 8 ? 'var(--warning)' : 'var(--success)'}">${r.usedPct}%</span></td>
      <!-- 实际报销数据 -->
      <td>${formatMoneyFull(r.totalUsedActual || 0)}</td>
      <td><span style="font-weight:700;color:${(r.usedActualPct || 0) >= 15 ? 'var(--danger)' : (r.usedActualPct || 0) >= 8 ? 'var(--warning)' : 'var(--success)'}">${r.usedActualPct || 0}%</span></td>
      <!-- 销售数据（万元） -->
      <td>${formatNumber(r.annualSales || 0)}</td>
      <td>${formatNumber(r.actualSales || 0)}</td>
      <td>${r.salesCompletionRate || 0}%</td>
      <!-- ROI -->
      <td>${r.roi || 0}</td>
      <td>${r.roiActual || 0}</td>
      <!-- 效能评级 -->
      <td><span class="${getRatingBadgeClass(r.rating)}">${r.rating}</span></td>
    </tr>
  `).join('');

  renderDashboardCharts();
}

function renderDashboardCharts() {
  const regions = appData.regions;
  
  // 1. 各运营中心使用积分占比（竖状图）
  destroyChart('regionBarChart');
  const ctx1 = document.getElementById('regionBarChart').getContext('2d');
  charts['regionBarChart'] = new Chart(ctx1, {
    type: 'bar',
    data: {
      labels: regions.map(r => r.name),
      datasets: [
        {
          label: '预提报积分',
          data: regions.map(r => r.totalUsed),
          backgroundColor: 'rgba(26, 86, 219, 0.8)',
          borderColor: 'rgba(26, 86, 219, 1)',
          borderWidth: 1,
          borderRadius: 4
        },
        {
          label: '实际报销积分',
          data: regions.map(r => r.totalUsedActual || 0),
          backgroundColor: 'rgba(22, 163, 74, 0.8)',
          borderColor: 'rgba(22, 163, 74, 1)',
          borderWidth: 1,
          borderRadius: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top',
          labels: { font: { size: 12 }, padding: 16 }
        },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatMoneyFull(ctx.parsed.y)}`
          }
        },
        datalabels: {
          display: true,
          anchor: 'end',
          align: 'top',
          offset: 4,
          color: '#1a1a1a',
          font: { size: 11, weight: 'bold' },
          formatter: (value) => formatMoney(value)
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { size: 11 } }
        },
        y: {
          beginAtZero: true,
          ticks: {
            callback: v => formatMoney(v),
            font: { size: 11 }
          },
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  });

  // 2. 全国各积分类目使用占比（竖状图）
  const cats = {
    '联合营销/大促': appData.regions.reduce((s, r) => s + (r.coMarketing || 0), 0),
    '对赌': appData.regions.reduce((s, r) => s + (r.gamble || 0), 0),
    '店中店装修': appData.regions.reduce((s, r) => s + (r.storeRenovation || 0), 0),
    '进场费': appData.regions.reduce((s, r) => s + (r.sceneFee || 0), 0)
  };
  
  const catLabels = Object.keys(cats).filter(k => cats[k] > 0);
  const catValues = catLabels.map(k => cats[k]);

  destroyChart('categoryBarChart2');
  if (catValues.length > 0) {
    const ctx2 = document.getElementById('categoryBarChart2').getContext('2d');
    charts['categoryBarChart2'] = new Chart(ctx2, {
      type: 'bar',
      data: {
        labels: catLabels,
        datasets: [{
          label: '积分使用金额（元）',
          data: catValues,
          backgroundColor: [
            'rgba(26, 86, 219, 0.8)',
            'rgba(249, 115, 22, 0.8)',
            'rgba(22, 163, 74, 0.8)',
            'rgba(220, 38, 38, 0.8)'
          ],
          borderWidth: 0,
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((ctx.parsed.y / total) * 100).toFixed(1) : 0;
                return `${ctx.label}: ${formatMoneyFull(ctx.parsed.y)} (${pct}%)`;
              }
            }
          },
          datalabels: {
            display: true,
            anchor: 'end',
            align: 'top',
            offset: 4,
            color: '#1a1a1a',
            font: { size: 11, weight: 'bold' },
            formatter: (value) => formatMoney(value)
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: 12 } }
          },
          y: {
            beginAtZero: true,
            ticks: {
              callback: v => formatMoney(v),
              font: { size: 11 }
            },
            grid: { color: 'rgba(0,0,0,0.05)' }
          }
        }
      }
    });
  }
}

// ============= LIST =============
function renderList() {
  listData = appData.regions;
  filterList();
}

function filterList() {
  const search = (document.getElementById('searchInput')?.value || '').toLowerCase();
  const rating = document.getElementById('ratingFilter')?.value || '';
  
  filteredList = listData.filter(r => {
    const matchName = r.name.toLowerCase().includes(search);
    const matchRating = !rating || r.rating === rating;
    return matchName && matchRating;
  });
  
  currentPage = 1;
  renderListPage();
}

function renderListPage() {
  const total = filteredList.length;
  const totalPages = Math.ceil(total / PAGE_SIZE);
  const start = (currentPage - 1) * PAGE_SIZE;
  const end = start + PAGE_SIZE;
  const pageData = filteredList.slice(start, end);

  const tbody = document.getElementById('listTableBody');
  const isAdmin = currentUser && currentUser.role === 'admin';
  
  tbody.innerHTML = pageData.map((r, i) => `
    <tr>
      <td><strong>${r.name}</strong></td>
      <td>${formatMoneyFull(r.annualBudget)}</td>
      <td>${formatMoneyFull(r.totalUsed)}</td>
      <td><span style="font-weight:700;color:${r.usedPct >= 15 ? 'var(--danger)' : r.usedPct >= 8 ? 'var(--warning)' : 'var(--success)'}">${r.usedPct}%</span></td>
      <td>${formatMoneyFull(r.coMarketing)}</td>
      <td>${formatMoneyFull(r.gamble)}</td>
      <td>${formatMoneyFull(r.storeRenovation)}</td>
      <td>${formatMoneyFull(r.sceneFee)}</td>
      <td>${formatNumber(r.annualSales || 0)}</td>
      <td>${r.salesPct || 0}%</td>
      <td><span class="${getRatingBadgeClass(r.rating)}">${r.rating}</span></td>
      <td>
        ${isAdmin ? `
        <div class="action-btns">
          <button class="btn-small btn-edit" onclick="openEditModal(${start + i})">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            编辑
          </button>
          <button class="btn-small btn-delete" onclick="confirmDelete(${start + i})">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2"/></svg>
            删除
          </button>
        </div>` : '<span style="color:var(--text-tertiary);font-size:12px">只读</span>'}
      </td>
    </tr>
  `).join('');

  // Pagination
  const pag = document.getElementById('pagination');
  if (totalPages <= 1) {
    pag.innerHTML = '';
    return;
  }
  
  let html = '';
  html += `<span>共 ${total} 条</span>`;
  if (currentPage > 1) html += `<button onclick="goPage(${currentPage - 1})">上一页</button>`;
  for (let i = 1; i <= totalPages; i++) {
    html += `<button class="${i === currentPage ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
  }
  if (currentPage < totalPages) html += `<button onclick="goPage(${currentPage + 1})">下一页</button>`;
  pag.innerHTML = html;
}

function goPage(page) {
  currentPage = page;
  renderListPage();
}

// ============= CATEGORY =============
function renderCategory() {
  const regions = appData.regions;
  const coMarketing = regions.reduce((s, r) => s + (r.coMarketing || 0), 0);
  const gamble = regions.reduce((s, r) => s + (r.gamble || 0), 0);
  const storeRenovation = regions.reduce((s, r) => s + (r.storeRenovation || 0), 0);
  const sceneFee = regions.reduce((s, r) => s + (r.sceneFee || 0), 0);

  document.getElementById('cat_coMarketing').textContent = formatMoney(coMarketing);
  document.getElementById('cat_gamble').textContent = formatMoney(gamble);
  document.getElementById('cat_storeRenovation').textContent = formatMoney(storeRenovation);
  document.getElementById('cat_sceneFee').textContent = formatMoney(sceneFee);

  renderCategoryCharts();
}

function renderCategoryCharts() {
  const regions = appData.regions;
  
  // Category bar chart
  destroyChart('categoryBarChart');
  const ctx3 = document.getElementById('categoryBarChart').getContext('2d');
  const catData = [
    { label: '联合营销/大促', value: regions.reduce((s, r) => s + (r.coMarketing || 0), 0), color: 'rgba(26, 86, 219, 0.8)' },
    { label: '对赌', value: regions.reduce((s, r) => s + (r.gamble || 0), 0), color: 'rgba(249, 115, 22, 0.8)' },
    { label: '店中店装修补贴', value: regions.reduce((s, r) => s + (r.storeRenovation || 0), 0), color: 'rgba(22, 163, 74, 0.8)' },
    { label: '进场费支持', value: regions.reduce((s, r) => s + (r.sceneFee || 0), 0), color: 'rgba(220, 38, 38, 0.8)' }
  ].filter(d => d.value > 0);

  charts['categoryBarChart'] = new Chart(ctx3, {
    type: 'bar',
    data: {
      labels: catData.map(d => d.label),
      datasets: [{
        label: '使用金额（万元）',
        data: catData.map(d => d.value),
        backgroundColor: catData.map(d => d.color),
        borderWidth: 0,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatMoneyFull(ctx.parsed.y)}`
          }
        },
        datalabels: {
          display: true,
          anchor: 'end',
          align: 'top',
          offset: 4,
          color: '#1a1a1a',
          font: { size: 11, weight: 'bold' },
          formatter: (value) => formatMoney(value)
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 12 } } },
        y: {
          beginAtZero: true,
          ticks: {
            callback: v => formatMoney(v),
            font: { size: 11 }
          },
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  });

  // Region stacked bar chart
  destroyChart('regionStackChart');
  const ctx4 = document.getElementById('regionStackChart').getContext('2d');
  charts['regionStackChart'] = new Chart(ctx4, {
    type: 'bar',
    data: {
      labels: regions.map(r => r.name),
      datasets: [
        {
          label: '联合营销',
          data: regions.map(r => r.coMarketing || 0),
          backgroundColor: 'rgba(26, 86, 219, 0.8)',
          borderWidth: 0
        },
        {
          label: '对赌',
          data: regions.map(r => r.gamble || 0),
          backgroundColor: 'rgba(249, 115, 22, 0.8)',
          borderWidth: 0
        },
        {
          label: '店中店装修',
          data: regions.map(r => r.storeRenovation || 0),
          backgroundColor: 'rgba(22, 163, 74, 0.8)',
          borderWidth: 0
        },
        {
          label: '进场费',
          data: regions.map(r => r.sceneFee || 0),
          backgroundColor: 'rgba(220, 38, 38, 0.8)',
          borderWidth: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', labels: { font: { size: 11 }, padding: 12 } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatMoneyFull(ctx.parsed.y)}`
          }
        },
        datalabels: {
          display: function(context) {
            // 只显示非零值
            return context.dataset.data[context.dataIndex] > 0;
          },
          anchor: 'center',
          align: 'center',
          color: 'white',
          font: { size: 10, weight: 'bold' },
          formatter: (value) => value > 0 ? formatMoney(value) : '',
          textShadowBlur: 4,
          textShadowColor: 'rgba(0,0,0,0.5)'
        }
      },
      scales: {
        x: {
          stacked: true,
          grid: { display: false },
          ticks: { font: { size: 11 } }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: {
            callback: v => formatMoney(v),
            font: { size: 11 }
          },
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  });
}

function destroyChart(id) {
  if (charts[id]) {
    charts[id].destroy();
    charts[id] = null;
  }
}

function destroyCharts() {
  Object.keys(charts).forEach(id => {
    if (charts[id]) {
      charts[id].destroy();
      charts[id] = null;
    }
  });
}

// ============= EDIT MODAL =============
let deleteTargetIndex = -1;

function openEditModal(index) {
  if (currentUser.role !== 'admin') {
    showToast('只读账号无权操作');
    return;
  }
  
  document.getElementById('editRegionIndex').value = index !== undefined ? index : -1;
  document.getElementById('modalTitle').textContent = index !== undefined ? '编辑区域数据' : '新增区域';
  
  if (index !== undefined) {
    const r = filteredList[index];
    document.getElementById('edit_name').value = r.name || '';
    document.getElementById('edit_annualBudget').value = r.annualBudget || 0;
    document.getElementById('edit_coMarketing').value = r.coMarketing || 0;
    document.getElementById('edit_gamble').value = r.gamble || 0;
    document.getElementById('edit_storeRenovation').value = r.storeRenovation || 0;
    document.getElementById('edit_sceneFee').value = r.sceneFee || 0;
    document.getElementById('edit_annualSales').value = r.annualSales || 0;
    document.getElementById('edit_h1Target').value = r.h1Target || 0;
    document.getElementById('edit_salesPct').value = r.salesPct || 0;
  } else {
    document.getElementById('edit_name').value = '';
    document.getElementById('edit_annualBudget').value = 0;
    document.getElementById('edit_coMarketing').value = 0;
    document.getElementById('edit_gamble').value = 0;
    document.getElementById('edit_storeRenovation').value = 0;
    document.getElementById('edit_sceneFee').value = 0;
    document.getElementById('edit_annualSales').value = 0;
    document.getElementById('edit_h1Target').value = 0;
    document.getElementById('edit_salesPct').value = 0;
  }

  document.getElementById('editModal').style.display = 'flex';
}

function closeModal() {
  document.getElementById('editModal').style.display = 'none';
}

function saveEdit() {
  const indexVal = document.getElementById('editRegionIndex').value;
  const idx = parseInt(indexVal);
  
  const region = {
    name: document.getElementById('edit_name').value.trim(),
    annualBudget: parseFloat(document.getElementById('edit_annualBudget').value) || 0,
    coMarketing: parseFloat(document.getElementById('edit_coMarketing').value) || 0,
    gamble: parseFloat(document.getElementById('edit_gamble').value) || 0,
    storeRenovation: parseFloat(document.getElementById('edit_storeRenovation').value) || 0,
    sceneFee: parseFloat(document.getElementById('edit_sceneFee').value) || 0,
    annualSales: parseFloat(document.getElementById('edit_annualSales').value) || 0,
    h1Target: parseFloat(document.getElementById('edit_h1Target').value) || 0,
    salesPct: parseFloat(document.getElementById('edit_salesPct').value) || 0
  };
  
  if (!region.name) {
    showToast('区域名称不能为空', 'error');
    return;
  }
  
  const computed = computeRegion(region);
  
  if (idx >= 0 && idx < filteredList.length) {
    // Find actual index in appData.regions
    const regionName = filteredList[idx].name;
    const actualIdx = appData.regions.findIndex(r => r.name === regionName);
    if (actualIdx >= 0) {
      appData.regions[actualIdx] = computed;
    }
  } else {
    appData.regions.push(computed);
  }
  
  saveData();
  closeModal();
  refreshAll();
  showToast('保存成功');
}

// ============= DELETE ============= 
function confirmDelete(index) {
  if (currentUser.role !== 'admin') {
    showToast('只读账号无权操作');
    return;
  }
  deleteTargetIndex = index;
  const r = filteredList[index];
  document.getElementById('confirmText').textContent = `确定要删除【${r.name}】区域的所有数据吗？此操作不可撤销。`;
  document.getElementById('confirmOkBtn').onclick = executeDelete;
  document.getElementById('confirmModal').style.display = 'flex';
}

function executeDelete() {
  if (deleteTargetIndex < 0) return;
  const r = filteredList[deleteTargetIndex];
  const actualIdx = appData.regions.findIndex(region => region.name === r.name);
  if (actualIdx >= 0) {
    appData.regions.splice(actualIdx, 1);
    saveData();
    closeConfirm();
    refreshAll();
    showToast(`已删除【${r.name}】`);
  }
}

function closeConfirm() {
  document.getElementById('confirmModal').style.display = 'none';
  deleteTargetIndex = -1;
}

// ============= ACCOUNTS =============
function renderAccounts() {
  const tbody = document.getElementById('accountsTableBody');
  const isAdmin = currentUser && currentUser.role === 'admin';
  
  tbody.innerHTML = users.map((u, i) => `
    <tr>
      <td><strong>${u.username}</strong></td>
      <td>${u.name || '-'}</td>
      <td><span class="badge ${u.role === 'admin' ? 'badge-excellent' : 'badge-good'}">${u.role === 'admin' ? '管理员' : '只读查看'}</span></td>
      <td><span class="badge ${u.status === 'active' ? 'badge-excellent' : 'badge-pending'}">${u.status === 'active' ? '启用' : '禁用'}</span></td>
      <td>
        ${isAdmin && u.username !== currentUser.username ? `
        <div class="action-btns">
          <button class="btn-small btn-edit" onclick="openEditAccount(${i})">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            编辑
          </button>
          <button class="btn-small btn-delete" onclick="toggleAccountStatus(${i})">
            ${u.status === 'active' ? '禁用' : '启用'}
          </button>
        </div>` : '<span style="color:var(--text-tertiary);font-size:12px">${isAdmin ? "当前账号" : "无权限"}</span>'}
      </td>
    </tr>
  `).join('');
}

function openAddAccount() {
  if (currentUser.role !== 'admin') {
    showToast('只读账号无权操作');
    return;
  }
  document.getElementById('editAccountIndex').value = -1;
  document.getElementById('accountModalTitle').textContent = '新增账号';
  document.getElementById('acc_username').value = '';
  document.getElementById('acc_password').value = '';
  document.getElementById('acc_name').value = '';
  document.getElementById('acc_role').value = 'viewer';
  document.getElementById('accountModal').style.display = 'flex';
}

function openEditAccount(i) {
  const u = users[i];
  document.getElementById('editAccountIndex').value = i;
  document.getElementById('accountModalTitle').textContent = '编辑账号';
  document.getElementById('acc_username').value = u.username;
  document.getElementById('acc_password').value = u.password;
  document.getElementById('acc_name').value = u.name || '';
  document.getElementById('acc_role').value = u.role;
  document.getElementById('accountModal').style.display = 'flex';
}

function closeAccountModal() {
  document.getElementById('accountModal').style.display = 'none';
}

function saveAccount() {
  const idxVal = document.getElementById('editAccountIndex').value;
  const idx = parseInt(idxVal);
  const username = document.getElementById('acc_username').value.trim();
  const password = document.getElementById('acc_password').value.trim();
  const name = document.getElementById('acc_name').value.trim();
  const role = document.getElementById('acc_role').value;
  
  if (!username || !password) {
    showToast('账号和密码不能为空', 'error');
    return;
  }
  
  // Check duplicate username
  if (idx < 0 && users.find(u => u.username === username)) {
    showToast('账号已存在', 'error');
    return;
  }
  
  const acc = { username, password, name, role, status: 'active' };
  
  if (idx >= 0) {
    users[idx] = { ...users[idx], ...acc };
  } else {
    users.push(acc);
  }
  
  saveUsers();
  closeAccountModal();
  renderAccounts();
  showToast('账号保存成功');
}

function toggleAccountStatus(i) {
  users[i].status = users[i].status === 'active' ? 'inactive' : 'active';
  saveUsers();
  renderAccounts();
  showToast('状态更新成功');
}

// ============= EXCEL UPLOAD =============
// Drag & Drop
const dropArea = document.getElementById('dropArea');
if (dropArea) {
  dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.style.background = '#dbeafe';
  });
  dropArea.addEventListener('dragleave', () => {
    dropArea.style.background = '';
  });
  dropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dropArea.style.background = '';
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
}

function handleFile(file) {
  if (!file) return;
  
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
    showToast('请上传 .xlsx 或 .xls 文件', 'error');
    return;
  }
  
  if (file.size > 10 * 1024 * 1024) {
    showToast('文件大小不能超过10MB', 'error');
    return;
  }
  
  const dropContent = document.getElementById('dropContent');
  dropContent.innerHTML = '<span style="color:var(--primary)">正在解析...</span>';
  
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      
      const parsed = parseExcelRows(rows);
      
      if (!parsed || parsed.length === 0) {
        showUploadError('未能识别到有效数据，请检查表格格式');
        dropContent.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg><span>点击选择文件或拖拽到此处</span><small>.xlsx / .xls 最大 10MB</small>`;
        return;
      }
      
      // Update data
      appData.regions = parsed.map(r => computeRegion(r));
      appData.lastUpdate = new Date().toISOString().slice(0, 10);
      saveData();
      
      refreshAll();
      
      const resultDiv = document.getElementById('uploadResult');
      resultDiv.style.display = 'block';
      resultDiv.innerHTML = `
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
          <svg style="width:20px;height:20px;color:var(--success)" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>
          <strong>上传成功！</strong>
        </div>
        <p>已解析 <strong>${parsed.length}</strong> 个区域的数据，数据已更新。</p>
        <p style="font-size:12px;margin-top:4px;color:var(--text-tertiary)">文件：${file.name} | 时间：${appData.lastUpdate}</p>
      `;
      
      dropContent.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg><span>点击选择文件或拖拽到此处</span><small>.xlsx / .xls 最大 10MB</small>`;
      
      showToast(`成功解析 ${parsed.length} 个区域`);
      
    } catch (err) {
      console.error(err);
      showUploadError('文件解析失败：' + err.message);
      dropContent.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg><span>点击选择文件或拖拽到此处</span><small>.xlsx / .xls 最大 10MB</small>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

function showUploadError(msg) {
  const resultDiv = document.getElementById('uploadResult');
  resultDiv.style.display = 'block';
  resultDiv.style.background = 'var(--danger-light)';
  resultDiv.style.color = 'var(--danger)';
  resultDiv.innerHTML = `<strong>解析失败：</strong>${msg}`;
}

function parseExcelRows(rows) {
  if (!rows || rows.length < 2) return [];
  
  // Find header row
  let headerRow = -1;
  let headerMap = {};
  
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i];
    const colMap = detectColumns(row);
    // 支持两种格式：汇总格式（有区域列）或明细格式（有大区列和项目类型列）
    if (colMap.name !== -1 || (colMap.region !== -1 && colMap.projectType !== -1)) {
      headerRow = i;
      headerMap = colMap;
      break;
    }
  }
  
  if (headerRow === -1) {
    // Try to guess: treat first row as header
    headerMap = detectColumns(rows[0]);
    headerRow = 0;
  }
  
  // 判断是否为明细格式（有大区和项目类型列）
  const isDetailFormat = headerMap.region !== -1 && headerMap.projectType !== -1;
  
  if (isDetailFormat) {
    return parseDetailExcelRows(rows, headerRow, headerMap);
  } else {
    return parseSummaryExcelRows(rows, headerRow, headerMap);
  }
}

// 解析汇总格式（每区域一行）
function parseSummaryExcelRows(rows, headerRow, headerMap) {
  const result = [];
  
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[headerMap.name] || row[headerMap.name] === '') continue;
    
    const name = String(row[headerMap.name]).trim();
    if (!name || name === '合计' || name === '总计' || name === '汇总') continue;
    
    const region = {
      name: name,
      annualBudget: toNum(headerMap.annualBudget >= 0 ? row[headerMap.annualBudget] : 0),
      coMarketing: toNum(headerMap.coMarketing >= 0 ? row[headerMap.coMarketing] : 0),
      gamble: toNum(headerMap.gamble >= 0 ? row[headerMap.gamble] : 0),
      storeRenovation: toNum(headerMap.storeRenovation >= 0 ? row[headerMap.storeRenovation] : 0),
      sceneFee: toNum(headerMap.sceneFee >= 0 ? row[headerMap.sceneFee] : 0),
      annualSales: toNum(headerMap.annualSales >= 0 ? row[headerMap.annualSales] : 0),
      h1Target: toNum(headerMap.h1Target >= 0 ? row[headerMap.h1Target] : 0),
      salesPct: toNum(headerMap.salesPct >= 0 ? row[headerMap.salesPct] : 0)
    };
    
    result.push(region);
  }
  
  return result;
}

// 解析明细格式（每项目一行，按区域和项目类型汇总）
function parseDetailExcelRows(rows, headerRow, headerMap) {
  const regionMap = new Map();
  
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 3) continue;
    
    // 获取区域名称
    const regionName = headerMap.region >= 0 ? String(row[headerMap.region] || '').trim() : '';
    if (!regionName || regionName === '合计' || regionName === '总计' || regionName === '汇总' || regionName === '全国') continue;
    
    // 获取项目类型
    const projectType = headerMap.projectType >= 0 ? String(row[headerMap.projectType] || '').trim() : '';
    
    // 获取积分金额（优先使用实际报销，其次使用申请金额）
    let amount = 0;
    if (headerMap.actualAmount >= 0) {
      amount = toNum(row[headerMap.actualAmount]);
    }
    if (amount === 0 && headerMap.applyAmount >= 0) {
      amount = toNum(row[headerMap.applyAmount]);
    }
    
    // 如果该区域还不存在，初始化
    if (!regionMap.has(regionName)) {
      regionMap.set(regionName, {
        name: regionName,
        annualBudget: 0,
        coMarketing: 0,
        gamble: 0,
        storeRenovation: 0,
        sceneFee: 0,
        annualSales: 0,
        h1Target: 0,
        salesPct: 0
      });
    }
    
    const region = regionMap.get(regionName);
    
    // 根据项目类型累加到对应类目
    const typeLower = projectType.toLowerCase();
    if (typeLower.includes('联合营销') || typeLower.includes('全国大促')) {
      region.coMarketing += amount;
    } else if (typeLower.includes('对赌')) {
      region.gamble += amount;
    } else if (typeLower.includes('店中店') || typeLower.includes('装修补贴')) {
      region.storeRenovation += amount;
    } else if (typeLower.includes('进场费')) {
      region.sceneFee += amount;
    }
  }
  
  return Array.from(regionMap.values());
}

function detectColumns(headerRow) {
  const map = {
    name: -1,
    region: -1,
    projectType: -1,
    applyAmount: -1,
    actualAmount: -1,
    annualBudget: -1,
    coMarketing: -1,
    gamble: -1,
    storeRenovation: -1,
    sceneFee: -1,
    annualSales: -1,
    h1Target: -1,
    salesPct: -1
  };
  
  if (!headerRow) return map;
  
  const keywords = {
    name: ['区域', '大区', '地区'],
    region: ['大区', '区域', '地区', '运营中心'],
    projectType: ['项目类型', '类型', '细项'],
    applyAmount: ['积分申请金额', '申请金额', '申请'],
    actualAmount: ['实际报销', '报销金额', '实际'],
    annualBudget: ['积分预算', '年度预算', '总预算'],
    coMarketing: ['联合营销', '全国大促', '大促'],
    gamble: ['对赌'],
    storeRenovation: ['店中店', '装修补贴', '装修'],
    sceneFee: ['进场费', '进场'],
    annualSales: ['年度销售', '销售额', '销售', '年销'],
    h1Target: ['h1', 'H1', '上半年目标', '上半年'],
    salesPct: ['业绩占比', '占比']
  };
  
  headerRow.forEach((cell, colIdx) => {
    if (!cell && cell !== 0) return;
    const cellStr = String(cell).replace(/\s/g, '').toLowerCase();
    
    Object.entries(keywords).forEach(([key, kws]) => {
      if (map[key] === -1) {
        for (const kw of kws) {
          if (cellStr.includes(kw.toLowerCase())) {
            map[key] = colIdx;
            break;
          }
        }
      }
    });
  });
  
  return map;
}

function toNum(val) {
  if (val === '' || val === null || val === undefined) return 0;
  const n = parseFloat(String(val).replace(/[,，%％]/g, ''));
  return isNaN(n) ? 0 : n;
}

// ============= TEMPLATE DOWNLOAD =============
function downloadTemplate(type) {
  let data = [];
  let filename = '';
  let sheetName = '';
  
  if (type === 'summary') {
    // 汇总格式模板
    filename = '积分提交模板-汇总格式.xlsx';
    sheetName = '区域汇总';
    data = [
      ['区域', '积分预算', '联合营销/全国大促', '对赌', '店中店装修补贴', '进场费支持', '年度销售额', 'H1目标', '业绩占比'],
      ['粤东', 593000, 32400, 33300, 0, 0, 8900, 27.3, 9.89],
      ['华东', 1163000, 26000, 0, 0, 0, 17450, 53.5, 19.39],
      ['西南', 933000, 0, 0, 0, 12500, 14000, 42.9, 15.56],
      ['华北', 473000, 21000, 0, 0, 0, 7100, 21.8, 7.89]
    ];
  } else if (type === 'detail') {
    // 明细格式模板
    filename = '积分提交模板-明细格式.xlsx';
    sheetName = '积分明细';
    data = [
      ['序号', '大区', '提报时间', '方案结束时间', '提报人', '归属经销商', '流水号', '项目类型', '细项', '积分申请金额', '实际报销', '备注'],
      [1, '华东运营中心', '2026-03-30', '2026-04-30', '张栋', '义务东鹏', 'RTM112-202603301387', '联合营销/全国大促', '义乌匠人', 3000, '', '新进驻N50匠人网点'],
      [2, '华东运营中心', '2026-03-30', '2026-04-30', '张栋', '金华东鹏', 'RTM112-202603301386', '联合营销/全国大促', '东鹏&风范装饰联合促销', 3000, '', 'N50战略客户'],
      [3, '西南运营中心', '2026-01-03', '2026-01-03', '', '四川德阳辰梓建材', 'RTM112-20260104919', '进场费支持', '德阳市旌阳区定艺装饰公司新店装修', 12500, '', '定艺装饰为N50头部装企'],
      [4, '东北运营中心', '', '', '', '沈阳长城陶瓷', '', '对赌', '', 21300, '', ''],
      [5, '赣皖运营中心', '2026-03-18', '2026-07-31', '刘跃明', '', 'RTM112-202603181269', '店中店装修补贴', '萍乡喜客喜新店装修用砖', 35000, '', '萍乡喜客喜新店装修']
    ];
  }
  
  // 创建工作簿
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  
  // 设置列宽
  const colWidths = data[0].map((_, idx) => {
    const maxLen = Math.max(...data.map(row => String(row[idx] || '').length));
    return { wch: Math.min(Math.max(maxLen + 2, 10), 40) };
  });
  ws['!cols'] = colWidths;
  
  // 添加工作表到工作簿
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  
  // 下载文件
  XLSX.writeFile(wb, filename);
  
  showToast(`已下载：${filename}`);
}

// ============= DETAIL PAGE =============
let detailItems = [];
let filteredDetailItems = [];
let detailCurrentPage = 1;
const DETAIL_PAGE_SIZE = 10;

function initDetailPage() {
  // 初始化区域下拉框
  const regionSelect = document.getElementById('detailRegionFilter');
  if (regionSelect && regionSelect.options.length <= 1) {
    appData.regions.forEach(r => {
      const option = document.createElement('option');
      option.value = r.name;
      option.textContent = r.name;
      regionSelect.appendChild(option);
    });
  }
  
  // 加载明细数据
  loadDetailItems();
  filterDetailList();
}

function loadDetailItems() {
  // 优先从 appData.detailItems 加载，如果没有则使用默认数据
  if (appData.detailItems && appData.detailItems.length > 0) {
    detailItems = appData.detailItems;
  } else {
    // 使用 Excel 中的默认明细数据
    detailItems = getDefaultDetailItems();
    appData.detailItems = detailItems;
    saveData();
  }
}

function getDefaultDetailItems() {
  return [
    { id: 1, region: '西南', submitDate: '2026-03-28', endDate: '2026-04-30', submitter: '', dealer: '成都浦利玛建材有限公司', serialNo: 'RTM112-202603281367', projectType: '联合营销/全国大促', item: '隆诚 4月高质活动', applyAmount: 0, actualAmount: '', remark: '非SAB类装企，且档期内所有瓷砖品牌均参与活动；活动参与方式为直降4%;同意执行，但不予以总部费用支持！' },
    { id: 2, region: '华东', submitDate: '2026-03-30', endDate: '2026-04-30', submitter: '张栋', dealer: '义务东鹏', serialNo: 'RTM112-202603301387', projectType: '联合营销/全国大促', item: '义乌匠人', applyAmount: 3000, actualAmount: '', remark: '新进驻N50匠人网点' },
    { id: 3, region: '华东', submitDate: '2026-03-30', endDate: '2026-04-30', submitter: '张栋', dealer: '金华东鹏', serialNo: 'RTM112-202603301386', projectType: '联合营销/全国大促', item: '东鹏&风范装饰联合促销（开门红以及工班大赛）', applyAmount: 3000, actualAmount: '', remark: 'N50需结合该月整体提货完成情况进行最终支持' },
    { id: 4, region: '鲁豫晋', submitDate: '2026-03-13', endDate: '2026-04-19', submitter: '山西运城悦霖建材', dealer: '山西运城悦霖建材', serialNo: 'RTM112-202603131229', projectType: '联合营销/全国大促', item: '九鼎装饰运城站', applyAmount: 3000, actualAmount: '', remark: '占有率保持第一，本月提货50万，若不达标不予兑现' },
    { id: 5, region: '鲁豫晋', submitDate: '2026-03-12', endDate: '2026-04-19', submitter: '山西运城悦霖建材', dealer: '山西运城悦霖建材', serialNo: 'RTM112-202603131227', projectType: '联合营销/全国大促', item: '业之峰装饰运城站', applyAmount: 3000, actualAmount: '', remark: '占有率保持第一，本月提货50万，若不达标不予兑现' },
    { id: 6, region: '华北', submitDate: '2026-03-02', endDate: '2026-06-30', submitter: '北京朗惠鸿泰科技', dealer: '北京朗惠鸿泰科技', serialNo: 'RTM112-202603021148', projectType: '联合营销/全国大促', item: '315北京被窝', applyAmount: 18000, actualAmount: '', remark: '' },
    { id: 7, region: '西南', submitDate: '2026-01-03', endDate: '2026-01-03', submitter: '', dealer: '四川德阳辰梓建材有限公司', serialNo: 'RTM112-20260104919', projectType: '进场费支持', item: '德阳市旌阳区定艺装饰公司新店装修', applyAmount: 12500, actualAmount: '', remark: '如期末未完成目标任务，按照相应比例扣回补贴积分；定艺装饰为N50头部装企' },
    { id: 8, region: '东北', submitDate: '2026-03-08', endDate: '2026-03-08', submitter: '辽宁沈阳兴华陶瓷', dealer: '辽宁沈阳兴华陶瓷', serialNo: 'RTM112-202603081189', projectType: '联合营销/全国大促', item: '哈尔滨爱林装饰', applyAmount: 5000, actualAmount: '', remark: '任务30单15万，按实际完成核销' },
    { id: 9, region: '东北', submitDate: '2026-03-05', endDate: '2026-03-25', submitter: '辽宁大连鹏东建筑', dealer: '辽宁大连鹏东建筑', serialNo: 'RTM112-202603051171', projectType: '联合营销/全国大促', item: '大连地区装企3月高值抢量活动：生活家', applyAmount: 8000, actualAmount: '', remark: '' },
    { id: 10, region: '赣皖', submitDate: '2026-03-10', endDate: '2026-03-10', submitter: '合肥东鹏', dealer: '合肥东鹏', serialNo: 'RTM112-202603101205', projectType: '联合营销/全国大促', item: '方林装饰高质', applyAmount: 10000, actualAmount: '', remark: 'N50战略客户活动激励金额以销售达成为准；需结合该月整体提货完成情况进行最终支持；' },
    { id: 11, region: '粤东', submitDate: '2026-03-12', endDate: '2026-03-31', submitter: '陶真', dealer: '', serialNo: 'RTM112-202603121215', projectType: '联合营销/全国大促', item: '佳美域', applyAmount: 6400, actualAmount: '', remark: '佳美域为B类头部，深圳占有率超60%，需要以活动形式保持高占' },
    { id: 12, region: '粤东', submitDate: '2026-03-12', endDate: '2026-03-31', submitter: '陈丽如', dealer: '', serialNo: 'RTM112-202603121223', projectType: '联合营销/全国大促', item: '鼎昇整装', applyAmount: 8000, actualAmount: '', remark: '' },
    { id: 13, region: '西南', submitDate: '', endDate: '', submitter: '', dealer: '成都浦利玛建材有限公司', serialNo: 'RTM112-202603121219', projectType: '联合营销/全国大促', item: '美居3月', applyAmount: 0, actualAmount: '', remark: '全屋东鹏-500两厨卫/客厅大地砖-300' },
    { id: 14, region: '华东', submitDate: '2026-03-10', endDate: '2026-04-12', submitter: '江苏南京创想东鹏', dealer: '江苏南京创想东鹏', serialNo: 'RTM112-202603101202', projectType: '联合营销/全国大促', item: '315金管家装饰', applyAmount: 5000, actualAmount: '', remark: '' },
    { id: 15, region: '华东', submitDate: '2026-03-10', endDate: '2026-04-12', submitter: '江苏南京创想东鹏', dealer: '江苏南京创想东鹏', serialNo: 'RTM112-202603101201', projectType: '联合营销/全国大促', item: '315鲸匠装饰', applyAmount: 5000, actualAmount: '', remark: '' },
    { id: 16, region: '华东', submitDate: '2026-03-10', endDate: '2026-04-12', submitter: '江苏南京创想东鹏', dealer: '江苏南京创想东鹏', serialNo: 'RTM112-202603101203', projectType: '联合营销/全国大促', item: '315我乐装饰', applyAmount: 5000, actualAmount: '', remark: '' },
    { id: 17, region: '华东', submitDate: '2026-03-10', endDate: '2026-04-12', submitter: '江苏南京创想东鹏', dealer: '江苏南京创想东鹏', serialNo: 'RTM112-202603101200', projectType: '联合营销/全国大促', item: '315锦华装饰', applyAmount: 5000, actualAmount: '', remark: '' },
    { id: 18, region: '东北', submitDate: '2026-03-05', endDate: '2026-03-25', submitter: '辽宁大连鹏东建筑', dealer: '辽宁大连鹏东建筑', serialNo: 'RTM112-202603051172', projectType: '联合营销/全国大促', item: '大连地区装企3月高值抢量活动：方林', applyAmount: 8000, actualAmount: '', remark: '' },
    { id: 19, region: '华东', submitDate: '2026-02-28', endDate: '2026-03-31', submitter: '张栋', dealer: '杭州东鹏', serialNo: 'RTM112-202602281136', projectType: '联合营销/全国大促', item: '浙江圣都', applyAmount: 18000, actualAmount: '', remark: '' },
    { id: 20, region: '西南', submitDate: '2025-12-25', endDate: '2026-01-31', submitter: '成都浦利玛建材有限公司', dealer: '成都浦利玛建材有限公司', serialNo: 'RTM112-20251225843', projectType: '联合营销/全国大促', item: '美居 26年元旦活动报备', applyAmount: 0, actualAmount: '', remark: '具体需等26年政策确认再评估2个空间选东鹏-200' },
    { id: 21, region: '全国', submitDate: '2026-02-03', endDate: '2026-06-30', submitter: '郭筱芊', dealer: '全国创艺活动服务商适用', serialNo: 'RTM112-202602031086', projectType: '联合营销/全国大促', item: '2026-鹏创美家|创艺x东鹏315开门红全国总对总营销活动', applyAmount: 0, actualAmount: '', remark: '免费升级（612 /715）' },
    { id: 22, region: '赣皖', submitDate: '2026-02-03', endDate: '2026-02-03', submitter: '郭筱芊', dealer: '江西丛一楼服务商', serialNo: 'RTM112-202602031088', projectType: '联合营销/全国大促', item: '江西丛一楼装饰315开年新春大促', applyAmount: 25000, actualAmount: '', remark: '活动套餐免费送砖' },
    { id: 23, region: '粤东', submitDate: '2026-03-11', endDate: '2026-03-31', submitter: '陶真', dealer: '', serialNo: 'RTM112-202603111210', projectType: '联合营销/全国大促', item: '贝壳 * 深圳圣都一周年庆典', applyAmount: 18000, actualAmount: '', remark: '3-4月沟通后为圣都东鹏专场' },
    { id: 24, region: '东北', submitDate: '', endDate: '', submitter: '', dealer: '沈阳长城陶瓷有限公司', serialNo: '', projectType: '对赌', item: '', applyAmount: 21300, actualAmount: '', remark: '' },
    { id: 25, region: '湘鄂', submitDate: '', endDate: '', submitter: '', dealer: '湖南铭颖建材销售有限公司', serialNo: '', projectType: '对赌', item: '', applyAmount: 6000, actualAmount: '', remark: '' },
    { id: 26, region: '粤东', submitDate: '', endDate: '', submitter: '', dealer: '石狮市三跃建材贸易有限公司', serialNo: '', projectType: '对赌', item: '', applyAmount: 33300, actualAmount: '', remark: '' },
    { id: 27, region: '鲁豫晋', submitDate: '2026-03-17', endDate: '2026-04-17', submitter: '山东济宁希金装饰', dealer: '山东济宁希金装饰', serialNo: 'RTM112-202603171255', projectType: '联合营销/全国大促', item: '3月份山东济宁生活家装饰', applyAmount: 3000, actualAmount: '', remark: '' },
    { id: 28, region: '西北', submitDate: '2026-03-19', endDate: '2026-04-12', submitter: '吕前君', dealer: '西安服务商', serialNo: 'RTM112-202603191271', projectType: '联合营销/全国大促', item: '峰光无限装饰3月', applyAmount: 3000, actualAmount: '', remark: '' },
    { id: 29, region: '赣皖', submitDate: '2026-03-18', endDate: '2026-07-31', submitter: '刘跃明', dealer: '', serialNo: 'RTM112-202603181269', projectType: '店中店装修补贴', item: '萍乡喜客喜新店装修用砖', applyAmount: 35000, actualAmount: '', remark: '萍乡喜客喜新店装修，需赞助店面用砖，约960m，预计费用3.5万。现由南昌东鹏310590提供装修用砖' },
    { id: 30, region: '西北', submitDate: '2026-03-17', endDate: '2026-03-20', submitter: '兰州万恒公司', dealer: '兰州万恒公司', serialNo: 'RTM112-202603171257', projectType: '联合营销/全国大促', item: '兰州壹品3月高质', applyAmount: 8000, actualAmount: '', remark: '' }
  ];
}

function filterDetailList() {
  const search = (document.getElementById('detailSearchInput')?.value || '').toLowerCase();
  const region = document.getElementById('detailRegionFilter')?.value || '';
  const type = document.getElementById('detailTypeFilter')?.value || '';
  
  filteredDetailItems = detailItems.filter(item => {
    const matchSearch = !search || 
      item.region.toLowerCase().includes(search) ||
      item.dealer.toLowerCase().includes(search) ||
      item.item.toLowerCase().includes(search) ||
      item.serialNo.toLowerCase().includes(search);
    const matchRegion = !region || item.region === region;
    const matchType = !type || item.projectType === type;
    return matchSearch && matchRegion && matchType;
  });
  
  // 更新区域汇总卡片
  updateRegionSummary(region);
  
  detailCurrentPage = 1;
  renderDetailTable();
}

function updateRegionSummary(regionName) {
  const summaryCard = document.getElementById('regionSummaryCard');
  
  if (!regionName) {
    summaryCard.style.display = 'none';
    return;
  }
  
  const region = appData.regions.find(r => r.name === regionName);
  if (!region) {
    summaryCard.style.display = 'none';
    return;
  }
  
  summaryCard.style.display = 'block';
  document.getElementById('summaryRegionName').textContent = region.name;
  document.getElementById('summaryRegionBadge').textContent = `总预算: ${formatMoneyFull(region.annualBudget)}`;
  document.getElementById('summaryUsed').textContent = formatMoneyFull(region.totalUsed);
  document.getElementById('summaryUsedPct').textContent = region.usedPct + '%';
  document.getElementById('summaryRoi').textContent = region.roi;
  document.getElementById('summaryRating').innerHTML = `<span class="${getRatingBadgeClass(region.rating)}">${region.rating}</span>`;
}

function renderDetailTable() {
  const tbody = document.getElementById('detailTableBody');
  const total = filteredDetailItems.length;
  const totalPages = Math.ceil(total / DETAIL_PAGE_SIZE);
  const start = (detailCurrentPage - 1) * DETAIL_PAGE_SIZE;
  const end = start + DETAIL_PAGE_SIZE;
  const pageData = filteredDetailItems.slice(start, end);
  
  document.getElementById('detailCount').textContent = `共 ${total} 条记录`;
  
  if (total === 0) {
    tbody.innerHTML = `<tr><td colspan="12" class="empty-tip">暂无数据，请选择区域或上传明细数据查看</td></tr>`;
    document.getElementById('detailPagination').innerHTML = '';
    return;
  }
  
  tbody.innerHTML = pageData.map((item, idx) => `
    <tr>
      <td>${item.id}</td>
      <td><strong>${item.region}</strong></td>
      <td>${item.submitDate || '-'}</td>
      <td>${item.endDate || '-'}</td>
      <td>${item.submitter || '-'}</td>
      <td>${item.dealer || '-'}</td>
      <td><code>${item.serialNo || '-'}</code></td>
      <td><span class="badge ${getProjectTypeBadgeClass(item.projectType)}">${item.projectType}</span></td>
      <td>${item.item || '-'}</td>
      <td>${item.applyAmount === 0 || item.applyAmount === '0' ? '-' : formatMoneyFull(item.applyAmount)}</td>
      <td>${item.actualAmount ? formatMoneyFull(item.actualAmount) : '-'}</td>
      <td title="${item.remark || ''}">${item.remark ? item.remark.substring(0, 20) + (item.remark.length > 20 ? '...' : '') : '-'}</td>
    </tr>
  `).join('');
  
  // Pagination
  const pag = document.getElementById('detailPagination');
  if (totalPages <= 1) {
    pag.innerHTML = '';
    return;
  }
  
  let html = `<span>共 ${total} 条</span>`;
  if (detailCurrentPage > 1) html += `<button onclick="goDetailPage(${detailCurrentPage - 1})">上一页</button>`;
  for (let i = 1; i <= totalPages; i++) {
    if (i === 1 || i === totalPages || (i >= detailCurrentPage - 1 && i <= detailCurrentPage + 1)) {
      html += `<button class="${i === detailCurrentPage ? 'active' : ''}" onclick="goDetailPage(${i})">${i}</button>`;
    } else if (i === detailCurrentPage - 2 || i === detailCurrentPage + 2) {
      html += `<span>...</span>`;
    }
  }
  if (detailCurrentPage < totalPages) html += `<button onclick="goDetailPage(${detailCurrentPage + 1})">下一页</button>`;
  pag.innerHTML = html;
}

function goDetailPage(page) {
  detailCurrentPage = page;
  renderDetailTable();
}

function getProjectTypeBadgeClass(type) {
  const map = {
    '联合营销/全国大促': 'badge-excellent',
    '对赌': 'badge-good',
    '店中店装修补贴': 'badge-watch',
    '进场费支持': 'badge-pending'
  };
  return map[type] || '';
}

// ============= TOAST =============
let toastTimer = null;
function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.style.display = 'block';
  toast.style.background = type === 'error' ? 'var(--danger)' : 'var(--text-primary)';
  
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.style.display = 'none';
  }, 2500);
}

// ============= CLOSE MODALS ON OVERLAY CLICK =============
document.getElementById('editModal').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});

document.getElementById('accountModal').addEventListener('click', function(e) {
  if (e.target === this) closeAccountModal();
});

document.getElementById('confirmModal').addEventListener('click', function(e) {
  if (e.target === this) closeConfirm();
});

// ============= START =============
init();
