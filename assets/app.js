/**
 * 东鹏整装渠道积分管理系统
 * 版本：2.0 | 数据驱动 | 支持Excel上传
 */

// ============= CONFIG =============
const DEFAULT_USERS = [
  { username: 'admin', password: 'dp2026', name: '系统管理员', role: 'admin', status: 'active' },
  { username: 'viewer', password: 'dp123', name: '只读用户', role: 'viewer', status: 'active' }
];

const DEFAULT_DATA = {
  title: "2026年东鹏整装渠道积分使用情况",
  lastUpdate: "2026-04-03",
  totalBudget: 6000000,
  note: "积分效能导向：用积分换业绩增量，ROI才是核心指标（万元销售消耗积分越低越优）",
  regions: [
    { name: "粤东", annualBudget: 593000, coMarketing: 32400, gamble: 33300, storeRenovation: 0, sceneFee: 0, totalUsed: 65700, annualSales: 8900, h1Target: 27.3, salesPct: 9.89 },
    { name: "粤西", annualBudget: 447000, coMarketing: 0, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 0, annualSales: 6700, h1Target: 20.5, salesPct: 7.44 },
    { name: "西南", annualBudget: 933000, coMarketing: 0, gamble: 0, storeRenovation: 0, sceneFee: 12500, totalUsed: 12500, annualSales: 14000, h1Target: 42.9, salesPct: 15.56 },
    { name: "华东", annualBudget: 1163000, coMarketing: 26000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 26000, annualSales: 17450, h1Target: 53.5, salesPct: 19.39 },
    { name: "华北", annualBudget: 473000, coMarketing: 21000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 21000, annualSales: 7100, h1Target: 21.8, salesPct: 7.89 },
    { name: "鲁豫晋", annualBudget: 507000, coMarketing: 3000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 3000, annualSales: 7600, h1Target: 23.3, salesPct: 8.44 },
    { name: "湘鄂", annualBudget: 697000, coMarketing: 0, gamble: 0, storeRenovation: 6000, sceneFee: 0, totalUsed: 6000, annualSales: 10450, h1Target: 32.0, salesPct: 11.61 },
    { name: "东北", annualBudget: 193000, coMarketing: 13000, gamble: 21300, storeRenovation: 0, sceneFee: 0, totalUsed: 34300, annualSales: 2900, h1Target: 8.9, salesPct: 3.22 },
    { name: "西北", annualBudget: 380000, coMarketing: 11000, gamble: 0, storeRenovation: 0, sceneFee: 0, totalUsed: 11000, annualSales: 5700, h1Target: 17.5, salesPct: 6.33 },
    { name: "赣皖", annualBudget: 613000, coMarketing: 35000, gamble: 0, storeRenovation: 35000, sceneFee: 0, totalUsed: 70000, annualSales: 9200, h1Target: 28.2, salesPct: 10.22 }
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
  const used = (r.coMarketing || 0) + (r.gamble || 0) + (r.storeRenovation || 0) + (r.sceneFee || 0);
  r.totalUsed = used;
  r.usedPct = r.annualBudget > 0 ? parseFloat(((used / r.annualBudget) * 100).toFixed(2)) : 0;
  r.roi = r.annualSales > 0 ? parseFloat((used / r.annualSales).toFixed(2)) : 0;
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

function formatMoney(val) {
  if (val === 0 || val === null || val === undefined) return '0';
  if (val >= 10000) return (val / 10000).toFixed(1) + '万';
  return val.toLocaleString();
}

function formatMoneyFull(val) {
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
  const totalUsed = regions.reduce((s, r) => s + r.totalUsed, 0);
  const totalBudget = data.totalBudget || regions.reduce((s, r) => s + r.annualBudget, 0);
  const totalSales = regions.reduce((s, r) => s + (r.annualSales || 0), 0);
  const roi = totalSales > 0 ? (totalUsed / totalSales).toFixed(1) : 0;
  const usedPct = totalBudget > 0 ? ((totalUsed / totalBudget) * 100).toFixed(1) : 0;

  document.getElementById('totalBudget').textContent = formatMoney(totalBudget);
  document.getElementById('totalUsed').textContent = formatMoney(totalUsed);
  document.getElementById('usedPct').textContent = `消耗率 ${usedPct}%`;
  document.getElementById('totalSales').textContent = totalSales.toLocaleString();
  document.getElementById('roiValue').textContent = `${roi} 元/万元`;
  document.getElementById('dataTime').textContent = `数据截止：${data.lastUpdate || '--'}`;

  // Efficiency table
  const tbody = document.getElementById('efficiencyTableBody');
  tbody.innerHTML = regions.map(r => `
    <tr>
      <td><strong>${r.name}</strong></td>
      <td>${formatMoneyFull(r.annualBudget)}</td>
      <td>${formatMoneyFull(r.totalUsed)}</td>
      <td><span style="font-weight:700;color:${r.usedPct >= 15 ? 'var(--danger)' : r.usedPct >= 8 ? 'var(--warning)' : 'var(--success)'}">${r.usedPct}%</span></td>
      <td>${(r.annualSales || 0).toLocaleString()}</td>
      <td>${r.salesPct || 0}%</td>
      <td>${r.h1Target || 0}</td>
      <td>${r.roi || 0}</td>
      <td><span class="${getRatingBadgeClass(r.rating)}">${r.rating}</span></td>
    </tr>
  `).join('');

  renderDashboardCharts();
}

function renderDashboardCharts() {
  const regions = appData.regions;
  
  // Compare chart: 消耗率 vs 业绩占比
  destroyChart('compareChart');
  const ctx1 = document.getElementById('compareChart').getContext('2d');
  charts['compareChart'] = new Chart(ctx1, {
    type: 'bar',
    data: {
      labels: regions.map(r => r.name),
      datasets: [
        {
          label: '积分消耗率(%)',
          data: regions.map(r => r.usedPct),
          backgroundColor: 'rgba(26, 86, 219, 0.7)',
          borderColor: 'rgba(26, 86, 219, 1)',
          borderWidth: 1,
          borderRadius: 4
        },
        {
          label: '业绩占比(%)',
          data: regions.map(r => r.salesPct || 0),
          backgroundColor: 'rgba(22, 163, 74, 0.7)',
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
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}%`
          }
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
            callback: v => v + '%',
            font: { size: 11 }
          },
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  });

  // Pie chart for categories
  const cats = {
    '联合营销/大促': appData.regions.reduce((s, r) => s + (r.coMarketing || 0), 0),
    '对赌': appData.regions.reduce((s, r) => s + (r.gamble || 0), 0),
    '店中店装修': appData.regions.reduce((s, r) => s + (r.storeRenovation || 0), 0),
    '进场费': appData.regions.reduce((s, r) => s + (r.sceneFee || 0), 0)
  };
  
  const catLabels = Object.keys(cats).filter(k => cats[k] > 0);
  const catValues = catLabels.map(k => cats[k]);

  destroyChart('categoryPieChart');
  if (catValues.length > 0) {
    const ctx2 = document.getElementById('categoryPieChart').getContext('2d');
    charts['categoryPieChart'] = new Chart(ctx2, {
      type: 'doughnut',
      data: {
        labels: catLabels,
        datasets: [{
          data: catValues,
          backgroundColor: [
            'rgba(26, 86, 219, 0.8)',
            'rgba(249, 115, 22, 0.8)',
            'rgba(22, 163, 74, 0.8)',
            'rgba(220, 38, 38, 0.8)'
          ],
          borderWidth: 2,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '60%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: { font: { size: 11 }, padding: 12 }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
                return `${ctx.label}: ${formatMoneyFull(ctx.parsed)} (${pct}%)`;
              }
            }
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
      <td>${(r.annualSales || 0).toLocaleString()}</td>
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
        label: '使用金额（元）',
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
  
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const row = rows[i];
    const colMap = detectColumns(row);
    if (colMap.name !== -1) {
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

function detectColumns(headerRow) {
  const map = {
    name: -1,
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
