// ========================================
// 绩效管理系统 - 主应用逻辑
// ========================================

let regionChart = null;
let statusChart = null;

document.addEventListener('DOMContentLoaded', function() {
  lucide.createIcons();
  initApp();
});

function initApp() {
  renderNavList();
  renderRegionFilter();
  renderHomePage();  // 默认显示首页
  updateTime();
  setInterval(updateTime, 1000);

  document.getElementById('menuToggle').addEventListener('click', toggleSidebar);
  document.getElementById('sidebarOverlay').addEventListener('click', toggleSidebar);

  document.getElementById('updateTarget').addEventListener('input', calculateRate);
  document.getElementById('updateCompleted').addEventListener('input', calculateRate);
}

// 当前显示的周期类型
let currentPeriod = 'quarter';

// 首页渲染
function renderHomePage() {
  const container = document.getElementById('projectsGrid');
  
  // 更新日期显示
  updateHeaderDate();
  
  // 渲染季度和年度数据
  renderQuarterData();
  renderYearData();
  
  // 渲染项目卡片
  const themeColors = ['blue', 'green', 'orange', 'red', 'purple', 'cyan', 'blue', 'green', 'orange'];
  
  container.innerHTML = PERFORMANCE_PROJECTS.map((project, index) => {
    const stats = AppData.calculateOverallStats(project.id);
    const rate = parseFloat(stats.overallRate);
    const status = stats.dangerCount > 0 ? 'danger' : (stats.warningCount > 0 ? 'warning' : 'normal');
    const theme = themeColors[index % themeColors.length];
    
    return `
      <div class="project-card theme-${theme}" onclick="showProjectDetail('${project.id}')">
        <div class="project-card-header">
          <div class="project-card-icon">
            <i data-lucide="${project.icon}"></i>
          </div>
          <span class="project-card-badge ${status}">${status === 'danger' ? '危险' : status === 'warning' ? '预警' : '正常'}</span>
        </div>
        <div class="project-card-name">${project.name}</div>
        <div class="project-card-desc">${project.code}</div>
        <div class="project-card-progress">
          <div class="project-card-progress-bar">
            <div class="project-card-progress-fill ${status}" style="width: ${Math.min(rate, 100)}%;"></div>
          </div>
        </div>
        <div class="project-card-stats">
          <div>
            <span class="label">完成进度：</span>
            <span class="rate">${stats.overallRate}%</span>
          </div>
          <div>
            <span class="label">${stats.totalCompleted.toLocaleString()} / ${stats.totalTarget.toLocaleString()}</span>
          </div>
        </div>
      </div>
    `;
  }).join('');
  
  lucide.createIcons();
}

// 切换季度/年度显示
function switchPeriod(period) {
  currentPeriod = period;
  
  // 更新标签样式
  document.querySelectorAll('.period-tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.period === period);
  });
  
  // 更新内容显示
  document.getElementById('quarterContent').style.display = period === 'quarter' ? 'block' : 'none';
  document.getElementById('yearContent').style.display = period === 'year' ? 'block' : 'none';
}

// 渲染季度数据
function renderQuarterData() {
  const currentMonth = new Date().getMonth() + 1;
  const currentQuarter = Math.ceil(currentMonth / 3);
  
  let totalTarget = 0;
  let totalCompleted = 0;
  const regionData = {};
  
  // 初始化各区域数据
  REGIONS.forEach(region => {
    regionData[region.id] = { name: region.name, target: 0, completed: 0 };
  });
  
  // 计算季度数据（当前季度月份的数据）
  PERFORMANCE_PROJECTS.forEach(project => {
    const projectData = AppData.getProjectData(project.id);
    if (projectData && projectData.regions) {
      projectData.regions.forEach(region => {
        if (regionData[region.id]) {
          // 季度目标 = 月度目标 × 3（简化计算）
          regionData[region.id].target += region.targetValue;
          regionData[region.id].completed += region.completedValue;
        }
      });
    }
  });
  
  // 计算总计
  Object.values(regionData).forEach(data => {
    totalTarget += data.target;
    totalCompleted += data.completed;
  });
  
  const rate = totalTarget > 0 ? (totalCompleted / totalTarget * 100).toFixed(1) : 0;
  
  // 更新概览数据
  document.getElementById('quarterTotal').textContent = totalTarget.toLocaleString();
  document.getElementById('quarterCompleted').textContent = totalCompleted.toLocaleString();
  document.getElementById('quarterRate').textContent = rate + '%';
  document.getElementById('quarterRemaining').textContent = (totalTarget - totalCompleted).toLocaleString();
  
  // 渲染区域网格
  const grid = document.getElementById('quarterRegionGrid');
  const sortedRegions = Object.entries(regionData)
    .map(([id, data]) => ({ id, ...data }))
    .sort((a, b) => b.completed - a.completed);
  
  grid.innerHTML = sortedRegions.map(region => {
    const regionRate = region.target > 0 ? (region.completed / region.target * 100).toFixed(0) : 0;
    return `
      <div class="region-item">
        <span class="region-name">${region.name.replace('运营中心', '')}</span>
        <span class="region-value highlight">${region.completed.toLocaleString()}</span>
      </div>
    `;
  }).join('');
}

// 渲染年度数据
function renderYearData() {
  let totalTarget = 0;
  let totalCompleted = 0;
  const regionData = {};
  
  // 初始化各区域数据
  REGIONS.forEach(region => {
    regionData[region.id] = { name: region.name, target: 0, completed: 0 };
  });
  
  // 计算年度累计数据
  PERFORMANCE_PROJECTS.forEach(project => {
    const projectData = AppData.getProjectData(project.id);
    if (projectData && projectData.regions) {
      projectData.regions.forEach(region => {
        if (regionData[region.id]) {
          regionData[region.id].target += region.targetValue;
          regionData[region.id].completed += region.completedValue;
        }
      });
    }
  });
  
  // 计算总计
  Object.values(regionData).forEach(data => {
    totalTarget += data.target;
    totalCompleted += data.completed;
  });
  
  const rate = totalTarget > 0 ? (totalCompleted / totalTarget * 100).toFixed(1) : 0;
  
  // 更新概览数据
  document.getElementById('yearTotal').textContent = totalTarget.toLocaleString();
  document.getElementById('yearCompleted').textContent = totalCompleted.toLocaleString();
  document.getElementById('yearRate').textContent = rate + '%';
  document.getElementById('yearRemaining').textContent = (totalTarget - totalCompleted).toLocaleString();
  
  // 渲染区域网格
  const grid = document.getElementById('yearRegionGrid');
  const sortedRegions = Object.entries(regionData)
    .map(([id, data]) => ({ id, ...data }))
    .sort((a, b) => b.completed - a.completed);
  
  grid.innerHTML = sortedRegions.map(region => {
    const regionRate = region.target > 0 ? (region.completed / region.target * 100).toFixed(0) : 0;
    return `
      <div class="region-item">
        <span class="region-name">${region.name.replace('运营中心', '')}</span>
        <span class="region-value highlight">${region.completed.toLocaleString()}</span>
      </div>
    `;
  }).join('');
}

// 更新头部日期显示
function updateHeaderDate() {
  const now = new Date();
  const currentQuarter = Math.ceil((now.getMonth() + 1) / 3);
  const dateStr = now.toLocaleDateString('zh-CN', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    weekday: 'long'
  }) + ` · 第${currentQuarter}季度`;
  document.getElementById('currentDate').textContent = dateStr;
}

// 计算到年底剩余天数
function calculateRemainingDays() {
  const now = new Date();
  const yearEnd = new Date(now.getFullYear(), 11, 31);
  const diffTime = yearEnd - now;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays > 0 ? diffDays : 0;
}

// 显示项目详情页
function showProjectDetail(projectId) {
  AppData.currentProject = PERFORMANCE_PROJECTS.find(p => p.id === projectId);
  
  document.getElementById('homePage').style.display = 'none';
  document.getElementById('detailPage').style.display = 'block';
  document.getElementById('currentProjectName').textContent = AppData.currentProject.name;
  
  renderNavList();
  renderOverviewCards();
  renderDataTables();
  initCharts();
  
  // 关闭移动端侧边栏
  if (window.innerWidth <= 992) {
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('sidebarOverlay').classList.remove('active');
  }
}

// 返回首页
function showHomePage() {
  document.getElementById('detailPage').style.display = 'none';
  document.getElementById('homePage').style.display = 'block';
  document.getElementById('currentProjectName').textContent = '总览';
  renderHomePage();
}

function updateTime() {
  const now = new Date();
  const timeStr = now.toLocaleString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
  document.getElementById('currentTime').textContent = timeStr;
  
  // 首页时实时更新剩余天数
  const homePage = document.getElementById('homePage');
  if (homePage && homePage.style.display !== 'none') {
    const remainingDays = calculateRemainingDays();
    document.getElementById('remainingTime').textContent = remainingDays;
  }
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('active');
}

function renderNavList() {
  const navList = document.getElementById('navList');
  navList.innerHTML = PERFORMANCE_PROJECTS.map(project => {
    const stats = AppData.calculateOverallStats(project.id);
    const statusClass = stats.dangerCount > 0 ? 'danger' : (stats.warningCount > 0 ? 'warning' : 'normal');
    const isActive = AppData.currentProject && AppData.currentProject.id === project.id;

    return `
      <div class="nav-item ${isActive ? 'active' : ''}" onclick="switchProject('${project.id}')">
        <i data-lucide="${project.icon}" class="nav-item-icon"></i>
        <div class="nav-item-content">
          <div class="nav-item-name">${project.name}</div>
          <div class="nav-item-code">${project.code}</div>
        </div>
        <div class="nav-item-status ${statusClass}"></div>
      </div>
    `;
  }).join('');
  lucide.createIcons();
}

function switchProject(projectId) {
  showProjectDetail(projectId);
}

function renderRegionFilter() {
  const regionFilter = document.getElementById('regionFilter');
  const updateRegion = document.getElementById('updateRegion');

  const options = REGIONS.map(r => `<option value="${r.id}">${r.name}</option>`).join('');
  regionFilter.innerHTML = '<option value="all">全部区域</option>' + options;
  updateRegion.innerHTML = '<option value="">请选择区域</option>' + options;

  updateRegion.addEventListener('change', function() {
    const regionId = this.value;
    const teamSelect = document.getElementById('updateTeam');
    if (regionId && TEAMS_BY_REGION[regionId]) {
      const teams = TEAMS_BY_REGION[regionId];
      teamSelect.innerHTML = '<option value="">请选择线组（可选）</option>' +
        teams.map(t => `<option value="${t.id}">${t.name} - ${t.leader}</option>`).join('');
    } else {
      teamSelect.innerHTML = '<option value="">请先选择区域</option>';
    }
  });
}

function renderOverviewCards() {
  const container = document.getElementById('overviewCards');
  const stats = AppData.calculateOverallStats(AppData.currentProject.id);
  const project = AppData.currentProject;

  const completionRate = parseFloat(stats.overallRate);
  const circumference = 2 * Math.PI * 32;
  const offset = circumference - (completionRate / 100) * circumference;

  container.innerHTML = `
    <div class="overview-card highlight">
      <div class="card-header">
        <span class="card-label">总体完成率</span>
        <div class="card-icon">
          <i data-lucide="target"></i>
        </div>
      </div>
      <div class="progress-ring">
        <svg class="progress-ring-circle" width="80" height="80">
          <circle class="progress-ring-bg" cx="40" cy="40" r="32"></circle>
          <circle class="progress-ring-progress" cx="40" cy="40" r="32"
            style="stroke-dasharray: ${circumference}; stroke-dashoffset: ${offset}; stroke: white;"></circle>
        </svg>
        <span class="progress-ring-text" style="color: white;">${stats.overallRate}%</span>
      </div>
      <div class="card-sublabel" style="color: rgba(255,255,255,0.8);">
        ${stats.totalCompleted.toLocaleString()} / ${stats.totalTarget.toLocaleString()} ${project.unit}
      </div>
    </div>

    <div class="overview-card">
      <div class="card-header">
        <span class="card-label">年度目标</span>
        <div class="card-icon">
          <i data-lucide="flag"></i>
        </div>
      </div>
      <div class="card-value">${stats.totalTarget.toLocaleString()}</div>
      <div class="card-sublabel">${project.unit} | ${stats.regionCount}个区域</div>
    </div>

    <div class="overview-card">
      <div class="card-header">
        <span class="card-label">已完成</span>
        <div class="card-icon">
          <i data-lucide="check-circle"></i>
        </div>
      </div>
      <div class="card-value text-success">${stats.totalCompleted.toLocaleString()}</div>
      <div class="card-sublabel">${stats.teamCount}个线组参与</div>
    </div>

    <div class="overview-card">
      <div class="card-header">
        <span class="card-label">预警统计</span>
        <div class="card-icon">
          <i data-lucide="alert-triangle"></i>
        </div>
      </div>
      <div class="card-value">
        <span class="text-success">${stats.normalCount}</span> /
        <span class="text-warning">${stats.warningCount}</span> /
        <span class="text-danger">${stats.dangerCount}</span>
      </div>
      <div class="card-sublabel">正常 / 预警 / 危险</div>
    </div>
  `;
  lucide.createIcons();
}

function renderDataTables() {
  const container = document.getElementById('dataTables');
  const data = AppData.getProjectData(AppData.currentProject.id);
  const regions = AppData.getFilteredRegions(AppData.currentProject.id);
  const project = AppData.currentProject;

  if (regions.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">
          <i data-lucide="inbox"></i>
        </div>
        <div class="empty-state-title">暂无数据</div>
        <div class="empty-state-text">当前筛选条件下没有数据</div>
        <button class="btn btn-secondary" onclick="clearFilters()">清除筛选</button>
      </div>
    `;
    lucide.createIcons();
    return;
  }

  let html = `
    <div class="data-section">
      <div class="section-header">
        <div class="section-title">
          <i data-lucide="building-2"></i>
          区域总负责人业绩拆解
          <span class="section-subtitle">| ${regions.length}个区域</span>
        </div>
      </div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width: 40px;"></th>
            <th>区域</th>
            <th>总负责人</th>
            <th>年度目标</th>
            <th>已完成</th>
            <th>完成率</th>
            <th>状态</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
  `;

  regions.forEach((region, index) => {
    const statusIcon = region.status === 'danger' ? '<span class="warning-flag" onclick="showWarning(\'region\', \'' + region.id + '\')">!</span>' :
                       (region.status === 'warning' ? '<span class="warning-flag" onclick="showWarning(\'region\', \'' + region.id + '\')">!</span>' : '');

    html += `
      <tr data-region="${region.id}">
        <td>
          <span class="expand-toggle" onclick="toggleExpand('${region.id}')" id="expand-${region.id}">
            <i data-lucide="chevron-right"></i>
          </span>
        </td>
        <td>
          <div class="cell-region">
            <div class="cell-region-info">
              <div class="cell-region-name">${region.name}</div>
            </div>
          </div>
        </td>
        <td><span class="cell-value">${region.director}</span></td>
        <td><span class="cell-value large">${region.targetValue.toLocaleString()}</span></td>
        <td><span class="cell-value large">${region.completedValue.toLocaleString()}</span></td>
        <td>
          <div class="flex items-center gap-md">
            <div class="progress-bar">
              <div class="progress-bar-fill ${region.status}" style="width: ${Math.min(region.completionRate, 100)}%;"></div>
            </div>
            <span class="cell-value font-mono">${region.completionRate}%</span>
          </div>
        </td>
        <td>
          <span class="status-badge ${region.status}">${getStatusText(region.status)}</span>
          ${statusIcon}
        </td>
        <td>
          <div class="cell-actions">
            <button class="btn btn-secondary btn-icon" onclick="openUpdateModal('${region.id}')" title="更新数据">
              <i data-lucide="edit-2"></i>
            </button>
          </div>
        </td>
      </tr>
      <tr class="expand-row" id="expand-row-${region.id}" style="display: none;">
        <td colspan="8">
          <div class="expand-content">
            <table class="sub-table">
              <thead>
                <tr>
                  <th>线组</th>
                  <th>负责人</th>
                  <th>城市</th>
                  <th>目标值</th>
                  <th>完成值</th>
                  <th>完成率</th>
                  <th>状态</th>
                  <th>趋势</th>
                  <th>操作</th>
                </tr>
              </thead>
              <tbody>
    `;

    region.teamData.forEach(team => {
      const trendIcon = team.trend === 'up' ? '<i data-lucide="trending-up" style="color: var(--success-color);"></i>' :
                       (team.trend === 'down' ? '<i data-lucide="trending-down" style="color: var(--danger-color);"></i>' :
                        '<i data-lucide="minus" style="color: var(--text-tertiary);"></i>');

      html += `
        <tr>
          <td><strong>${team.name}</strong></td>
          <td>${team.leader}</td>
          <td>${team.city}</td>
          <td>${team.targetValue.toLocaleString()}</td>
          <td>${team.completedValue.toLocaleString()}</td>
          <td>
            <div class="flex items-center gap-sm">
              <div class="progress-bar" style="width: 80px;">
                <div class="progress-bar-fill ${team.status}" style="width: ${Math.min(team.completionRate, 100)}%;"></div>
              </div>
              <span class="font-mono">${team.completionRate}%</span>
            </div>
          </td>
          <td><span class="status-badge ${team.status}">${getStatusText(team.status)}</span></td>
          <td>${trendIcon}</td>
          <td>
            <button class="btn btn-secondary btn-icon" onclick="openUpdateModal('${region.id}', '${team.id}')" title="更新数据">
              <i data-lucide="edit-2"></i>
            </button>
          </td>
        </tr>
      `;
    });

    const teamTotalTarget = region.teamData.reduce((a, b) => a + b.targetValue, 0);
    const teamTotalCompleted = region.teamData.reduce((a, b) => a + b.completedValue, 0);
    const teamTotalRate = teamTotalTarget > 0 ? (teamTotalCompleted / teamTotalTarget * 100).toFixed(2) : 0;

    html += `
              </tbody>
              <tfoot>
                <tr style="background: #FAFBFC; font-weight: 600;">
                  <td colspan="3">线组汇总</td>
                  <td>${teamTotalTarget.toLocaleString()}</td>
                  <td>${teamTotalCompleted.toLocaleString()}</td>
                  <td>${teamTotalRate}%</td>
                  <td colspan="3"></td>
                </tr>
                <tr style="background: #FFF9E6; font-weight: 600;">
                  <td colspan="3">总负责人额外承担</td>
                  <td colspan="2">${region.directorExtraTarget.toLocaleString()}</td>
                  <td colspan="3" style="color: var(--text-tertiary);">（约15%挂在总负责人身上）</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  container.innerHTML = html;
  lucide.createIcons();
}

function toggleExpand(regionId) {
  const row = document.getElementById(`expand-row-${regionId}`);
  const toggle = document.getElementById(`expand-${regionId}`);
  const isHidden = row.style.display === 'none';

  row.style.display = isHidden ? 'table-row' : 'none';
  toggle.classList.toggle('expanded', isHidden);
}

function getStatusText(status) {
  const texts = { normal: '正常', warning: '预警', danger: '危险' };
  return texts[status] || status;
}

function initCharts() {
  const regionCtx = document.getElementById('regionChart').getContext('2d');
  const statusCtx = document.getElementById('statusChart').getContext('2d');

  regionChart = new Chart(regionCtx, {
    type: 'bar',
    data: {
      labels: [],
      datasets: [{
        label: '完成率',
        data: [],
        backgroundColor: [],
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          ticks: { callback: value => value + '%' }
        }
      }
    }
  });

  statusChart = new Chart(statusCtx, {
    type: 'doughnut',
    data: {
      labels: ['正常', '预警', '危险'],
      datasets: [{
        data: [0, 0, 0],
        backgroundColor: ['#52C41A', '#FAAD14', '#FF4D4F']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });

  updateCharts();
}

function updateCharts() {
  const stats = AppData.calculateOverallStats(AppData.currentProject.id);
  const data = AppData.getProjectData(AppData.currentProject.id);

  const labels = data.regions.map(r => r.name.replace('运营中心', ''));
  const rates = data.regions.map(r => r.completionRate);
  const colors = data.regions.map(r => {
    if (r.completionRate < 50) return '#FF4D4F';
    if (r.completionRate < 80) return '#FAAD14';
    return '#52C41A';
  });

  regionChart.data.labels = labels;
  regionChart.data.datasets[0].data = rates;
  regionChart.data.datasets[0].backgroundColor = colors;
  regionChart.update();

  statusChart.data.datasets[0].data = [stats.normalCount, stats.warningCount, stats.dangerCount];
  statusChart.update();
}

function handleMonthChange(month) {
  AppData.setMonth(parseInt(month));
  renderOverviewCards();
  renderDataTables();
  updateCharts();
}

function handleRegionFilter(regionId) {
  AppData.currentRegionFilter = regionId;
  renderDataTables();
}

function handleStatusFilter(status) {
  AppData.currentStatusFilter = status;
  document.querySelectorAll('.status-tag').forEach(tag => {
    tag.classList.toggle('active', tag.dataset.status === status);
  });
  renderDataTables();
}

function clearFilters() {
  AppData.currentRegionFilter = 'all';
  AppData.currentStatusFilter = 'all';
  document.getElementById('regionFilter').value = 'all';
  document.querySelectorAll('.status-tag').forEach(tag => {
    tag.classList.toggle('active', tag.dataset.status === 'all');
  });
  renderDataTables();
}

function refreshData() {
  AppData.setMonth(AppData.currentMonth);
  renderOverviewCards();
  renderDataTables();
  updateCharts();
}

function openUpdateModal(regionId, teamId) {
  const modal = document.getElementById('updateModal');
  modal.classList.add('active');

  if (regionId) {
    document.getElementById('updateRegion').value = regionId;
    document.getElementById('updateRegion').dispatchEvent(new Event('change'));

    const region = AppData.getProjectData(AppData.currentProject.id).regions.find(r => r.id === regionId);

    if (teamId) {
      document.getElementById('updateTeam').value = teamId;
      const team = region.teamData.find(t => t.id === teamId);
      document.getElementById('updateTarget').value = team.targetValue;
      document.getElementById('updateCompleted').value = team.completedValue;
    } else {
      document.getElementById('updateTarget').value = region.targetValue;
      document.getElementById('updateCompleted').value = region.completedValue;
    }
    calculateRate();
  }
}

function closeUpdateModal() {
  document.getElementById('updateModal').classList.remove('active');
}

function calculateRate() {
  const target = parseFloat(document.getElementById('updateTarget').value) || 0;
  const completed = parseFloat(document.getElementById('updateCompleted').value) || 0;
  const rate = target > 0 ? ((completed / target) * 100).toFixed(2) : 0;
  document.getElementById('calculatedRate').textContent = rate + '%';
}

function saveUpdate() {
  const regionId = document.getElementById('updateRegion').value;
  const teamId = document.getElementById('updateTeam').value || null;
  const target = parseFloat(document.getElementById('updateTarget').value) || 0;
  const completed = parseFloat(document.getElementById('updateCompleted').value) || 0;

  if (!regionId) {
    alert('请选择区域');
    return;
  }

  if (target > 0) {
    AppData.updateData(AppData.currentProject.id, regionId, teamId, target, 'targetValue');
  }
  if (completed >= 0) {
    AppData.updateData(AppData.currentProject.id, regionId, teamId, completed, 'completedValue');
  }

  closeUpdateModal();
  renderNavList();
  renderOverviewCards();
  renderDataTables();
  updateCharts();
}

function showWarning(type, id) {
  const modal = document.getElementById('warningModal');
  const body = document.getElementById('warningDetailBody');

  let content = '';
  if (type === 'region') {
    const region = AppData.getProjectData(AppData.currentProject.id).regions.find(r => r.id === id);
    content = `
      <div class="warning-detail">
        <div class="warning-detail-header">
          <i data-lucide="alert-triangle"></i>
          ${region.status === 'danger' ? '危险预警' : '预警提醒'}
        </div>
        <div class="warning-detail-content">
          <p><strong>${region.name}</strong> 完成率低于${region.status === 'danger' ? '50%' : '80%'}，需要关注！</p>
          <p>当前完成率：<strong>${region.completionRate}%</strong></p>
          <p>目标值：${region.targetValue.toLocaleString()}</p>
          <p>完成值：${region.completedValue.toLocaleString()}</p>
          <p>缺口：${(region.targetValue - region.completedValue).toLocaleString()}</p>
        </div>
      </div>
      <p style="color: var(--text-tertiary);">请及时与区域负责人沟通，了解原因并制定改进措施。</p>
    `;
  }

  body.innerHTML = content;
  modal.classList.add('active');
  lucide.createIcons();
}

function closeWarningModal() {
  document.getElementById('warningModal').classList.remove('active');
}

function markWarningHandled() {
  alert('已标记为已处理');
  closeWarningModal();
}

document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') {
    closeUpdateModal();
    closeWarningModal();
  }
});

document.getElementById('updateModal').addEventListener('click', function(e) {
  if (e.target === this) closeUpdateModal();
});
document.getElementById('warningModal').addEventListener('click', function(e) {
  if (e.target === this) closeWarningModal();
});
