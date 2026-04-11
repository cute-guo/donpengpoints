// 绩效管理系统数据 - 基于【含线组】2026年整装市场中心绩效数据.xlsx
// 数据结构：9个绩效项目 × 多个运营中心 × 多个线组

const PERFORMANCE_PROJECTS = [
  {
    id: "P001",
    name: "355目标",
    code: "355线",
    description: "355计划目标达成",
    icon: "target",
    unit: "万元"
  },
  {
    id: "P002",
    name: "N50目标",
    code: "N50线",
    description: "N50目标达成",
    icon: "award",
    unit: "万元"
  },
  {
    id: "P003",
    name: "差距城市目标",
    code: "差距线",
    description: "差距城市目标达成",
    icon: "map",
    unit: "万元"
  },
  {
    id: "P004",
    name: "高质抢量活动",
    code: "活动线",
    description: "高质抢量活动执行",
    icon: "zap",
    unit: "场次"
  },
  {
    id: "P005",
    name: "头部开拓",
    code: "头部线",
    description: "头部装企开拓",
    icon: "building",
    unit: "家"
  },
  {
    id: "P006",
    name: "头部网点开拓",
    code: "网点线",
    description: "头部网点开拓",
    icon: "home",
    unit: "个"
  },
  {
    id: "P007",
    name: "星链设计师活动",
    code: "设计线",
    description: "星链设计师活动",
    icon: "users",
    unit: "场次"
  },
  {
    id: "P008",
    name: "新品任务",
    code: "新品线",
    description: "新品推广任务",
    icon: "package",
    unit: "万元"
  },
  {
    id: "P009",
    name: "总对总活动",
    code: "总对总线",
    description: "总对总活动执行",
    icon: "globe",
    unit: "场次"
  }
];

// 运营中心（区域）信息
const REGIONS = [
  { id: "R01", name: "华东运营中心", director: "张总", teamCount: 6 },
  { id: "R02", name: "华南运营中心", director: "待定", teamCount: 1 },
  { id: "R03", name: "华北运营中心", director: "王总", teamCount: 1 },
  { id: "R04", name: "西南运营中心", director: "吕总", teamCount: 5 },
  { id: "R05", name: "湘鄂运营中心", director: "邓总", teamCount: 1 },
  { id: "R06", name: "粤东运营中心", director: "黄总", teamCount: 5 },
  { id: "R07", name: "粤西运营中心", director: "待定", teamCount: 4 },
  { id: "R08", name: "西北运营中心", director: "待定", teamCount: 4 },
  { id: "R09", name: "东北运营中心", director: "待定", teamCount: 1 },
  { id: "R10", name: "赣皖运营中心", director: "刘总", teamCount: 3 },
  { id: "R11", name: "鲁豫晋运营中心", director: "周总", teamCount: 4 }
];

// 区域-线组对应关系（基于Excel真实数据）
const TEAMS_BY_REGION = {
  "R01": [
    { id: "T01", name: "浙江一组", leader: "待定", city: "浙江" },
    { id: "T02", name: "浙江二组", leader: "待定", city: "浙江" },
    { id: "T03", name: "江苏组", leader: "待定", city: "江苏" },
    { id: "T04", name: "江浙组", leader: "待定", city: "江浙" },
    { id: "T05", name: "上海直营", leader: "待定", city: "上海" },
    { id: "T06", name: "上海组", leader: "待定", city: "上海" }
  ],
  "R02": [
    { id: "T07", name: "华南整装部", leader: "待定", city: "华南" }
  ],
  "R03": [
    { id: "T08", name: "华北整装部", leader: "待定", city: "华北" }
  ],
  "R04": [
    { id: "T09", name: "广西组", leader: "李辉", city: "广西" },
    { id: "T10", name: "四川组", leader: "吴华富", city: "四川" },
    { id: "T11", name: "云贵组", leader: "刘昌盛", city: "云贵" },
    { id: "T12", name: "重庆组", leader: "吕中清", city: "重庆" },
    { id: "T13", name: "西南整装部", leader: "吕伟东", city: "西南" }
  ],
  "R05": [
    { id: "T14", name: "湘鄂整装部", leader: "邓振辉/邓超", city: "湘鄂" }
  ],
  "R06": [
    { id: "T15", name: "福建组", leader: "待定", city: "福建" },
    { id: "T16", name: "粤东组", leader: "黄宗明", city: "粤东" },
    { id: "T17", name: "深圳1组", leader: "待定", city: "深圳" },
    { id: "T18", name: "深圳2组", leader: "待定", city: "深圳" },
    { id: "T19", name: "粤东整装部", leader: "待定", city: "粤东" }
  ],
  "R07": [
    { id: "T20", name: "广州组", leader: "方灿烽", city: "广州" },
    { id: "T21", name: "佛山组", leader: "待定", city: "佛山" },
    { id: "T22", name: "粤海组", leader: "待定", city: "粤海" },
    { id: "T23", name: "粤西整装部", leader: "待定", city: "粤西" }
  ],
  "R08": [
    { id: "T24", name: "西北组", leader: "待定", city: "西北" },
    { id: "T25", name: "西安直营", leader: "待定", city: "西安" },
    { id: "T26", name: "西安服务商", leader: "待定", city: "西安" },
    { id: "T27", name: "西北整装部", leader: "待定", city: "西北" }
  ],
  "R09": [
    { id: "T28", name: "东北整装部", leader: "待定", city: "东北" }
  ],
  "R10": [
    { id: "T29", name: "江西组", leader: "刘跃明", city: "江西" },
    { id: "T30", name: "安徽组", leader: "待定", city: "安徽" },
    { id: "T31", name: "赣皖整装部", leader: "待定", city: "赣皖" }
  ],
  "R11": [
    { id: "T32", name: "山东组", leader: "待定", city: "山东" },
    { id: "T33", name: "河南组", leader: "谭武宏", city: "河南" },
    { id: "T34", name: "山西组", leader: "周永文", city: "山西" },
    { id: "T35", name: "鲁豫晋整装部", leader: "待定", city: "鲁豫晋" }
  ]
};

// 生成模拟数据（基于Excel真实结构）
function generateMockData(projectId, month = 3) {
  const project = PERFORMANCE_PROJECTS.find(p => p.id === projectId);
  const regionData = [];

  // 根据不同项目设置不同的目标基数
  const baseTargets = {
    "P001": { region: 8000, team: 1600 },
    "P002": { region: 3000, team: 600 },
    "P003": { region: 5000, team: 1000 },
    "P004": { region: 50, team: 10 },
    "P005": { region: 30, team: 6 },
    "P006": { region: 100, team: 20 },
    "P007": { region: 60, team: 12 },
    "P008": { region: 2000, team: 400 },
    "P009": { region: 50, team: 10 }
  };

  const base = baseTargets[projectId] || { region: 5000, team: 1000 };

  REGIONS.forEach(region => {
    const regionVariance = 0.7 + Math.random() * 0.6;
    const regionTarget = Math.round(base.region * regionVariance);
    const regionCompleted = Math.round(regionTarget * (0.2 + Math.random() * 0.75));
    const regionRate = regionTarget > 0 ? (regionCompleted / regionTarget * 100).toFixed(2) : 0;

    let regionStatus = 'normal';
    if (regionRate < 50) regionStatus = 'danger';
    else if (regionRate < 80) regionStatus = 'warning';

    const teams = TEAMS_BY_REGION[region.id] || [];
    const teamData = [];
    let teamTotalTarget = 0;

    teams.forEach(team => {
      const teamVariance = 0.6 + Math.random() * 0.8;
      const teamTarget = Math.round(base.team * teamVariance);
      const teamCompleted = Math.round(teamTarget * (0.15 + Math.random() * 0.8));
      const teamRate = teamTarget > 0 ? (teamCompleted / teamTarget * 100).toFixed(2) : 0;

      let teamStatus = 'normal';
      if (teamRate < 50) teamStatus = 'danger';
      else if (teamRate < 80) teamStatus = 'warning';

      const monthlyData = [];
      for (let m = 1; m <= 12; m++) {
        const mTarget = Math.round(teamTarget / 12 * (0.8 + Math.random() * 0.4));
        const mCompleted = m <= month ? Math.round(mTarget * (0.2 + Math.random() * 0.75)) : 0;
        monthlyData.push({
          month: m,
          target: mTarget,
          completed: mCompleted,
          rate: mTarget > 0 ? (mCompleted / mTarget * 100).toFixed(2) : 0
        });
      }

      const q1Data = monthlyData.slice(0, 3);
      const q2Data = monthlyData.slice(3, 6);
      const q3Data = monthlyData.slice(6, 9);
      const q4Data = monthlyData.slice(9, 12);

      teamData.push({
        id: team.id,
        name: team.name,
        leader: team.leader,
        city: team.city,
        targetValue: teamTarget,
        completedValue: teamCompleted,
        completionRate: parseFloat(teamRate),
        status: teamStatus,
        trend: Math.random() > 0.5 ? 'up' : (Math.random() > 0.5 ? 'down' : 'stable'),
        monthlyData: monthlyData,
        quarterlyData: {
          Q1: { target: q1Data.reduce((a, b) => a + b.target, 0), completed: q1Data.reduce((a, b) => a + b.completed, 0) },
          Q2: { target: q2Data.reduce((a, b) => a + b.target, 0), completed: q2Data.reduce((a, b) => a + b.completed, 0) },
          Q3: { target: q3Data.reduce((a, b) => a + b.target, 0), completed: q3Data.reduce((a, b) => a + b.completed, 0) },
          Q4: { target: q4Data.reduce((a, b) => a + b.target, 0), completed: q4Data.reduce((a, b) => a + b.completed, 0) }
        }
      });

      teamTotalTarget += teamTarget;
    });

    const monthlyRegionData = [];
    for (let m = 1; m <= 12; m++) {
      const mTarget = Math.round(regionTarget / 12 * (0.8 + Math.random() * 0.4));
      const mCompleted = m <= month ? Math.round(mTarget * (0.2 + Math.random() * 0.75)) : 0;
      monthlyRegionData.push({
        month: m,
        target: mTarget,
        completed: mCompleted,
        rate: mTarget > 0 ? (mCompleted / mTarget * 100).toFixed(2) : 0
      });
    }

    regionData.push({
      id: region.id,
      name: region.name,
      director: region.director,
      teams: teams,
      targetValue: regionTarget,
      completedValue: regionCompleted,
      completionRate: parseFloat(regionRate),
      status: regionStatus,
      trend: Math.random() > 0.5 ? 'up' : (Math.random() > 0.5 ? 'down' : 'stable'),
      teamTotalTarget: teamTotalTarget,
      directorExtraTarget: regionTarget - teamTotalTarget,
      teamData: teamData,
      monthlyData: monthlyRegionData
    });
  });

  return {
    project: project,
    month: month,
    updateTime: new Date().toISOString(),
    regions: regionData
  };
}

const AppData = {
  currentProject: null,
  currentMonth: 3,
  currentRegionFilter: 'all',
  currentStatusFilter: 'all',
  allData: {},

  init() {
    PERFORMANCE_PROJECTS.forEach(project => {
      this.allData[project.id] = generateMockData(project.id, this.currentMonth);
    });
    this.currentProject = PERFORMANCE_PROJECTS[0];
  },

  getProjectData(projectId) {
    if (!this.allData[projectId]) {
      this.allData[projectId] = generateMockData(projectId, this.currentMonth);
    }
    return this.allData[projectId];
  },

  updateData(projectId, regionId, teamId, newValue, field = 'completedValue') {
    const data = this.allData[projectId];
    if (teamId) {
      const team = data.regions.find(r => r.id === regionId).teamData.find(t => t.id === teamId);
      if (field === 'completedValue') {
        team.completedValue = newValue;
        team.completionRate = parseFloat((newValue / team.targetValue * 100).toFixed(2));
      } else if (field === 'targetValue') {
        team.targetValue = newValue;
        team.completionRate = parseFloat((team.completedValue / newValue * 100).toFixed(2));
      }
      team.status = this.calculateStatus(team.completionRate);
    } else {
      const region = data.regions.find(r => r.id === regionId);
      if (field === 'completedValue') {
        region.completedValue = newValue;
        region.completionRate = parseFloat((newValue / region.targetValue * 100).toFixed(2));
      } else if (field === 'targetValue') {
        region.targetValue = newValue;
        region.completionRate = parseFloat((region.completedValue / newValue * 100).toFixed(2));
      }
      region.status = this.calculateStatus(region.completionRate);
    }
    return data;
  },

  calculateStatus(rate) {
    if (rate < 50) return 'danger';
    if (rate < 80) return 'warning';
    return 'normal';
  },

  calculateOverallStats(projectId) {
    const data = this.allData[projectId];
    const totals = {
      totalTarget: 0,
      totalCompleted: 0,
      regionCount: data.regions.length,
      warningCount: 0,
      dangerCount: 0,
      normalCount: 0,
      teamCount: 0
    };

    data.regions.forEach(region => {
      totals.totalTarget += region.targetValue;
      totals.totalCompleted += region.completedValue;
      if (region.status === 'danger') totals.dangerCount++;
      else if (region.status === 'warning') totals.warningCount++;
      else totals.normalCount++;
      totals.teamCount += region.teamData.length;
    });

    totals.overallRate = totals.totalTarget > 0
      ? (totals.totalCompleted / totals.totalTarget * 100).toFixed(2)
      : 0;

    return totals;
  },

  setMonth(month) {
    this.currentMonth = month;
    Object.keys(this.allData).forEach(projectId => {
      this.allData[projectId] = generateMockData(projectId, month);
    });
  },

  getFilteredRegions(projectId) {
    let data = this.allData[projectId];
    if (!data) return [];

    let regions = data.regions;

    if (this.currentRegionFilter !== 'all') {
      regions = regions.filter(r => r.id === this.currentRegionFilter);
    }

    if (this.currentStatusFilter !== 'all') {
      regions = regions.filter(r => r.status === this.currentStatusFilter);
    }

    return regions;
  }
};

const MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

AppData.init();
