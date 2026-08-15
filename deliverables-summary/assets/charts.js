(function() {
  var style = getComputedStyle(document.documentElement);
  var accent = style.getPropertyValue('--accent').trim();
  var accent2 = style.getPropertyValue('--accent2').trim();
  var ink = style.getPropertyValue('--ink').trim();
  var muted = style.getPropertyValue('--muted').trim();
  var rule = style.getPropertyValue('--rule').trim();
  var bg2 = style.getPropertyValue('--bg2').trim();

  var baseAxis = {
    axisLine: { lineStyle: { color: rule } },
    axisTick: { show: false },
    axisLabel: { color: muted, fontSize: 11 }
  };

  // --- 图 1: 八大类完全支持率 ---
  var cats = ['A 基础防火墙\n(32点)', 'B 应用识别\n(19点)', 'C 威胁防护\n(22点)', 'D 内容安全\n(10点)', 'E VPN\n(14点)', 'F 管理与运维\n(21点)', 'G 硬件与平台\n(13点)', 'H 合规与生态\n(8点)'];
  var tsRate = [90.6, 68.4, 90.9, 70.0, 42.9, 66.7, 38.5, 0.0];
  var paRate = [78.1, 84.2, 59.1, 60.0, 71.4, 95.2, 69.2, 25.0];
  var c1 = echarts.init(document.getElementById('chart-cat-rate'), null, { renderer: 'svg' });
  c1.setOption({
    animation: false,
    tooltip: { trigger: 'axis', appendToBody: true, valueFormatter: function(v) { return v + '%'; } },
    legend: { top: 0, textStyle: { color: ink, fontSize: 12 } },
    grid: { left: 8, right: 8, top: 40, bottom: 4, containLabel: true },
    xAxis: Object.assign({ type: 'category', data: cats, axisLabel: { color: muted, fontSize: 10.5, interval: 0, lineHeight: 14 } }, baseAxis),
    yAxis: Object.assign({ type: 'value', max: 100, axisLabel: { color: muted, fontSize: 11, formatter: '{value}%' }, splitLine: { lineStyle: { color: rule } } }, {}),
    series: [
      { name: '天融信 完全支持率', type: 'bar', data: tsRate, barMaxWidth: 26, itemStyle: { color: accent }, label: { show: true, position: 'top', fontSize: 10, color: accent, formatter: '{c}%' } },
      { name: 'Palo Alto 完全支持率', type: 'bar', data: paRate, barMaxWidth: 26, itemStyle: { color: accent2 }, label: { show: true, position: 'top', fontSize: 10, color: accent2, formatter: '{c}%' } }
    ]
  });
  window.addEventListener('resize', function() { c1.resize(); });

  // --- 图 2: 支持度分布 ---
  var lv = ['完全支持', '部分支持', '需订阅授权', '待验证', '不支持'];
  var tsSup = [94, 19, 0, 15, 11];
  var paSup = [101, 12, 10, 6, 10];
  var c2 = echarts.init(document.getElementById('chart-support'), null, { renderer: 'svg' });
  c2.setOption({
    animation: false,
    tooltip: { trigger: 'axis', appendToBody: true },
    legend: { top: 0, textStyle: { color: ink, fontSize: 12 } },
    grid: { left: 8, right: 8, top: 40, bottom: 4, containLabel: true },
    xAxis: Object.assign({ type: 'category', data: lv }, baseAxis),
    yAxis: Object.assign({ type: 'value', splitLine: { lineStyle: { color: rule } } }, {}),
    series: [
      { name: '天融信', type: 'bar', data: tsSup, barMaxWidth: 40, itemStyle: { color: accent }, label: { show: true, position: 'top', fontSize: 10.5, color: accent } },
      { name: 'Palo Alto', type: 'bar', data: paSup, barMaxWidth: 40, itemStyle: { color: accent2 }, label: { show: true, position: 'top', fontSize: 10.5, color: accent2 } }
    ]
  });
  window.addEventListener('resize', function() { c2.resize(); });

  // --- 图 3: 差异定性分布（横向条形） ---
  var diffNames = ['部分差异', '天融信优势(待验证)', '天融信优势', '待验证', '实现路径不同', 'PA 优势', '仅命名不同'];
  var diffVals = [2, 2, 16, 17, 25, 31, 46];
  var diffColors = { '部分差异': muted, '天融信优势(待验证)': accent, '天融信优势': accent, '待验证': muted, '实现路径不同': muted, 'PA 优势': accent2, '仅命名不同': muted };
  var c3 = echarts.init(document.getElementById('chart-diff'), null, { renderer: 'svg' });
  c3.setOption({
    animation: false,
    tooltip: { trigger: 'axis', appendToBody: true },
    grid: { left: 8, right: 40, top: 10, bottom: 4, containLabel: true },
    xAxis: Object.assign({ type: 'value', splitLine: { lineStyle: { color: rule } } }, {}),
    yAxis: Object.assign({ type: 'category', data: diffNames, axisLabel: { color: ink, fontSize: 12 } }, baseAxis),
    series: [{
      type: 'bar',
      data: diffNames.map(function(n, i) {
        return { value: diffVals[i], itemStyle: { color: diffColors[n] } };
      }),
      barMaxWidth: 22,
      label: { show: true, position: 'right', fontSize: 11, color: ink, formatter: '{c} 点' }
    }]
  });
  window.addEventListener('resize', function() { c3.resize(); });
})();
