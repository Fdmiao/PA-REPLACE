(function () {
  var style = getComputedStyle(document.documentElement);
  var accent = style.getPropertyValue('--accent').trim();
  var accent2 = style.getPropertyValue('--accent2').trim();
  var ink = style.getPropertyValue('--ink').trim();
  var muted = style.getPropertyValue('--muted').trim();
  var rule = style.getPropertyValue('--rule').trim();
  var bg2 = style.getPropertyValue('--bg2').trim();

  if (window.mermaid) {
    mermaid.initialize({ startOnLoad: true, theme: 'neutral', securityLevel: 'loose' });
  }

  // ===== 图1：八大类完全支持率对比 =====
  var el1 = document.getElementById('chart-cat-rate');
  if (el1) {
    var c1 = echarts.init(el1, null, { renderer: 'svg' });
    var cats = ['A 基础网络', 'B 应用识别', 'C 威胁防护', 'D 内容安全', 'E VPN加密', 'F 管理运维', 'G 硬件平台', 'H 合规生态'];
    var tsRates = [90.6, 68.4, 90.9, 70.0, 42.9, 66.7, 38.5, 0.0];
    var paRates = [78.1, 84.2, 59.1, 60.0, 71.4, 95.2, 69.2, 25.0];
    c1.setOption({
      animation: false,
      tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, appendToBody: true, valueFormatter: function (v) { return v + '%'; } },
      legend: { data: ['天融信完全支持率', 'PA完全支持率'], top: 0, textStyle: { color: muted } },
      grid: { left: 50, right: 20, top: 40, bottom: 40 },
      xAxis: { type: 'category', data: cats, axisLabel: { color: muted, interval: 0, rotate: 20 }, axisLine: { lineStyle: { color: rule } } },
      yAxis: { type: 'value', max: 100, axisLabel: { color: muted, formatter: '{value}%' }, splitLine: { lineStyle: { color: rule } } },
      series: [
        { name: '天融信完全支持率', type: 'bar', data: tsRates, itemStyle: { color: accent }, barWidth: '30%', label: { show: true, position: 'top', color: ink, fontSize: 10, formatter: '{c}%' } },
        { name: 'PA完全支持率', type: 'bar', data: paRates, itemStyle: { color: accent2 }, barWidth: '30%', label: { show: true, position: 'top', color: ink, fontSize: 10, formatter: '{c}%' } }
      ]
    });
    window.addEventListener('resize', function () { c1.resize(); });
  }

  // ===== 图2：差异定性分布 =====
  var el2 = document.getElementById('chart-diff-dist');
  if (el2) {
    var c2 = echarts.init(el2, null, { renderer: 'svg' });
    c2.setOption({
      animation: false,
      tooltip: { trigger: 'item', appendToBody: true, formatter: '{b}: {c} 点 ({d}%)' },
      legend: { orient: 'vertical', right: 10, top: 'center', textStyle: { color: muted } },
      series: [{
        type: 'pie', radius: ['42%', '68%'], center: ['38%', '50%'],
        itemStyle: { borderColor: bg2, borderWidth: 2 },
        label: { color: ink, formatter: '{b}\n{c} 点' },
        data: [
          { value: 46, name: '仅命名不同', itemStyle: { color: accent } },
          { value: 31, name: 'PA优势', itemStyle: { color: accent2 } },
          { value: 25, name: '实现路径不同', itemStyle: { color: muted } },
          { value: 18, name: '天融信优势(含待验证)', itemStyle: { color: '#2f7d4f' } },
          { value: 17, name: '待验证', itemStyle: { color: '#94a3b8' } },
          { value: 2, name: '部分差异', itemStyle: { color: accent + '99' } }
        ]
      }]
    });
    window.addEventListener('resize', function () { c2.resize(); });
  }

  // ===== 图3：58条需求结论分布 =====
  var el3 = document.getElementById('chart-req-concl');
  if (el3) {
    var c3 = echarts.init(el3, null, { renderer: 'svg' });
    c3.setOption({
      animation: false,
      tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, appendToBody: true, valueFormatter: function (v) { return v + ' 条'; } },
      grid: { left: 190, right: 40, top: 20, bottom: 30 },
      xAxis: { type: 'value', axisLabel: { color: muted }, splitLine: { lineStyle: { color: rule } } },
      yAxis: {
        type: 'category',
        data: ['待查资料（规格/资质）', 'PA优势·TS待实测', '待实测（机制有据）', '天融信缺口', 'PA优势', '双方支持·TS占优', '双方支持'],
        axisLabel: { color: ink }, axisLine: { lineStyle: { color: rule } }
      },
      series: [{
        type: 'bar', barWidth: '60%',
        data: [
          { value: 11, itemStyle: { color: '#94a3b8' } },
          { value: 10, itemStyle: { color: accent2 } },
          { value: 13, itemStyle: { color: accent + '99' } },
          { value: 4, itemStyle: { color: '#b91c1c' } },
          { value: 1, itemStyle: { color: accent2 } },
          { value: 4, itemStyle: { color: '#2f7d4f' } },
          { value: 15, itemStyle: { color: accent } }
        ],
        label: { show: true, position: 'right', color: ink, formatter: '{c} 条' }
      }]
    });
    window.addEventListener('resize', function () { c3.resize(); });
  }
})();
