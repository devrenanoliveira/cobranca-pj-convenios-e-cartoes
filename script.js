document.addEventListener('DOMContentLoaded', () => {
  let appData = null;
  let chartVolume = null;
  let chartCategory = null;

  // Carrega os dados do arquivo JSON
  fetch('data.json')
    .then(response => response.json())
    .then(data => {
      appData = data;
      initDashboard();
    })
    .catch(err => console.error('Erro ao carregar data.json:', err));

  function initDashboard() {
    renderKPIs();
    renderFilters();
    renderCharts();
    setupEventListeners();
  }

  function renderKPIs() {
    const { totalVolume, avgPrice, activeDealers, totalTransactions } = appData.summary;
    
    document.getElementById('kpi-vol').textContent = `R$ ${(totalVolume / 1e6).toFixed(2)}M`;
    document.getElementById('kpi-price').textContent = `R$ ${avgPrice.toFixed(2)}`;
    document.getElementById('kpi-dealers').textContent = activeDealers;
    document.getElementById('kpi-tx').textContent = totalTransactions.toLocaleString('pt-BR');
  }

  function renderFilters() {
    const regionContainer = document.getElementById('region-chips');
    const categoryContainer = document.getElementById('category-chips');

    if (regionContainer) {
      regionContainer.innerHTML = appData.regions
        .map(r => `<button class="chip active" data-type="region" data-val="${r}">${r}</button>`)
        .join('');
    }

    if (categoryContainer) {
      categoryContainer.innerHTML = appData.categories
        .map(c => `<button class="chip active" data-type="category" data-val="${c}">${c}</button>`)
        .join('');
    }
  }

  function renderCharts() {
    const ctxVol = document.getElementById('chartVolume')?.getContext('2d');
    const ctxCat = document.getElementById('chartCategory')?.getContext('2d');

    if (ctxVol) {
      const labels = appData.monthlyPerformance.map(m => m.month);
      const values = appData.monthlyPerformance.map(m => m.volume / 1000);

      chartVolume = new Chart(ctxVol, {
        type: 'line',
        data: {
          labels,
          datasets: [{
            label: 'Volume (R$ mil)',
            data: values,
            borderColor: '#3B82F6',
            backgroundColor: 'rgba(59, 130, 246, 0.1)',
            fill: true,
            tension: 0.3
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false } },
          scales: {
            x: { ticks: { color: '#94A3B8' }, grid: { color: '#334155' } },
            y: { ticks: { color: '#94A3B8' }, grid: { color: '#334155' } }
          }
        }
      });
    }

    if (ctxCat) {
      // Agrupamento simples por categoria baseado nos dealers
      const catTotals = {};
      appData.dealers.forEach(d => {
        catTotals[d.category] = (catTotals[d.category] || 0) + d.salesVolume;
      });

      chartCategory = new Chart(ctxCat, {
        type: 'doughnut',
        data: {
          labels: Object.keys(catTotals),
          datasets: [{
            data: Object.values(catTotals),
            backgroundColor: ['#3B82F6', '#10B981', '#F59E0B', '#EC4899']
          }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { labels: { color: '#E2E8F0' }, position: 'bottom' }
          }
        }
      });
    }
  }

  function setupEventListeners() {
    document.addEventListener('click', (e) => {
      if (e.target.classList.contains('chip')) {
        e.target.classList.toggle('active');
        // Adicione aqui a lógica de filtragem dinâmica se necessário
      }
    });
  }
});
