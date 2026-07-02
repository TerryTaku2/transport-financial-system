/* Transport Fleet ERP — Client JS */

document.addEventListener('DOMContentLoaded', function () {

  // ── Feather icons ──
  if (typeof feather !== 'undefined') feather.replace();

  // ── Auto-dismiss flash alerts ──
  setTimeout(function () {
    document.querySelectorAll('.alert').forEach(function (el) {
      el.style.transition = 'opacity .4s';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 400);
    });
  }, 4500);

  // ── Fuel log: auto-compute total ──
  var litersInput = document.getElementById('liters');
  var cplInput    = document.getElementById('cost_per_liter');
  var totalSpan   = document.getElementById('total_preview');
  function updateFuelTotal() {
    if (!litersInput || !cplInput || !totalSpan) return;
    var l = parseFloat(litersInput.value) || 0;
    var c = parseFloat(cplInput.value) || 0;
    totalSpan.textContent = '$' + (l * c).toFixed(2);
  }
  if (litersInput) litersInput.addEventListener('input', updateFuelTotal);
  if (cplInput)    cplInput.addEventListener('input', updateFuelTotal);

  // ── Maintenance log: auto-compute total ──
  var partsInput = document.getElementById('parts_cost');
  var laborInput = document.getElementById('labor_cost');
  var maintTotal = document.getElementById('maint_total_preview');
  function updateMaintTotal() {
    if (!partsInput || !laborInput || !maintTotal) return;
    var p = parseFloat(partsInput.value) || 0;
    var l = parseFloat(laborInput.value) || 0;
    maintTotal.textContent = '$' + (p + l).toFixed(2);
  }
  if (partsInput) partsInput.addEventListener('input', updateMaintTotal);
  if (laborInput) laborInput.addEventListener('input', updateMaintTotal);

  // ── Confirm delete ──
  document.querySelectorAll('form[data-confirm]').forEach(function (form) {
    form.addEventListener('submit', function (e) {
      var msg = form.getAttribute('data-confirm') || 'Are you sure?';
      if (!confirm(msg)) e.preventDefault();
    });
  });

  // ── Dashboard charts ──
  initDashboardCharts();
});

function initDashboardCharts() {
  var ctx7 = document.getElementById('chart7Days');
  if (ctx7 && typeof Chart !== 'undefined' && window.REVENUE_7DAYS) {
    var data = window.REVENUE_7DAYS;
    new Chart(ctx7, {
      type: 'line',
      data: {
        labels: data.map(d => d.date),
        datasets: [{
          label: 'Revenue (USD)',
          data: data.map(d => d.revenue),
          borderColor: '#2563eb',
          backgroundColor: 'rgba(37,99,235,0.08)',
          borderWidth: 2.5,
          pointBackgroundColor: '#2563eb',
          pointRadius: 4,
          fill: true,
          tension: 0.3,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          y: {
            beginAtZero: true,
            grid: { color: '#f1f5f9' },
            ticks: {
              callback: v => '$' + v.toLocaleString(),
              font: { size: 11 },
            }
          },
          x: {
            grid: { display: false },
            ticks: { font: { size: 11 } }
          }
        }
      }
    });
  }

  // Monthly revenue vs expenses
  var ctxMonthly = document.getElementById('chartMonthly');
  if (ctxMonthly && typeof Chart !== 'undefined') {
    fetch('/api/revenue/monthly')
      .then(r => r.json())
      .then(data => {
        new Chart(ctxMonthly, {
          type: 'bar',
          data: {
            labels: data.map(d => d.month),
            datasets: [
              {
                label: 'Revenue',
                data: data.map(d => d.revenue),
                backgroundColor: 'rgba(37,99,235,0.8)',
                borderRadius: 4,
              },
              {
                label: 'Expenses',
                data: data.map(d => d.expenses),
                backgroundColor: 'rgba(220,38,38,0.7)',
                borderRadius: 4,
              },
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { position: 'top', labels: { font: { size: 11 } } }
            },
            scales: {
              y: {
                beginAtZero: true,
                grid: { color: '#f1f5f9' },
                ticks: { callback: v => '$' + v.toLocaleString(), font: { size: 11 } }
              },
              x: { grid: { display: false }, ticks: { font: { size: 11 } } }
            }
          }
        });
      });
  }

  // Expense breakdown doughnut
  var ctxDonut = document.getElementById('chartExpenses');
  if (ctxDonut && typeof Chart !== 'undefined') {
    fetch('/api/expenses/breakdown')
      .then(r => r.json())
      .then(data => {
        new Chart(ctxDonut, {
          type: 'doughnut',
          data: {
            labels: ['Fuel', 'Maintenance'],
            datasets: [{
              data: [data.fuel, data.maintenance],
              backgroundColor: ['#f59e0b', '#ef4444'],
              borderWidth: 0,
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { position: 'bottom', labels: { font: { size: 11 } } },
              tooltip: {
                callbacks: {
                  label: ctx => ' $' + ctx.raw.toFixed(2)
                }
              }
            },
            cutout: '68%',
          }
        });
      });
  }

  // Vehicle performance bar
  var ctxVeh = document.getElementById('chartVehicles');
  if (ctxVeh && typeof Chart !== 'undefined') {
    fetch('/api/vehicles/performance')
      .then(r => r.json())
      .then(data => {
        if (!data.length) return;
        new Chart(ctxVeh, {
          type: 'bar',
          data: {
            labels: data.map(d => d.vehicle),
            datasets: [{
              label: 'Revenue (USD)',
              data: data.map(d => d.revenue),
              backgroundColor: 'rgba(22,163,74,0.8)',
              borderRadius: 4,
            }]
          },
          options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: {
                beginAtZero: true,
                ticks: { callback: v => '$' + v.toLocaleString(), font: { size: 11 } },
                grid: { color: '#f1f5f9' }
              },
              y: { grid: { display: false }, ticks: { font: { size: 11 } } }
            }
          }
        });
      });
  }
}
