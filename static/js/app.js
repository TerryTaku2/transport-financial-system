/* Transport Fleet ERP — Client JS */

document.addEventListener('DOMContentLoaded', function () {

  // ── Feather icons ──
  if (typeof feather !== 'undefined') feather.replace();

  // ── Sidebar: remember scroll position across page loads, so navigating
  // to a link further down a section doesn't dump you back at the top on
  // the next page and force you to scroll down again to find it. ──
  var sidebar = document.getElementById('sidebar');
  if (sidebar) {
    var savedScroll = sessionStorage.getItem('sidebarScrollTop');
    if (savedScroll !== null) {
      sidebar.scrollTop = parseInt(savedScroll, 10) || 0;
    } else {
      var activeLink = sidebar.querySelector('.nav-item.active');
      if (activeLink) activeLink.scrollIntoView({ block: 'center' });
    }
    sidebar.addEventListener('scroll', function () {
      sessionStorage.setItem('sidebarScrollTop', sidebar.scrollTop);
    });
  }

  // ── Mobile nav: hamburger toggle + backdrop ──
  var sidebarToggle = document.getElementById('sidebarToggle');
  var sidebarOverlay = document.getElementById('sidebarOverlay');
  var mobileNavQuery = window.matchMedia('(max-width: 768px)');

  function openSidebar() {
    if (sidebar) sidebar.classList.add('open');
    if (sidebarOverlay) sidebarOverlay.classList.add('show');
    document.body.classList.add('sidebar-locked');
    if (sidebarToggle) sidebarToggle.setAttribute('aria-expanded', 'true');
  }
  function closeSidebar() {
    if (sidebar) sidebar.classList.remove('open');
    if (sidebarOverlay) sidebarOverlay.classList.remove('show');
    document.body.classList.remove('sidebar-locked');
    if (sidebarToggle) sidebarToggle.setAttribute('aria-expanded', 'false');
  }
  if (sidebarToggle) {
    sidebarToggle.addEventListener('click', function () {
      if (sidebar && sidebar.classList.contains('open')) closeSidebar();
      else openSidebar();
    });
  }
  if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebar);
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeSidebar();
  });
  if (sidebar) {
    sidebar.querySelectorAll('.nav-item').forEach(function (link) {
      link.addEventListener('click', function () {
        if (mobileNavQuery.matches) closeSidebar();
      });
    });
  }
  // Dropping below/above the mobile breakpoint (rotation, resize) shouldn't
  // leave the sidebar open+locked when the layout no longer needs a toggle.
  mobileNavQuery.addEventListener('change', function (e) {
    if (!e.matches) closeSidebar();
  });

  // ── Auto-dismiss flash alerts ──
  setTimeout(function () {
    document.querySelectorAll('.alert').forEach(function (el) {
      el.style.transition = 'opacity .4s';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 400);
    });
  }, 4500);

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

  // ── Store purchase: auto-compute total ──
  var purQty  = document.getElementById('pur_qty');
  var purCost = document.getElementById('pur_cost');
  var purTotal = document.getElementById('pur_total_preview');
  function updatePurchaseTotal() {
    if (!purQty || !purCost || !purTotal) return;
    var q = parseFloat(purQty.value) || 0;
    var c = parseFloat(purCost.value) || 0;
    purTotal.textContent = '$' + (q * c).toFixed(2);
  }
  if (purQty)  purQty.addEventListener('input', updatePurchaseTotal);
  if (purCost) purCost.addEventListener('input', updatePurchaseTotal);

  // ── Store sale: prefill price/stock cap from selected part, auto-compute total ──
  var salePart  = document.getElementById('sale_part');
  var saleQty   = document.getElementById('sale_qty');
  var salePrice = document.getElementById('sale_price');
  var saleTotal = document.getElementById('sale_total_preview');
  var saleStockHint = document.getElementById('sale_stock_hint');
  function updateSaleTotal() {
    if (!saleQty || !salePrice || !saleTotal) return;
    var q = parseFloat(saleQty.value) || 0;
    var p = parseFloat(salePrice.value) || 0;
    saleTotal.textContent = '$' + (q * p).toFixed(2);
  }
  function onSalePartChange() {
    if (!salePart) return;
    var opt = salePart.options[salePart.selectedIndex];
    var price = opt ? parseFloat(opt.getAttribute('data-price')) || 0 : 0;
    var stock = opt ? parseInt(opt.getAttribute('data-stock'), 10) || 0 : 0;
    if (salePrice && opt && opt.value) salePrice.value = price.toFixed(2);
    if (saleQty && opt && opt.value) saleQty.setAttribute('max', stock);
    if (saleStockHint) saleStockHint.textContent = opt && opt.value ? stock + ' in stock' : '';
    updateSaleTotal();
  }
  if (salePart)  salePart.addEventListener('change', onSalePartChange);
  if (saleQty)   saleQty.addEventListener('input', updateSaleTotal);
  if (salePrice) salePrice.addEventListener('input', updateSaleTotal);

  // ── Store sale: hide the free-text Customer field once a company vehicle is picked ──
  var saleVehicle = document.getElementById('sale_vehicle');
  var saleCustomerGroup = document.getElementById('sale_customer_group');
  var saleCustomer = document.getElementById('sale_customer');
  function onSaleVehicleChange() {
    if (!saleVehicle || !saleCustomerGroup) return;
    var toVehicle = !!saleVehicle.value;
    saleCustomerGroup.style.display = toVehicle ? 'none' : '';
    if (toVehicle && saleCustomer) saleCustomer.value = '';
  }
  if (saleVehicle) saleVehicle.addEventListener('change', onSaleVehicleChange);

  // ── Confirm delete ──
  document.querySelectorAll('form[data-confirm]').forEach(function (form) {
    form.addEventListener('submit', function (e) {
      var msg = form.getAttribute('data-confirm') || 'Are you sure?';
      if (!confirm(msg)) e.preventDefault();
    });
  });

  // ── Dashboard charts ──
  initDashboardCharts();

  // ── Searchable driver/vehicle selects ──
  enhanceSearchableSelects();
});

// Turns every plain <select data-refdata="drivers|vehicles|franchise_vehicles">
// into a type-to-filter combobox, so a fleet with a long driver/vehicle list
// doesn't force scrolling through a native dropdown. The original <select> is
// kept in the DOM (invisible, overlaid) so form submission, `required`
// validation, and offline.js's repopulateSelects() all keep working exactly
// as before — this only changes how the value gets picked.
function enhanceSearchableSelects() {
  document.querySelectorAll('select[data-refdata="drivers"], select[data-refdata="vehicles"], select[data-refdata="franchise_vehicles"]').forEach(function (select) {
    if (select.dataset.searchEnhanced) return;
    select.dataset.searchEnhanced = '1';
    buildSearchableSelect(select);
  });
}

function buildSearchableSelect(select) {
  var wrap = document.createElement('div');
  wrap.className = 'searchable-select';
  select.parentNode.insertBefore(wrap, select);

  var input = document.createElement('input');
  input.type = 'text';
  input.className = 'searchable-select-input';
  input.autocomplete = 'off';
  input.placeholder = 'Type to search…';
  wrap.appendChild(input);

  var clearBtn = document.createElement('button');
  clearBtn.type = 'button';
  clearBtn.className = 'searchable-select-clear';
  clearBtn.innerHTML = '&times;';
  clearBtn.setAttribute('aria-label', 'Clear selection');
  wrap.appendChild(clearBtn);

  var toggleBtn = document.createElement('button');
  toggleBtn.type = 'button';
  toggleBtn.className = 'searchable-select-toggle';
  toggleBtn.innerHTML = '&#9662;';
  toggleBtn.setAttribute('aria-label', 'Show all options');
  wrap.appendChild(toggleBtn);

  wrap.appendChild(select);
  select.classList.add('searchable-select-native');
  select.setAttribute('tabindex', '-1');

  var dropdown = document.createElement('div');
  dropdown.className = 'searchable-select-dropdown';
  wrap.appendChild(dropdown);

  var activeIndex = -1;
  var visibleOptions = [];

  function realOptions() {
    return Array.prototype.slice.call(select.options).filter(function (o) { return o.value !== ''; });
  }

  function labelFor(value) {
    var opt = Array.prototype.slice.call(select.options).filter(function (o) { return o.value === value; })[0];
    return opt ? opt.textContent : '';
  }

  function syncInputFromSelect() {
    input.value = select.value ? labelFor(select.value) : '';
    clearBtn.style.display = select.value ? '' : 'none';
  }

  function closeDropdown() {
    dropdown.style.display = 'none';
    activeIndex = -1;
  }

  function updateActive(rows) {
    rows.forEach(function (r, i) { r.classList.toggle('active', i === activeIndex); });
    if (activeIndex >= 0 && rows[activeIndex]) rows[activeIndex].scrollIntoView({ block: 'nearest' });
  }

  function renderDropdown(filter) {
    var q = (filter || '').trim().toLowerCase();
    dropdown.innerHTML = '';
    visibleOptions = realOptions().filter(function (o) {
      return !q || o.textContent.toLowerCase().indexOf(q) !== -1;
    });
    if (!visibleOptions.length) {
      var empty = document.createElement('div');
      empty.className = 'searchable-select-empty';
      empty.textContent = 'No matches';
      dropdown.appendChild(empty);
    } else {
      visibleOptions.forEach(function (o) {
        var row = document.createElement('div');
        row.className = 'searchable-select-option';
        row.textContent = o.textContent;
        row.addEventListener('mousedown', function (e) {
          e.preventDefault();
          selectValue(o.value);
        });
        dropdown.appendChild(row);
      });
    }
    activeIndex = -1;
    dropdown.style.display = 'block';
  }

  function selectValue(value) {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
    syncInputFromSelect();
    closeDropdown();
  }

  input.addEventListener('focus', function () { renderDropdown(''); });
  input.addEventListener('input', function () { renderDropdown(input.value); });
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closeDropdown(); syncInputFromSelect(); return; }
    if (dropdown.style.display === 'none') return;
    var rows = dropdown.querySelectorAll('.searchable-select-option');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIndex = Math.min(activeIndex + 1, rows.length - 1);
      updateActive(rows);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIndex = Math.max(activeIndex - 1, 0);
      updateActive(rows);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIndex >= 0 && visibleOptions[activeIndex]) selectValue(visibleOptions[activeIndex].value);
    }
  });

  clearBtn.addEventListener('click', function () {
    selectValue('');
    input.focus();
  });

  toggleBtn.addEventListener('mousedown', function (e) {
    e.preventDefault();
    if (dropdown.style.display === 'none' || !dropdown.style.display) {
      input.focus();
      renderDropdown(input.value === labelFor(select.value) ? '' : input.value);
    } else {
      closeDropdown();
      syncInputFromSelect();
    }
  });

  document.addEventListener('click', function (e) {
    if (!wrap.contains(e.target)) { closeDropdown(); syncInputFromSelect(); }
  });

  // offline.js rewrites select.options wholesale when refdata is refreshed
  // (see repopulateSelects in offline.js); pick up the new label + value.
  select.addEventListener('refdata-repopulated', syncInputFromSelect);

  syncInputFromSelect();
}

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
