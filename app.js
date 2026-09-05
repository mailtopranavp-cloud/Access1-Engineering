/**
 * Application Controller for Design Department Weekly Report Dashboard
 * Supports:
 * - Provision for adding new projects and removing existing ones
 * - Notes & Observations mapped directly from Column I
 * - Monthly report chart formatted as clustered column (grouped bars)
 * - OSR Resource Utilization Y-axis labeled "Number of OSR Resources"
 * - Key Personnel, Modelers, and Schedule updated dynamically on file upload
 * - Formula: Period of Extension = Planned end Date – Present Date (formatted as X Months)
 * - Report Date value mapped by extracting date stamp from Excel file name
 */

(function () {
  'use strict';

  // --- State ---
  let projects = {};
  let currentProjectKey = 'PH497-SG1';
  let currentDateKey = '2026-09-05';
  let activeView = 'overview';
  let chartInstances = {};
  let drawingFilterStatus = 'ALL';
  let drawingSearchQuery = '';

  // Palette matching PowerPoint mockup
  const STATUS_COLORS = {
    approved: '#1F5070',          // Deep Blue
    approvedWithResub: '#E07A26', // Amber / Orange
    rejected: '#C0392B',          // Crimson Red
    underPrep: '#0EA5E9',        // Sky / Teal
    pending: '#8E44AD',           // Purple
    notStarted: '#16A34A',       // Green
    submitted: '#2563EB',         // Royal Blue
    extension: '#EA580C'          // Vibrant Orange for extension period
  };

  // --- Chart.js Custom Plugins ---

  // Bar Tip Label Plugin: Draws count directly above each bar
  const barTipLabelsPlugin = {
    id: 'barTipLabels',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      chart.data.datasets.forEach((dataset, i) => {
        const meta = chart.getDatasetMeta(i);
        if (!meta.hidden) {
          meta.data.forEach((bar, index) => {
            const val = dataset.data[index];
            if (val !== undefined && val !== null && val > 0) {
              ctx.save();
              ctx.textAlign = 'center';
              ctx.textBaseline = 'bottom';
              ctx.font = 'bold 12px sans-serif';
              ctx.fillStyle = '#0F172A';
              ctx.fillText(String(val), bar.x, bar.y - 4);
              ctx.restore();
            }
          });
        }
      });
    }
  };

  /**
   * Initializes data from localStorage or default seed
   */
  function initData() {
    try {
      const saved = localStorage.getItem('company_dashboard_projects');
      if (saved) {
        projects = JSON.parse(saved);
      }
    } catch (e) {
      console.warn('Could not read from localStorage', e);
    }

    if (!projects || Object.keys(projects).length === 0) {
      projects = window.DEFAULT_PROJECTS || {};
      saveData();
    }

    if (!projects[currentProjectKey]) {
      currentProjectKey = Object.keys(projects)[0] || 'PH497-SG1';
    }

    const proj = projects[currentProjectKey];
    if (proj && proj.history) {
      const dates = Object.keys(proj.history).sort();
      if (dates.length > 0) {
        currentDateKey = dates[dates.length - 1];
      }
    }

    loadSavedHeadings();
  }

  function saveData() {
    try {
      localStorage.setItem('company_dashboard_projects', JSON.stringify(projects));
    } catch (e) {
      console.warn('Could not write to localStorage', e);
    }
  }

  // --- Editable Headings Persistence ---

  function loadSavedHeadings() {
    try {
      const savedHeadings = JSON.parse(localStorage.getItem('dashboard_custom_headings') || '{}');
      Object.keys(savedHeadings).forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.textContent = savedHeadings[id];
      });
    } catch (e) {
      console.warn('Could not load custom headings', e);
    }
  }

  function saveHeading(el) {
    if (!el || !el.id) return;
    try {
      const savedHeadings = JSON.parse(localStorage.getItem('dashboard_custom_headings') || '{}');
      savedHeadings[el.id] = el.textContent.trim();
      localStorage.setItem('dashboard_custom_headings', JSON.stringify(savedHeadings));
    } catch (e) {
      console.warn('Could not save heading', e);
    }
  }

  // --- Chart Lifecycle ---

  function destroyChart(key) {
    if (chartInstances[key]) {
      chartInstances[key].destroy();
      delete chartInstances[key];
    }
  }

  function destroyAllCharts() {
    Object.keys(chartInstances).forEach((key) => {
      chartInstances[key].destroy();
    });
    chartInstances = {};
  }

  /**
   * Returns latest or specified report state for a project
   */
  function getProjectReport(proj, dateKey) {
    if (!proj) return null;
    if (proj.history && dateKey && proj.history[dateKey]) {
      return {
        ...proj,
        ...proj.history[dateKey],
        dateKey: dateKey
      };
    }
    if (proj.history) {
      const dates = Object.keys(proj.history).sort();
      if (dates.length > 0) {
        const latest = dates[dates.length - 1];
        return {
          ...proj,
          ...proj.history[latest],
          dateKey: latest
        };
      }
    }
    return proj;
  }

  // --- Project Addition & Removal ---

  function deleteProject(key) {
    const projName = projects[key]?.displayName || key;
    if (confirm(`Are you sure you want to remove project "${projName}" from the dashboard?`)) {
      delete projects[key];
      saveData();
      if (currentProjectKey === key) {
        currentProjectKey = Object.keys(projects)[0] || '';
      }
      if (activeView === 'overview') {
        renderOverview();
      } else {
        if (currentProjectKey) {
          navigateToDetail(currentProjectKey);
        } else {
          navigateToOverview();
        }
      }
    }
  }

  function createManualProject(num, name, pm, totalDwg) {
    const projectNumber = (num || 'PRJ').trim();
    const projectName = (name || 'New Project').trim();
    const displayName = `${projectNumber}-${projectName}`;
    const today = new Date();
    const dateKey = today.toISOString().split('T')[0];
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    const dateOfReport = `${dd}/${mm}/${yyyy}`;
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const monthYearLabel = `${monthNames[today.getMonth()]}, ${String(yyyy).slice(-2)}`;

    const total = parseInt(totalDwg, 10) || 50;
    const approved = Math.round(total * 0.4);
    const underPrep = Math.round(total * 0.3);
    const pending = Math.round(total * 0.2);
    const notStarted = total - (approved + underPrep + pending);
    const completionPct = ((approved / total) * 100).toFixed(1);

    const plannedEndDate = `30/${String((today.getMonth() + 6) % 12 + 1).padStart(2, '0')}/${yyyy}`;
    const presentDate = dateOfReport;
    const periodOfExtension = window.ExcelParser.calculatePeriodOfExtension(plannedEndDate, presentDate);

    projects[displayName] = {
      projectNumber,
      projectName,
      displayName,
      projectLead: pm || 'Pranav',
      projectInCharge: 'Rejoe',
      bimCoordinator: 'Ajmal',
      projectDc: 'Romeo',
      procurement: 'Osama',
      estimation: 'Abdul',
      plannedStartDate: `01/${mm}/${yyyy}`,
      plannedEndDate: plannedEndDate,
      actualStartDate: dateOfReport,
      presentDate: presentDate,
      periodOfExtension: periodOfExtension,
      modelers: [
        { name: "OSR 1 (Lead Modeler)", role: "BIM Specialist", dateOfJoin: dateOfReport, active: true },
        { name: "OSR 2 (Senior Modeler)", role: "Revit Architectural", dateOfJoin: dateOfReport, active: true }
      ],
      notes: [
        "Project initialized in dashboard.",
        "Awaiting submittal schedule from client."
      ],
      history: {
        [dateKey]: {
          fileName: `${projectNumber}_${dateKey}.xlsx`,
          dateOfReport,
          presentDate,
          plannedEndDate,
          periodOfExtension,
          monthYearLabel,
          totalDwg: total,
          resourceCount: 2,
          statusSummary: {
            approved,
            approvedWithResub: 0,
            rejected: 0,
            underPrep,
            pending,
            submitted: 0,
            notStarted
          },
          completionPct: parseFloat(completionPct),
          drawings: []
        }
      }
    };

    saveData();
    navigateToDetail(displayName);
  }

  // --- Rendering: Overview (Page 1) ---

  function renderOverview() {
    destroyAllCharts();
    const grid = document.getElementById('projects-grid');
    if (!grid) return;
    grid.innerHTML = '';

    const query = (document.getElementById('search-projects')?.value || '').toLowerCase().trim();

    const projectKeys = Object.keys(projects).filter((key) => {
      const p = projects[key];
      if (!query) return true;
      return (
        p.displayName?.toLowerCase().includes(query) ||
        p.projectNumber?.toLowerCase().includes(query) ||
        p.projectLead?.toLowerCase().includes(query) ||
        p.projectName?.toLowerCase().includes(query)
      );
    });

    if (projectKeys.length === 0) {
      grid.innerHTML = `
        <div style="grid-column: 1/-1; text-align: center; padding: 48px; background: #FFF; border-radius: 16px; border: 1px dashed #CBD5E1;">
          <h3 style="color: #64748B;">No matching projects found</h3>
          <p style="color: #94A3B8; font-size: 14px; margin-top: 6px;">Use "+ Add Project" or "Refresh / Upload Excel" to create one.</p>
        </div>
      `;
      return;
    }

    projectKeys.forEach((key) => {
      const rawProj = projects[key];
      const proj = getProjectReport(rawProj, currentDateKey);
      const summary = proj.statusSummary || {
        approved: 0,
        approvedWithResub: 0,
        rejected: 0,
        underPrep: 0,
        pending: 0,
        notStarted: 0,
        submitted: 0
      };

      const totalDwg = proj.totalDwg || Object.values(summary).reduce((a, b) => a + b, 0);
      const approvedCount = summary.approved || 0;
      const completionPct = totalDwg > 0 ? ((approvedCount / totalDwg) * 100).toFixed(1) : '0.0';

      const card = document.createElement('div');
      card.className = 'project-card';
      const chartId = `chart-donut-${key.replace(/[^a-zA-Z0-9_-]/g, '_')}`;

      card.innerHTML = `
        <div class="card-header">
          <div class="card-header-left">
            <a class="card-header-link" data-project="${key}">
              ${proj.displayName || key}
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
            </a>
          </div>
          <div class="card-header-right">
            <div class="card-pm">PL: ${proj.projectLead || 'Pranav'}</div>
            <button class="delete-project-btn" data-delete-project="${key}" title="Remove Project">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/></svg>
            </button>
          </div>
        </div>

        <div class="card-body">
          <div class="stats-list">
            <div class="stats-row stats-row-total">
              <span>Total Number of DWG</span>
              <span class="stats-val">: ${totalDwg}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.approved}"></span>Approved</span>
              <span class="stats-val">: ${summary.approved || 0}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.approvedWithResub}"></span>Approved with Resub.</span>
              <span class="stats-val">: ${summary.approvedWithResub || 0}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.rejected}"></span>Rejected</span>
              <span class="stats-val">: ${summary.rejected || 0}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.underPrep}"></span>Under Preparation</span>
              <span class="stats-val">: ${summary.underPrep || 0}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.pending}"></span>Pending(Notes)</span>
              <span class="stats-val">: ${summary.pending || 0}</span>
            </div>
            <div class="stats-row">
              <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.notStarted}"></span>Not Started</span>
              <span class="stats-val">: ${summary.notStarted || 0}</span>
            </div>
            ${
              summary.submitted
                ? `
              <div class="stats-row">
                <span class="stats-label"><span class="status-dot" style="background: ${STATUS_COLORS.submitted}"></span>Submitted/Review</span>
                <span class="stats-val">: ${summary.submitted}</span>
              </div>`
                : ''
            }
          </div>

          <div class="card-chart-container">
            <canvas id="${chartId}"></canvas>
            <div class="chart-center-pct">
              <span class="pct-val">${completionPct}%</span>
              <span class="pct-sub">Completed</span>
            </div>
          </div>
        </div>

        <div class="card-footer">
          <span>Date: ${proj.dateOfReport || '05/09/2026'}</span>
          <span class="view-details-btn" data-project="${key}">View Details &rarr;</span>
        </div>
      `;

      card.querySelectorAll('[data-project]').forEach((el) => {
        el.addEventListener('click', (e) => {
          e.preventDefault();
          navigateToDetail(key);
        });
      });

      card.querySelector('[data-delete-project]').addEventListener('click', (e) => {
        e.stopPropagation();
        deleteProject(key);
      });

      grid.appendChild(card);

      setTimeout(() => {
        const ctx = document.getElementById(chartId);
        if (ctx) {
          const chartData = [
            summary.approved || 0,
            summary.approvedWithResub || 0,
            summary.rejected || 0,
            summary.underPrep || 0,
            summary.pending || 0,
            summary.notStarted || 0
          ];
          const bgColors = [
            STATUS_COLORS.approved,
            STATUS_COLORS.approvedWithResub,
            STATUS_COLORS.rejected,
            STATUS_COLORS.underPrep,
            STATUS_COLORS.pending,
            STATUS_COLORS.notStarted
          ];

          if (summary.submitted) {
            chartData.push(summary.submitted);
            bgColors.push(STATUS_COLORS.submitted);
          }

          chartInstances[chartId] = new Chart(ctx, {
            type: 'doughnut',
            data: {
              datasets: [
                {
                  data: chartData,
                  backgroundColor: bgColors,
                  borderWidth: 2,
                  borderColor: '#FFFFFF',
                  hoverOffset: 3
                }
              ]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              cutout: '58%',
              plugins: {
                legend: { display: false },
                tooltip: {
                  callbacks: {
                    label: function (c) {
                      const labels = [
                        'Approved',
                        'Appr. w/ Resub',
                        'Rejected',
                        'Under Prep',
                        'Pending(Notes)',
                        'Not Started',
                        'Submitted'
                      ];
                      const val = c.raw;
                      const pct = totalDwg > 0 ? Math.round((val / totalDwg) * 100) : 0;
                      return ` ${labels[c.dataIndex] || ''}: ${val} (${pct}%)`;
                    }
                  }
                }
              }
            }
          });
        }
      }, 0);
    });
  }

  // --- Rendering: Project In-Depth Detail (Page 2) ---

  function renderProjectDetail() {
    destroyAllCharts();
    const rawProj = projects[currentProjectKey];
    if (!rawProj) {
      navigateToOverview();
      return;
    }

    const proj = getProjectReport(rawProj, currentDateKey);

    // Update Title Banner
    const titleElem = document.getElementById('detail-project-title');
    if (titleElem) {
      titleElem.textContent = `Project Overview (${proj.projectNumber || ''} ${proj.projectName || ''})`;
    }

    // Correction 2: Populate Date Selector with file name date stamps and file names
    const dateSelect = document.getElementById('detail-date-select');
    if (dateSelect && rawProj.history) {
      dateSelect.innerHTML = '';
      const dates = Object.keys(rawProj.history).sort().reverse();
      dates.forEach((d) => {
        const h = rawProj.history[d];
        const opt = document.createElement('option');
        opt.value = d;
        // Display Report Date and File Name
        const fileTag = h.fileName ? ` (${h.fileName})` : '';
        opt.textContent = `${h.dateOfReport || d}${fileTag}`;
        if (d === currentDateKey) opt.selected = true;
        dateSelect.appendChild(opt);
      });
      dateSelect.style.display = dates.length > 1 ? 'inline-block' : 'none';
    }

    // Key Personnel Block
    document.getElementById('val-incharge').textContent = proj.projectInCharge || 'Rejoe';
    document.getElementById('val-lead').textContent = proj.projectLead || 'Pranav';
    document.getElementById('val-bim').textContent = proj.bimCoordinator || 'Ajmal';
    document.getElementById('val-dc').textContent = proj.projectDc || 'Romeo';
    document.getElementById('val-procurement').textContent = proj.procurement || 'Osama';
    document.getElementById('val-estimation').textContent = proj.estimation || 'Abdul';

    // Modelers List (Column C & D)
    const modelersContainer = document.getElementById('detail-modelers-list');
    if (modelersContainer) {
      modelersContainer.innerHTML = '';
      const modelers = proj.modelers || [];
      if (modelers.length === 0) {
        modelersContainer.innerHTML = '<p style="color: #94A3B8; font-size: 13px;">No modelers listed in Column C.</p>';
      } else {
        modelers.forEach((m) => {
          const item = document.createElement('div');
          item.className = 'modeler-item';
          item.innerHTML = `
            <span class="modeler-name">${m.name}</span>
            <span class="modeler-date">${m.dateOfJoin}</span>
          `;
          modelersContainer.appendChild(item);
        });
      }
    }

    // Correction 1: Dynamic Calculation of Period of Extension = Planned end Date – Present Date
    const plannedEndDate = proj.plannedEndDate || '06/06/2026';
    const presentDate = proj.presentDate || proj.dateOfReport || '05/09/2026';
    const computedExtension = window.ExcelParser.calculatePeriodOfExtension(plannedEndDate, presentDate);

    document.getElementById('val-plan-start').textContent = proj.plannedStartDate || '01/01/2026';
    document.getElementById('val-plan-end').textContent = plannedEndDate;
    document.getElementById('val-act-start').textContent = proj.actualStartDate || '04/04/2026';
    document.getElementById('val-present-date').textContent = presentDate;
    document.getElementById('val-extension').textContent = computedExtension;

    // Notes List (Mapped from Column I)
    const notesContainer = document.getElementById('detail-notes-list');
    if (notesContainer) {
      notesContainer.innerHTML = '';
      const notes = proj.notes || [];
      if (notes.length === 0) {
        notesContainer.innerHTML = '<li style="color: #94A3B8;">No notes or observations entered in Column I of the Excel file.</li>';
      } else {
        notes.forEach((note) => {
          const li = document.createElement('li');
          li.textContent = note;
          notesContainer.appendChild(li);
        });
      }
    }

    // Render Charts
    renderMonthlyChart(rawProj);
    renderUtilizationChart(rawProj);
    renderProgrammeChart(rawProj);

    // Render Drawing Register Table
    renderDrawingsTable(proj);
  }

  /**
   * Monthly Report Chart formatted as CLUSTERED COLUMN (Grouped Bars)
   */
  function renderMonthlyChart(rawProj) {
    const ctx = document.getElementById('chart-monthly-report');
    if (!ctx) return;
    destroyChart('monthly');

    const historyObj = rawProj.history || {};
    const sortedDates = Object.keys(historyObj).sort();

    const labels = sortedDates.map((dk) => {
      const h = historyObj[dk];
      return h.monthYearLabel || window.ExcelParser.formatMonthYear(h.dateOfReport || dk);
    });

    const approvedData = [];
    const resubData = [];
    const rejectedData = [];
    const prepData = [];
    const pendingData = [];
    const notStartedData = [];
    const submittedData = [];

    sortedDates.forEach((dk) => {
      const s = historyObj[dk].statusSummary || {};
      approvedData.push(s.approved || 0);
      resubData.push(s.approvedWithResub || 0);
      rejectedData.push(s.rejected || 0);
      prepData.push(s.underPrep || 0);
      pendingData.push(s.pending || 0);
      notStartedData.push(s.notStarted || 0);
      submittedData.push(s.submitted || 0);
    });

    chartInstances['monthly'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Approved',
            data: approvedData,
            backgroundColor: STATUS_COLORS.approved
          },
          {
            label: 'Appr. W Re.',
            data: resubData,
            backgroundColor: STATUS_COLORS.approvedWithResub
          },
          {
            label: 'Rejected',
            data: rejectedData,
            backgroundColor: STATUS_COLORS.rejected
          },
          {
            label: 'Under Prep.',
            data: prepData,
            backgroundColor: STATUS_COLORS.underPrep
          },
          {
            label: 'Pending',
            data: pendingData,
            backgroundColor: STATUS_COLORS.pending
          },
          {
            label: 'Not Started',
            data: notStartedData,
            backgroundColor: STATUS_COLORS.notStarted
          },
          {
            label: 'Submitted',
            data: submittedData,
            backgroundColor: STATUS_COLORS.submitted
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        categoryPercentage: 0.85,
        barPercentage: 0.9,
        plugins: {
          legend: {
            position: 'top',
            labels: {
              boxWidth: 12,
              font: { size: 11, weight: 'bold' },
              color: '#334155'
            }
          },
          tooltip: {
            mode: 'index',
            intersect: false
          }
        },
        scales: {
          x: {
            stacked: false, // CLUSTERED COLUMN
            grid: { display: false },
            ticks: { font: { weight: 'bold' }, color: '#475569' }
          },
          y: {
            stacked: false, // CLUSTERED COLUMN
            beginAtZero: true,
            title: {
              display: true,
              text: 'Total Number of Drawings',
              font: { weight: 'bold', size: 12 }
            },
            grid: { color: '#F1F5F9' },
            ticks: { precision: 0 }
          }
        }
      }
    });
  }

  /**
   * OSR Resource Utilization Chart
   * - Y-axis label: "Number of OSR Resources"
   */
  function renderUtilizationChart(rawProj) {
    const ctx = document.getElementById('chart-utilization');
    if (!ctx) return;
    destroyChart('utilization');

    const historyObj = rawProj.history || {};
    const sortedDates = Object.keys(historyObj).sort();

    const labels = sortedDates.map((dk) => {
      const h = historyObj[dk];
      return h.monthYearLabel || window.ExcelParser.formatMonthYear(h.dateOfReport || dk);
    });

    const resourceCounts = sortedDates.map((dk) => {
      const h = historyObj[dk];
      return h.resourceCount !== undefined ? h.resourceCount : (rawProj.modelers ? rawProj.modelers.length : 4);
    });

    const maxCount = Math.max(...resourceCounts, 5);

    chartInstances['utilization'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Number of OSR Resources',
            data: resourceCounts,
            backgroundColor: '#1F5070',
            barPercentage: 0.5,
            categoryPercentage: 0.7
          }
        ]
      },
      plugins: [barTipLabelsPlugin],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 20 }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (c) => ` Number of OSR Resources: ${c.raw}`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { weight: 'bold' }, color: '#475569' }
          },
          y: {
            beginAtZero: true,
            suggestedMax: maxCount + 1,
            title: {
              display: true,
              text: 'Number of OSR Resources',
              font: { weight: 'bold', size: 12 }
            },
            ticks: { stepSize: 1, precision: 0 },
            grid: { color: '#E2E8F0' }
          }
        }
      }
    });
  }

  /**
   * Engineering Programme Line Chart
   */
  function renderProgrammeChart(rawProj) {
    const ctx = document.getElementById('chart-programme');
    if (!ctx) return;
    destroyChart('programme');

    const plannedStart = rawProj.plannedStartDate || '01/01/2026';
    const plannedEnd = rawProj.plannedEndDate || '06/06/2026';
    const actualStart = rawProj.actualStartDate || '04/04/2026';
    const presentDate = rawProj.presentDate || '05/09/2026';
    const extension = window.ExcelParser.calculatePeriodOfExtension(plannedEnd, presentDate);

    const months = [
      "Jan, 26", "Feb, 26", "Mar, 26", "Apr, 26", "May, 26", "Jun, 26",
      "Jul, 26", "Aug, 26", "Sep, 26", "Oct, 26"
    ];

    const plannedData = [
      0, 20, 40, 60, 80, 100, null, null, null, null
    ];

    const historyObj = rawProj.history || {};
    const actualData = new Array(months.length).fill(null);

    actualData[3] = 0;
    actualData[4] = 25;
    actualData[5] = 50;
    actualData[6] = 68;

    Object.keys(historyObj).forEach((dk) => {
      const h = historyObj[dk];
      const mLabel = h.monthYearLabel || window.ExcelParser.formatMonthYear(h.dateOfReport || dk);
      const mIdx = months.indexOf(mLabel);
      if (mIdx !== -1) {
        const approved = h.statusSummary?.approved || 0;
        const total = h.totalDwg || 1;
        const pct = Math.round((approved / total) * 100);
        actualData[mIdx] = pct;
      }
    });

    const extensionData = new Array(months.length).fill(null);
    extensionData[5] = actualData[5] || 50;
    extensionData[6] = actualData[6] || 68;
    extensionData[7] = actualData[7] || 82;
    extensionData[8] = actualData[8] || 69;

    chartInstances['programme'] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: months,
        datasets: [
          {
            label: `Planned Date (${plannedStart} to ${plannedEnd})`,
            data: plannedData,
            borderColor: '#1F5070',
            backgroundColor: '#1F5070',
            borderWidth: 2.5,
            tension: 0.2,
            pointRadius: 4
          },
          {
            label: `Actual Date (${actualStart} to ${presentDate})`,
            data: actualData,
            borderColor: '#2563EB',
            backgroundColor: '#2563EB',
            borderWidth: 2.5,
            tension: 0.2,
            pointRadius: 4
          },
          {
            label: `Period of Extension (${extension})`,
            data: extensionData,
            borderColor: STATUS_COLORS.extension,
            backgroundColor: 'rgba(234, 88, 12, 0.1)',
            borderWidth: 3,
            borderDash: [5, 4],
            tension: 0.2,
            pointRadius: 5,
            pointBackgroundColor: STATUS_COLORS.extension
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              font: { weight: 'bold', size: 11 },
              color: '#334155'
            }
          },
          tooltip: {
            callbacks: {
              label: (c) => ` ${c.dataset.label.split('(')[0]}: ${c.raw}%`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { weight: '600' }, color: '#475569' }
          },
          y: {
            beginAtZero: true,
            max: 100,
            title: {
              display: true,
              text: 'Percentage of Completion (%)',
              font: { weight: 'bold', size: 12 }
            },
            ticks: {
              callback: (v) => `${v}%`,
              stepSize: 20
            },
            grid: { color: '#F1F5F9' }
          }
        }
      }
    });
  }

  /**
   * Drawing Register Table
   */
  function renderDrawingsTable(proj) {
    const tbody = document.getElementById('drawings-table-body');
    const countBadge = document.getElementById('drawings-count-badge');
    if (!tbody) return;

    tbody.innerHTML = '';
    const drawings = proj.drawings || [];

    const filtered = drawings.filter((dwg) => {
      const matchesSearch =
        !drawingSearchQuery || dwg.drawingNo.toLowerCase().includes(drawingSearchQuery.toLowerCase());
      const matchesFilter =
        drawingFilterStatus === 'ALL' || dwg.status.toLowerCase() === drawingFilterStatus.toLowerCase();
      return matchesSearch && matchesFilter;
    });

    if (countBadge) {
      countBadge.textContent = `Showing ${filtered.length} of ${drawings.length} drawings`;
    }

    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="3" style="text-align: center; padding: 24px; color: #94A3B8;">
            No drawings match current filters.
          </td>
        </tr>
      `;
      return;
    }

    filtered.slice(0, 500).forEach((dwg, index) => {
      const tr = document.createElement('tr');
      const badgeClass = getBadgeClass(dwg.status);
      tr.innerHTML = `
        <td style="color: #64748B; width: 60px;">#${index + 1}</td>
        <td style="font-family: monospace; font-weight: 600;">${dwg.drawingNo}</td>
        <td><span class="badge ${badgeClass}">${dwg.status}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }

  function getBadgeClass(status) {
    const s = String(status).toLowerCase();
    if (s.includes('resub')) return 'badge-resub';
    if (s.startsWith('appr')) return 'badge-approved';
    if (s.includes('rej')) return 'badge-rejected';
    if (s.includes('prep')) return 'badge-prep';
    if (s.includes('pend')) return 'badge-pending';
    if (s.includes('review') || s.includes('submi')) return 'badge-submitted';
    return 'badge-notstarted';
  }

  // --- Routing / Navigation ---

  function navigateToOverview() {
    activeView = 'overview';
    document.getElementById('view-overview').classList.add('active');
    document.getElementById('view-detail').classList.remove('active');
    window.location.hash = 'overview';
    renderOverview();
  }

  function navigateToDetail(projectKey) {
    if (projectKey && projects[projectKey]) {
      currentProjectKey = projectKey;
    }
    activeView = 'detail';
    document.getElementById('view-overview').classList.remove('active');
    document.getElementById('view-detail').classList.add('active');
    window.location.hash = `detail?project=${encodeURIComponent(currentProjectKey)}`;
    renderProjectDetail();
  }

  // --- Excel Upload & Refresh Handler ---

  async function handleFilesUpload(files) {
    if (!files || files.length === 0) return;
    const statusMsg = document.getElementById('upload-status-msg');
    if (statusMsg) {
      statusMsg.style.display = 'block';
      statusMsg.textContent = `Processing ${files.length} file(s)...`;
      statusMsg.style.color = '#1E3A8A';
    }

    let loadedCount = 0;
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (!file.name.match(/\.xlsx$/i)) continue;

      try {
        const buffer = await file.arrayBuffer();
        const parsed = await window.ExcelParser.parseWorkbookBuffer(buffer, file.name);

        let pKey = parsed.displayName;
        const matchedKey = Object.keys(projects).find((k) => 
          k.toLowerCase() === pKey.toLowerCase() ||
          projects[k].projectNumber?.toLowerCase() === parsed.projectNumber?.toLowerCase()
        );
        if (matchedKey) {
          pKey = matchedKey;
        }

        if (!projects[pKey]) {
          projects[pKey] = {
            projectNumber: parsed.projectNumber,
            projectName: parsed.projectName,
            displayName: parsed.displayName,
            history: {}
          };
        }

        // Live update all project properties
        projects[pKey].projectNumber = parsed.projectNumber;
        projects[pKey].projectName = parsed.projectName;
        projects[pKey].displayName = pKey;
        projects[pKey].projectLead = parsed.projectLead;
        projects[pKey].projectInCharge = parsed.projectInCharge;
        projects[pKey].bimCoordinator = parsed.bimCoordinator;
        projects[pKey].projectDc = parsed.projectDc;
        projects[pKey].procurement = parsed.procurement;
        projects[pKey].estimation = parsed.estimation;
        projects[pKey].plannedStartDate = parsed.plannedStartDate;
        projects[pKey].plannedEndDate = parsed.plannedEndDate;
        projects[pKey].actualStartDate = parsed.actualStartDate;
        projects[pKey].presentDate = parsed.presentDate;
        projects[pKey].periodOfExtension = parsed.periodOfExtension;
        projects[pKey].modelers = parsed.modelers;
        projects[pKey].notes = parsed.notes;

        // Store file-specific snapshot in history
        const dateKey = parsed.dateKey || 'latest';
        if (!projects[pKey].history) projects[pKey].history = {};

        projects[pKey].history[dateKey] = {
          fileName: file.name,
          dateOfReport: parsed.dateOfReport,
          monthYearLabel: parsed.monthYearLabel,
          totalDwg: parsed.totalDwg,
          resourceCount: parsed.resourceCount,
          statusSummary: parsed.statusSummary,
          completionPct: parsed.completionPct,
          modelers: parsed.modelers,
          notes: parsed.notes,
          projectInCharge: parsed.projectInCharge,
          projectLead: parsed.projectLead,
          bimCoordinator: parsed.bimCoordinator,
          projectDc: parsed.projectDc,
          procurement: parsed.procurement,
          estimation: parsed.estimation,
          plannedStartDate: parsed.plannedStartDate,
          plannedEndDate: parsed.plannedEndDate,
          actualStartDate: parsed.actualStartDate,
          presentDate: parsed.presentDate,
          periodOfExtension: parsed.periodOfExtension,
          drawings: parsed.drawings
        };

        currentProjectKey = pKey;
        currentDateKey = dateKey;
        loadedCount++;
      } catch (err) {
        console.error('Error parsing file:', file.name, err);
      }
    }

    saveData();

    if (statusMsg) {
      statusMsg.textContent = `Successfully updated ${loadedCount} project report(s)!`;
      statusMsg.style.color = '#16A34A';
    }

    setTimeout(() => {
      closeModal('upload-modal');
      closeModal('add-project-modal');
      if (activeView === 'overview') {
        renderOverview();
      } else {
        renderProjectDetail();
      }
    }, 1000);
  }

  // --- Modal Helpers ---

  function openModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.add('active');
  }

  function closeModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.remove('active');
  }

  // --- Export to CSV ---
  function exportDrawingsToCsv() {
    const rawProj = projects[currentProjectKey];
    if (!rawProj) return;
    const proj = getProjectReport(rawProj, currentDateKey);
    const drawings = proj.drawings || [];

    let csvContent = 'data:text/csv;charset=utf-8,Row,Drawing Number,Status\r\n';
    drawings.forEach((d, idx) => {
      csvContent += `${idx + 1},"${d.drawingNo}","${d.status}"\r\n`;
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `${proj.displayName || 'project'}_Drawings_${proj.dateKey || 'export'}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // --- Event Listeners & Bootstrapping ---

  function bindEvents() {
    document.getElementById('back-to-overview')?.addEventListener('click', (e) => {
      e.preventDefault();
      navigateToOverview();
    });

    document.getElementById('search-projects')?.addEventListener('input', () => {
      renderOverview();
    });

    document.getElementById('btn-refresh-data')?.addEventListener('click', () => openModal('upload-modal'));
    document.getElementById('btn-upload-more')?.addEventListener('click', () => openModal('upload-modal'));
    document.getElementById('modal-close-btn')?.addEventListener('click', () => closeModal('upload-modal'));
    document.getElementById('modal-cancel-btn')?.addEventListener('click', () => closeModal('upload-modal'));

    document.getElementById('btn-add-project')?.addEventListener('click', () => openModal('add-project-modal'));
    document.getElementById('add-modal-close-btn')?.addEventListener('click', () => closeModal('add-project-modal'));
    document.getElementById('add-modal-cancel-btn')?.addEventListener('click', () => closeModal('add-project-modal'));

    document.getElementById('btn-confirm-add-project')?.addEventListener('click', () => {
      const num = document.getElementById('input-project-num')?.value;
      const name = document.getElementById('input-project-name')?.value;
      const pm = document.getElementById('input-project-pm')?.value;
      const dwg = document.getElementById('input-project-dwg')?.value;
      if (!num || !name) {
        alert('Please provide at least Project Number and Name.');
        return;
      }
      createManualProject(num, name, pm, dwg);
      closeModal('add-project-modal');
    });

    document.getElementById('btn-delete-current-project')?.addEventListener('click', () => {
      if (currentProjectKey) {
        deleteProject(currentProjectKey);
      }
    });

    document.querySelectorAll('.editable-heading').forEach((heading) => {
      heading.setAttribute('contenteditable', 'true');
      heading.addEventListener('blur', () => saveHeading(heading));
      heading.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          heading.blur();
        }
      });
    });

    // Date selector (Page 2): Switch between verified file date stamps
    document.getElementById('detail-date-select')?.addEventListener('change', (e) => {
      currentDateKey = e.target.value;
      renderProjectDetail();
    });

    document.getElementById('filter-drawing-status')?.addEventListener('change', (e) => {
      drawingFilterStatus = e.target.value;
      const proj = getProjectReport(projects[currentProjectKey], currentDateKey);
      renderDrawingsTable(proj);
    });

    document.getElementById('search-drawings')?.addEventListener('input', (e) => {
      drawingSearchQuery = e.target.value;
      const proj = getProjectReport(projects[currentProjectKey], currentDateKey);
      renderDrawingsTable(proj);
    });

    document.getElementById('btn-export-csv')?.addEventListener('click', exportDrawingsToCsv);

    document.getElementById('btn-print-report')?.addEventListener('click', () => {
      window.print();
    });

    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    if (dropZone && fileInput) {
      dropZone.addEventListener('click', () => fileInput.click());

      fileInput.addEventListener('change', (e) => {
        handleFilesUpload(e.target.files);
      });

      ['dragenter', 'dragover'].forEach((eventName) => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          dropZone.classList.add('dragover');
        });
      });

      ['dragleave', 'drop'].forEach((eventName) => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          dropZone.classList.remove('dragover');
        });
      });

      dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFilesUpload(files);
      });
    }

    window.addEventListener('hashchange', handleHashRouting);
  }

  function handleHashRouting() {
    const hash = window.location.hash || '#overview';
    if (hash.startsWith('#detail')) {
      const match = hash.match(/project=([^&]+)/);
      if (match) {
        navigateToDetail(decodeURIComponent(match[1]));
      } else {
        navigateToDetail(currentProjectKey);
      }
    } else {
      navigateToOverview();
    }
  }

  // --- Bootstrap ---
  window.addEventListener('DOMContentLoaded', () => {
    // Check for read-only review mode query parameter: ?mode=review or ?review=true
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('mode') === 'review' || urlParams.get('review') === 'true') {
      document.body.classList.add('review-mode');
    }
    initData();
    bindEvents();
    handleHashRouting();
  });
})();
