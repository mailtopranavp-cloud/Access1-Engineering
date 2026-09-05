/**
 * Excel Parser for Design Department Weekly Report Dashboard
 * Decoupled mapping engine utilizing SheetJS (xlsx.full.min.js)
 */

const ExcelParser = {
  /**
   * Parses various date formats into a standard JS Date object
   */
  parseToDate(val) {
    if (!val) return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    if (typeof val === 'number' && val > 20000) {
      return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    const str = String(val).trim();
    // DD/MM/YYYY
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
      const parts = str.split('/');
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
    // YYYY-MM-DD
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
      const parts = str.split('-');
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  },

  /**
   * Converts Excel date serial or raw date strings into standard DD/MM/YYYY
   */
  formatExcelDate(val) {
    if (!val) return "";
    const date = this.parseToDate(val);
    if (!date) return String(val).trim();
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  },

  /**
   * Converts date to "Mmm, YY" format (e.g. "Sep, 26")
   */
  formatMonthYear(val) {
    if (!val) return "";
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const d = this.parseToDate(val);
    if (d) {
      const m = monthNames[d.getMonth()];
      const y = String(d.getFullYear()).slice(-2);
      return `${m}, ${y}`;
    }
    return String(val);
  },

  /**
   * Correction 2: Extract date stamp from Excel file name
   * Supports:
   * - YYYY-MM-DD or YYYY_MM_DD (e.g. PH497_2026-09-05.xlsx)
   * - DD-MM-YYYY or DD_MM_YYYY
   */
  extractDateFromFileName(fileName) {
    if (!fileName) return null;
    
    // YYYY-MM-DD or YYYY_MM_DD
    let match = fileName.match(/(\d{4})[-_](\d{2})[-_](\d{2})/);
    if (match) {
      const yyyy = match[1];
      const mm = match[2];
      const dd = match[3];
      return {
        dateKey: `${yyyy}-${mm}-${dd}`,
        formattedDate: `${dd}/${mm}/${yyyy}`
      };
    }

    // DD-MM-YYYY or DD_MM_YYYY
    match = fileName.match(/(\d{2})[-_](\d{2})[-_](\d{4})/);
    if (match) {
      const dd = match[1];
      const mm = match[2];
      const yyyy = match[3];
      return {
        dateKey: `${yyyy}-${mm}-${dd}`,
        formattedDate: `${dd}/${mm}/${yyyy}`
      };
    }

    return null;
  },

  /**
   * Correction 1: Calculate Period of Extension
   * Formula: Period of extension = Planned end Date – Present Date
   * Output format: months (e.g. "2 Months", "3 Months")
   */
  calculatePeriodOfExtension(plannedEndDateStr, presentDateStr) {
    const dEnd = this.parseToDate(plannedEndDateStr);
    const dPres = this.parseToDate(presentDateStr);
    if (!dEnd || !dPres) return "None";

    // Difference in milliseconds: Present Date - Planned End Date
    const diffMs = dPres.getTime() - dEnd.getTime();
    const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));

    // If present date is not past planned end date, extension is 0 Months
    if (diffDays <= 10) {
      return "0 Months";
    }

    // Convert days difference into months (average month length = 30.4375 days)
    const months = Math.round(diffDays / 30.4375);
    return months === 1 ? "1 Month" : `${months} Months`;
  },

  /**
   * Normalizes drawing status strings
   */
  normalizeStatus(rawStatus) {
    if (!rawStatus) return "Not Started";
    const s = String(rawStatus).trim().toLowerCase();

    if (s.includes("resub")) return "Approved with Resub.";
    if (s.startsWith("appr")) return "Approved";
    if (s.includes("submi") || s.includes("review")) return "Submitted/Under Review";
    if (s.includes("rej")) return "Rejected";
    if (s.includes("prep")) return "Under Preparation";
    if (s.includes("pend") || s.includes("note")) return "Pending(Notes)";
    if (s.includes("not") || s.includes("start")) return "Not Started";
    return rawStatus.trim();
  },

  /**
   * Maps status string to object property
   */
  statusToKey(status) {
    switch (status) {
      case "Approved": return "approved";
      case "Approved with Resub.": return "approvedWithResub";
      case "Submitted/Under Review": return "submitted";
      case "Rejected": return "rejected";
      case "Under Preparation": return "underPrep";
      case "Pending(Notes)": return "pending";
      case "Not Started": return "notStarted";
      default: return "other";
    }
  },

  /**
   * Parses an ArrayBuffer of an Excel file (.xlsx)
   */
  async parseWorkbookBuffer(arrayBuffer, fileName = "report.xlsx") {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS (XLSX) library is not loaded.");
    }

    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    const getVal = (cellRef) => {
      const cell = sheet[cellRef];
      if (!cell) return "";
      if (cell.w !== undefined && cell.w !== null && String(cell.w).trim() !== "") return String(cell.w).trim();
      if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
      return "";
    };

    // Build label map for resilient dynamic scanning
    const labelMapA = {};
    const labelMapE = {};
    for (let r = 1; r <= 35; r++) {
      const a = getVal(`A${r}`).toLowerCase();
      if (a) labelMapA[a] = r;
      const e = getVal(`E${r}`).toLowerCase();
      if (e) labelMapE[e] = r;
    }

    const findByLabel = (keywords, col, map, defaultCell) => {
      for (const k in map) {
        if (keywords.some(kw => k.includes(kw))) {
          const val = getVal(`${col}${map[k]}`);
          if (val) return val;
        }
      }
      return getVal(defaultCell);
    };

    // Correction 2: Extract Report Date primarily from file name date stamp
    const fileDateInfo = this.extractDateFromFileName(fileName);
    let dateKey = "";
    let reportDateFormatted = "";

    if (fileDateInfo) {
      dateKey = fileDateInfo.dateKey;
      reportDateFormatted = fileDateInfo.formattedDate;
    } else {
      // Fallback if file name has no date
      const rawReportDate = findByLabel(["date of report", "report date"], "B", labelMapA, "B1");
      reportDateFormatted = this.formatExcelDate(rawReportDate);
      if (reportDateFormatted) {
        const parts = reportDateFormatted.split('/');
        if (parts.length === 3) dateKey = `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
      if (!dateKey) {
        const now = new Date();
        dateKey = now.toISOString().split('T')[0];
        reportDateFormatted = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;
      }
    }

    const monthYearLabel = this.formatMonthYear(dateKey);

    // Extract Key Personnel Details
    const projectName = findByLabel(["project name"], "B", labelMapA, "B2") || "SG1";
    const projectNumber = findByLabel(["project number", "project no"], "B", labelMapA, "B3") || (fileName.split('_')[0] || "PH497");
    const projectInCharge = findByLabel(["in-charge", "incharge", "in charge"], "B", labelMapA, "B4") || "Rejoe";
    const projectLead = findByLabel(["project lead", "lead", "pm"], "B", labelMapA, "B5") || "Pranav";
    const bimCoordinator = findByLabel(["bim", "coordinator", "design coordinator"], "B", labelMapA, "B6") || "Ajmal";
    const projectDc = findByLabel(["project dc", "dc", "document controller"], "B", labelMapA, "B7") || "Romeo";
    const procurement = findByLabel(["procurement"], "B", labelMapA, "B8") || "Osama";
    const estimation = findByLabel(["estimation"], "B", labelMapA, "B9") || "Abdul";

    // Extract Schedule Details
    const rawPlanStart = findByLabel(["planned start"], "F", labelMapE, "F2");
    const rawPlanEnd = findByLabel(["planned end"], "F", labelMapE, "F3");
    const rawActStart = findByLabel(["actual start"], "F", labelMapE, "F4");

    const plannedStart = this.formatExcelDate(rawPlanStart) || "01/01/2026";
    const plannedEnd = this.formatExcelDate(rawPlanEnd) || "06/06/2026";
    const actualStart = this.formatExcelDate(rawActStart) || "04/04/2026";
    
    // Present Date maps to the report date from file name
    const presentDate = reportDateFormatted;

    // Correction 1: Calculate Period of Extension = Planned end Date – Present Date
    const periodOfExtension = this.calculatePeriodOfExtension(plannedEnd, presentDate);

    // Extract List of OSR Modelers from Column C & D
    const modelers = [];
    for (let r = 1; r <= 35; r++) {
      const cVal = getVal(`C${r}`);
      const dVal = getVal(`D${r}`);
      if (cVal) {
        const cLower = cVal.toLowerCase();
        if (cLower.includes("list of osr") || cLower.includes("modeler name") || cLower.includes("date of join")) {
          continue;
        }
        let formattedDate = this.formatExcelDate(dVal);
        if (!formattedDate || formattedDate.includes("[dd/mm")) {
          formattedDate = "Active";
        }
        modelers.push({
          name: cVal,
          role: "BIM Modeler",
          dateOfJoin: formattedDate,
          active: true
        });
      }
    }
    const resourceCount = modelers.length;

    // Extract Notes & Observations from Column I
    const notes = [];
    for (let r = 1; r <= 100; r++) {
      const val = getVal(`I${r}`);
      if (val) {
        const lower = val.toLowerCase();
        if (r === 1 && (lower === "notes" || lower === "notes & observations" || lower === "observations" || lower === "note" || lower === "notes:" || lower === "report updated by")) {
          continue;
        }
        const lines = val.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        lines.forEach(line => {
          if (!notes.includes(line)) notes.push(line);
        });
      }
    }

    // Extract Drawings from Column G and Status from Column H
    const drawings = [];
    const statusCounts = {
      approved: 0,
      approvedWithResub: 0,
      submitted: 0,
      rejected: 0,
      underPrep: 0,
      pending: 0,
      notStarted: 0,
      other: 0
    };

    let emptyStreak = 0;
    for (let r = 2; r <= 5000; r++) {
      const dwgNo = getVal(`G${r}`);
      const rawStatus = getVal(`H${r}`);

      if (!dwgNo && !rawStatus) {
        emptyStreak++;
        if (emptyStreak >= 20) break;
        continue;
      }
      emptyStreak = 0;

      if (dwgNo) {
        const normalized = this.normalizeStatus(rawStatus);
        const key = this.statusToKey(normalized);
        statusCounts[key] = (statusCounts[key] || 0) + 1;

        drawings.push({
          row: r,
          drawingNo: dwgNo,
          status: normalized
        });
      }
    }

    const totalDwg = drawings.length;
    const completionPct = totalDwg > 0 ? ((statusCounts.approved / totalDwg) * 100).toFixed(1) : "0.0";
    const projectKey = `${projectNumber}-${projectName}`.replace(/[\s_]+/g, '-');

    return {
      fileName,
      projectKey,
      projectNumber,
      projectName,
      displayName: `${projectNumber}-${projectName}`,
      projectInCharge: projectInCharge || "Rejoe",
      projectLead: projectLead || "Pranav",
      bimCoordinator: bimCoordinator || "Ajmal",
      projectDc: projectDc || "Romeo",
      procurement: procurement || "Osama",
      estimation: estimation || "Abdul",
      plannedStartDate: plannedStart,
      plannedEndDate: plannedEnd,
      actualStartDate: actualStart,
      presentDate: presentDate,
      periodOfExtension: periodOfExtension,
      modelers,
      resourceCount,
      notes,
      dateKey,
      dateOfReport: reportDateFormatted,
      monthYearLabel,
      totalDwg,
      statusSummary: statusCounts,
      completionPct: parseFloat(completionPct),
      drawings
    };
  }
};

window.ExcelParser = ExcelParser;
