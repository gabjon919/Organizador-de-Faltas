/* global XLSX */

(function () {
  "use strict";

  const ABSENCE_COLUMNS = [47, 54, 60];
  const DAY_PATTERN = /^\d{2}\/\d{2}\s/;

  const state = {
    file: null,
    employees: [],
    workbook: null,
  };

  const elements = {
    uploadZone: document.getElementById("uploadZone"),
    fileInput: document.getElementById("fileInput"),
    browseButton: document.getElementById("browseButton"),
    changeFileButton: document.getElementById("changeFileButton"),
    uploadTitle: document.getElementById("uploadTitle"),
    uploadHint: document.getElementById("uploadHint"),
    fileSummary: document.getElementById("fileSummary"),
    fileName: document.getElementById("fileName"),
    fileDetails: document.getElementById("fileDetails"),
    generateButton: document.getElementById("generateButton"),
    downloadButton: document.getElementById("downloadButton"),
    alert: document.getElementById("alert"),
    results: document.getElementById("results"),
    employeeCount: document.getElementById("employeeCount"),
    departmentCount: document.getElementById("departmentCount"),
    dayCount: document.getElementById("dayCount"),
    hourCount: document.getElementById("hourCount"),
    tableDescription: document.getElementById("tableDescription"),
    previewBody: document.getElementById("previewBody"),
    emptyState: document.getElementById("emptyState"),
    searchInput: document.getElementById("searchInput"),
  };

  function clean(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function parseCSV(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let quoted = false;

    for (let i = 0; i < text.length; i += 1) {
      const character = text[i];
      const next = text[i + 1];

      if (character === '"') {
        if (quoted && next === '"') {
          cell += '"';
          i += 1;
        } else {
          quoted = !quoted;
        }
      } else if (character === "," && !quoted) {
        row.push(cell);
        cell = "";
      } else if ((character === "\n" || character === "\r") && !quoted) {
        if (character === "\r" && next === "\n") i += 1;
        row.push(cell);
        if (row.some((value) => clean(value) !== "")) rows.push(row);
        row = [];
        cell = "";
      } else {
        cell += character;
      }
    }

    if (cell || row.length) {
      row.push(cell);
      if (row.some((value) => clean(value) !== "")) rows.push(row);
    }

    return rows;
  }

  function firstValue(row, start, end) {
    return row.slice(start, end).map(clean).find(Boolean) || "";
  }

  function toMinutes(value) {
    const match = clean(value).match(/^(\d{1,3}):(\d{2})$/);
    if (!match) return 0;
    return Number(match[1]) * 60 + Number(match[2]);
  }

  function formatHours(minutes) {
    const safeMinutes = Math.max(0, Number(minutes) || 0);
    const hours = Math.floor(safeMinutes / 60);
    const remainingMinutes = safeMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(remainingMinutes).padStart(2, "0")}`;
  }

  function parseEmployees(rows) {
    const employees = [];

    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      if (clean(row[0]) !== "Funcionário:") continue;

      const employee = {
        name: firstValue(row, 1, 48),
        department: "Sem departamento",
        role: "Não informado",
        daysToDiscount: 0,
        absenceMinutes: 0,
        hasData: false,
      };

      for (let headerIndex = index + 1; headerIndex < Math.min(index + 4, rows.length); headerIndex += 1) {
        const headerRow = rows[headerIndex];
        if (clean(headerRow[0]) === "Departamento:") {
          employee.department = firstValue(headerRow, 1, 48) || "Sem departamento";
          employee.role = firstValue(headerRow, 49, 73) || "Não informado";
        }
      }

      for (let dataIndex = index + 1; dataIndex < rows.length; dataIndex += 1) {
        const dataRow = rows[dataIndex];
        const firstCell = clean(dataRow[0]);
        if (firstCell === "Funcionário:") break;

        if (DAY_PATTERN.test(firstCell)) {
          const observation = clean(dataRow[68]);
          const normalizedObservation = observation.toLocaleLowerCase("pt-BR");
          const isFullAbsence = normalizedObservation === "falta";
          const hasAbsenceInObservation = normalizedObservation.includes("falta");
          if (isFullAbsence) {
            employee.daysToDiscount += 1;
          }
          if (!hasAbsenceInObservation) {
            employee.absenceMinutes += ABSENCE_COLUMNS.reduce(
              (total, columnIndex) => total + toMinutes(dataRow[columnIndex]),
              0,
            );
          }
          employee.hasData = true;
        }
      }

      if (employee.name && employee.hasData) employees.push(employee);
    }

    return employees;
  }

  function uniqueSheetName(department, usedNames) {
    const invalidCharacters = /[\\/?*[\]:]/g;
    const base = (clean(department).replace(invalidCharacters, "-") || "Sem departamento").slice(0, 31);
    let name = base;
    let suffix = 2;

    while (usedNames.has(name)) {
      const ending = ` (${suffix})`;
      name = `${base.slice(0, 31 - ending.length)}${ending}`;
      suffix += 1;
    }

    usedNames.add(name);
    return name;
  }

  function groupByDepartment(employees) {
    return employees.reduce((groups, employee) => {
      const department = employee.department || "Sem departamento";
      if (!groups[department]) groups[department] = [];
      groups[department].push(employee);
      return groups;
    }, {});
  }

  function escapeHTML(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function renderPreview(employees) {
    const query = clean(elements.searchInput.value).toLocaleLowerCase("pt-BR");
    const filtered = employees.filter((employee) => {
      const searchable = `${employee.name} ${employee.role} ${employee.department}`.toLocaleLowerCase("pt-BR");
      return searchable.includes(query);
    });

    elements.previewBody.innerHTML = filtered
      .map(
        (employee) => `
          <tr>
            <td>${escapeHTML(employee.name)}</td>
            <td>${escapeHTML(employee.role)}</td>
            <td>${escapeHTML(employee.department)}</td>
            <td class="numeric">${employee.daysToDiscount}</td>
            <td class="numeric hours-cell">${formatHours(employee.absenceMinutes)}</td>
          </tr>
        `,
      )
      .join("");

    elements.emptyState.classList.toggle("hidden", filtered.length !== 0);
  }

  function buildWorkbook(employees) {
    const workbook = XLSX.utils.book_new();
    const groups = groupByDepartment(employees);
    const usedSheetNames = new Set();

    Object.entries(groups)
      .sort(([first], [second]) => first.localeCompare(second, "pt-BR"))
      .forEach(([department, departmentEmployees]) => {
        const sheetRows = [
          ["Funcionário", "Cargo", "Dias a descontar", "Horas a descontar"],
          ...departmentEmployees
            .slice()
            .sort((first, second) => first.name.localeCompare(second.name, "pt-BR"))
            .map((employee) => [
              employee.name,
              employee.role,
              employee.daysToDiscount,
              formatHours(employee.absenceMinutes),
            ]),
        ];

        const sheet = XLSX.utils.aoa_to_sheet(sheetRows);
        sheet["!cols"] = [
          { wch: 38 },
          { wch: 34 },
          { wch: 19 },
          { wch: 20 },
        ];
        sheet["!autofilter"] = { ref: `A1:D${sheetRows.length}` };
        XLSX.utils.book_append_sheet(workbook, sheet, uniqueSheetName(department, usedSheetNames));
      });

    return workbook;
  }

  function showAlert(message) {
    elements.alert.textContent = message;
    elements.alert.classList.remove("hidden");
  }

  function hideAlert() {
    elements.alert.classList.add("hidden");
    elements.alert.textContent = "";
  }

  function resetResults() {
    state.employees = [];
    state.workbook = null;
    elements.results.classList.add("hidden");
    elements.searchInput.value = "";
    elements.previewBody.innerHTML = "";
    hideAlert();
  }

  function updateFileUI(file) {
    state.file = file;
    elements.uploadTitle.textContent = file.name;
    elements.uploadHint.textContent = "Arquivo selecionado. Clique em gerar para continuar.";
    elements.fileName.textContent = file.name;
    elements.fileDetails.textContent = `${(file.size / 1024).toFixed(1)} KB · CSV`;
    elements.fileSummary.classList.remove("hidden");
    elements.generateButton.disabled = false;
    elements.uploadZone.classList.add("has-file");
  }

  function chooseFile(file) {
    if (!file) return;
    if (!file.name.toLocaleLowerCase("pt-BR").endsWith(".csv")) {
      showAlert("Selecione um arquivo no formato CSV.");
      return;
    }
    hideAlert();
    resetResults();
    updateFileUI(file);
  }

  function processFile() {
    if (!state.file) return;

    const reader = new FileReader();
    elements.generateButton.disabled = true;
    elements.generateButton.querySelector("span").textContent = "Lendo relatório...";
    hideAlert();

    reader.onload = () => {
      try {
        const rows = parseCSV(String(reader.result || "").replace(/^\uFEFF/, ""));
        const employees = parseEmployees(rows);
        if (!employees.length) {
          throw new Error("Não encontrei blocos de funcionários no CSV.");
        }

        state.employees = employees;
        state.workbook = buildWorkbook(employees);
        const departments = Object.keys(groupByDepartment(employees));
        const totalDays = employees.reduce((total, employee) => total + employee.daysToDiscount, 0);
        const totalMinutes = employees.reduce((total, employee) => total + employee.absenceMinutes, 0);

        elements.employeeCount.textContent = employees.length.toLocaleString("pt-BR");
        elements.departmentCount.textContent = departments.length.toLocaleString("pt-BR");
        elements.dayCount.textContent = totalDays.toLocaleString("pt-BR");
        elements.hourCount.textContent = formatHours(totalMinutes);
        elements.tableDescription.textContent = `${employees.length} funcionário${employees.length === 1 ? "" : "s"} · ${departments.length} departamento${departments.length === 1 ? "" : "s"}`;
        renderPreview(employees);
        elements.results.classList.remove("hidden");
        elements.generateButton.querySelector("span").textContent = "Gerar planilha XLSX";
        elements.generateButton.disabled = false;
        elements.results.scrollIntoView({ behavior: "smooth", block: "start" });
      } catch (error) {
        showAlert(error instanceof Error ? error.message : "Não foi possível ler o arquivo.");
        elements.generateButton.querySelector("span").textContent = "Gerar planilha XLSX";
        elements.generateButton.disabled = false;
      }
    };

    reader.onerror = () => {
      showAlert("Não foi possível abrir o CSV. Tente selecionar o arquivo novamente.");
      elements.generateButton.querySelector("span").textContent = "Gerar planilha XLSX";
      elements.generateButton.disabled = false;
    };

    reader.readAsText(state.file, "UTF-8");
  }

  function downloadWorkbook() {
    if (!state.workbook || typeof XLSX === "undefined") {
      showAlert("A biblioteca de exportação ainda não carregou. Atualize a página e tente novamente.");
      return;
    }

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(state.workbook, `descontos_frequencia_${date}.xlsx`);
  }

  elements.uploadZone.addEventListener("click", (event) => {
    if (event.target !== elements.browseButton) elements.fileInput.click();
  });
  elements.uploadZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      elements.fileInput.click();
    }
  });
  elements.browseButton.addEventListener("click", () => elements.fileInput.click());
  elements.changeFileButton.addEventListener("click", () => elements.fileInput.click());
  elements.fileInput.addEventListener("change", () => chooseFile(elements.fileInput.files[0]));
  elements.generateButton.addEventListener("click", processFile);
  elements.downloadButton.addEventListener("click", downloadWorkbook);
  elements.searchInput.addEventListener("input", () => renderPreview(state.employees));

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.uploadZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.uploadZone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.uploadZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.uploadZone.classList.remove("dragover");
    });
  });
  elements.uploadZone.addEventListener("drop", (event) => {
    chooseFile(event.dataTransfer.files[0]);
  });
})();
