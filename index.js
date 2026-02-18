const INGENIEROS = [
  "ROGERD FLORES",
  "ALEX BERROCAL",
  "DIEGO PRADA",
  "ANGEL HUAMANI",
  "MAURO DIAZ",
  "JAIRO PALACIOS",
  "WILTON MAZA",
  "JOSUE HUAMAN",
  "RICARDO ESTUPIÑAN",
  "ANTONY SAAVEDRA",
  "JOHN HUAMANI",
  "ROBERT OTINIANO",
  "MARLON TORRES",
  "WILLIAM RODRIGUEZ",
  "ROBERT AMADOR",
  "KEVIN PALERMO"
];

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear();
}

function fillEngineers(select, selectedName) {
  if (!select) return;
  select.innerHTML = "";
  INGENIEROS.forEach(function (name) {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    if (name === selectedName) option.selected = true;
    select.appendChild(option);
  });
}

const btnManana = document.getElementById("btn-manana");
const btnNoche = document.getElementById("btn-noche");
const shiftRange = document.getElementById("shift-range");
const dateTag = document.getElementById("current-date-tag");
const horaInicioFull = document.getElementById("hora-inicio-full");
const horaCierreFull = document.getElementById("hora-cierre-full");
const ingenieroSaliente = document.getElementById("ingeniero-saliente");
const ingenieroEntrante = document.getElementById("ingeniero-entrante");
const btnActualizarReporte = document.getElementById("btn-actualizar-reporte");
const btnAbrirEnviar = document.getElementById("btn-abrir-enviar");
const btnCerrarEnviar = document.getElementById("btn-cerrar-enviar");
const btnConfirmarEnviar = document.getElementById("btn-confirmar-enviar");
const excelFileInput = document.getElementById("excel-file-input");
const mailModal = document.getElementById("mail-modal");
const mailPara = document.getElementById("mail-para");
const mailCc = document.getElementById("mail-cc");
const countProcesoN1 = document.getElementById("count-proceso-n1");
const countProcesoN0 = document.getElementById("count-proceso-n0");
const countFinalizadosN1 = document.getElementById("count-finalizados-n1");
const countFinalizadosN0 = document.getElementById("count-finalizados-n0");
const countDesestimadosN1 = document.getElementById("count-desestimados-n1");
const countDesestimadosN0 = document.getElementById("count-desestimados-n0");
let turnoSeleccionado = null;
let horaInicioManual = false;
let horaCierreManual = false;
let procesoObservers = [];
let finalizadosObservers = [];
let desestimadosObservers = [];

fillEngineers(ingenieroSaliente, " - ");
fillEngineers(ingenieroEntrante, " - ");

function applyShift(turno, currentDate) {
  const dateText = formatDate(currentDate || new Date());
  if (turno === "noche") {
    if (btnNoche) btnNoche.classList.add("active");
    if (btnManana) btnManana.classList.remove("active");
    if (shiftRange) shiftRange.textContent = "Desde 18:00 hasta 06:00";
    if (horaInicioFull && !horaInicioManual) horaInicioFull.value = dateText + " 18:00";
    if (horaCierreFull && !horaCierreManual) horaCierreFull.value = dateText + " 06:00";
  } else {
    if (btnManana) btnManana.classList.add("active");
    if (btnNoche) btnNoche.classList.remove("active");
    if (shiftRange) shiftRange.textContent = "Desde 06:00 hasta 18:00";
    if (horaInicioFull && !horaInicioManual) horaInicioFull.value = dateText + " 06:00";
    if (horaCierreFull && !horaCierreManual) horaCierreFull.value = dateText + " 18:00";
  }
}

if (btnManana) {
  btnManana.addEventListener("click", function () {
    turnoSeleccionado = "manana";
    applyShift("manana", new Date());
  });
}
if (btnNoche) {
  btnNoche.addEventListener("click", function () {
    turnoSeleccionado = "noche";
    applyShift("noche", new Date());
  });
}
if (horaInicioFull) {
  horaInicioFull.addEventListener("input", function () {
    horaInicioManual = true;
  });
}
if (horaCierreFull) {
  horaCierreFull.addEventListener("input", function () {
    horaCierreManual = true;
  });
}

function refreshNow() {
  const now = new Date();
  if (dateTag) dateTag.textContent = "[" + formatDate(now) + "]";

  if (!turnoSeleccionado) {
    const hour = now.getHours();
    turnoSeleccionado = hour >= 18 || hour < 6 ? "noche" : "manana";
  }
  applyShift(turnoSeleccionado, now);
}

function refreshAtNextMinute() {
  const now = new Date();
  const delay = (60 - now.getSeconds()) * 1000 - now.getMilliseconds();
  setTimeout(function () {
    refreshNow();
    setInterval(refreshNow, 60000);
  }, Math.max(delay, 0));
}

function enableEditableEstadoActual() {
  const styleId = "editable-estado-actual-style";
  if (!document.getElementById(styleId)) {
    const style = document.createElement("style");
    style.id = styleId;
    style.textContent =
      ".editable-estado-actual{background:#fff;border:1px solid #b9b9b9;padding:6px 8px;min-height:32px;outline:none;}" +
      ".editable-estado-actual:focus{border-color:#c4221a;box-shadow:0 0 0 1px #c4221a inset;}";
    document.head.appendChild(style);
  }

  const rows = document.querySelectorAll("tr.detail-row");
  rows.forEach(function (row) {
    const labelCell = row.querySelector("td.detail-label");
    const textCell = row.querySelector("td.detail-text");
    if (!labelCell || !textCell) return;
    if (labelCell.textContent.trim().toLowerCase() !== "estado actual") return;

    textCell.contentEditable = "true";
    textCell.spellcheck = true;
    textCell.classList.add("editable-estado-actual");
    textCell.setAttribute("title", "Puedes editar este estado manualmente");
  });
}

function createEmptyTable(tableClass, ariaLabel, headers) {
  const table = document.createElement("table");
  table.className = "incident-table " + tableClass;
  table.setAttribute("aria-label", ariaLabel);
  const thead = document.createElement("thead");
  const tr = document.createElement("tr");
  headers.forEach(function (h) {
    const th = document.createElement("th");
    if (h.width) th.style.width = h.width;
    th.textContent = h.text;
    tr.appendChild(th);
  });
  thead.appendChild(tr);
  table.appendChild(thead);
  table.appendChild(document.createElement("tbody"));
  return table;
}

function clearSectionsOnLoad() {
  const sectionProceso = document.querySelector("section.rojo");
  const sectionFinal = document.querySelector("section.verde");
  const sectionDesestimado = document.querySelector("section.morado");

  if (sectionProceso) {
    sectionProceso.querySelectorAll("table.tabla-proceso-incidente").forEach(function (t) { t.remove(); });
    sectionProceso.appendChild(
      createEmptyTable("tabla-proceso-incidente", "Incidentes masivos en proceso", [
        { text: "Aplicación", width: "16%" },
        { text: "Ticket", width: "12%" },
        { text: "Fuente", width: "12%" },
        { text: "Nivel", width: "10%" },
        { text: "Inicio del Incidente (Fecha y Hora)", width: "16%" },
        { text: "Situation Manager", width: "16%" },
        { text: "Estado", width: "10%" },
        { text: "Ultimo Correo de Gestión", width: "8%" }
      ])
    );
  }

  if (sectionFinal) {
    sectionFinal.querySelectorAll("table.tabla-finalizados-incidente").forEach(function (t) { t.remove(); });
    sectionFinal.appendChild(
      createEmptyTable("tabla-finalizados-incidente", "Incidentes masivos finalizados", [
        { text: "Aplicación", width: "16%" },
        { text: "Ticket", width: "12%" },
        { text: "Fuente", width: "14%" },
        { text: "Nivel", width: "8%" },
        { text: "Inicio del Incidente (Fecha y Hora)", width: "14%" },
        { text: "Fin del Incidente (Fecha y Hora)", width: "14%" },
        { text: "Situation Manager", width: "14%" }
      ])
    );
  }

  if (sectionDesestimado) {
    sectionDesestimado.querySelectorAll("table.tabla-desestimados-incidente").forEach(function (t) { t.remove(); });
    sectionDesestimado.appendChild(
      createEmptyTable("tabla-desestimados-incidente", "Incidentes desestimados", [
        { text: "Aplicación", width: "16%" },
        { text: "Ticket", width: "12%" },
        { text: "Fuente", width: "14%" },
        { text: "Nivel", width: "8%" },
        { text: "Inicio del Incidente (Fecha y Hora)", width: "14%" },
        { text: "Fin del Incidente (Fecha y Hora)", width: "14%" },
        { text: "Situation Manager", width: "14%" }
      ])
    );
  }
}

function enableEditableProcesoMainFields() {
  const styleId = "editable-proceso-main-style";
  if (!document.getElementById(styleId)) {
    const style = document.createElement("style");
    style.id = styleId;
    style.textContent =
      ".editable-proceso-main{background:#fff;border:1px solid #b9b9b9;padding:6px 8px;min-height:28px;outline:none;}" +
      ".editable-proceso-main:focus{border-color:#c4221a;box-shadow:0 0 0 1px #c4221a inset;}";
    document.head.appendChild(style);
  }

  const rows = document.querySelectorAll("section.rojo table.tabla-proceso-incidente tbody tr:not(.detail-row)");
  rows.forEach(function (row) {
    const estadoCell = row.querySelector("td:nth-child(7)");
    const correoCell = row.querySelector("td:nth-child(8)");
    [estadoCell, correoCell].forEach(function (cell) {
      if (!cell) return;
      cell.contentEditable = "true";
      cell.spellcheck = true;
      cell.classList.add("editable-proceso-main");
      cell.setAttribute("title", "Campo editable manualmente");
    });
  });
}

function disconnectObservers(list) {
  list.forEach(function (obs) { obs.disconnect(); });
  return [];
}

function normalizeKey(key) {
  return String(key || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function pickValue(row, candidateKeys) {
  const entries = Object.entries(row);
  for (let i = 0; i < entries.length; i += 1) {
    const k = normalizeKey(entries[i][0]);
    for (let j = 0; j < candidateKeys.length; j += 1) {
      if (k.includes(candidateKeys[j])) {
        const value = String(entries[i][1] || "").trim();
        if (value) return value;
      }
    }
  }
  return "";
}

function normalizeNivel(raw) {
  const value = String(raw || "").toUpperCase().trim();
  if (value.includes("N1") || value === "1") return "N1";
  if (value.includes("N0") || value === "0") return "N0";
  return value || "N0";
}

function formatDateTimeFromDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return (
    pad2(date.getUTCDate()) + "/" +
    pad2(date.getUTCMonth() + 1) + "/" +
    date.getUTCFullYear() + " " +
    pad2(date.getUTCHours()) + ":" +
    pad2(date.getUTCMinutes())
  );
}

function formatDateTimeLocal(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return (
    pad2(date.getDate()) + "/" +
    pad2(date.getMonth() + 1) + "/" +
    date.getFullYear() + " " +
    pad2(date.getHours()) + ":" +
    pad2(date.getMinutes())
  );
}

function normalizeHumanDateTime(text) {
  const value = String(text || "").trim();
  if (!value) return "";

  const m1 = value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2}))?$/);
  if (m1) {
    const dd = pad2(Number(m1[1]));
    const mm = pad2(Number(m1[2]));
    const yyyy = m1[3].length === 2 ? "20" + m1[3] : m1[3];
    const hh = pad2(Number(m1[4] || 0));
    const min = pad2(Number(m1[5] || 0));
    return dd + "/" + mm + "/" + yyyy + " " + hh + ":" + min;
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    return formatDateTimeLocal(parsed);
  }

  return value;
}

function normalizeExcelDateTime(raw) {
  if (raw === null || raw === undefined) return "";
  if (raw instanceof Date) return formatDateTimeLocal(raw);

  const text = String(raw).trim();
  if (!text) return "";

  const numeric = typeof raw === "number" ? raw : Number(text);
  const isExcelSerial = Number.isFinite(numeric) && numeric > 20000 && numeric < 100000;
  if (!isExcelSerial) return normalizeHumanDateTime(text);

  if (window.XLSX && window.XLSX.SSF) {
    return normalizeHumanDateTime(window.XLSX.SSF.format("dd/mm/yyyy hh:mm", numeric));
  }

  const excelEpochUtc = Date.UTC(1899, 11, 30);
  const date = new Date(excelEpochUtc + numeric * 86400000);
  return formatDateTimeFromDate(date);
}

function classifyEstado(raw, finValue, rawNivel) {
  const estado = String(raw || "").toLowerCase();
  const nivel = String(rawNivel || "").toLowerCase();
  if (nivel.includes("desestim")) return "desestimados";
  if (estado.includes("desestim")) return "desestimados";
  if (estado.includes("superad") || estado.includes("finaliz") || estado.includes("cerrad")) return "finalizados";
  if (estado.includes("proceso")) return "proceso";
  if (String(finValue || "").trim()) return "finalizados";
  return "proceso";
}

function mapExcelRowsToIncidents(rows) {
  const out = {
    proceso: [],
    finalizados: [],
    desestimados: []
  };

  rows.forEach(function (row) {
    const aplicacion = pickValue(row, ["aplicacion", "sistema", "servicio"]);
    const ticket = pickValue(row, ["ticket", "incidente"]);
    if (!aplicacion && !ticket) return;

    const fuente = pickValue(row, ["fuente", "origen", "equipo"]);
    const rawNivel = pickValue(row, ["nivel"]);
    const nivel = normalizeNivel(rawNivel);
    const inicio = normalizeExcelDateTime(pickValue(row, ["inicio", "fecha inicio"]));
    const fin = normalizeExcelDateTime(pickValue(row, ["fin", "fecha fin", "superado"]));
    const manager = pickValue(row, ["situation manager", "manager", "responsable"]);
    const estado = pickValue(row, ["estado", "status", "estatus"]);
    const ultimoCorreo = normalizeExcelDateTime(pickValue(row, ["ultimo correo", "ultimo", "correo"]));
    const sintoma = pickValue(row, ["sintoma"]);
    const estadoActual = pickValue(row, ["estado actual"]);
    const causa = pickValue(row, ["causa"]);
    const solucion = pickValue(row, ["solucion"]);

    const incident = {
      aplicacion: aplicacion || "-",
      ticket: ticket || "-",
      fuente: fuente || "-",
      nivel: nivel || "N0",
      inicio: inicio || "-",
      fin: fin || "",
      manager: manager || "-",
      estado: (estado || "EN PROCESO").toUpperCase(),
      ultimoCorreo: ultimoCorreo || "-",
      details: []
    };

    const bucket = classifyEstado(estado, fin, rawNivel);
    if (sintoma) incident.details.push({ label: "Síntoma", text: sintoma });
    if (bucket === "proceso") {
      incident.details.push({ label: "Estado actual", text: estadoActual || "Pendiente de actualización." });
    } else {
      if (causa) incident.details.push({ label: "Causa", text: causa });
      if (solucion) incident.details.push({ label: "Solución", text: solucion });
    }

    out[bucket].push(incident);
  });

  return out;
}

function createCell(text, colSpan) {
  const td = document.createElement("td");
  td.textContent = text || "";
  if (colSpan) td.colSpan = colSpan;
  return td;
}

function createHeaderCells(table, headers) {
  const thead = document.createElement("thead");
  const tr = document.createElement("tr");
  headers.forEach(function (h) {
    const th = document.createElement("th");
    if (h.width) th.style.width = h.width;
    th.textContent = h.text;
    tr.appendChild(th);
  });
  thead.appendChild(tr);
  table.appendChild(thead);
}

function renderProcesoSection(incidents) {
  const section = document.querySelector("section.rojo");
  if (!section) return;
  section.querySelectorAll("table.tabla-proceso-incidente").forEach(function (t) { t.remove(); });

  incidents.forEach(function (item) {
    const table = document.createElement("table");
    table.className = "incident-table tabla-proceso-incidente";
    table.setAttribute("aria-label", "Incidentes masivos en proceso");
    createHeaderCells(table, [
      { text: "Aplicación", width: "16%" },
      { text: "Ticket", width: "12%" },
      { text: "Fuente", width: "12%" },
      { text: "Nivel", width: "10%" },
      { text: "Inicio del Incidente (Fecha y Hora)", width: "16%" },
      { text: "Situation Manager", width: "16%" },
      { text: "Estado", width: "10%" },
      { text: "Ultimo Correo de Gestión", width: "8%" }
    ]);

    const tbody = document.createElement("tbody");
    const main = document.createElement("tr");
    main.appendChild(createCell(item.aplicacion));
    main.appendChild(createCell(item.ticket));
    main.appendChild(createCell(item.fuente));
    main.appendChild(createCell(item.nivel));
    main.appendChild(createCell(item.inicio));
    main.appendChild(createCell(item.manager));
    main.appendChild(createCell(item.estado));
    main.appendChild(createCell(item.ultimoCorreo));
    tbody.appendChild(main);

    item.details.forEach(function (d) {
      const tr = document.createElement("tr");
      tr.className = "detail-row";
      const label = document.createElement("td");
      label.className = "detail-label";
      label.textContent = d.label;
      const text = document.createElement("td");
      text.className = "detail-text";
      text.colSpan = 7;
      text.textContent = d.text;
      tr.appendChild(label);
      tr.appendChild(text);
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    section.appendChild(table);
  });
}

function renderFinalOrDesestimadosSection(sectionSelector, tableClass, ariaLabel, incidents) {
  const section = document.querySelector(sectionSelector);
  if (!section) return;
  section.querySelectorAll("table." + tableClass).forEach(function (t) { t.remove(); });

  if (incidents.length === 0) {
    const empty = document.createElement("table");
    empty.className = "incident-table " + tableClass;
    empty.setAttribute("aria-label", ariaLabel);
    const tbody = document.createElement("tbody");
    const tr = document.createElement("tr");
    tr.appendChild(createCell("Sin incidentes por el momento.", 7));
    tbody.appendChild(tr);
    empty.appendChild(tbody);
    section.appendChild(empty);
    return;
  }

  incidents.forEach(function (item) {
    const table = document.createElement("table");
    table.className = "incident-table " + tableClass;
    table.setAttribute("aria-label", ariaLabel);
    createHeaderCells(table, [
      { text: "Aplicación", width: "16%" },
      { text: "Ticket", width: "12%" },
      { text: "Fuente", width: "14%" },
      { text: "Nivel", width: "8%" },
      { text: "Inicio del Incidente (Fecha y Hora)", width: "14%" },
      { text: "Fin del Incidente (Fecha y Hora)", width: "14%" },
      { text: "Situation Manager", width: "14%" }
    ]);

    const tbody = document.createElement("tbody");
    const main = document.createElement("tr");
    main.appendChild(createCell(item.aplicacion));
    main.appendChild(createCell(item.ticket));
    main.appendChild(createCell(item.fuente));
    main.appendChild(createCell(item.nivel));
    main.appendChild(createCell(item.inicio));
    main.appendChild(createCell(item.fin || "-"));
    main.appendChild(createCell(item.manager));
    tbody.appendChild(main);

    item.details.forEach(function (d) {
      const tr = document.createElement("tr");
      tr.className = "detail-row";
      const label = document.createElement("td");
      label.className = "detail-label";
      label.textContent = d.label;
      const text = document.createElement("td");
      text.className = "detail-text";
      text.colSpan = 6;
      text.textContent = d.text;
      tr.appendChild(label);
      tr.appendChild(text);
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    section.appendChild(table);
  });
}

async function syncFromExcelFile(file) {
  if (!file) return;
  if (!window.XLSX) {
    alert("No se pudo cargar la librería de Excel (XLSX).");
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = window.XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  const data = mapExcelRowsToIncidents(rows);

  renderProcesoSection(data.proceso);
  renderFinalOrDesestimadosSection(
    "section.verde",
    "tabla-finalizados-incidente",
    "Incidentes masivos finalizados",
    data.finalizados
  );
  renderFinalOrDesestimadosSection(
    "section.morado",
    "tabla-desestimados-incidente",
    "Incidentes desestimados",
    data.desestimados
  );

  refreshReportNow();
  watchProcesoTableCounters();
  watchFinalizadosTableCounters();
  watchDesestimadosTableCounters();
}

function updateProcesoLevelCounters() {
  let n1 = 0;
  let n0 = 0;
  const tables = document.querySelectorAll("section.rojo table.tabla-proceso-incidente");
  if (tables.length === 0) {
    if (countProcesoN1) countProcesoN1.textContent = "0";
    if (countProcesoN0) countProcesoN0.textContent = "0";
    return;
  }

  tables.forEach(function (table) {
    const rows = table.querySelectorAll("tbody tr:not(.detail-row)");
    rows.forEach(function (row) {
      const levelCell = row.querySelector("td:nth-child(4)");
      if (!levelCell) return;
      const level = levelCell.textContent.trim().toUpperCase();
      if (level === "N1") n1 += 1;
      if (level === "N0") n0 += 1;
    });
  });

  if (countProcesoN1) countProcesoN1.textContent = String(n1);
  if (countProcesoN0) countProcesoN0.textContent = String(n0);
}

function watchProcesoTableCounters() {
  procesoObservers = disconnectObservers(procesoObservers);
  const tableBodies = document.querySelectorAll("section.rojo table.tabla-proceso-incidente tbody");
  if (tableBodies.length === 0) return;
  tableBodies.forEach(function (tableBody) {
    const observer = new MutationObserver(function () {
      updateProcesoLevelCounters();
    });
    observer.observe(tableBody, {
      childList: true,
      subtree: true,
      characterData: true
    });
    procesoObservers.push(observer);
  });
}

function updateFinalizadosLevelCounters() {
  let n1 = 0;
  let n0 = 0;
  const tables = document.querySelectorAll("section.verde table.tabla-finalizados-incidente");
  if (tables.length === 0) {
    if (countFinalizadosN1) countFinalizadosN1.textContent = "0";
    if (countFinalizadosN0) countFinalizadosN0.textContent = "0";
    return;
  }

  tables.forEach(function (table) {
    const rows = table.querySelectorAll("tbody tr:not(.detail-row)");
    rows.forEach(function (row) {
      const levelCell = row.querySelector("td:nth-child(4)");
      if (!levelCell) return;
      const level = levelCell.textContent.trim().toUpperCase();
      if (level === "N1") n1 += 1;
      if (level === "N0") n0 += 1;
    });
  });

  if (countFinalizadosN1) countFinalizadosN1.textContent = String(n1);
  if (countFinalizadosN0) countFinalizadosN0.textContent = String(n0);
}

function watchFinalizadosTableCounters() {
  finalizadosObservers = disconnectObservers(finalizadosObservers);
  const tableBodies = document.querySelectorAll("section.verde table.tabla-finalizados-incidente tbody");
  if (tableBodies.length === 0) return;
  tableBodies.forEach(function (tableBody) {
    const observer = new MutationObserver(function () {
      updateFinalizadosLevelCounters();
    });
    observer.observe(tableBody, {
      childList: true,
      subtree: true,
      characterData: true
    });
    finalizadosObservers.push(observer);
  });
}

function updateDesestimadosLevelCounters() {
  let n1 = 0;
  let n0 = 0;
  const tables = document.querySelectorAll("section.morado table.tabla-desestimados-incidente");
  if (tables.length === 0) {
    if (countDesestimadosN1) countDesestimadosN1.textContent = "0";
    if (countDesestimadosN0) countDesestimadosN0.textContent = "0";
    return;
  }

  tables.forEach(function (table) {
    const rows = table.querySelectorAll("tbody tr:not(.detail-row)");
    rows.forEach(function (row) {
      const levelCell = row.querySelector("td:nth-child(4)");
      if (!levelCell) return;
      const level = levelCell.textContent.trim().toUpperCase();
      if (level === "N1") n1 += 1;
      if (level === "N0") n0 += 1;
    });
  });

  if (countDesestimadosN1) countDesestimadosN1.textContent = String(n1);
  if (countDesestimadosN0) countDesestimadosN0.textContent = String(n0);
}

function watchDesestimadosTableCounters() {
  desestimadosObservers = disconnectObservers(desestimadosObservers);
  const tableBodies = document.querySelectorAll("section.morado table.tabla-desestimados-incidente tbody");
  if (tableBodies.length === 0) return;
  tableBodies.forEach(function (tableBody) {
    const observer = new MutationObserver(function () {
      updateDesestimadosLevelCounters();
    });
    observer.observe(tableBody, {
      childList: true,
      subtree: true,
      characterData: true
    });
    desestimadosObservers.push(observer);
  });
}

function refreshReportNow() {
  refreshNow();
  enableEditableEstadoActual();
  enableEditableProcesoMainFields();
  updateProcesoLevelCounters();
  updateFinalizadosLevelCounters();
  updateDesestimadosLevelCounters();
}

function openMailModal() {
  if (!mailModal) return;
  mailModal.classList.add("show");
  mailModal.setAttribute("aria-hidden", "false");
  if (mailPara) mailPara.focus();
}

function closeMailModal() {
  if (!mailModal) return;
  mailModal.classList.remove("show");
  mailModal.setAttribute("aria-hidden", "true");
}

function splitEmails(raw) {
  if (!raw) return [];
  return raw
    .split(",")
    .map(function (value) { return value.trim(); })
    .filter(function (value) { return value.length > 0; });
}

function getCurrentTurnoLabel() {
  return turnoSeleccionado === "noche" ? "Turno Noche" : "Turno Mañana";
}

function getCurrentShiftPhrase() {
  if (turnoSeleccionado === "noche") {
    return "Desde hoy las 18Hrs hasta las 6Hrs";
  }
  return "Desde hoy las 6Hrs hasta las 18Hrs";
}

function sanitizeHeader(text) {
  return String(text || "").replace(/[\r\n]+/g, " ").trim();
}

function replaceControlWithText(control) {
  const text =
    control.tagName === "SELECT"
      ? (control.options[control.selectedIndex] ? control.options[control.selectedIndex].text : "")
      : (control.value || "");
  const span = document.createElement("span");
  span.textContent = text;
  span.style.display = "inline-block";
  span.style.padding = "2px 6px";
  span.style.border = "1px solid #bfbfbf";
  span.style.minWidth = "210px";
  span.style.height = "24px";
  span.style.lineHeight = "18px";
  span.style.background = "#fff";
  span.style.color = "#000";
  control.replaceWith(span);
}

function blobToDataUrl(blob) {
  return new Promise(function (resolve, reject) {
    const reader = new FileReader();
    reader.onload = function () { resolve(String(reader.result || "")); };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function inlineImagesAsDataUri(root) {
  const images = Array.from(root.querySelectorAll("img"));
  await Promise.all(images.map(async function (img) {
    const src = img.getAttribute("src");
    if (!src || src.startsWith("data:")) return;
    try {
      const absoluteUrl = new URL(src, window.location.href).href;
      const response = await fetch(absoluteUrl);
      if (!response.ok) return;
      const blob = await response.blob();
      const dataUrl = await blobToDataUrl(blob);
      if (dataUrl) img.setAttribute("src", dataUrl);
    } catch (error) {
      console.warn("No se pudo incrustar imagen en correo:", src, error);
    }
  }));
}

async function buildEmailHtmlSnapshot() {
  const container = document.querySelector(".container");
  if (!container) return "<p>Sin contenido.</p>";
  const clone = container.cloneNode(true);

  const actions = clone.querySelector(".actions");
  if (actions) actions.remove();

  clone.querySelectorAll("select, input.hora-input").forEach(function (control) {
    replaceControlWithText(control);
  });
  await inlineImagesAsDataUri(clone);

  const styleTag = document.querySelector("style");
  const styles = styleTag ? styleTag.textContent : "";

  return (
    "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<style>" + styles + "</style>" +
    "</head><body>" + clone.outerHTML + "</body></html>"
  );
}

function downloadEml(toList, ccList, subject, htmlBody) {
  const headers = [
    "To: " + sanitizeHeader(toList.join(",")),
    ccList.length ? "Cc: " + sanitizeHeader(ccList.join(",")) : "",
    "Subject: " + sanitizeHeader(subject),
    "MIME-Version: 1.0",
    "Content-Type: text/html; charset=UTF-8",
    "Content-Transfer-Encoding: 8bit",
    ""
  ].filter(Boolean);

  const emlContent = headers.join("\r\n") + "\r\n" + htmlBody;
  const blob = new Blob([emlContent], { type: "message/rfc822;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const now = new Date();
  const filename =
    "Reporte_Monitoring_" +
    now.getFullYear() +
    pad2(now.getMonth() + 1) +
    pad2(now.getDate()) + "_" +
    pad2(now.getHours()) +
    pad2(now.getMinutes()) +
    ".eml";
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function readIncidentRows(sectionSelector) {
  const incidents = [];

  const tables = Array.from(document.querySelectorAll(sectionSelector));
  tables.forEach(function (table) {
    const rows = Array.from(table.querySelectorAll("tbody tr"));
    let current = null;

    rows.forEach(function (row) {
      if (!row.classList.contains("detail-row")) {
        const cells = row.querySelectorAll("td");
        const isProcesoRow = cells.length >= 8;
        current = {
          aplicacion: cells[0] ? cells[0].textContent.trim() : "",
          ticket: cells[1] ? cells[1].textContent.trim() : "",
          fuente: cells[2] ? cells[2].textContent.trim() : "",
          nivel: cells[3] ? cells[3].textContent.trim() : "",
          inicio: cells[4] ? cells[4].textContent.trim() : "",
          fin: !isProcesoRow && cells[5] ? cells[5].textContent.trim() : "",
          manager: isProcesoRow
            ? (cells[5] ? cells[5].textContent.trim() : "")
            : (cells[6] ? cells[6].textContent.trim() : ""),
          estado: isProcesoRow && cells[6] ? cells[6].textContent.trim() : "",
          ultimoCorreo: isProcesoRow && cells[7] ? cells[7].textContent.trim() : "",
          details: []
        };
        incidents.push(current);
        return;
      }

      if (!current) return;
      const labelCell = row.querySelector("td.detail-label");
      const textCell = row.querySelector("td.detail-text");
      if (!labelCell || !textCell) return;
      current.details.push({
        label: labelCell.textContent.trim(),
        text: textCell.textContent.trim()
      });
    });
  });

  return incidents;
}

function buildReportBody() {
  const proceso = readIncidentRows("section.rojo table.tabla-proceso-incidente");
  const finalizados = readIncidentRows("section.verde table.tabla-finalizados-incidente");
  const desestimados = readIncidentRows("section.morado table.tabla-desestimados-incidente");
  const lines = [];

  lines.push("REPORTE DE GESTION DE INCIDENTES - MONITORING");
  lines.push(getCurrentTurnoLabel());
  lines.push(getCurrentShiftPhrase());
  lines.push("Fecha: " + formatDate(new Date()));
  lines.push("");

  lines.push("INCIDENTES MASIVOS EN PROCESO");
  if (proceso.length === 0) {
    lines.push("Sin incidentes.");
  } else {
    proceso.forEach(function (item, index) {
      lines.push((index + 1) + ". " + item.aplicacion);
      lines.push("Ticket: " + item.ticket);
      lines.push("Fuente: " + item.fuente);
      lines.push("Nivel: " + item.nivel);
      lines.push("Inicio: " + item.inicio);
      lines.push("Situation Manager: " + item.manager);
      lines.push("Estado: " + item.estado);
      lines.push("Ultimo Correo de Gestion: " + item.ultimoCorreo);
      item.details.forEach(function (detail) {
        lines.push(detail.label + ": " + detail.text);
      });
      lines.push("");
    });
  }

  lines.push("INCIDENTES MASIVOS FINALIZADOS");
  if (finalizados.length === 0) {
    lines.push("Sin incidentes.");
  } else {
    finalizados.forEach(function (item, index) {
      lines.push((index + 1) + ". " + item.aplicacion);
      lines.push("Ticket: " + item.ticket);
      lines.push("Fuente: " + item.fuente);
      lines.push("Nivel: " + item.nivel);
      lines.push("Inicio: " + item.inicio);
      lines.push("Fin: " + item.fin);
      lines.push("Situation Manager: " + item.manager);
      item.details.forEach(function (detail) {
        lines.push(detail.label + ": " + detail.text);
      });
      lines.push("");
    });
  }

  lines.push("INCIDENTES DESESTIMADOS");
  if (desestimados.length === 0) {
    lines.push("Sin incidentes.");
  } else {
    desestimados.forEach(function (item, index) {
      lines.push((index + 1) + ". " + item.aplicacion);
      lines.push("Ticket: " + item.ticket);
      lines.push("Fuente: " + item.fuente);
      lines.push("Nivel: " + item.nivel);
      lines.push("Inicio: " + item.inicio);
      lines.push("Fin: " + item.fin);
      lines.push("Situation Manager: " + item.manager);
      item.details.forEach(function (detail) {
        lines.push(detail.label + ": " + detail.text);
      });
      lines.push("");
    });
  }

  return lines.join("\n");
}

async function openMailClientFromModal() {
  const toList = splitEmails(mailPara ? mailPara.value : "");
  const ccList = splitEmails(mailCc ? mailCc.value : "");
  if (toList.length === 0) {
    alert("Ingresa al menos un correo en 'Para'.");
    if (mailPara) mailPara.focus();
    return;
  }

  const subject =
    "[Reporte de Gestión de Incidente - Monitoring]" +
    "[" + getCurrentTurnoLabel() + "]" +
    "[" + getCurrentShiftPhrase() + "]" +
    "[" + formatDate(new Date()) + "]";
  const htmlBody = await buildEmailHtmlSnapshot();
  downloadEml(toList, ccList, subject, htmlBody);
  closeMailModal();
  alert("Se descargó el correo con diseño completo (.eml).");
}

if (btnAbrirEnviar) {
  btnAbrirEnviar.addEventListener("click", openMailModal);
}
if (btnActualizarReporte) {
  btnActualizarReporte.addEventListener("click", function () {
    if (!excelFileInput) {
      refreshReportNow();
      return;
    }
    excelFileInput.value = "";
    excelFileInput.click();
  });
}
if (btnCerrarEnviar) {
  btnCerrarEnviar.addEventListener("click", closeMailModal);
}
if (btnConfirmarEnviar) {
  btnConfirmarEnviar.addEventListener("click", openMailClientFromModal);
}
if (mailModal) {
  mailModal.addEventListener("click", function (event) {
    if (event.target === mailModal) closeMailModal();
  });
}
document.addEventListener("keydown", function (event) {
  if (event.key === "Escape") closeMailModal();
});

if (excelFileInput) {
  excelFileInput.addEventListener("change", function (event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    syncFromExcelFile(file).catch(function (error) {
      console.error(error);
      alert("No se pudo sincronizar el Excel. Verifica el formato de columnas.");
    });
  });
}

clearSectionsOnLoad();
enableEditableEstadoActual();
enableEditableProcesoMainFields();
updateProcesoLevelCounters();
watchProcesoTableCounters();
updateFinalizadosLevelCounters();
watchFinalizadosTableCounters();
updateDesestimadosLevelCounters();
watchDesestimadosTableCounters();
refreshNow();
refreshAtNextMinute();
