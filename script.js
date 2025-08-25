/* =======================================================
   Nordil-BH - script.js (TradePro -> Dashboard)
   - Importa XLS/XLSX/XLSM/CSV via SheetJS
   - Layout fixo por posição (linha 5 = cabeçalho)
   - Cálculos automáticos e filtros avançados (análises)
   - Gráficos (Status, Análises, Horas Extras)
   - Novidades:
     * ACÚMULO de histórico (merge por chave ID|Data ou Nome|Data)
     * Análise "Atraso > tolerância"
   ======================================================= */

let dadosPonto = [];
let dadosFiltrados = [];
let tabelaBody = null;

// ===================== HELPERS DE DATA/HORA =====================
function pad2(n) { return String(n).padStart(2, "0"); }

// Converte somente a fração do serial Excel em HH:MM:SS (sem fuso/locale)
function excelSerialFractionToHHMMSS(value) {
  const num = Number(value);
  if (!isFinite(num)) return "";
  const sign = num < 0 ? "-" : "";
  const abs = Math.abs(num);
  const fraction = ((abs % 1) + 1) % 1; // garante [0,1)
  let totalSeconds = Math.round(fraction * 86400);
  if (totalSeconds === 86400) totalSeconds = 0; // 24:00:00 -> 00:00:00
  const hh = Math.floor(totalSeconds / 3600);
  const mm = Math.floor((totalSeconds % 3600) / 60);
  const ss = totalSeconds % 60;
  return `${sign}${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
}

// Converte parte inteira do serial Excel em DD/MM/AAAA (formato brasileiro)
function excelSerialToDateBR(value) {
  const num = Number(value);
  if (!isFinite(num)) return "";
  const days = Math.floor(num);
  const base = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
  base.setUTCDate(base.getUTCDate() + days);
  const yyyy = base.getUTCFullYear();
  const mm = pad2(base.getUTCMonth() + 1);
  const dd = pad2(base.getUTCDate());
  return `${dd}/${mm}/${yyyy}`;
}

// Excel serial -> "YYYY-MM-DD HH:MM:SS"
function excelSerialToISO(n) {
  const epoch = (n - 25569) * 86400 * 1000;
  const d = new Date(Math.round(epoch));
  const yyyy = d.getFullYear();
  const mm = pad2(d.getMonth() + 1);
  const dd = pad2(d.getDate());
  const hh = pad2(d.getHours());
  const mi = pad2(d.getMinutes());
  const ss = pad2(d.getSeconds());
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function normalizeDate(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) {
    const yyyy = v.getFullYear();
    const mm = pad2(v.getMonth() + 1);
    const dd = pad2(v.getDate());
    return `${dd}/${mm}/${yyyy}`;
  }
  if (typeof v === "number") return excelSerialToDateBR(v);
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return "";
    if (s.includes(" ")) return s.split(" ")[0];

    // número como string (pode vir com vírgula decimal do Excel)
    if (/^-?\d+(?:[\.,]\d+)?$/.test(s)) {
      const num = parseFloat(s.replace(',', '.'));
      return excelSerialToDateBR(num);
    }

    // dd/mm/aaaa
    let m = s.match(/^(\d{2})[/\-](\d{2})[/\-](\d{4})$/);
    if (m) return `${m[1]}/${m[2]}/${m[3]}`;

    // m/d/yy|yyyy
    m = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})$/);
    if (m) {
      const yyyy = m[3].length === 2 ? `20${m[3]}` : m[3];
      return `${pad2(m[1])}/${pad2(m[2])}/${yyyy}`;
    }

    // ISO (YYYY-MM-DD) -> DD/MM/AAAA
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[3]}/${m[2]}/${m[1]}`;

    return s;
  }
  return String(v);
}

function normalizeTime(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) {
    const hh = pad2(v.getHours());
    const mi = pad2(v.getMinutes());
    const ss = pad2(v.getSeconds());
    return `${hh}:${mi}:${ss}`;
  }
  if (typeof v === "number") return excelSerialFractionToHHMMSS(v);
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return "";
    if (s.includes(" ")) return s.split(" ")[1];
    // número como string (pode vir com vírgula decimal do Excel)
    if (/^-?\d+(?:[\.,]\d+)?$/.test(s)) {
      const num = parseFloat(s.replace(',', '.'));
      return excelSerialFractionToHHMMSS(num);
    }
    const parts = s.split(":").map(x => x.trim());
    if (parts.length >= 2) {
      const hh = pad2(parts[0] || "00");
      const mm = pad2(parts[1] || "00");
      const ss = pad2(parts[2] || "00");
      return `${hh}:${mm}:${ss}`;
    }
    return s;
  }
  return String(v);
}

function timeToSec(t) {
  if (t == null || t === "") return null;
  const s = String(t).trim();
  if (!s) return null;
  const neg = s.startsWith("-");
  const parts = (neg || s.startsWith("+")) ? s.slice(1).split(":") : s.split(":");
  if (parts.length < 2) return null;
  const hh = parseInt(parts[0], 10) || 0;
  const mm = parseInt(parts[1], 10) || 0;
  const ss = parseInt(parts[2] || "0", 10) || 0;
  let val = hh * 3600 + mm * 60 + ss;
  return neg ? -val : val;
}

function secToHHMMSS(sec) {
  if (sec == null) return "";
  const sign = sec < 0 ? "-" : "";
  const a = Math.abs(sec);
  const hh = Math.floor(a / 3600);
  const mm = Math.floor((a % 3600) / 60);
  const ss = a % 60;
  return `${sign}${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
}

// ===================== CÁLCULOS (Tempo/Jornada/Diferença) =====================
function computeDurations(row) {
  const e = timeToSec(row.entrada);
  const a = timeToSec(row.almoco);
  const r = timeToSec(row.retorno);
  const s = timeToSec(row.saida);

  const jp  = timeToSec(row.jornadaPlanejada);
  const tol = timeToSec(row.tolerancia);

  let lunchSec = null;
  if (a != null && r != null && r >= a) lunchSec = r - a;

  let workSec = null;
  if (e != null && s != null) workSec = s - e - (lunchSec || 0);

  if ((!row.tempoAlmoco || row.tempoAlmoco === "-") && lunchSec != null)
    row.tempoAlmoco = secToHHMMSS(lunchSec);

  if ((!row.jornada || row.jornada === "-") && workSec != null)
    row.jornada = secToHHMMSS(workSec);

  if (workSec != null && jp != null) {
    let diffSec = workSec - jp;
    if (tol != null && Math.abs(diffSec) <= tol) diffSec = 0;
    row.diferenca = secToHHMMSS(diffSec);
  }
  return row;
}

// ===================== GRÁFICOS =====================
let statusChart = null;
let analysesChart = null;
let overtimeChart = null;

function ensureCharts() {
  const statusEl = document.getElementById("statusChart");
  const analEl   = document.getElementById("analisesChart");
  const otEl     = document.getElementById("overtimeChart");
  if (!statusEl || !analEl) return;

  if (!statusChart) {
    statusChart = new Chart(statusEl.getContext("2d"), {
      type: "doughnut",
      data: { labels: ["Presente","Incompleto","Falta"], datasets: [{ data:[0,0,0] }] },
      options: { responsive:true, plugins:{ legend:{ position:"bottom" } } }
    });
  }
  if (!analysesChart) {
    analysesChart = new Chart(analEl.getContext("2d"), {
      type: "bar",
      data: {
        labels: [
          "Menos de 4 batidas",
          "Intervalo < 1h",
          "Intervalo > 2h",
          "Jornada > 8h",
          "Jornada > 10h",
          "Interjornada < 11h",
          "Atraso > tolerância" // NOVO
        ],
        datasets: [{ label: "Qtd. colaboradores", data: [0,0,0,0,0,0,0] }]
      },
      options: { responsive:true, plugins:{ legend:{ display:false } }, scales:{ y:{ beginAtZero:true, precision:0 } } }
    });
  }
  if (!overtimeChart && otEl) {
    overtimeChart = new Chart(otEl.getContext("2d"), {
      type: "bar",
      data: {
        labels: [],
        datasets: [
          { label: "Saldo (h)", data: [] } // total líquido do mês (positivo + negativo)
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { position: "bottom" } },
        scales: { y: { beginAtZero: true } }
      }
    });
  }
}

function updateCharts(dados) {
  ensureCharts();
  if (!statusChart) return;

  const presentes   = dados.filter(d => d.status === "presente").length;
  const incompletos = dados.filter(d => d.status === "incompleto").length;
  const faltas      = dados.filter(d => d.status === "falta").length;
  statusChart.data.datasets[0].data = [presentes, incompletos, faltas];
  statusChart.update();

  // Horas extras por mês (SALDO): usa "Diferença" se houver; senão JE-JP com tolerância
  if (overtimeChart) {
    const byMonthNet = new Map(); // soma algébrica do mês (positivos + negativos)

    dados.forEach(d => {
      const dataVal = d.data || "";
      let key = "";
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(dataVal)) {
        const [dd, mm, yyyy] = dataVal.split("/");
        key = `${yyyy}-${mm}`; // YYYY-MM
      } else if (/^\d{4}-\d{2}-\d{2}/.test(dataVal)) {
        key = dataVal.slice(0, 7);
      } else if (/^\d{4}-\d{2}$/.test(dataVal)) {
        key = dataVal;
      } else {
        return;
      }

      // 1) Preferir "Diferença"
      let diffSec = timeToSec(d.diferenca);

      // 2) Fallback: JE - JP respeitando "Tolerância"
      if (diffSec == null) {
        const jp  = timeToSec(d.jornadaPlanejada);
        const je  = d.jornadaSec != null ? d.jornadaSec : timeToSec(d.jornada);
        const tol = timeToSec(d.tolerancia);
        if (jp != null && je != null) {
          let tmp = je - jp;
          if (tol != null && Math.abs(tmp) <= tol) tmp = 0;
          diffSec = tmp;
        }
      }

      if (diffSec != null) {
        byMonthNet.set(key, (byMonthNet.get(key) || 0) + diffSec);
      }
    });

    const labels = [...byMonthNet.keys()].sort();
    const saldo  = labels.map(k => +(((byMonthNet.get(k) || 0) / 3600).toFixed(2))); // horas com sinal

    overtimeChart.data.labels = labels;
    overtimeChart.data.datasets[0].data = saldo;
    overtimeChart.update();
  }
}

function updateAnalisesChart(dados) {
  ensureCharts();
  if (!analysesChart) return;

  const setMenos4    = new Set();
  const setIntLt1    = new Set();
  const setIntGt2    = new Set();
  const setJorGt8    = new Set();
  const setJorGt10   = new Set();
  const setInterLt11 = new Set();
  const setAtrasoTol = new Set(); // NOVO

  dados.forEach(d => {
    // chave única por colaborador (prefere ID)
    const chave = (d.id != null && String(d.id).trim() !== "") ? String(d.id).trim() : (d.nome || "");

    if ((d.batidasCount || 0) < 4) setMenos4.add(chave);
    if (d.intervaloAlmocoSec != null && d.intervaloAlmocoSec < 3600) setIntLt1.add(chave);
    if (d.intervaloAlmocoSec != null && d.intervaloAlmocoSec > 7200) setIntGt2.add(chave);
    if (d.jornadaSec != null && d.jornadaSec > 8*3600) setJorGt8.add(chave);
    if (d.jornadaSec != null && d.jornadaSec > 10*3600) setJorGt10.add(chave);
    if (d.interjornadaSec != null && d.interjornadaSec < 11*3600) setInterLt11.add(chave);

    // Atraso > tolerância (entrada executada - planejada > tolerância)
    const e  = timeToSec(d.entrada);
    const ep = timeToSec(d.entradaPlanejada);
    const tol = timeToSec(d.tolerancia);
    if (e != null && ep != null && tol != null && (e - ep) > tol) setAtrasoTol.add(chave);
  });

  analysesChart.data.datasets[0].data = [
    setMenos4.size,
    setIntLt1.size,
    setIntGt2.size,
    setJorGt8.size,
    setJorGt10.size,
    setInterLt11.size,
    setAtrasoTol.size // NOVO
  ];
  analysesChart.update();
}

// ===================== LAYOUT FIXO TRADEPRO =====================
function buildFixedTradeProMap() {
  return {
    id: 0, nome: 1, data: 2,
    entrada: 3, almoco: 4, retorno: 5, saida: 6,
    tempoAlmoco: 7, jornada: 8,
    entradaPlanejada: 9, saidaPlanejada: 10, tempoAlmocoPlanejada: 11, jornadaPlanejada: 12,
    tolerancia: 13, diferenca: 14, observacoes: 15
  };
}

// Tenta inferir o mapeamento pelas legendas da linha de cabeçalho
function inferTradeProMapFromHeader(rows, headerRowIndex) {
  const headerRow = rows[headerRowIndex] || [];
  const groupRow  = rows[headerRowIndex - 1] || []; // linha com "Executado" / "Planejado"
  if (!Array.isArray(headerRow)) return null;
  const norm = (v) => (v == null ? "" : String(v).toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim());

  // aceita "Planejado" na linha de grupo OU "(P)" / "planejado" no próprio header
  const isPlanejado = (i) => {
    const g = norm(groupRow[i] || '');
    const h = norm(headerRow[i] || '');
    return g.includes('planejado') || /\(p\)|planejado/.test(h);
  };

  const findFirst = (frags) => {
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(headerRow[i]);
      if (!h) continue;
      if (frags.every(f => h.includes(f))) return i;
    }
    return -1;
  };
  const findPlanejado = (frags) => {
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(headerRow[i]);
      if (!h) continue;
      if (frags.every(f => h.includes(f)) && isPlanejado(i)) return i;
    }
    return -1;
  };
  const findExecutado = (frags) => {
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(headerRow[i]);
      if (!h) continue;
      if (frags.every(f => h.includes(f)) && !isPlanejado(i)) return i;
    }
    return -1;
  };

  const id   = findFirst(['id']);
  const nome = findFirst(['colaborador']) >= 0 ? findFirst(['colaborador']) : findFirst(['funcionario']);
  const data = findFirst(['data']);

  const entrada = findExecutado(['inicio']);
  const almoco  = findExecutado(['almoco']);
  const retorno = findExecutado(['retorno']);
  const saida   = findExecutado(['saida']);
  const tempoAlmoco = findExecutado(['tempo','almoco']);
  const jornada     = findExecutado(['jornada']);

  const entradaPlanejada = findPlanejado(['inicio']);
  const saidaPlanejada   = findPlanejado(['saida']);
  const tempoAlmocoPlanejada = findPlanejado(['tempo','almoco']);
  const jornadaPlanejada     = findPlanejado(['jornada']);

  const tolerancia = findFirst(['tolerancia']);
  const diferenca  = findFirst(['diferenca']);
  const observacoes= findFirst(['observ']);

  const found = { id, nome, data, entrada, almoco, retorno, saida, tempoAlmoco, jornada,
                  entradaPlanejada, saidaPlanejada, tempoAlmocoPlanejada, jornadaPlanejada,
                  tolerancia, diferenca, observacoes };

  const values = Object.values(found);
  const ok = values.filter(v => typeof v === 'number' && v >= 0).length;
  if (ok >= Math.ceil(values.length * 0.8)) return found;
  return null;
}

// ===================== ENRIQUECER MÉTRICAS (batidas/interjornada) =====================
function enrichDerivedMetrics(dados) {
  const byColab = new Map();

  dados.forEach(d => {
    d.batidasCount = ["entrada","almoco","retorno","saida"].reduce((acc, k) => acc + (timeToSec(d[k]) != null ? 1 : 0), 0);

    const tAlm = timeToSec(d.tempoAlmoco);
    if (tAlm != null) {
      d.intervaloAlmocoSec = tAlm;
    } else {
      const a = timeToSec(d.almoco);
      const r = timeToSec(d.retorno);
      d.intervaloAlmocoSec = (a != null && r != null && r >= a) ? (r - a) : null;
      if (d.intervaloAlmocoSec != null) d.tempoAlmoco = secToHHMMSS(d.intervaloAlmocoSec);
    }

    const j = timeToSec(d.jornada);
    if (j != null) {
      d.jornadaSec = j;
    } else {
      const e = timeToSec(d.entrada);
      const s = timeToSec(d.saida);
      d.jornadaSec = (e != null && s != null) ? (s - e - (d.intervaloAlmocoSec || 0)) : null;
      if (d.jornadaSec != null) d.jornada = secToHHMMSS(d.jornadaSec);
    }

    const key = d.nome || "";
    if (!byColab.has(key)) byColab.set(key, []);
    byColab.get(key).push(d);
  });

  for (const [, arr] of byColab) {
    arr.sort((a,b) => (a.data || "").localeCompare(b.data || ""));
    let prev = null;
    arr.forEach(rec => {
      rec.interjornadaSec = null;
      if (prev) {
        const e = timeToSec(rec.entrada);
        const s = timeToSec(prev.saida);
        if (e != null && s != null) {
          let diff = e - s;
          if (diff < 0) diff += 24 * 3600; // ajusta virada de dia
          rec.interjornadaSec = diff;
        }
      }
      prev = rec;
    });
  }
}

// ===================== CHAVE DO REGISTRO (para merge do histórico) =====================
function chaveRegistro(d) {
  const id = (d.id != null && String(d.id).trim() !== "") ? String(d.id).trim() : null;
  const nome = (d.nome || "").trim();
  const data = (d.data || "").trim();
  // Preferir ID quando existir; fallback para Nome
  const pessoa = id || nome;
  return `${pessoa}|${data}`;
}

// ===================== PROCESSAR DADOS (com ACÚMULO) =====================
function processarDados(rows) {
  if (!rows || rows.length < 6) {
    alert("Arquivo não contém dados suficientes.");
    return;
  }

  const headerRowIndex = 4;           // 5ª linha
  const startDataIndex = headerRowIndex + 1;
  const map = inferTradeProMapFromHeader(rows, headerRowIndex) || buildFixedTradeProMap();

  const novos = rows
    .slice(startDataIndex)
    .filter(line => (line || []).some(v => (v ?? "").toString().trim() !== "")) // limpa totalmente vazias
    .filter(line => { // precisa ter ID e Nome
      const id = (line[map.id] ?? "").toString().trim();
      const nome = (line[map.nome] ?? "").toString().trim();
      return id !== "" && nome !== "";
    })
    .map((l) => {
      const gv = (idx) => (idx != null ? l[idx] : "");
      const o = {
        id:               gv(map.id) || "N/A",
        nome:             gv(map.nome) || "N/A",
        data:             normalizeDate(gv(map.data)),
        entrada:          normalizeTime(gv(map.entrada)),
        almoco:           normalizeTime(gv(map.almoco)),
        retorno:          normalizeTime(gv(map.retorno)),
        saida:            normalizeTime(gv(map.saida)),
        tempoAlmoco:      normalizeTime(gv(map.tempoAlmoco)),
        jornada:          normalizeTime(gv(map.jornada)),
        entradaPlanejada: normalizeTime(gv(map.entradaPlanejada)),
        saidaPlanejada:   normalizeTime(gv(map.saidaPlanejada)),
        tempoAlmocoPlanejada: normalizeTime(gv(map.tempoAlmocoPlanejada)),
        jornadaPlanejada:     normalizeTime(gv(map.jornadaPlanejada)),
        tolerancia: (normalizeTime(gv(map.tolerancia)) || "00:10:00"),
        diferenca:  (gv(map.diferenca) || "").toString(),
        observacoes:(gv(map.observacoes) || "").toString()
      };

      o.status = (!o.entrada && !o.almoco && !o.retorno && !o.saida)
        ? "falta"
        : (!o.entrada || !o.saida) ? "incompleto" : "presente";

      // Recalcula diferença com base em jornada executada vs planejada, respeitando tolerância
      const computed = computeDurations(o);
      const jp = timeToSec(computed.jornadaPlanejada);
      const je = timeToSec(computed.jornada);
      const tol = timeToSec(computed.tolerancia);
      if (jp != null && je != null) {
        let diffSec = je - jp;
        if (tol != null && Math.abs(diffSec) <= tol) diffSec = 0;
        computed.diferenca = secToHHMMSS(diffSec);
      }
      return computed;
    });

  // === ACÚMULO: mescla "novos" no "dadosPonto" sem apagar histórico ===
  const mapa = new Map(dadosPonto.map(d => [chaveRegistro(d), d]));
  novos.forEach(n => {
    mapa.set(chaveRegistro(n), n); // atualiza se já existir, insere se novo
  });
  dadosPonto = [...mapa.values()];

  // enriquecer métricas derivadas para o conjunto completo
  enrichDerivedMetrics(dadosPonto);

  preencherSelects(dadosPonto);
  aplicarFiltros();
}

// ===================== RENDER / SELECTS / FILTROS =====================
function renderizarTabela(dados) {
  const tbody = tabelaBody || document.querySelector("#tabelaPonto tbody");
  if (!tbody) return;

  tbody.innerHTML = dados.map(d => `
    <tr class="status-${d.status}">
      <td>${d.id}</td>
      <td>${d.nome}</td>
      <td>${d.data || "-"}</td>
      <td>${d.entrada || "-"}</td>
      <td>${d.almoco || "-"}</td>
      <td>${d.retorno || "-"}</td>
      <td>${d.saida || "-"}</td>
      <td>${d.tempoAlmoco || "-"}</td>
      <td>${d.jornada || "-"}</td>
      <td>${d.entradaPlanejada || "-"}</td>
      <td>${d.saidaPlanejada || "-"}</td>
      <td>${d.jornadaPlanejada || "-"}</td>
      <td>${d.tolerancia || "-"}</td>
      <td>${d.diferenca || "-"}</td>
      <td>${d.observacoes || "-"}</td>
      <td>${d.status}</td>
    </tr>
  `).join("");
}

function preencherSelects(dados) {
  const nomes = [...new Set(dados.map(d => d.nome).filter(Boolean))].sort();
  // Extrai mês e ano do formato DD/MM/AAAA
  const meses = [...new Set(dados.map(d => {
    const data = d.data || "";
    if (data.includes("/")) {
      return data.split("/")[1]; // mês está na posição 1
    }
    return "";
  }).filter(Boolean))].sort();
  
  const anos  = [...new Set(dados.map(d => {
    const data = d.data || "";
    if (data.includes("/")) {
      return data.split("/")[2]; // ano está na posição 2
    }
    return "";
  }).filter(Boolean))].sort();

  const selColab = document.getElementById("filtroColaborador");
  const selMes   = document.getElementById("filtroMes");
  const selAno   = document.getElementById("filtroAno");

  if (selColab) selColab.innerHTML = `<option value="">Todos</option>${nomes.map(n => `<option>${n}</option>`).join("")}`;
  if (selMes)   selMes.innerHTML   = `<option value="">Todos</option>${meses.map(m => `<option>${m}</option>`).join("")}`;
  if (selAno)   selAno.innerHTML   = `<option value="">Todos</option>${anos.map(a => `<option>${a}</option>`).join("")}`;
}

function atualizarCards(dados) {
  const totalFunc = document.getElementById("totalFuncionarios");
  const arqImp    = document.getElementById("arquivosImportados");
  const jr        = document.getElementById("totalJornadas");
  const faltas    = document.getElementById("faltas");
  const incomps   = document.getElementById("incompletas");

  const keys = new Set(dados.map(d => {
    const id = (d.id != null && String(d.id).trim() !== "") ? String(d.id).trim() : null;
    return id || (d.nome || "").trim();
  }).filter(Boolean));

  if (totalFunc) totalFunc.textContent = String(keys.size);
  if (arqImp)    arqImp.textContent    = dados.length > 0 ? "1" : "0"; // mantém simples
  if (jr)        jr.textContent        = String(dados.filter(d => d.entrada && d.saida).length);
  if (faltas)    faltas.textContent    = String(dados.filter(d => d.status === "falta").length);
  if (incomps)   incomps.textContent   = String(dados.filter(d => d.status === "incompleto").length);
}

function aplicarFiltros() {
  const selects = {
    mes: document.getElementById("filtroMes"),
    ano: document.getElementById("filtroAno"),
    colaborador: document.getElementById("filtroColaborador"),
    status: document.getElementById("filtroStatus"),
    analise: document.getElementById("filtroAnalise")
  };
  const filtroNome = document.getElementById("filtroNome");
  const filtroData = document.getElementById("filtroData");

  const filtro = {
    nome: (filtroNome?.value || "").toLowerCase(),
    data: (filtroData?.value || "").trim(),
    mes: selects.mes?.value || "",
    ano: selects.ano?.value || "",
    colaborador: selects.colaborador?.value || "",
    status: selects.status?.value || ""
  };
  const analise = selects.analise?.value || "";

  // normaliza data exata (DD/MM/AAAA)
  let dataFiltroBR = "";
  if (filtro.data) {
    const m = filtro.data.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) dataFiltroBR = `${m[1]}/${m[2]}/${m[3]}`;
  }

  dadosFiltrados = dadosPonto.filter((d) => {
    const dataBR = d.data || "";
    const partes = dataBR.split("/");
    const ano = partes.length === 3 ? partes[2] : "";
    const mes = partes.length === 3 ? partes[1] : "";

    const passaBasicos =
      (!filtro.nome || (d.nome || "").toLowerCase().includes(filtro.nome)) &&
      (!filtro.mes || (mes && mes === filtro.mes)) &&
      (!filtro.ano || (ano && ano === filtro.ano)) &&
      (!filtro.colaborador || d.nome === filtro.colaborador) &&
      (!filtro.status || d.status === filtro.status) &&
      (!dataFiltroBR || (dataBR && dataBR === dataFiltroBR));

    if (!passaBasicos) return false;

    switch (analise) {
      case "menos4batidas":     return (d.batidasCount || 0) < 4;
      case "intervaloLt1h":     return d.intervaloAlmocoSec != null && d.intervaloAlmocoSec < 3600;
      case "intervaloGt2h":     return d.intervaloAlmocoSec != null && d.intervaloAlmocoSec > 7200;
      case "jornadaGt8h":       return d.jornadaSec != null && d.jornadaSec > 8*3600;
      case "jornadaGt10h":      return d.jornadaSec != null && d.jornadaSec > 10*3600;
      case "interjornadaLt11h": return d.interjornadaSec != null && d.interjornadaSec < 11*3600;
      case "atrasoTol": { // NOVO: precisa adicionar esta opção no select se quiser filtrar
        const e  = timeToSec(d.entrada);
        const ep = timeToSec(d.entradaPlanejada);
        const tol = timeToSec(d.tolerancia);
        return (e != null && ep != null && tol != null && (e - ep) > tol);
      }
      default: return true;
    }
  });

  renderizarTabela(dadosFiltrados);
  atualizarCards(dadosFiltrados);
  updateCharts(dadosFiltrados);
  updateAnalisesChart(dadosFiltrados);
}

// ===================== IMPORTAÇÃO (SheetJS) =====================
// Preenche células mescladas do Excel no array "rows"
function expandMergedCells(ws, rows) {
  const merges = ws && ws["!merges"] ? ws["!merges"] : [];
  merges.forEach(m => {
    const topLeft = (rows[m.s.r] || [])[m.s.c];
    for (let r = m.s.r; r <= m.e.r; r++) {
      if (!rows[r]) rows[r] = [];
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (rows[r][c] == null || rows[r][c] === "") rows[r][c] = topLeft;
      }
    }
  });
}

function lerArquivo(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const wsname = wb.SheetNames[0];
    const ws = wb.Sheets[wsname];

    // Mantém vazios como "" para não “comer” colunas
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

    // Expande merges (Executado/Planejado e títulos mesclados)
    if (ws && ws["!merges"] && ws["!merges"].length) expandMergedCells(ws, rows);

    processarDados(rows);
  };
  reader.readAsArrayBuffer(file);
}

// ===================== BOOT (Dashboard) =====================
if (window.location.pathname.includes("dashboard.html")) {
  tabelaBody = document.querySelector("#tabelaPonto tbody");

  const input  = document.getElementById("csvFile");
  const btnImp = document.getElementById("importarCSV");
  const limparFiltro = document.getElementById("limparFiltro");
  const exportarBtn  = document.getElementById("exportarDados");
  const limparTudoBtn = document.getElementById("limparTudo");

  if (btnImp) {
    btnImp.addEventListener("click", () => {
      const file = input?.files?.[0];
      if (!file) { alert("Selecione um arquivo .xls/.xlsx/.xlsm/.csv"); return; }
      lerArquivo(file);
    });
  }

  // Listeners de filtros (texto/select)
  ["filtroNome","filtroData","filtroMes","filtroAno","filtroColaborador","filtroStatus"]
    .map(id => document.getElementById(id))
    .forEach(el => el && el.addEventListener("input", aplicarFiltros));

  const filtroAnalise = document.getElementById("filtroAnalise");
  if (filtroAnalise) filtroAnalise.addEventListener("change", aplicarFiltros);

  if (limparFiltro) {
    limparFiltro.addEventListener("click", () => {
      ["filtroNome","filtroData","filtroMes","filtroAno","filtroColaborador","filtroStatus","filtroAnalise"]
        .map(id => document.getElementById(id))
        .forEach(el => { if (el) el.value = ""; });
      aplicarFiltros();
    });
  }

  if (exportarBtn) {
    exportarBtn.addEventListener("click", () => {
      const header = [
        "ID","Colaborador","Data","Início (Executado)","Almoço","Retorno","Saída",
        "Tempo Almoço","Jornada","Início (Planejado)","Saída (Planejado)","Tempo Almoço (P)","Jornada (P)",
        "Tolerância","Diferença","Observações","Status"
      ];
      const body = dadosFiltrados.length ? dadosFiltrados : dadosPonto;
      const data = [header].concat(body.map(d => [
        d.id,d.nome,d.data,d.entrada,d.almoco,d.retorno,d.saida,
        d.tempoAlmoco,d.jornada,d.entradaPlanejada,d.saidaPlanejada,d.tempoAlmocoPlanejada,d.jornadaPlanejada,
        d.tolerancia,d.diferenca,d.observacoes,d.status
      ]));
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Ponto");
      XLSX.writeFile(wb, "nordil_ponto_export.xlsx");
    });
  }

  if (limparTudoBtn) {
    limparTudoBtn.addEventListener("click", () => {
      dadosPonto = [];
      dadosFiltrados = [];
      renderizarTabela([]);
      atualizarCards([]);
      updateCharts([]);
      if (document.getElementById("csvFile")) document.getElementById("csvFile").value = "";
      ["filtroNome","filtroData","filtroMes","filtroAno","filtroColaborador","filtroStatus","filtroAnalise"]
        .map(id => document.getElementById(id))
        .forEach(el => { if (el) el.value = ""; });
    });
  }

  // inicializa UI em branco
  dadosFiltrados = [];
  renderizarTabela(dadosFiltrados);
  atualizarCards(dadosFiltrados);
  updateCharts(dadosFiltrados);
  updateAnalisesChart(dadosFiltrados);
}
