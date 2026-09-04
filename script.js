/**
 * Dashboard Operacional - Lógica e Integração DOM / Chart.js / SheetJS
 */

let chartRegiaoInstance = null;
let chartTipoInstance = null;
let baseDadosGlobal = [];

const CIDADES_SP_ABC = [
  "SAO PAULO", "SANTO ANDRE", "SAO BERNARDO DO CAMPO", 
  "DIADEMA", "MAUA", "OSASCO", "GUARULHOS"
];

const CIDADES_JAGUARIUNA_ENTORNO = [
  "JAGUARIUNA", "CAMPINAS", "HOLAMBRA", "AMPARO", 
  "PEDREIRA", "MOGI MIRIM", "MOGI GUACU", "PAULINIA", 
  "SERRA NEGRA", "ITAPIRA"
];

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("excelFile");
  const statusFilter = document.getElementById("statusFilter");
  const startDate = document.getElementById("startDate");
  const endDate = document.getElementById("endDate");
  const btnClear = document.getElementById("btnClear");

  if (fileInput) fileInput.addEventListener("change", carregarArquivoExcel);
  if (statusFilter) statusFilter.addEventListener("change", aplicarFiltrosEAtualizar);
  if (startDate) startDate.addEventListener("change", aplicarFiltrosEAtualizar);
  if (endDate) endDate.addEventListener("change", aplicarFiltrosEAtualizar);
  if (btnClear) btnClear.addEventListener("click", limparFiltrosDatas);
});

function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();
}

/**
 * 1. LEITURA APONTADA PARA A COLUNA W (ÍNDICE 22)
 */
function carregarArquivoExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const isCsv = file.name.toLowerCase().endsWith('.csv');
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      let workbook;
      if (isCsv) {
        const csvContent = e.target.result;
        workbook = XLSX.read(csvContent, { type: "string", raw: true });
      } else {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: "array" });
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonMatriz = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (!jsonMatriz || jsonMatriz.length <= 1) {
        alert("A planilha selecionada está vazia ou não possui linhas suficientes.");
        return;
      }

      const cabecalho = jsonMatriz[0].map(c => normalizarTexto(c));
      const linhas = jsonMatriz.slice(1);

      // Coluna W no Excel equivale ao índice 22 em arrays (A=0, B=1 ... W=22)
      const idxStatus = 22; 
      
      let idxCidade = cabecalho.findIndex(c => c.includes("CIDADE") || c.includes("MUNIC"));
      let idxData = cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND") || c.includes("PREV"));

      if (idxCidade === -1) idxCidade = 2;
      if (idxData === -1) idxData = 21;

      baseDadosGlobal = linhas.map((linha, index) => {
        const statusOriginal = String(linha[idxStatus] || "").trim();
        
        return {
          id: linha[0] || `PED-${index + 1}`,
          cidade: normalizarTexto(linha[idxCidade]),
          data: formatarDataIso(linha[idxData]),
          statusRaw: statusOriginal,
          statusNorm: normalizarTexto(statusOriginal)
        };
      }).filter(item => item.cidade || item.statusNorm);

      aplicarFiltrosEAtualizar();

    } catch (erro) {
      console.error("Erro ao processar o arquivo:", erro);
      alert("Houve um erro ao ler o arquivo CSV/Excel.");
    }
  };

  if (isCsv) {
    reader.readAsText(file, "UTF-8");
  } else {
    reader.readAsArrayBuffer(file);
  }
}

/**
 * 2. FILTRAGEM DE ACORDO COM O STATUS DA COLUNA W
 */
function aplicarFiltrosEAtualizar() {
  const statusFiltro = document.getElementById("statusFilter")?.value || "TODOS";
  const dataInicio = document.getElementById("startDate")?.value;
  const dataFim = document.getElementById("endDate")?.value;

  const dadosFiltrados = baseDadosGlobal.filter(item => {
    const st = item.statusNorm;
    let atendeStatus = false;

    switch (statusFiltro) {
      case "AGENDADO_COLETADO":
        atendeStatus = (st === "AGENDADO" || st === "COLETADO");
        break;
      case "AGENDADO":
        atendeStatus = (st === "AGENDADO");
        break;
      case "COLETADO":
        atendeStatus = (st === "COLETADO");
        break;
      case "ENTREGUE":
        atendeStatus = (st === "ENTREGUE");
        break;
      case "FINALIZADO":
        atendeStatus = (st === "FINALIZADO");
        break;
      case "PENDENTE_CONFERENCIA":
      case "PENDENTE DE CONFERENCIA":
        atendeStatus = (st.includes("PENDENTE") || st.includes("CONFERENCIA"));
        break;
      case "PROGRAMADO":
        atendeStatus = (st === "PROGRAMADO");
        break;
      case "TODOS":
      default:
        atendeStatus = true;
        break;
    }

    let atendeData = true;
    if (dataInicio && item.data) atendeData = atendeData && item.data >= dataInicio;
    if (dataFim && item.data) atendeData = atendeData && item.data <= dataFim;

    return atendeStatus && atendeData;
  });

  const resumo = processarDadosLogistica(dadosFiltrados);
  renderizarKPIs(resumo);
  renderizarPainelRegioes(resumo);
  renderizarGraficos(resumo, dadosFiltrados);
  renderizarTabelaAgenda(dadosFiltrados);
}

/**
 * 3. PROCESSAMENTO EXATO DE TODOS OS STATUS
 */
function processarDadosLogistica(dados) {
  let resumo = {
    totalPedidos: dados.length,
    agendados: 0,
    coletas: 0,
    entregas: 0,
    finalizados: 0,
    pendentesConferencia: 0,
    programados: 0,
    regioes: { spAbc: 0, jaguariunaEntorno: 0, outras: 0 }
  };

  dados.forEach(item => {
    const st = item.statusNorm;

    if (st === "AGENDADO") resumo.agendados++;
    else if (st === "COLETADO") resumo.coletas++;
    else if (st === "ENTREGUE") resumo.entregas++;
    else if (st === "FINALIZADO") resumo.finalizados++;
    else if (st.includes("PENDENTE") || st.includes("CONFERENCIA")) resumo.pendentesConferencia++;
    else if (st === "PROGRAMADO") resumo.programados++;

    if (CIDADES_SP_ABC.includes(item.cidade)) resumo.regioes.spAbc++;
    else if (CIDADES_JAGUARIUNA_ENTORNO.includes(item.cidade)) resumo.regioes.jaguariunaEntorno++;
    else resumo.regioes.outras++;
  });

  return resumo;
}

/**
 * 4. RENDERIZAÇÃO NA INTERFACE (DOM)
 */
function renderizarKPIs(resumo) {
  const container = document.getElementById("painelKPIs");
  if (!container) return;

  container.innerHTML = `
    <div class="kpi-card">
      <span>Total de Pedidos</span>
      <strong>${resumo.totalPedidos}</strong>
      <small>Base filtrada</small>
    </div>
    <div class="kpi-card">
      <span>Agendados / Programados</span>
      <strong>${resumo.agendados + resumo.programados}</strong>
      <small>Agend: ${resumo.agendados} | Prog: ${resumo.programados}</small>
    </div>
    <div class="kpi-card">
      <span>Coletas / Pendentes</span>
      <strong>${resumo.coletas + resumo.pendentesConferencia}</strong>
      <small>Colet: ${resumo.coletas} | Conf: ${resumo.pendentesConferencia}</small>
    </div>
    <div class="kpi-card">
      <span>Entregues / Finalizados</span>
      <strong>${resumo.entregas + resumo.finalizados}</strong>
      <small>Entr: ${resumo.entregas} | Fin: ${resumo.finalizados}</small>
    </div>
  `;
}

function renderizarPainelRegioes(resumo) {
  const container = document.getElementById("painelRegioes");
  if (!container) return;

  const total = resumo.totalPedidos || 1;
  const pctSp = ((resumo.regioes.spAbc / total) * 100).toFixed(1);
  const pctJag = ((resumo.regioes.jaguariunaEntorno / total) * 100).toFixed(1);
  const pctOutras = ((resumo.regioes.outras / total) * 100).toFixed(1);

  container.innerHTML = `
    <div class="card">
      <h3>Distribuição por Região Geográfica</h3>
      <div class="regiao-linha">
        <span>São Paulo / Grande ABC</span>
        <div>${resumo.regioes.spAbc} <span>(${pctSp}%)</span></div>
      </div>
      <div class="regiao-linha">
        <span>Jaguariúna e Entorno</span>
        <div>${resumo.regioes.jaguariunaEntorno} <span>(${pctJag}%)</span></div>
      </div>
      <div class="regiao-linha">
        <span>Demais Regiões</span>
        <div>${resumo.regioes.outras} <span>(${pctOutras}%)</span></div>
      </div>
    </div>
  `;
}

function renderizarGraficos(resumo, dados) {
  // Registra o plugin de DataLabels se estiver disponível no HTML
  if (typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
  }

  // 1. GRÁFICO POR REGIÃO (DOUGHNUT)
  const ctxRegiao = document.getElementById("chartRegiao")?.getContext("2d");
  if (ctxRegiao) {
    if (chartRegiaoInstance) chartRegiaoInstance.destroy();

    chartRegiaoInstance = new Chart(ctxRegiao, {
      type: "doughnut",
      data: {
        labels: ["São Paulo / ABC", "Jaguariúna / Entorno", "Outras Regiões"],
        datasets: [{
          data: [resumo.regioes.spAbc, resumo.regioes.jaguariunaEntorno, resumo.regioes.outras],
          backgroundColor: ["#2563eb", "#10b981", "#64748b"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          datalabels: {
            color: "#ffffff",
            font: {
              family: "'Inter', sans-serif",
              weight: "bold",
              size: 13
            },
            formatter: (value) => (value > 0 ? value : "")
          }
        }
      }
    });
  }

  // 2. GRÁFICO POR TIPO DE OPERAÇÃO (BAR)
  const ctxTipo = document.getElementById("chartTipo")?.getContext("2d");
  if (ctxTipo) {
    if (chartTipoInstance) chartTipoInstance.destroy();

    chartTipoInstance = new Chart(ctxTipo, {
      type: "bar",
      data: {
        labels: ["Agendado", "Coletado", "Entregue", "Finalizado", "Pend. Conf.", "Programado"],
        datasets: [{
          label: "Quantidade",
          data: [
            resumo.agendados, 
            resumo.coletas, 
            resumo.entregas, 
            resumo.finalizados, 
            resumo.pendentesConferencia, 
            resumo.programados
          ],
          backgroundColor: ["#3b82f6", "#0284c7", "#10b981", "#059669", "#f59e0b", "#8b5cf6"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: "end",
            align: "top",
            color: "#0f172a",
            font: {
              family: "'Inter', sans-serif",
              weight: "bold",
              size: 12
            },
            formatter: (value) => (value > 0 ? value : "")
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            grace: "12%" // Espaço superior para os rótulos não cortarem
          }
        }
      }
    });
  }
}

function renderizarTabelaAgenda(dados) {
  const tbody = document.querySelector("#tabelaAgenda tbody");
  if (!tbody) return;

  const agenda = {};
  dados.forEach(item => {
    const dataChave = item.data || "Sem Data Agendada";
    if (!agenda[dataChave]) agenda[dataChave] = { qtd: 0, statusSet: new Set() };
    agenda[dataChave].qtd++;
    if (item.statusRaw) agenda[dataChave].statusSet.add(item.statusRaw);
  });

  const datasOrdenadas = Object.keys(agenda).sort();

  tbody.innerHTML = datasOrdenadas.map(dt => `
    <tr>
      <td><strong>${formatarDataExibicao(dt)}</strong></td>
      <td><strong>${agenda[dt].qtd}</strong></td>
      <td><small>${Array.from(agenda[dt].statusSet).join(", ")}</small></td>
    </tr>
  `).join("");
}

/**
 * 5. UTILS
 */
function limparFiltrosDatas() {
  document.getElementById("startDate").value = "";
  document.getElementById("endDate").value = "";
  document.getElementById("statusFilter").value = "TODOS";
  aplicarFiltrosEAtualizar();
}

function formatarDataIso(valor) {
  if (!valor) return "";
  if (typeof valor === "number") {
    const dateObj = XLSX.SSF.parse_date_code(valor);
    if (dateObj) return `${dateObj.y}-${String(dateObj.m).padStart(2, '0')}-${String(dateObj.d).padStart(2, '0')}`;
  }
  const str = String(valor).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) return str.substring(0, 10);
  
  const matchPt = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (matchPt) {
    const d = matchPt[1].padStart(2, '0');
    const m = matchPt[2].padStart(2, '0');
    const y = matchPt[3];
    return `${y}-${m}-${d}`;
  }
  return "";
}

function formatarDataExibicao(dataIso) {
  if (!dataIso || !dataIso.includes("-")) return dataIso;
  const [y, m, d] = dataIso.split("-");
  return `${d}/${m}/${y}`;
}