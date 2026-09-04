/**
 * Dashboard Operacional - Lógica e Integração Dom / Chart.js / SheetJS
 */

// Instâncias globais dos gráficos para permitir atualização/destruição
let chartRegiaoInstance = null;
let chartTipoInstance = null;

// Base de dados bruta importada do arquivo
let baseDadosGlobal = [];

// Listas de cidades para agrupamento regional
const CIDADES_SP_ABC = [
  "SAO PAULO", "SANTO ANDRE", "SAO BERNARDO DO CAMPO", 
  "DIADEMA", "MAUA", "OSASCO", "GUARULHOS"
];

const CIDADES_JAGUARIUNA_ENTORNO = [
  "JAGUARIUNA", "CAMPINAS", "HOLAMBRA", "AMPARO", 
  "PEDREIRA", "MOGI MIRIM", "MOGI GUACU", "PAULINIA", 
  "SERRA NEGRA", "ITAPIRA"
];

// Event Listeners assim que o DOM estiver pronto
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("excelFile");
  const statusFilter = document.getElementById("statusFilter");
  const startDate = document.getElementById("startDate");
  const endDate = document.getElementById("endDate");

  if (fileInput) fileInput.addEventListener("change", carregarArquivoExcel);
  if (statusFilter) statusFilter.addEventListener("change", aplicarFiltrosEAtualizar);
  if (startDate) startDate.addEventListener("change", aplicarFiltrosEAtualizar);
  if (endDate) endDate.addEventListener("change", aplicarFiltrosEAtualizar);
});

/**
 * 1. LEITURA DO ARQUIVO EXCEL / CSV
 */
function carregarArquivoExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Converte a planilha em matriz JS (com cabeçalho na 1ª linha)
    const jsonMatriz = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonMatriz.length <= 1) {
      alert("A planilha selecionada está vazia ou contém apenas cabeçalhos.");
      return;
    }

    // Separa cabeçalho dos dados
    const cabecalho = jsonMatriz[0].map(c => String(c || "").toUpperCase().trim());
    const linhas = jsonMatriz.slice(1);

    // Mapeia índices por nome ou usa padrões (Coluna C=2, V=21, W=22)
    const idxCidade = cabecalho.findIndex(c => c.includes("CIDADE")) !== -1 ? cabecalho.findIndex(c => c.includes("CIDADE")) : 2;
    const idxData = cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND")) !== -1 ? cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND")) : 21;
    const idxStatus = cabecalho.findIndex(c => c.includes("STATUS") || c.includes("TIPO") || c.includes("OPER")) !== -1 ? cabecalho.findIndex(c => c.includes("STATUS") || c.includes("TIPO") || c.includes("OPER")) : 22;

    // Normalização da base global
    baseDadosGlobal = linhas.map((linha, index) => ({
      id: linha[0] || `PED-${index + 1}`,
      cidade: String(linha[idxCidade] || "").toUpperCase().trim(),
      data: formatarDataIso(linha[idxData]),
      status: String(linha[idxStatus] || "").toUpperCase().trim()
    })).filter(item => item.cidade || item.status);

    aplicarFiltrosEAtualizar();
  };

  reader.readAsArrayBuffer(file);
}

/**
 * 2. FILTRAGEM DE DADOS
 */
function aplicarFiltrosEAtualizar() {
  const statusFiltro = document.getElementById("statusFilter")?.value || "TODOS";
  const dataInicio = document.getElementById("startDate")?.value;
  const dataFim = document.getElementById("endDate")?.value;

  const dadosFiltrados = baseDadosGlobal.filter(item => {
    // Filtro por Status
    let atendeStatus = true;
    if (statusFiltro === "AGENDADO_COLETADO") {
      atendeStatus = item.status.includes("AGEND") || item.status.includes("COLET");
    } else if (statusFiltro !== "TODOS") {
      atendeStatus = item.status.includes(statusFiltro);
    }

    // Filtro por Período
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
 * 3. PROCESSAMENTO DOS INDICADORES
 */
function processarDadosLogistica(dados) {
  let resumo = {
    totalPedidos: dados.length,
    coletas: 0,
    recebimentos: 0,
    entregas: 0,
    outrosTipos: 0,
    regioes: { spAbc: 0, jaguariunaEntorno: 0, outras: 0 }
  };

  dados.forEach(item => {
    // Status / Operação
    if (item.status.includes("COLET")) resumo.coletas++;
    else if (item.status.includes("RECEB")) resumo.recebimentos++;
    else if (item.status.includes("ENTREG")) resumo.entregas++;
    else resumo.outrosTipos++;

    // Regiões
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
      <span>Coletas</span>
      <strong>${resumo.coletas}</strong>
      <small>Operações de Coleta</small>
    </div>
    <div class="kpi-card">
      <span>Recebimentos</span>
      <strong>${resumo.recebimentos}</strong>
      <small>Entradas no CD</small>
    </div>
    <div class="kpi-card">
      <span>Entregas</span>
      <strong>${resumo.entregas}</strong>
      <small>Destino Final</small>
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
  // 1. Gráfico de Regiões (Doughnut)
  const ctxRegiao = document.getElementById("chartRegiao")?.getContext("2d");
  if (ctxRegiao) {
    if (chartRegiaoInstance) chartRegiaoInstance.destroy();

    chartRegiaoInstance = new Chart(ctxRegiao, {
      type: "doughnut",
      data: {
        labels: ["São Paulo / ABC", "Jaguariúna / Entorno", "Outras Regiões"],
        datasets: [{
          data: [resumo.regioes.spAbc, resumo.regioes.jaguariunaEntorno, resumo.regioes.outras],
          backgroundColor: ["#007bff", "#28a745", "#6c757d"]
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  // 2. Gráfico de Tipos (Bar)
  const ctxTipo = document.getElementById("chartTipo")?.getContext("2d");
  if (ctxTipo) {
    if (chartTipoInstance) chartTipoInstance.destroy();

    chartTipoInstance = new Chart(ctxTipo, {
      type: "bar",
      data: {
        labels: ["Coletas", "Recebimentos", "Entregas", "Outros"],
        datasets: [{
          label: "Volume de Pedidos",
          data: [resumo.coletas, resumo.recebimentos, resumo.entregas, resumo.outrosTipos],
          backgroundColor: ["#17a2b8", "#ffc107", "#28a745", "#6c757d"]
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}

function renderizarTabelaAgenda(dados) {
  const tbody = document.querySelector("#tabelaAgenda tbody");
  if (!tbody) return;

  // Agrupa pedidos por data
  const agenda = {};
  dados.forEach(item => {
    const dataChave = item.data || "Sem Data Agendada";
    if (!agenda[dataChave]) agenda[dataChave] = { qtd: 0, statusSet: new Set() };
    agenda[dataChave].qtd++;
    if (item.status) agenda[dataChave].statusSet.add(item.status);
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
 * 5. FUNÇÕES UTILITÁRIAS E LIMPEZA
 */
function limparFiltrosDatas() {
  document.getElementById("startDate").value = "";
  document.getElementById("endDate").value = "";
  document.getElementById("statusFilter").value = "AGENDADO_COLETADO";
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
  if (str.match(/^\d{2}\/\d{2}\/\d{4}/)) {
    const [d, m, y] = str.split("/");
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return str;
}

function formatarDataExibicao(dataIso) {
  if (!dataIso || !dataIso.includes("-")) return dataIso;
  const [y, m, d] = dataIso.split("-");
  return `${d}/${m}/${y}`;
}