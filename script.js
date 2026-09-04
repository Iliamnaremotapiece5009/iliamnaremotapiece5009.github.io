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

/**
 * Função utilitária para normalizar textos (remove acentos e padroniza para caixa alta)
 */
function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();
}

/**
 * 1. LEITURA APRIMORADA DE ARQUIVOS (XLSX, XLS E CSV)
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

      // Converte cabeçalho para maiúsculas e remove espaços
      const cabecalho = jsonMatriz[0].map(c => normalizarTexto(c));
      const linhas = jsonMatriz.slice(1);

      // Busca colunas por palavras-chave flexíveis
      let idxCidade = cabecalho.findIndex(c => c.includes("CIDADE") || c.includes("MUNIC"));
      let idxData = cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND") || c.includes("PREV"));
      let idxStatus = cabecalho.findIndex(c => c.includes("STATUS") || c.includes("SITUAC") || c.includes("TIPO") || c.includes("ETAPA"));

      // Fallback para posições padrão caso não encontre pelo nome
      if (idxCidade === -1) idxCidade = 2;
      if (idxData === -1) idxData = 21;
      if (idxStatus === -1) idxStatus = 22;

      baseDadosGlobal = linhas.map((linha, index) => ({
        id: linha[0] || `PED-${index + 1}`,
        cidade: normalizarTexto(linha[idxCidade]),
        data: formatarDataIso(linha[idxData]),
        status: normalizarTexto(linha[idxStatus])
      })).filter(item => item.cidade || item.status);

      aplicarFiltrosEAtualizar();

    } catch (erro) {
      console.error("Erro ao processar o arquivo:", erro);
      alert("Houve um erro ao ler o arquivo CSV/Excel. Verifique o console (F12).");
    }
  };

  if (isCsv) {
    reader.readAsText(file, "UTF-8");
  } else {
    reader.readAsArrayBuffer(file);
  }
}

/**
 * 2. FILTRAGEM DE DADOS (CORRIGIDA E SEM DUPLICAÇÃO DE FUNÇÃO)
 */
function aplicarFiltrosEAtualizar() {
  const statusFiltro = document.getElementById("statusFilter")?.value || "TODOS";
  const dataInicio = document.getElementById("startDate")?.value;
  const dataFim = document.getElementById("endDate")?.value;

  const dadosFiltrados = baseDadosGlobal.filter(item => {
    const st = item.status;
    let atendeStatus = false;

    switch (statusFiltro) {
      case "AGENDADO_COLETADO":
        atendeStatus = /AGEND|COLET/.test(st);
        break;
      case "AGENDADO":
        atendeStatus = /AGEND/.test(st);
        break;
      case "COLETADO":
        atendeStatus = /COLET/.test(st);
        break;
      case "FINALIZADO":
        atendeStatus = /FINALIZ|CONCLU|ENCERR|ATEND/.test(st);
        break;
      case "ENTREGUE":
        atendeStatus = /ENTREG/.test(st);
        break;
      case "PROGRAMADO":
        atendeStatus = /PROGRAM/.test(st);
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
    if (item.status.includes("COLET")) resumo.coletas++;
    else if (item.status.includes("RECEB")) resumo.recebimentos++;
    else if (item.status.includes("ENTREG") || item.status.includes("FINALIZ")) resumo.entregas++;
    else resumo.outrosTipos++;

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
      <span>Entregas / Finalizados</span>
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
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  const ctxTipo = document.getElementById("chartTipo")?.getContext("2d");
  if (ctxTipo) {
    if (chartTipoInstance) chartTipoInstance.destroy();

    chartTipoInstance = new Chart(ctxTipo, {
      type: "bar",
      data: {
        labels: ["Coletas", "Recebimentos", "Entregas/Finalizados", "Outros"],
        datasets: [{
          label: "Volume de Pedidos",
          data: [resumo.coletas, resumo.recebimentos, resumo.entregas, resumo.outrosTipos],
          backgroundColor: ["#0284c7", "#f59e0b", "#10b981", "#64748b"]
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
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
 * 5. UTILS
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