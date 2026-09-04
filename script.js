/**
 * LEITURA APRIMORADA DE ARQUIVOS (XLSX, XLS E CSV)
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
        // Trata CSVs com acentuação/caracteres do Brasil em UTF-8
        const csvContent = e.target.result;
        workbook = XLSX.read(csvContent, { type: "string", raw: true });
      } else {
        // Trata planilhas binárias Excel
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: "array" });
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Converte para matriz
      const jsonMatriz = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (!jsonMatriz || jsonMatriz.length <= 1) {
        alert("A planilha selecionada está vazia ou não possui linhas suficientes.");
        return;
      }

      // Mapeamento dinâmico dos dados
      const cabecalho = jsonMatriz[0].map(c => String(c || "").toUpperCase().trim());
      const linhas = jsonMatriz.slice(1);

      // Localização das colunas
      const idxCidade = cabecalho.findIndex(c => c.includes("CIDADE")) !== -1 ? cabecalho.findIndex(c => c.includes("CIDADE")) : COLUNA_CIDADE;
      const idxData = cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND")) !== -1 ? cabecalho.findIndex(c => c.includes("DATA") || c.includes("AGEND")) : COLUNA_DATA_AGENDAMENTO;
      const idxStatus = cabecalho.findIndex(c => c.includes("STATUS") || c.includes("TIPO") || c.includes("OPER")) !== -1 ? cabecalho.findIndex(c => c.includes("STATUS") || c.includes("TIPO") || c.includes("OPER")) : COLUNA_STATUS;

      baseDadosGlobal = linhas.map((linha, index) => ({
        id: linha[0] || `PED-${index + 1}`,
        cidade: String(linha[idxCidade] || "").toUpperCase().trim(),
        data: formatarDataIso(linha[idxData]),
        status: String(linha[idxStatus] || "").toUpperCase().trim()
      })).filter(item => item.cidade || item.status);

      // Renderiza imediatamente na tela
      aplicarFiltrosEAtualizar();

    } catch (erro) {
      console.error("Erro ao processar o arquivo:", erro);
      alert("Houve um erro ao ler o arquivo CSV/Excel. Verifique o console (F12) para detalhes.");
    }
  };

  // CSV deve ser lido como Texto; Excel (.xlsx) como ArrayBuffer
  if (isCsv) {
    reader.readAsText(file, "UTF-8");
  } else {
    reader.readAsArrayBuffer(file);
  }
}