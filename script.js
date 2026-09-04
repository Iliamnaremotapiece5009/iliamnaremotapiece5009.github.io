/**
 * Script de Tratamento e Leitura de Dados Logísticos
 * Mapeamento das colunas:
 * - Coluna V (Índice 21): DATA_AGENDAMENTO
 * - Coluna W (Índice 22): STATUS / TIPO
 */

// Índices das colunas (Base Zero)
const COLUNA_CIDADE = 2;            // Coluna C: CIDADE
const COLUNA_DATA_AGENDAMENTO = 21; // Coluna V: DATA_AGENDAMENTO
const COLUNA_STATUS = 22;           // Coluna W: STATUS (OU TIPO DE OPERAÇÃO)

// Lista de cidades para agrupamento regional
const CIDADES_SP_ABC = [
  "SAO PAULO", 
  "SANTO ANDRE", 
  "SAO BERNARDO DO CAMPO", 
  "DIADEMA", 
  "MAUA", 
  "OSASCO", 
  "GUARULHOS"
];

const CIDADES_JAGUARIUNA_ENTORNO = [
  "JAGUARIUNA", 
  "CAMPINAS", 
  "HOLAMBRA", 
  "AMPARO", 
  "PEDREIRA", 
  "MOGI MIRIM", 
  "MOGI GUACU", 
  "PAULINIA", 
  "SERRA NEGRA", 
  "ITAPIRA"
];

/**
 * Função principal para processar a base de dados
 * @param {Array<Array<any>>} dados - Matriz com as linhas e colunas da planilha (sem cabeçalho)
 */
function processarDadosLogistica(dados) {
  let resumo = {
    totalPedidos: dados.length,
    coletas: 0,
    recebimentos: 0,
    entregas: 0,
    regioes: {
      spAbc: 0,
      jaguariunaEntorno: 0,
      outras: 0
    }
  };

  dados.forEach(linha => {
    // 1. Processamento de Status / Tipo de Operação (Coluna W - Índice 22)
    const status = String(linha[COLUNA_STATUS] || "").toUpperCase().trim();
    
    if (status.includes("COLET") || status === "COLETA") {
      resumo.coletas++;
    } else if (status.includes("RECEB") || status === "RECEBIMENTO") {
      resumo.recebimentos++;
    } else if (status.includes("ENTREG") || status === "ENTREGA") {
      resumo.entregas++;
    }

    // 2. Classificação por Região Geográfica (Coluna C - Cidade)
    const cidade = String(linha[COLUNA_CIDADE] || "").toUpperCase().trim();

    if (CIDADES_SP_ABC.includes(cidade)) {
      resumo.regioes.spAbc++;
    } else if (CIDADES_JAGUARIUNA_ENTORNO.includes(cidade)) {
      resumo.regioes.jaguariunaEntorno++;
    } else {
      resumo.regioes.outras++;
    }
  });

  return resumo;
}

/**
 * Função para calcular percentuais de cada região
 * @param {Object} regioes - Objeto com a contagem das regiões
 * @param {number} total - Total de pedidos
 */
function calcularPercentuaisRegiao(regioes, total) {
  if (total === 0) return { spAbc: "0%", jaguariunaEntorno: "0%", outras: "0%" };

  return {
    spAbc: ((regioes.spAbc / total) * 100).toFixed(1) + "%",
    jaguariunaEntorno: ((regioes.jaguariunaEntorno / total) * 100).toFixed(1) + "%",
    outras: ((regioes.outras / total) * 100).toFixed(1) + "%"
  };
}

// ==========================================
// Exemplo de Teste / Uso Local no VSCode:
// ==========================================

const baseExemplo = [
  ["PED-1001", "Logística Silva", "SAO PAULO", "SP", "ENTREGA", 15, 120.5, 4500, "CD SP 1", "Carlos", "ABC-1234", "ROTA SP 01", -23.55, -46.63, "", "", "", "ALTA", "LOG", "08:00", "NENHUMA", "2026-09-01", "ENTREGUE"],
  ["PED-1002", "Distribuidora Mogi", "MOGI MIRIM", "SP", "COLETA", 40, 450, 12800, "CD INT", "Roberto", "DEF-5678", "ROTA INT 02", -22.43, -46.95, "", "", "", "MEDIA", "TRANS", "13:00", "NENHUMA", "2026-09-01", "COLETADO"],
  ["PED-1003", "Indústria Campinas", "CAMPINAS", "SP", "RECEBIMENTO", 8, 35, 1200, "CD INT", "Marcos", "GHI-9012", "ROTA INT 01", -22.90, -47.06, "", "", "", "BAIXA", "LOG", "08:00", "NENHUMA", "2026-09-02", "RECEBIDO"],
  ["PED-1004", "Comércio Jaguariúna", "JAGUARIUNA", "SP", "ENTREGA", 22, 210, 6700, "CD INT", "Marcos", "GHI-9012", "ROTA INT 01", -22.70, -46.98, "", "", "", "ALTA", "LOG", "10:00", "NENHUMA", "2026-09-02", "ENTREGUE"],
  ["PED-1005", "Varejo Santo André", "SANTO ANDRE", "SP", "ENTREGA", 10, 85, 2300, "CD SP 1", "Carlos", "ABC-1234", "ROTA ABC", -23.65, -46.53, "", "", "", "MEDIA", "TRANS", "08:00", "NENHUMA", "2026-09-03", "EM TRANSITO"]
];

const resultado = processarDadosLogistica(baseExemplo);
const percentuais = calcularPercentuaisRegiao(resultado.regioes, resultado.totalPedidos);

console.log("=== RESUMO LOGÍSTICO ===");
console.log("Total de Pedidos:", resultado.totalPedidos);
console.log("Coletas:", resultado.coletas);
console.log("Recebimentos:", resultado.recebimentos);
console.log("Entregas:", resultado.entregas);
console.log("\n=== DISTRIBUIÇÃO REGIONAL ===");
console.log(`São Paulo / ABC: ${resultado.regioes.spAbc} (${percentuais.spAbc})`);
console.log(`Jaguariúna / Entorno: ${resultado.regioes.jaguariunaEntorno} (${percentuais.jaguariunaEntorno})`);
console.log(`Outras Regiões: ${resultado.regioes.outras} (${percentuais.outras})`);