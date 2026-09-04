/**

* ============================================================
* DASHBOARD OPERACIONAL
* ============================================================
*
* COLUNAS PRINCIPAIS DA BASE:
*
* V = DATA DE AGENDAMENTO
* W = STATUS
*
* PEDIDO_   = Número do pedido
* CIDADE    = Cidade
* TIPO      = Tipo da operação
* FORNECEDOR = Fornecedor
*
* ============================================================
  */

const COLUNA_DATA = 21;      // V
const COLUNA_STATUS = 22;    // W

const COLUNA_PEDIDO = "PEDIDO_";
const COLUNA_CIDADE = "CIDADE";
const COLUNA_TIPO = "TIPO";
const COLUNA_FORNECEDOR = "FORNECEDOR";

/* ============================================================
VARIÁVEIS
============================================================ */

let rawData = [];
let pedidosConsolidados = [];

let instanceChartRegiao = null;
let instanceChartTipo = null;

/* ============================================================
INICIALIZAÇÃO
============================================================ */

document.addEventListener("DOMContentLoaded", function () {

```
const fileInput = document.getElementById("excelFile");
const statusFilter = document.getElementById("statusFilter");
const regionFilter = document.getElementById("regionFilter");
const startDate = document.getElementById("startDate");
const endDate = document.getElementById("endDate");
const btnClear = document.getElementById("btnClear");

if (fileInput) {
    fileInput.addEventListener(
        "change",
        carregarArquivo
    );
}

if (statusFilter) {
    statusFilter.addEventListener(
        "change",
        processarEAtualizar
    );
}

if (regionFilter) {
    regionFilter.addEventListener(
        "change",
        processarEAtualizar
    );
}

if (startDate) {
    startDate.addEventListener(
        "change",
        processarEAtualizar
    );
}

if (endDate) {
    endDate.addEventListener(
        "change",
        processarEAtualizar
    );
}

if (btnClear) {
    btnClear.addEventListener(
        "click",
        limparFiltros
    );
}

if (typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
}
```

});

/* ============================================================
NORMALIZAR TEXTO
============================================================ */

function normalizarTexto(valor) {

```
if (
    valor === undefined ||
    valor === null
) {
    return "";
}

return String(valor)
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(
        /[\u0300-\u036f]/g,
        ""
    );
```

}

/* ============================================================
LOCALIZAR COLUNA PELO NOME
============================================================ */

function encontrarColuna(objeto, nomeProcurado) {

```
const alvo =
    normalizarTexto(nomeProcurado);

return Object.keys(objeto).find(
    function (chave) {

        return (
            normalizarTexto(chave) ===
            alvo
        );

    }
);
```

}

/* ============================================================
LEITURA DO ARQUIVO
============================================================ */

function carregarArquivo(event) {

```
const file =
    event.target.files[0];

if (!file) {
    return;
}

const reader =
    new FileReader();

reader.onload =
    function (evt) {

        try {

            const data =
                new Uint8Array(
                    evt.target.result
                );

            const workbook =
                XLSX.read(
                    data,
                    {
                        type: "array",
                        cellDates: true
                    }
                );

            const primeiraAba =
                workbook.SheetNames[0];

            const worksheet =
                workbook.Sheets[
                    primeiraAba
                ];

            rawData =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        raw: true,
                        defval: ""
                    }
                );

            if (
                !rawData ||
                rawData.length === 0
            ) {

                alert(
                    "A planilha selecionada está vazia."
                );

                return;
            }


            /*
             * Busca V = Data
             * W = Status
             */

            rawData =
                rawData.map(
                    function (item, index) {

                        const celulaData =
                            worksheet[
                                XLSX.utils.encode_cell({
                                    r: index + 1,
                                    c: COLUNA_DATA
                                })
                            ];

                        const celulaStatus =
                            worksheet[
                                XLSX.utils.encode_cell({
                                    r: index + 1,
                                    c: COLUNA_STATUS
                                })
                            ];

                        return {

                            ...item,

                            DATA_V:
                                celulaData
                                    ? celulaData.v
                                    : "",

                            STATUS_W:
                                celulaStatus
                                    ? celulaStatus.v
                                    : "",

                            LINHA_ORIGINAL:
                                index + 2

                        };

                    }
                );


            montarFiltroStatus();

            processarEAtualizar();

        }
        catch (erro) {

            console.error(
                "Erro ao ler arquivo:",
                erro
            );

            alert(
                "Erro ao ler a planilha.\n\n" +
                "Verifique se o arquivo é um XLSX, XLS ou CSV válido."
            );

        }

    };

reader.readAsArrayBuffer(file);
```

}

/* ============================================================
FILTRO DE STATUS
============================================================ */

function montarFiltroStatus() {

```
const select =
    document.getElementById(
        "statusFilter"
    );

if (!select) {
    return;
}

const statusEncontrados = [];


rawData.forEach(
    function (item) {

        const status =
            normalizarTexto(
                item.STATUS_W
            );

        if (
            status &&
            !statusEncontrados.includes(
                status
            )
        ) {

            statusEncontrados.push(
                status
            );

        }

    }
);


statusEncontrados.sort();


/*
 * Mantém somente:
 *
 * TODOS
 * AGENDADO
 * COLETADO
 * ENTREGUE
 * FINALIZADO
 * PENDENTE DE CONFERÊNCIA
 * PROGRAMADO
 *
 * Não existe mais:
 * AGENDADO E COLETADO
 */

select.innerHTML = "";


const todos =
    document.createElement(
        "option"
    );

todos.value =
    "TODOS";

todos.textContent =
    "Todos os Status";

todos.selected =
    true;

select.appendChild(
    todos
);


const ordemStatus = [
    "AGENDADO",
    "COLETADO",
    "ENTREGUE",
    "FINALIZADO",
    "PENDENTE DE CONFERENCIA",
    "PROGRAMADO"
];


ordemStatus.forEach(
    function (statusDesejado) {

        const encontrado =
            statusEncontrados.find(
                function (status) {

                    if (
                        statusDesejado ===
                        "PENDENTE DE CONFERENCIA"
                    ) {

                        return (
                            status.includes("PENDENTE") ||
                            status.includes("CONFERENCIA")
                        );

                    }

                    return (
                        status ===
                        statusDesejado
                    );

                }
            );


        if (encontrado) {

            const option =
                document.createElement(
                    "option"
                );

            option.value =
                statusDesejado;

            option.textContent =
                statusDesejado;

            select.appendChild(
                option
            );

        }

    }
);
```

}

/* ============================================================
CLASSIFICAÇÃO REGIONAL
============================================================ */

function classificarRegiao(cidade) {

```
const c =
    normalizarTexto(
        cidade
    );


/*
 * SÃO PAULO / ABC
 */

const saoPauloABC = [

    "SAO PAULO",
    "SANTO ANDRE",
    "SAO BERNARDO DO CAMPO",
    "SAO CAETANO DO SUL",
    "DIADEMA",
    "MAUA",
    "RIBEIRAO PIRES",
    "RIO GRANDE DA SERRA"

];


if (
    saoPauloABC.includes(c)
) {

    return "SAO_PAULO_ABC";

}


/*
 * JAGUARIÚNA / ENTORNO
 */

const jaguariunaEntorno = [

    "JAGUARIUNA",
    "CAMPINAS",
    "HOLAMBRA",
    "PEDREIRA",
    "AMPARO",
    "SANTO ANTONIO DE POSSE",
    "MOGI MIRIM",
    "MOGI GUACU",
    "PAULINIA",
    "VALINHOS",
    "VINHEDO",
    "ARTUR NOGUEIRA",
    "COSMOPOLIS",
    "MORUNGABA",
    "ITATIBA",
    "MONTE ALEGRE DO SUL",
    "PINHALZINHO",
    "SERRA NEGRA",
    "LINDOIA",
    "AGUAS DE LINDOIA"

];


if (
    jaguariunaEntorno.includes(c)
) {

    return "JAGUARIUNA_ENTORNO";

}


return "OUTRAS_REGIOES";
```

}

/* ============================================================
NOME DA REGIÃO
============================================================ */

function nomeRegiao(regiao) {

```
switch (regiao) {

    case "SAO_PAULO_ABC":
        return "São Paulo / ABC";

    case "JAGUARIUNA_ENTORNO":
        return "Jaguariúna / Entorno";

    case "OUTRAS_REGIOES":
        return "Outras Regiões";

    default:
        return "Não informado";

}
```

}

/* ============================================================
PEGAR PEDIDO
============================================================ */

function obterPedido(item) {

```
const chave =
    encontrarColuna(
        item,
        COLUNA_PEDIDO
    );

if (
    chave &&
    item[chave] !== undefined &&
    item[chave] !== null
) {

    return String(
        item[chave]
    ).trim();

}

/*
 * Tenta PEDIDO sem _
 */

const chavePedido =
    Object.keys(item).find(
        function (key) {

            const nome =
                normalizarTexto(
                    key
                );

            return (
                nome === "PEDIDO" ||
                nome === "PEDIDO_"
            );

        }
    );


if (chavePedido) {

    return String(
        item[chavePedido] || ""
    ).trim();

}


return "";
```

}

/* ============================================================
PEGAR CIDADE
============================================================ */

function obterCidade(item) {

```
const chave =
    encontrarColuna(
        item,
        COLUNA_CIDADE
    );

return chave
    ? String(
        item[chave] || ""
    ).trim()
    : "";
```

}

/* ============================================================
PEGAR TIPO
============================================================ */

function obterTipo(item) {

```
const chave =
    encontrarColuna(
        item,
        COLUNA_TIPO
    );

return chave
    ? String(
        item[chave] || ""
    ).trim()
    : "";
```

}

/* ============================================================
PEGAR FORNECEDOR
============================================================ */

function obterFornecedor(item) {

```
const chave =
    encontrarColuna(
        item,
        COLUNA_FORNECEDOR
    );

if (chave) {

    return String(
        item[chave] || ""
    ).trim();

}


/*
 * Procura variações
 */

const chaveFornecedor =
    Object.keys(item).find(
        function (key) {

            const nome =
                normalizarTexto(
                    key
                );

            return (
                nome.includes(
                    "FORNECEDOR"
                )
            );

        }
    );


if (chaveFornecedor) {

    return String(
        item[chaveFornecedor] || ""
    ).trim();

}


return "Não informado";
```

}

/* ============================================================
CONSOLIDAR PEDIDOS
============================================================ */

function consolidarPedidos(dados) {

```
const mapa = {};


dados.forEach(
    function (item, index) {

        let pedido =
            obterPedido(item);


        if (!pedido) {

            pedido =
                "LINHA_" +
                item.LINHA_ORIGINAL ||
                index;

        }


        if (!mapa[pedido]) {

            mapa[pedido] = {

                pedido:
                    pedido,

                cidade:
                    obterCidade(item),

                tipo:
                    obterTipo(item),

                fornecedor:
                    obterFornecedor(item),

                data:
                    item.DATA_V,

                status:
                    item.STATUS_W,

                regiao:
                    classificarRegiao(
                        obterCidade(item)
                    ),

                quantidadeLinhas:
                    1,

                linhas:
                    [item]

            };

        }
        else {

            const registro =
                mapa[pedido];

            registro.quantidadeLinhas++;

            registro.linhas.push(
                item
            );


            if (
                !registro.cidade
            ) {

                registro.cidade =
                    obterCidade(item);

            }


            if (
                !registro.tipo
            ) {

                registro.tipo =
                    obterTipo(item);

            }


            if (
                !registro.fornecedor
            ) {

                registro.fornecedor =
                    obterFornecedor(item);

            }

        }

    }
);


return Object.values(
    mapa
);
```

}

/* ============================================================
PROCESSAMENTO PRINCIPAL
============================================================ */

function processarEAtualizar() {

```
if (
    !rawData ||
    rawData.length === 0
) {

    return;

}


const filtroStatus =
    document.getElementById(
        "statusFilter"
    )?.value ||
    "TODOS";


const filtroRegiao =
    document.getElementById(
        "regionFilter"
    )?.value ||
    "TODAS";


const dataInicial =
    document.getElementById(
        "startDate"
    )?.value ||
    "";


const dataFinal =
    document.getElementById(
        "endDate"
    )?.value ||
    "";


const timestampInicial =
    dataInicial
        ? obterTimestampZerado(
            dataInicial
        )
        : null;


const timestampFinal =
    dataFinal
        ? obterTimestampZerado(
            dataFinal
        )
        : null;


const dadosFiltrados =
    rawData.filter(
        function (item) {

            const status =
                normalizarTexto(
                    item.STATUS_W
                );


            /*
             * STATUS
             */

            if (
                filtroStatus !==
                "TODOS"
            ) {

                if (
                    filtroStatus ===
                    "PENDENTE DE CONFERENCIA" ||
                    filtroStatus ===
                    "PENDENTE_DE_CONFERENCIA"
                ) {

                    if (
                        !status.includes(
                            "PENDENTE"
                        ) &&
                        !status.includes(
                            "CONFERENCIA"
                        )
                    ) {

                        return false;

                    }

                }
                else {

                    if (
                        status !==
                        normalizarTexto(
                            filtroStatus
                        )
                    ) {

                        return false;

                    }

                }

            }


            /*
             * REGIÃO
             */

            if (
                filtroRegiao !==
                "TODAS"
            ) {

                const regiao =
                    classificarRegiao(
                        obterCidade(item)
                    );

                if (
                    regiao !==
                    filtroRegiao
                ) {

                    return false;

                }

            }


            /*
             * DATA
             */

            if (
                timestampInicial !== null ||
                timestampFinal !== null
            ) {

                const timestampItem =
                    obterTimestampZerado(
                        item.DATA_V
                    );


                if (
                    timestampItem === null
                ) {

                    return false;

                }


                if (
                    timestampInicial !== null &&
                    timestampItem <
                    timestampInicial
                ) {

                    return false;

                }


                if (
                    timestampFinal !== null &&
                    timestampItem >
                    timestampFinal
                ) {

                    return false;

                }

            }


            return true;

        }
    );


pedidosConsolidados =
    consolidarPedidos(
        dadosFiltrados
    );


atualizarKPIs(
    pedidosConsolidados
);


atualizarDistribuicaoRegional(
    pedidosConsolidados
);


atualizarGraficoCidades(
    pedidosConsolidados
);


atualizarGraficoTipos(
    pedidosConsolidados
);


atualizarTabelaDatas(
    dadosFiltrados
);
```

}

/* ============================================================
KPIs
============================================================ */

function atualizarKPIs(pedidos) {

```
const total =
    pedidos.length;

let agendados = 0;
let coletas = 0;
let entregas = 0;
let finalizados = 0;
let pendentes = 0;
let programados = 0;


pedidos.forEach(
    function (pedido) {

        const status =
            normalizarTexto(
                pedido.status
            );


        if (
            status ===
            "AGENDADO"
        ) {

            agendados++;

        }
        else if (
            status ===
            "COLETADO"
        ) {

            coletas++;

        }
        else if (
            status ===
            "ENTREGUE"
        ) {

            entregas++;

        }
        else if (
            status ===
            "FINALIZADO"
        ) {

            finalizados++;

        }
        else if (
            status.includes(
                "PENDENTE"
            ) ||
            status.includes(
                "CONFERENCIA"
            )
        ) {

            pendentes++;

        }
        else if (
            status ===
            "PROGRAMADO"
        ) {

            programados++;

        }

    }
);


const painel =
    document.getElementById(
        "painelKPIs"
    );

if (!painel) {
    return;
}


painel.innerHTML = `

    <div class="kpi-card">

        <span>
            TOTAL DE PEDIDOS
        </span>

        <strong>
            ${total}
        </strong>

        <small>
            Base filtrada
        </small>

    </div>


    <div class="kpi-card">

        <span>
            AGENDADOS
        </span>

        <strong>
            ${agendados}
        </strong>

        <small>
            ${percentual(
                agendados,
                total
            )}
        </small>

    </div>


    <div class="kpi-card">

        <span>
            COLETADOS
        </span>

        <strong>
            ${coletas}
        </strong>

        <small>
            ${percentual(
                coletas,
                total
            )}
        </small>

    </div>


    <div class="kpi-card">

        <span>
            ENTREGUES / FINALIZADOS
        </span>

        <strong>
            ${entregas + finalizados}
        </strong>

        <small>
            Entregues: ${entregas}
            |
            Finalizados: ${finalizados}
        </small>

    </div>

`;
```

}

/* ============================================================
PERCENTUAL
============================================================ */

function percentual(valor, total) {

```
if (!total) {
    return "0.0%";
}

return (
    (
        valor /
        total
    ) *
    100
).toFixed(1) +
"%";
```

}

/* ============================================================
DISTRIBUIÇÃO REGIONAL
============================================================ */

function atualizarDistribuicaoRegional(
pedidos
) {

```
const painel =
    document.getElementById(
        "painelRegioes"
    );

if (!painel) {
    return;
}


const total =
    pedidos.length;


let saoPaulo = 0;
let jaguariuna = 0;
let outras = 0;


pedidos.forEach(
    function (pedido) {

        if (
            pedido.regiao ===
            "SAO_PAULO_ABC"
        ) {

            saoPaulo++;

        }
        else if (
            pedido.regiao ===
            "JAGUARIUNA_ENTORNO"
        ) {

            jaguariuna++;

        }
        else {

            outras++;

        }

    }
);


painel.innerHTML = `

    <div class="card">

        <h3>
            Distribuição Regional
        </h3>


        <div class="regiao-linha">

            <span>
                São Paulo / ABC
            </span>

            <div>

                <strong>
                    ${saoPaulo}
                </strong>

                <small>
                    ${percentual(
                        saoPaulo,
                        total
                    )}
                </small>

            </div>

        </div>


        <div class="regiao-linha">

            <span>
                Jaguariúna / Entorno
            </span>

            <div>

                <strong>
                    ${jaguariuna}
                </strong>

                <small>
                    ${percentual(
                        jaguariuna,
                        total
                    )}
                </small>

            </div>

        </div>


        <div class="regiao-linha">

            <span>
                Outras Regiões
            </span>

            <div>

                <strong>
                    ${outras}
                </strong>

                <small>
                    ${percentual(
                        outras,
                        total
                    )}
                </small>

            </div>

        </div>

    </div>

`;
```

}

/* ============================================================
GRÁFICO DE CIDADES
============================================================ */

function atualizarGraficoCidades(
pedidos
) {

```
const contagem = {};


pedidos.forEach(
    function (pedido) {

        const cidade =
            normalizarTexto(
                pedido.cidade
            ) ||
            "NAO INFORMADO";


        contagem[cidade] =
            (
                contagem[cidade] ||
                0
            ) + 1;

    }
);


const cidades =
    Object.entries(
        contagem
    ).sort(
        function (a, b) {

            return (
                b[1] -
                a[1]
            );

        }
    );


const labels =
    cidades.map(
        function (item) {
            return item[0];
        }
    );


const valores =
    cidades.map(
        function (item) {
            return item[1];
        }
    );


renderizarGraficoBarras(
    labels,
    valores
);
```

}

/* ============================================================
GRÁFICO DE BARRAS
============================================================ */

function renderizarGraficoBarras(
labels,
valores
) {

```
const canvas =
    document.getElementById(
        "chartRegiao"
    );

if (!canvas) {
    return;
}


const ctx =
    canvas.getContext(
        "2d"
    );


if (
    instanceChartRegiao
) {

    instanceChartRegiao.destroy();

}


instanceChartRegiao =
    new Chart(
        ctx,
        {

            type: "bar",

            data: {

                labels: labels,

                datasets: [

                    {

                        label:
                            "Qtd. de Pedidos",

                        data:
                            valores,

                        backgroundColor:
                            "#2563eb",

                        borderRadius:
                            5

                    }

                ]

            },


            options: {

                responsive:
                    true,

                maintainAspectRatio:
                    false,

                plugins: {

                    legend: {
                        display: false
                    },

                    datalabels: {

                        anchor:
                            "end",

                        align:
                            "top",

                        color:
                            "#0f172a",

                        font: {

                            family:
                                "'Inter', sans-serif",

                            weight:
                                "bold",

                            size:
                                12

                        },

                        formatter:
                            function (
                                value
                            ) {

                                return value >
                                    0
                                    ? value
                                    : "";

                            }

                    }

                },


                scales: {

                    y: {

                        beginAtZero:
                            true,

                        ticks: {

                            precision:
                                0

                        },

                        grace:
                            "15%"

                    }

                }

            }

        }
    );
```

}

/* ============================================================
GRÁFICO POR TIPO
============================================================ */

function atualizarGraficoTipos(
pedidos
) {

```
const contagem = {};


pedidos.forEach(
    function (pedido) {

        const tipo =
            normalizarTexto(
                pedido.tipo
            ) ||
            "OUTROS";


        contagem[tipo] =
            (
                contagem[tipo] ||
                0
            ) + 1;

    }
);


renderizarGraficoTipos(
    Object.keys(
        contagem
    ),
    Object.values(
        contagem
    )
);
```

}

/* ============================================================
GRÁFICO DE TIPOS
============================================================ */

function renderizarGraficoTipos(
labels,
valores
) {

```
const canvas =
    document.getElementById(
        "chartTipo"
    );

if (!canvas) {
    return;
}


const ctx =
    canvas.getContext(
        "2d"
    );


if (
    instanceChartTipo
) {

    instanceChartTipo.destroy();

}


instanceChartTipo =
    new Chart(
        ctx,
        {

            type: "doughnut",

            data: {

                labels:
                    labels,

                datasets: [

                    {

                        data:
                            valores,

                        backgroundColor: [

                            "#2563eb",
                            "#10b981",
                            "#f59e0b",
                            "#8b5cf6",
                            "#ef4444",
                            "#64748b",
                            "#06b6d4"

                        ]

                    }

                ]

            },


            options: {

                responsive:
                    true,

                maintainAspectRatio:
                    false,

                plugins: {

                    legend: {

                        position:
                            "top"

                    },


                    datalabels: {

                        color:
                            "#ffffff",

                        font: {

                            family:
                                "'Inter', sans-serif",

                            weight:
                                "bold",

                            size:
                                13

                        },

                        formatter:
                            function (
                                value
                            ) {

                                return value >
                                    0
                                    ? value
                                    : "";

                            }

                    }

                }

            }

        }
    );
```

}

/* ============================================================
TABELA INTERATIVA
============================================================ */

function atualizarTabelaDatas(
dadosFiltrados
) {

```
const tbody =
    document.querySelector(
        "#tabelaAgenda tbody"
    );

if (!tbody) {
    return;
}


/*
 * Agrupa por:
 *
 * DATA
 * REGIÃO
 * STATUS
 */

const grupos = {};


dadosFiltrados.forEach(
    function (item, index) {

        const data =
            obterTimestampZerado(
                item.DATA_V
            );


        const dataTexto =
            data !== null
                ? formatarData(data)
                : "Sem Data";


        const cidade =
            obterCidade(item);


        const regiao =
            classificarRegiao(
                cidade
            );


        const status =
            normalizarTexto(
                item.STATUS_W
            ) ||
            "SEM STATUS";


        const chave =
            dataTexto +
            "|" +
            regiao +
            "|" +
            status;


        if (!grupos[chave]) {

            grupos[chave] = {

                data:
                    dataTexto,

                timestamp:
                    data || 9999999999999,

                regiao:
                    regiao,

                status:
                    status,

                pedidos:
                    [],

                quantidade:
                    0

            };

        }


        grupos[chave]
            .pedidos
            .push(item);


        grupos[chave]
            .quantidade++;

    }
);


const gruposOrdenados =
    Object.values(
        grupos
    ).sort(
        function (a, b) {

            return (
                a.timestamp -
                b.timestamp
            );

        }
    );


if (
    gruposOrdenados.length === 0
) {

    tbody.innerHTML = `

        <tr>

            <td
                colspan="4"
                class="empty-table"
            >
                Nenhum pedido encontrado
                para os filtros selecionados.
            </td>

        </tr>

    `;

    return;

}


let html = "";


gruposOrdenados.forEach(
    function (grupo, index) {

        const grupoId =
            "grupo-" +
            index;


        html += `

            <tr class="agenda-row">

                <td>

                    <strong>
                        ${grupo.data}
                    </strong>

                </td>


                <td>

                    <span
                        class="region-badge"
                    >
                        ${nomeRegiao(
                            grupo.regiao
                        )}
                    </span>

                </td>


                <td>

                    <span
                        class="status-badge"
                    >
                        ${grupo.status}
                    </span>

                </td>


                <td>

                    <button
                        type="button"
                        class="btn-expand"
                        onclick="alternarDetalhes('${grupoId}')"
                    >

                        <strong>
                            ${grupo.quantidade}
                        </strong>

                        <span>
                            pedidos
                        </span>

                        <span
                            class="arrow"
                            id="arrow-${grupoId}"
                        >
                            ▾
                        </span>

                    </button>

                </td>

            </tr>


            <tr
                id="${grupoId}"
                class="details-row"
                style="display:none;"
            >

                <td
                    colspan="4"
                >

                    <div
                        class="details-container"
                    >

                        <table
                            class="details-table"
                        >

                            <thead>

                                <tr>

                                    <th>
                                        Nº Pedido
                                    </th>

                                    <th>
                                        Fornecedor
                                    </th>

                                    <th>
                                        Cidade
                                    </th>

                                    <th>
                                        Tipo
                                    </th>

                                    <th>
                                        Status
                                    </th>

                                    <th>
                                        Linha
                                    </th>

                                </tr>

                            </thead>

                            <tbody>

                                ${grupo.pedidos.map(
                                    function (item) {

                                        return `

                                            <tr>

                                                <td>

                                                    <strong>
                                                        ${escapeHtml(
                                                            obterPedido(item) ||
                                                            "Não informado"
                                                        )}
                                                    </strong>

                                                </td>

                                                <td>
                                                    ${escapeHtml(
                                                        obterFornecedor(item)
                                                    )}
                                                </td>

                                                <td>
                                                    ${escapeHtml(
                                                        obterCidade(item) ||
                                                        "Não informado"
                                                    )}
                                                </td>

                                                <td>
                                                    ${escapeHtml(
                                                        obterTipo(item) ||
                                                        "Não informado"
                                                    )}
                                                </td>

                                                <td>
                                                    ${escapeHtml(
                                                        String(
                                                            item.STATUS_W ||
                                                            "Sem status"
                                                        )
                                                    )}
                                                </td>

                                                <td>
                                                    ${item.LINHA_ORIGINAL || "-"}
                                                </td>

                                            </tr>

                                        `;

                                    }
                                ).join("")}

                            </tbody>

                        </table>

                    </div>

                </td>

            </tr>

        `;

    }
);


tbody.innerHTML =
    html;
```

}

/* ============================================================
ABRIR / FECHAR DETALHES
============================================================ */

function alternarDetalhes(
grupoId
) {

```
const linha =
    document.getElementById(
        grupoId
    );

const seta =
    document.getElementById(
        "arrow-" +
        grupoId
    );


if (!linha) {
    return;
}


if (
    linha.style.display ===
    "none"
) {

    linha.style.display =
        "table-row";

    if (seta) {
        seta.textContent =
            "▴";
    }

}
else {

    linha.style.display =
        "none";

    if (seta) {
        seta.textContent =
            "▾";
    }

}
```

}

/* ============================================================
DATA
============================================================ */

function obterTimestampZerado(
valor
) {

```
if (
    valor === undefined ||
    valor === null ||
    valor === ""
) {

    return null;

}


let dia = null;
let mes = null;
let ano = null;


/*
 * Date
 */

if (
    valor instanceof Date
) {

    dia =
        valor.getDate();

    mes =
        valor.getMonth();

    ano =
        valor.getFullYear();

}


/*
 * Número serial do Excel
 */

else if (
    typeof valor === "number"
) {

    const dataExcel =
        XLSX.SSF.parse_date_code(
            valor
        );


    if (!dataExcel) {
        return null;
    }


    dia =
        dataExcel.d;

    mes =
        dataExcel.m - 1;

    ano =
        dataExcel.y;

}


/*
 * Texto
 */

else {

    let texto =
        String(valor)
            .trim();


    texto =
        texto
            .split(" ")[0]
            .split("T")[0];


    /*
     * DD/MM/YYYY
     */

    if (
        texto.includes("/")
    ) {

        const partes =
            texto.split("/");


        if (
            partes.length === 3
        ) {

            dia =
                parseInt(
                    partes[0],
                    10
                );

            mes =
                parseInt(
                    partes[1],
                    10
                ) - 1;

            ano =
                parseInt(
                    partes[2],
                    10
                );

        }

    }


    /*
     * YYYY-MM-DD
     */

    else if (
        texto.includes("-")
    ) {

        const partes =
            texto.split("-");


        if (
            partes.length === 3
        ) {

            ano =
                parseInt(
                    partes[0],
                    10
                );

            mes =
                parseInt(
                    partes[1],
                    10
                ) - 1;

            dia =
                parseInt(
                    partes[2],
                    10
                );

        }

    }


    /*
     * YYYYMMDD
     */

    else if (
        /^\d{8}$/.test(
            texto
        )
    ) {

        ano =
            parseInt(
                texto.substring(
                    0,
                    4
                ),
                10
            );

        mes =
            parseInt(
                texto.substring(
                    4,
                    6
                ),
                10
            ) - 1;

        dia =
            parseInt(
                texto.substring(
                    6,
                    8
                ),
                10
            );

    }

}


if (
    dia === null ||
    mes === null ||
    ano === null ||
    isNaN(dia) ||
    isNaN(mes) ||
    isNaN(ano)
) {

    return null;

}


const data =
    new Date(
        ano,
        mes,
        dia,
        0,
        0,
        0,
        0
    );


if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes ||
    data.getDate() !== dia
) {

    return null;

}


return data.getTime();
```

}

/* ============================================================
FORMATAR DATA
============================================================ */

function formatarData(
timestamp
) {

```
if (!timestamp) {
    return "";
}


const data =
    new Date(
        timestamp
    );


const dia =
    String(
        data.getDate()
    ).padStart(
        2,
        "0"
    );


const mes =
    String(
        data.getMonth() + 1
    ).padStart(
        2,
        "0"
    );


const ano =
    data.getFullYear();


return (
    dia +
    "/" +
    mes +
    "/" +
    ano
);
```

}

/* ============================================================
ESCAPE HTML
============================================================ */

function escapeHtml(valor) {

```
return String(
    valor || ""
)
    .replace(
        /&/g,
        "&amp;"
    )
    .replace(
        /</g,
        "&lt;"
    )
    .replace(
        />/g,
        "&gt;"
    )
    .replace(
        /"/g,
        "&quot;"
    )
    .replace(
        /'/g,
        "&#039;"
    );
```

}

/* ============================================================
LIMPAR FILTROS
============================================================ */

function limparFiltros() {

```
const status =
    document.getElementById(
        "statusFilter"
    );

const regiao =
    document.getElementById(
        "regionFilter"
    );

const inicio =
    document.getElementById(
        "startDate"
    );

const fim =
    document.getElementById(
        "endDate"
    );


if (status) {
    status.value =
        "TODOS";
}


if (regiao) {
    regiao.value =
        "TODAS";
}


if (inicio) {
    inicio.value =
        "";
}


if (fim) {
    fim.value =
        "";
}


processarEAtualizar();
```

}
