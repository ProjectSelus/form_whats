// ==========================================
// AGENDA DE OPERAÇÕES
// Formato obrigatório da data: AAAA-MM-DD
// Adicionar um dia a mais no fim da operação, pois os dados do resumo diario são finalizadoa penas no dia posterior.
// ==========================================
const agendaOperacoes = [
    {
        nome: "Op. Integrada Mulher Segura",
        inicio: "2026-09-01",
        fim: "2027-01-01"
    },
    {
        nome: "Op. PRODETOP 3B",
        inicio: "2026-09-03",
        fim: "2026-09-05"
    },
    {
        nome: "Op. PROVIAVIDAS II 3A",
        inicio: "2026-11-14",
        fim: "2026-11-26"
    }
    // Para adicionar mais, basta colocar uma vírgula no fechamento da chave } e criar um novo bloco.
];

// ==========================================
// LÓGICA DE ATIVAÇÃO AUTOMÁTICA
// Não é necessário alterar o código abaixo
// ==========================================
function obterOperacoesVigentes() {
    // Captura a data exata do navegador/dispositivo do usuário
    const hoje = new Date();
    const ano = hoje.getFullYear();
    const mes = String(hoje.getMonth() + 1).padStart(2, '0');
    const dia = String(hoje.getDate()).padStart(2, '0');
    const dataHojeStr = `${ano}-${mes}-${dia}`;

    // Filtra apenas as operações onde a data de hoje está entre o início e o fim
    const ativas = agendaOperacoes.filter(op => {
        return dataHojeStr >= op.inicio && dataHojeStr <= op.fim;
    });

    // Se nenhuma operação estiver ativa hoje, retorna 0
    if (ativas.length === 0) {
        return "0";
    }

    // Monta a string final no padrão "02: Nome da Op - Outra Op"
    const nomes = ativas.map(op => op.nome).join(" - ");
    const quantidade = ativas.length < 10 ? `0${ativas.length}` : ativas.length;

    return `${quantidade}: ${nomes}`;
}

// Exporta a variável para o index.html ler
const OPERACOES_VIGENTES = obterOperacoesVigentes();
