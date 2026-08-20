import matplotlib.pyplot as plt


def generate_pdf_report(
    perfil,
    salario,
    aporte_mensal,
    taxa_anual_media,
    filename="relatorio_investimentos.pdf",
):
    # Dados base para simulação no relatório
    ativos = [
        "RENDA FIXA",
        "IMOBILIÁRIO",
        "MULTIMERCADO",
        "AÇÕES BR",
        "INTERNACIONAL",
        "CRIPTOMOEDAS",
    ]

    # Pesos padrão para o perfil Moderado como exemplo visual
    pesos = [0.35, 0.32, 0.08, 0.10, 0.10, 0.05]
    valores_alocados = [aporte_mensal * p for p in pesos]

    anos = [2, 5, 10, 20, 30]
    # Cálculo aproximado de juros compostos para o PDF
    taxa_mensal = (1 + taxa_anual_media) ** (1 / 12) - 1
    patrimonio = [
        aporte_mensal * (((1 + taxa_mensal) ** (a * 12) - 1) / taxa_mensal)
        for a in anos
    ]

    # Configuração do Layout do PDF
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(
        f"Relatório de Investimentos - Perfil: {perfil.upper()}",
        fontsize=16,
        fontweight="bold",
        color="#1f4e78",
    )

    # Gráfico 1: Distribuição de Alocação
    ax1.barh(ativos, valores_alocados, color="#2b579a")
    ax1.set_title(
        "Aporte Mensal por Classe de Ativo (R$)", fontsize=12, fontweight="bold"
    )
    ax1.set_xlabel("Valor em Reais (R$)")
    ax1.grid(axis="x", linestyle="--", alpha=0.7)

    # Gráfico 2: Projeção de Longo Prazo
    ax2.plot(anos, patrimonio, marker="o", color="#2ca02c", linewidth=2.5, markersize=6)
    ax2.set_title("Patrimônio Acumulado (Longo Prazo)", fontsize=12, fontweight="bold")
    ax2.set_xlabel("Anos")
    ax2.set_ylabel("Patrimônio Total (R$)")
    ax2.grid(True, linestyle="--", alpha=0.7)

    plt.tight_layout()

    # Salva como PDF executivo
    pdf_filename = "relatorio_investimentos_executivo.pdf"
    plt.savefig(filename, format="pdf", dpi=300)
    plt.close()
    print(f"Relatório PDF gerado: {filename}")
