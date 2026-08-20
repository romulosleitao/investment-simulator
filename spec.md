"""
=============================================================================
SPECIFICATION (SPEC.PY) - SIMULADOR DE INVESTIMENTOS MULTICLASSE
=============================================================================
Este documento define formalmente as regras de negócio, premissas matemáticas,
fontes de dados e arquitetura de cálculo do simulador financeiro.
"""

PROJECT_SPEC = {
    "name": "Simulador de Investimentos Multiclasse",
    "version": "2.0.0",
    "architecture": "Modular (Specification-Driven Development)",
    
    # -------------------------------------------------------------------------
    # 1. PREMISSAS MACROECONÔMICAS E DE ATIVOS (Backend Database)
    # -------------------------------------------------------------------------
    "assets_database": {
        "description": "Define as classes de ativos, pesos por perfil de risco e retornos esperados anuais.",
        "assets": [
            {"tipo": "RENDA FIXA", "retorno_aa": 0.10},
            {"tipo": "IMOBILIÁRIO", "retorno_aa": 0.11},
            {"tipo": "MULTIMERCADO", "retorno_aa": 0.105},
            {"tipo": "AÇÕES BR", "retorno_aa": 0.12},
            {"tipo": "INTERNACIONAL", "retorno_aa": 0.115},
            {"tipo": "CRIPTOMOEDAS", "retorno_aa": 0.18}
        ],
        "profiles": {
            "Conservador": {"RENDA FIXA": 0.50, "IMOBILIÁRIO": 0.30, "MULTIMERCADO": 0.10, "AÇÕES BR": 0.10, "INTERNACIONAL": 0.00, "CRIPTOMOEDAS": 0.00},
            "Moderado":    {"RENDA FIXA": 0.35, "IMOBILIÁRIO": 0.32, "MULTIMERCADO": 0.08, "AÇÕES BR": 0.10, "INTERNACIONAL": 0.10, "CRIPTOMOEDAS": 0.05},
            "Agressivo":   {"RENDA FIXA": 0.10, "IMOBILIÁRIO": 0.20, "MULTIMERCADO": 0.05, "AÇÕES BR": 0.35, "INTERNACIONAL": 0.20, "CRIPTOMOEDAS": 0.10}
        }
    },

    # -------------------------------------------------------------------------
    # 2. REGRAS DE CÁLCULO E FLUXO DE DADOS (Business Rules)
    # -------------------------------------------------------------------------
    "business_rules": {
        
        "blended_rate_formula": {
            "description": "Taxa de Retorno Ponderada da Carteira (ao mês).",
            "formula": "Somatório de (Peso do Ativo * Retorno Mensal do Ativo)",
            "conversion_to_monthly": "R_am = (1 + R_aa)^(1/12) - 1"
        },
        
        "monthly_cash_flow": {
            "description": "Extrato de Acúmulo Mensal (Mês a Mês).",
            "step_1_aporte": "O usuário define o Salário (ex: R$ 2.000,00) e o Percentual de Poupança (ex: 30%), resultando no Aporte Mensal Base (ex: R$ 600,00).",
            "step_2_rendimento_mes": "Rendimento do Mês = (Saldo Inicial + Aporte Mensal) * Taxa Mensal Ponderada do Perfil.",
            "step_3_saldo_final": "Saldo Final do Mês = Saldo Inicial + Aporte Mensal + Rendimento do Mês."
        },
        
        "long_term_projections": {
            "description": "Projeções de Cenários de Longo Prazo (Anos).",
            "horizons": [2, 5, 10, 20, 30],
            "total_invested": "Aporte Mensal * Anos * 12 (Capital puro aportado pelo usuário)",
            "future_value_patrimonio": "Função VF do Excel aplicada com a Taxa Ponderada Mensal, Prazo em Meses e Aporte."
        },
        
        "asset_allocation": {
            "description": "Distribuição do valor mensal aportado por classe de ativo.",
            "methodology": "Busca dinâmica via PROCX cruzando [Perfil Selecionado] & '-' & [Tipo de Ativo] na base de dados, multiplicando o percentual resultante pelo Valor Total Investido no Mês."
        }
    },

    # -------------------------------------------------------------------------
    # 3. ESTRUTURA VISUAL DA PLANILHA (UI / Layout Spec)
    # -------------------------------------------------------------------------
    "layout_specification": {
        "sheet_1": "APP (Dashboard Interativo)",
        "blocks_order": [
            "Bloco 1: Configurações Globais (Salário, Rendimento Base, Sugestão 30%)",
            "Bloco 2: Simulador de Aporte e Seleção de Perfil (Com Validação de Dados)",
            "Bloco 3: Alocação de Ativos por Perfil (Tabela com PROCX e SOMA totalizadora)",
            "Bloco 4: Projeções de Cenários de Longo Prazo (Tabela baseada no extrato e juros compostos)"
        ],
        "sheet_2": "Database_Profiles (Tabela de Apoio e Matriz de Dados de Ativos)"
    }
}

if __name__ == "__main__":
    print(f"Especificação carregada com sucesso para: {PROJECT_SPEC['name']}")
    print(f"Arquitetura baseada em: {PROJECT_SPEC['architecture']}")