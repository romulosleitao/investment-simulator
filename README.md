```markdown
# 📊 Simulador de Investimentos Multiclasse (Specification-Driven Modular Architecture)

Simulador financeiro profissional desenvolvido em Python utilizando a biblioteca `openpyxl`, estruturado sob o conceito de desenvolvimento orientado a especificações (*Specification-Driven Development*). O projeto foi concebido e arquitetado em uma colaboração de engenharia de software entre o desenvolvedor e a inteligência artificial **Gemini**, garantindo modularidade, manutenibilidade e total compatibilidade nativa com o **Excel na Nuvem (Excel Online / Office 365 Web)**.

---

## 🛠️ Arquitetura do Projeto

O código foi totalmente desacoplado do modelo monolítico tradicional, sendo dividido em camadas de responsabilidade única (Orientação a Objetos):

```text
/investment-simulator
│
├── main.py                  # Orquestrador principal (executa a geração completa)
├── spec.md                  # Documento de Especificação Funcional e Regras de Negócio
├── investmentsimulator.xlsx # Planilha gerada pronta para uso na nuvem
├── pyrightconfig.json       # Configuração de linting e caminhos do editor
└── src/
    ├── __init__.py          # Inicializador do pacote Python
    ├── styles.py            # Gerenciamento centralizado de cores, fontes e preenchimentos
    ├── database_builder.py  # Constrói a aba de backend com a matriz de perfis e ativos
    └── dashboard_builder.py # Constrói o painel APP (fórmulas internacionais, validações e cenários)

```

---

## 📐 Regras de Negócio Implementadas (`spec.md`)

1. **Entradas Macro e Capacidade de Aporte:**
* O usuário define o **Salário Base** (ex: R$ 2.000,00) e a **Taxa de Poupança Sugerida** (ex: 30%), resultando no aporte mensal base.


2. **Backend de Ativos e Perfis (`Database_Profiles`):**
* Contém 6 classes de ativos principais: *Renda Fixa, Imobiliário, Multimercado, Ações BR, Internacional e Criptomoedas*.
* Três perfis de risco (*Conservador, Moderado e Agressivo*) com pesos percentuais específicos e taxas de retorno anual de mercado associadas.


3. **Motor de Rentabilidade Ponderada (Blended Rate):**
* A taxa média mensal da carteira é calculada dinamicamente cruzando os pesos alocados com os retornos esperados de cada ativo por meio da função internacional `SUMPRODUCT` dividida por 12.


4. **Alocação Dinâmica por Ativo:**
* Utiliza a função `XLOOKUP` em conjunto com chaves compostas (`[Perfil]-[Ativo]`) para buscar instantaneamente os percentuais na aba de backend.


5. **Projeções de Longo Prazo (Cenários):**
* Utiliza a função financeira de Valor Futuro (`FV`) para projetar o acúmulo de patrimônio com juros compostos em horizontes de 2, 5, 10, 20 e 30 anos.

---
## 🚀 Processo de Construção e Execução

### Pré-requisitos

* Python instalado na máquina.
* Biblioteca `openpyxl` instalada (`pip install openpyxl`).

### Como gerar a planilha:

No terminal, na pasta raiz do projeto, execute o orquestrador:

```bash
python main.py

```

O script compilará todas as camadas, aplicará as formatações corporativas em tons de azul e salvará o arquivo atualizado na raiz.

---

## 🌐 Compatibilidade com Excel Online (Nuvem / OneDrive)

As fórmulas geradas pelo script utilizam o padrão internacional de engenharia exigido pelo Excel na nuvem (`XLOOKUP`, `SUMPRODUCT`, `FV`, `SUM`) com separadores baseados em **vírgulas (`,`)**, garantindo abertura imediata, sem erros de sintaxe e com cálculo nativo em tempo real no OneDrive.

-----------------

---
Dentro do processo de criação foram utilizadas Gemini como assistente para o processo de construção de arquivos. 
```
