from src.database_builder import DatabaseBuilder
from src.dashboard_builder import DashboardBuilder
from src.styles import AppStyles
from src.pdf_generator import generate_pdf_report
import openpyxl


def main():
    wb = openpyxl.Workbook()
    styles = AppStyles()

    # 1. Constrói o banco de dados interno
    db_builder = DatabaseBuilder(wb, styles)
    db_builder.build()

    # 2. Constrói o painel/dashboard com gráficos embutidos
    dash_builder = DashboardBuilder(wb, styles)
    dash_builder.build()

    # Salva o Excel
    output_excel = "investment_simulator.xlsx"
    wb.save(output_excel)
    print(f"Sucesso absoluto! Simulador gerado em '{output_excel}'.")

    # 3. Definição dos perfis para os relatórios em PDF
    perfis = [
        {"nome": "Conservador", "taxa": 0.09},
        {"nome": "Moderado", "taxa": 0.11},
        {"nome": "Agressivo", "taxa": 0.13},
    ]

    # 4. Loop para gerar os 3 relatórios em PDF automaticamente
    for p in perfis:
        generate_pdf_report(
            perfil=p["nome"],
            salario=2000,
            aporte_mensal=600,
            taxa_anual_media=p["taxa"],
            filename=f"relatorio_{p['nome'].lower()}.pdf",
        )
    print("Relatórios PDF gerados com sucesso para todos os perfis!")


if __name__ == "__main__":
    main()
