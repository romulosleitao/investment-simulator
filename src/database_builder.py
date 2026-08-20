class DatabaseBuilder:
    def __init__(self, wb, styles):
        self.wb = wb
        self.styles = styles

        # Premissas baseadas estritamente no spec.md
        self.assets_data = {
            "RENDA FIXA": 0.10,
            "IMOBILIÁRIO": 0.11,
            "MULTIMERCADO": 0.105,
            "AÇÕES BR": 0.12,
            "INTERNACIONAL": 0.115,
            "CRIPTOMOEDAS": 0.18,
        }

        self.profiles_data = {
            "Conservador": {
                "RENDA FIXA": 0.50,
                "IMOBILIÁRIO": 0.30,
                "MULTIMERCADO": 0.10,
                "AÇÕES BR": 0.10,
                "INTERNACIONAL": 0.00,
                "CRIPTOMOEDAS": 0.00,
            },
            "Moderado": {
                "RENDA FIXA": 0.35,
                "IMOBILIÁRIO": 0.32,
                "MULTIMERCADO": 0.08,
                "AÇÕES BR": 0.10,
                "INTERNACIONAL": 0.10,
                "CRIPTOMOEDAS": 0.05,
            },
            "Agressivo": {
                "RENDA FIXA": 0.10,
                "IMOBILIÁRIO": 0.20,
                "MULTIMERCADO": 0.05,
                "AÇÕES BR": 0.35,
                "INTERNACIONAL": 0.20,
                "CRIPTOMOEDAS": 0.10,
            },
        }

    def build(self):
        ws_db = self.wb.active
        ws_db.title = "Database_Profiles"
        ws_db.views.sheetView[0].showGridLines = True

        headers_db = ["CHAVE", "PERFIL", "TIPO DE INVESTIMENTO", "%", "RETORNO AA"]
        for col_idx, h in enumerate(headers_db, start=1):
            cell = ws_db.cell(row=1, column=col_idx, value=h)
            cell.font = self.styles.header_font
            cell.fill = self.styles.header_fill

        row_idx = 2
        for profile, weights in self.profiles_data.items():
            for asset, weight in weights.items():
                key = f"{profile}-{asset}"
                retorno = self.assets_data[asset]

                ws_db.cell(row=row_idx, column=1, value=key).font = (
                    self.styles.regular_font
                )
                ws_db.cell(row=row_idx, column=2, value=profile).font = (
                    self.styles.regular_font
                )
                ws_db.cell(row=row_idx, column=3, value=asset).font = (
                    self.styles.regular_font
                )

                c_w = ws_db.cell(row=row_idx, column=4, value=weight)
                c_w.font = self.styles.regular_font
                c_w.number_format = "0.0%"

                c_r = ws_db.cell(row=row_idx, column=5, value=retorno)
                c_r.font = self.styles.regular_font
                c_r.number_format = "0.0%"
                row_idx += 1

        # Perfis únicos para Validação de Dados na coluna F
        perfis_unicos = list(self.profiles_data.keys())
        ws_db["F1"] = "Perfis_Unicos"
        for idx, perfil in enumerate(perfis_unicos, start=2):
            ws_db[f"F{idx}"] = perfil

        return ws_db
