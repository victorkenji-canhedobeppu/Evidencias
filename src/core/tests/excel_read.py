# Ficheiro de teste corrigido: excel_read_test.py
# VERSÃO FINAL: A filtragem horizontal para a aba "Trado" agora deteta
# e processa corretamente todas as colunas de medição (MEDIÇÃO, MEDIÇÃO_2, etc.).

import datetime
from datetime import datetime
import pandas as pd
import numpy as np

SHEET_COLUMN_MAP = {
    # O seu SHEET_COLUMN_MAP permanece o mesmo.
    # Certifique-se de que ele inclui todos os cabeçalhos que você quer no output final.
    "PERCUSSÃO": [
        "KM",
        "SONDAGEM A PERCUSSÃO",
        "PROFUNDIDADE EXECUTADA (m)",
        "MEDIÇÃO",
        "NORTE",
        "LESTE",
    ],
    "TRADO": [
        "KM",
        "SONDAGEM A TRADO",
        "NORTE",
        "LESTE",
        "PROFUNDIDADE EXECUTADA (m)",
        "MEDIÇÃO",
        "UMIDADE NATURAL",
        "MEDIÇÃO.1",
        "DENSIDADE IN SITU",
        "MEDIÇÃO.2",
        "LL-LP",
        "MEDIÇÃO.3",
        "ANÁLISE GRANULOMÉTRICA COM PENEIRAMENTO E SEDIMENTAÇÃO",
        "MEDIÇÃO.4",
        "COMPACTAÇÃO E CBR (5 PONTOS) ENERGIA PROCTOR NORMAL",
        "MEDIÇÃO.5",
        "COMPACTAÇÃO E CBR (5 PONTOS) ENERGIA PROCTOR INTERMEDIÁRIO",
        "MEDIÇÃO.6",
        "MCT-PASTILHAS",
        "MEDIÇÃO.7",
        "MÓDULO DE RESILIÊNCIA",
        "MEDIÇÃO.8",
    ],
    "POÇO": ["KM", "POÇO DE INSPEÇÃO", "MEDIÇÃO", "NORTE", "LESTE"],
    "MISTA": ["KM", "SONDAGEM MISTA", "MEDIÇÃO", "NORTE", "LESTE"],
}

# --- Dados do Filtro (Exemplo) ---
measurement_year = 2025
measurement_month = 1  # Junho
print(f"Filtrando pela medição: Mês {measurement_month:02d}/{measurement_year}")

# --- Caminho para o Ficheiro ---
# --- Caminho para o Ficheiro ---
# --- Caminho para o Ficheiro ---
excel_path = r"D:\AT-Victor\Arquivos\Novo-Teste\MSV-163MS-007-009-FAD-EXE-ID-Z9-001-R07_Evidencias\Sondagens\Medição_Sondagens.xlsx"

# --- Início da Lógica ---
try:
    excel_file = pd.ExcelFile(excel_path)
    for sheet_name in excel_file.sheet_names:
        if sheet_name.upper() in [s.upper() for s in SHEET_COLUMN_MAP.keys()]:

            # --- Leitura e Construção do Cabeçalho ---
            header_df = pd.read_excel(
                excel_file, sheet_name=sheet_name, header=None, nrows=7
            )
            row6, row7 = header_df.iloc[5], header_df.iloc[6]

            final_headers_raw = [
                h7 if pd.notna(h7) else h6 for h6, h7 in zip(row6, row7)
            ]
            counts = {}
            final_headers = []
            for item in final_headers_raw:
                item_str = (
                    str(item) if pd.notna(item) else f"Unnamed_{len(final_headers)}"
                )
                counts[item_str] = counts.get(item_str, 0) + 1
                if counts[item_str] > 1:
                    final_headers.append(f"{item_str}_{counts[item_str]}")
                else:
                    final_headers.append(item_str)

            data_df = pd.read_excel(
                excel_file, sheet_name=sheet_name, header=None, skiprows=7
            )
            data_df.columns = final_headers[: len(data_df.columns)]

            # --- Limpeza de Dados (Linha TOTAL) ---
            total_row_indices = np.where(
                data_df.apply(
                    lambda row: row.astype(str)
                    .str.contains("TOTAL", case=False, na=False)
                    .any(),
                    axis=1,
                )
            )
            if total_row_indices[0].size > 0:
                data_df = data_df.iloc[: total_row_indices[0][0]]

            # --- LÓGICA DE FILTRAGEM CORRIGIDA E ROBUSTA ---

            if sheet_name.upper() == "TRADO":
                print(
                    f"\n>>> Aplicando FILTRAGEM HORIZONTAL para a aba: '{sheet_name}'"
                )

                # Cria uma máscara para saber que linhas manter no final
                final_mask = pd.Series(False, index=data_df.index)

                for col_name in data_df.columns:
                    if str(col_name).upper().startswith("MEDIÇÃO"):
                        col_index = data_df.columns.get_loc(col_name)
                        if col_index > 0:
                            medicao_col_as_datetime = pd.to_datetime(
                                data_df[col_name], errors="coerce"
                            )

                            # Máscara para as datas que correspondem ao filtro
                            date_match_mask = (
                                medicao_col_as_datetime.dt.year == measurement_year
                            ) & (medicao_col_as_datetime.dt.month == measurement_month)

                            # Atualiza a máscara final: uma linha é mantida se QUALQUER medição for válida
                            final_mask = final_mask | date_match_mask.fillna(False)

                            # Anula o valor da coluna anterior onde a data NÃO corresponde
                            data_df.iloc[~date_match_mask, col_index - 1] = np.nan

                # Filtra o DataFrame para manter apenas as linhas com pelo menos uma medição válida
                data_df = data_df[final_mask]

            else:  # Para "Percussão", "Poço", "Mista"
                print(f"\n>>> Aplicando FILTRAGEM VERTICAL para a aba: '{sheet_name}'")
                if "MEDIÇÃO" in data_df.columns:
                    medicao_col_as_datetime = pd.to_datetime(
                        data_df["MEDIÇÃO"], errors="coerce"
                    )
                    mask = (medicao_col_as_datetime.dt.year == measurement_year) & (
                        medicao_col_as_datetime.dt.month == measurement_month
                    )
                    data_df = data_df[mask]
                else:
                    print(f"  > Aviso: Coluna 'MEDIÇÃO' não encontrada.")

            if data_df.empty:
                print(
                    f"  > Nenhum dado encontrado para a medição na aba '{sheet_name}'."
                )
                continue

            # --- Limpeza e Seleção Final ---

            # Pega a lista de todas as colunas que queremos mostrar para esta aba
            cols_to_show = SHEET_COLUMN_MAP[sheet_name.upper()]

            # Filtra a lista para incluir apenas as colunas que realmente existem no DataFrame processado
            final_cols = [col for col in cols_to_show if col in data_df.columns]

            # Seleciona apenas essas colunas
            final_df = data_df[final_cols]

            final_df.dropna(how="all").to_excel("output_path.xlsx", index=False)

            print(f"\n  --- Conteúdo da Aba: '{sheet_name}' (Processado) ---")
            # dropna(how='all') remove apenas as linhas que ficaram completamente vazias
            print(final_df.dropna(how="all"))
            print("-" * 50)

except FileNotFoundError:
    print(f"ERRO: O ficheiro não foi encontrado em '{excel_path}'")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
    import traceback

    traceback.print_exc()
