import pandas as pd

excel_path = r"D:\AT-Victor\Arquivos\Novo-Teste\MSV-163MS-007-009-FAD-EXE-ID-Z9-001-R07_Evidencias\Sondagens\Medição_Sondagens.xlsx"

excel_file = pd.ExcelFile(excel_path)
for sheet_name in excel_file.sheet_names:
    if sheet_name == "TRADO":
        # lê só o cabeçalho (linha 6 do Excel) ─ nenhuma linha de dados
        header_only = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=5, nrows=0
        )
        headers = header_only.columns.tolist()
        print(headers)  # ['Coluna1', 'Coluna2', 'Coluna3', ...]
