# src/core/doc_appender.py
# VERSÃO FINAL E DEFINITIVA: Usa win32com para programaticamente criar e aplicar
# o esquema de numeração hierárquico, garantindo que funcione em qualquer documento.

from math import ceil
import os
import fitz
import tempfile
import docx
import pandas as pd
import numpy as np
from docx.shared import Pt
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH
import win32com.client as win32
from win32com.client import constants as c
from tkinter import messagebox

from config.settings import COLUMN_RENAME_MAP, DISCIPLINE_SHEET_TYPES, SHEET_COLUMN_MAP


class DocxAppender:
    def __init__(
        self,
        docx_path: str,
        measurement_month: int,
        measurement_year: int,
    ):
        self.docx_path = os.path.abspath(docx_path)
        self.measurement_month = measurement_month
        self.measurement_year = measurement_year
        self.all_excel_data = pd.DataFrame()

    def _find_files(self, folder_path: str):
        if not os.path.isdir(folder_path):
            return []
        supported_files = []
        for f in os.listdir(folder_path):
            if f.lower().endswith((".jpeg", ".jpg", ".png", ".pdf", ".xlsx")):
                supported_files.append(os.path.join(folder_path, f))
        return sorted(supported_files)

    def _apply_heading_numbering(self, word_app, doc):
        """
        Constrói e aplica um esquema de numeração hierárquico aos estilos de Título.
        Esta função é a chave para garantir que a numeração funcione sempre.
        """
        try:
            # Pega na galeria de listas de múltiplos níveis
            gallery = word_app.ListGalleries(c.wdOutlineNumberGallery)

            # Adiciona um novo template de lista ao documento. Isto "reseta" a formatação.
            list_template = doc.ListTemplates.Add(True)

            text_position_pt = 1.5 * 28.3464567

            lvl1 = list_template.ListLevels(1)
            lvl1.NumberFormat = "%1"  # Formato "NúmeroDoNível1.NúmeroDoNível2"
            lvl1.TrailingCharacter = c.wdTrailingTab
            lvl1.NumberStyle = c.wdListNumberStyleArabic
            lvl1.LinkedStyle = "Título 1"
            lvl1.NumberPosition = 0
            lvl1.TextPosition = text_position_pt
            lvl1.TabPosition = text_position_pt

            # Define o Nível 2 para estar ligado ao estilo "Título 2"
            lvl2 = list_template.ListLevels(2)
            lvl2.NumberFormat = "%1.%2"  # Formato "NúmeroDoNível1.NúmeroDoNível2"
            lvl2.TrailingCharacter = c.wdTrailingTab
            lvl2.NumberStyle = c.wdListNumberStyleArabic
            lvl2.LinkedStyle = "Título 2"
            lvl2.NumberPosition = 0
            lvl2.TextPosition = text_position_pt
            lvl2.TabPosition = text_position_pt

            # Define o Nível 3 para estar ligado ao estilo "Título 3"
            lvl3 = list_template.ListLevels(3)
            lvl3.NumberFormat = "%1.%2.%3"  # Formato "N1.N2.N3"
            lvl3.TrailingCharacter = c.wdTrailingTab
            lvl3.NumberStyle = c.wdListNumberStyleArabic
            lvl3.LinkedStyle = "Título 3"
            lvl3.NumberPosition = 0
            lvl3.TextPosition = text_position_pt
            lvl3.TabPosition = text_position_pt

            print("Numeração dos estilos de título foi configurada com sucesso.")
            return True
        except Exception as e:
            print(f"Não foi possível configurar a numeração dos títulos. Erro: {e}")
            messagebox.showerror(
                "Erro de Configuração",
                f"Não foi possível configurar a numeração automática no Word.\n\nDetalhe: {e}",
            )
            return False

    def _convert_pdf_to_png(self, pdf_path):
        """
        Converte a primeira página de um PDF para um ficheiro PNG,
        salvando-o na mesma pasta com o mesmo nome.
        """
        try:
            # Constrói o caminho para o ficheiro de saída
            output_png_path = os.path.splitext(pdf_path)[0] + ".png"

            # Se a imagem já existir, não a converte novamente
            if os.path.exists(output_png_path):
                print(
                    f"Imagem PNG já existe para {os.path.basename(pdf_path)}. A usar a existente."
                )
                return output_png_path

            doc = fitz.open(pdf_path)
            page = doc.load_page(0)

            # Aumenta o zoom para melhor qualidade da imagem (dpi)
            zoom = 2
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix)
            doc.close()

            # Salva a imagem no caminho de saída
            pix.save(output_png_path)
            print(f"PDF convertido para: {output_png_path}")
            return output_png_path

        except Exception as e:
            print(f"Erro ao converter PDF {os.path.basename(pdf_path)}: {e}")
            return None

    def _define_magin_title(self, path, margin_left, margin_right):

        try:
            doc = docx.Document(path)
            styles = doc.styles

            # Nomes dos style de título que serão modificados
            styles_name_title = ["Heading 1", "Heading 2", "Heading 3"]

            for style_name in styles_name_title:
                if style_name in styles:
                    estilo = styles[style_name]
                    para_format = estilo.paragraph_format
                    para_format.left_indent = Pt(margin_left)
                    para_format.right_indent = Pt(margin_right)
                    print(
                        f"Margens do estilo '{style_name}' atualizadas para {margin_left}pt (esquerda) e {margin_right}pt (direita)."
                    )
                else:
                    print(
                        f"Aviso: O estilo '{style_name}' não foi encontrado no documento."
                    )

            # Salva o documento com as modificações
            doc.save(path)
            print("\nDocumento salvo com sucesso!")

        except FileNotFoundError:
            print(f"Erro: O arquivo '{path}' não foi encontrado.")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

    def _centralize_image(self, caminho_arquivo):
        try:
            doc = docx.Document(caminho_arquivo)
            centralized_paragraphs = 0

            for paragraph in doc.paragraphs:
                # A forma mais robusta de detectar uma imagem é verificar a tag '<w:drawing>'
                # no XML do parágrafo.
                if "<w:drawing>" in paragraph._p.xml:
                    # 1. Zera os recuos esquerdo e direito do parágrafo
                    # Isso faz o parágrafo se expandir por toda a largura entre as margens da página
                    paragraph_format = paragraph.paragraph_format
                    paragraph_format.left_indent = Pt(0)
                    paragraph_format.right_indent = Pt(0)

                    # 2. Define o alinhamento do parágrafo como centralizado
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    centralized_paragraphs += 1

            if centralized_paragraphs > 0:
                doc.save(caminho_arquivo)
                print(
                    f"\nOperação concluída. {centralized_paragraphs} parágrafo(s) com imagem foram centralizados."
                )
            else:
                print(
                    "\nOperação concluída. Nenhuma imagem 'Em Linha com o Texto' foi encontrada."
                )

        except FileNotFoundError:
            print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        except Exception as e:
            print(f"Ocorreu um erro inesperado: {e}")

    def _process_excel_file(
        self,
        excel_path: str,
        measurement_month: int,
        measurement_year: int,
        discipline: str,
    ) -> dict:

        processed_data = {}
        valid_discipline = DISCIPLINE_SHEET_TYPES.get(discipline, [])
        print(f"Processando abas do tipo: '{valid_discipline}'")

        try:
            excel_file = pd.ExcelFile(excel_path)
            for sheet_name in excel_file.sheet_names:
                if sheet_name in valid_discipline:
                    if sheet_name.upper() in SHEET_COLUMN_MAP:
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
                                str(item)
                                if pd.notna(item)
                                else f"Unnamed_{len(final_headers)}"
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

                        if sheet_name.upper() in (
                            "TRADO",
                            "SHELBY",
                            "DENISON",
                            "BLOCO",
                            "POÇO",
                        ):
                            # 1. Filtragem Horizontal (anula os dados fora do período)
                            for col_name in data_df.columns:
                                if str(col_name).upper().startswith("MEDIÇÃO"):
                                    col_idx = data_df.columns.get_loc(col_name)
                                    if col_idx > 0:
                                        dt_col = pd.to_datetime(
                                            data_df[col_name], errors="coerce"
                                        )
                                        null_rows = (
                                            dt_col.dt.year != measurement_year
                                        ) | (dt_col.dt.month != measurement_month)
                                        data_df.iloc[null_rows, col_idx - 1] = np.nan

                            # 2. Filtragem Vertical (remove linhas sem NENHUMA medição válida)
                            any_medicao_valid = pd.Series(False, index=data_df.index)
                            for col_name in data_df.columns:
                                if str(col_name).upper().startswith("MEDIÇÃO"):
                                    dt_col = pd.to_datetime(
                                        data_df[col_name], errors="coerce"
                                    )
                                    valid_rows_mask = (
                                        dt_col.dt.year == measurement_year
                                    ) & (dt_col.dt.month == measurement_month)
                                    any_medicao_valid = (
                                        any_medicao_valid
                                        | valid_rows_mask.fillna(False)
                                    )
                            data_df = data_df[any_medicao_valid]
                        else:
                            if "MEDIÇÃO" in data_df.columns:
                                medicao_col_as_datetime = pd.to_datetime(
                                    data_df["MEDIÇÃO"], errors="coerce"
                                )
                                mask = (
                                    medicao_col_as_datetime.dt.year == measurement_year
                                ) & (
                                    medicao_col_as_datetime.dt.month
                                    == measurement_month
                                )
                                data_df = data_df[mask]

                        if not data_df.empty:
                            # --- LÓGICA DE RENOMEAÇÃO APLICADA AQUI ---
                            final_cols = [
                                col
                                for col in SHEET_COLUMN_MAP[sheet_name.upper()]
                                if col in data_df.columns
                            ]
                            df_to_add = data_df[final_cols].dropna(how="all")

                            # Renomeia as colunas usando o novo mapa
                            df_to_add.rename(columns=COLUMN_RENAME_MAP, inplace=True)

                            processed_data[sheet_name] = df_to_add
        except Exception as e:
            print(f"Erro ao processar o arquivo Excel '{excel_path}': {e}")

        return processed_data

    # --- NOVA FUNÇÃO DE APOIO PARA CRIAR TABELAS ---
    def _add_dataframe_as_table(self, doc, df: pd.DataFrame, word_app):
        if df.empty:
            return

        p = doc.Paragraphs.Add()
        table = doc.Tables.Add(p.Range, NumRows=df.shape[0] + 1, NumColumns=df.shape[1])
        table.Style = "Tabela com grade"
        table.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter
        table.Borders.Enable = True

        # Formatação do Cabeçalho
        header_row = table.Rows(1)
        header_row.HeadingFormat = True
        header_row.Shading.BackgroundPatternColor = c.wdColorDarkBlue
        header_font = header_row.Range.Font
        header_font.Size = 8
        header_font.Name = "Calibri"
        header_font.ColorIndex = c.wdWhite
        header_font.Bold = True
        header_row.Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

        for j, col_name in enumerate(df.columns):
            cell = table.Cell(Row=1, Column=j + 1)
            cell.Range.Text = str(col_name)
            # --- MELHORIA: Desativa a quebra de linha automática para a célula do cabeçalho ---
            cell.WordWrap = False
            cell.VerticalAlignment = c.wdCellAlignVerticalCenter
        # Preenche as linhas de dados com conversão segura
        for i, row_data in enumerate(df.itertuples(index=False)):
            for j, val in enumerate(row_data):
                cell = table.Cell(i + 2, j + 1)

                # Converte o valor para string de forma segura ANTES de o passar ao Word
                if pd.isna(val):
                    cell_text = ""
                elif isinstance(val, (datetime, pd.Timestamp)):
                    cell_text = val.strftime("%d/%m/%Y")
                else:
                    cell_text = str(val)

                cell.Range.Text = cell_text
                cell.Range.Font.Size = 8
                cell.Range.Font.Name = "Calibri"
                cell.WordWrap = False
                cell.VerticalAlignment = c.wdCellAlignVerticalCenter

        table.AutoFitBehavior(c.wdAutoFitContent)
        print("Tabela inserida no documento com sucesso.")

    def append_measurement(
        self,
        disciplines: list,
        user_texts: dict,
        base_evidence_path: str,
        measurement_month: int,
        measurement_year: int,
    ):
        word_app = None
        doc = None
        try:
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False
            doc = word_app.Documents.Open(self.docx_path)

            # --- 1. APLICA A CONFIGURAÇÃO DE NUMERAÇÃO ---
            if not self._apply_heading_numbering(word_app, doc):
                return  # Interrompe se a configuração da numeração falhar

            # --- 2. ADICIONA O CONTEÚDO ---
            selection = word_app.Selection
            selection.EndKey(Unit=6)  # Move para o final

            today_date = datetime.now().strftime("%m/%Y")

            # Adiciona o Título 2
            selection.Style = "Título 2"
            selection.TypeText(Text=f"MEDIÇÃO ({today_date})")
            selection.TypeParagraph()

            for discipline in disciplines:
                # Adiciona o Título 3
                selection.Style = "Título 3"
                selection.TypeText(Text=discipline)
                selection.TypeParagraph()

                # Adiciona o texto do utilizador
                user_text = user_texts.get(discipline, "")
                if user_text:
                    selection.Style = "Normal"
                    selection.TypeText(Text=user_text)
                    selection.TypeParagraph()

                # Adiciona ficheiros
                discipline_folder_path = os.path.join(base_evidence_path, discipline)
                files_to_add = self._find_files(discipline_folder_path)

                excel_files = [
                    f for f in files_to_add if f.lower().endswith((".xlsx", ".xls"))
                ]

                if not files_to_add:
                    selection.Style = "Normal"
                    selection.TypeText(
                        Text="Nenhuma evidência encontrada para esta disciplina."
                    )
                    selection.TypeParagraph()
                else:
                    for file_path in files_to_add:
                        if file_path.lower().endswith((".jpeg", ".jpg", ".png")):
                            selection.InlineShapes.AddPicture(
                                FileName=os.path.abspath(file_path),
                                LinkToFile=False,
                                SaveWithDocument=True,
                            )
                            selection.TypeParagraph()
                        elif file_path.lower().endswith(".pdf"):
                            png_path = self._convert_pdf_to_png(file_path)
                            selection.InlineShapes.AddPicture(
                                FileName=os.path.abspath(png_path),
                                LinkToFile=False,
                                SaveWithDocument=True,
                            )
                            selection.TypeParagraph()
                if excel_files:
                    p_excel_header = doc.Paragraphs.Add()
                    p_excel_header.Range.Text = "\nDados das Planilhas:"
                    p_excel_header.Range.Bold = True
                    for excel_path in excel_files:
                        try:
                            self.all_excel_data = self._process_excel_file(
                                excel_path,
                                measurement_month,
                                measurement_year,
                                discipline,
                            )
                            for sheet_name, final_df in self.all_excel_data.items():
                                p_sheet_name = doc.Paragraphs.Add()
                                p_sheet_name.Range.Text = f"Tabela da Aba: {sheet_name}"
                                p_sheet_name.Range.Bold = True
                                p_sheet_name.Range.Font.Name = "Calibri"
                                p_sheet_name.Range.InsertParagraphAfter()
                                if sheet_name.upper() == "TRADO":
                                    key_cols = ["Sondagem"]
                                    actual_keys = [
                                        c for c in key_cols if c in final_df.columns
                                    ]
                                    data_cols = [
                                        c
                                        for c in final_df.columns
                                        if c not in actual_keys
                                    ]
                                    MAX_DATA_COLS = 9
                                    num_chunks = ceil(len(data_cols) / MAX_DATA_COLS)
                                    for i in range(num_chunks):
                                        chunk_cols = data_cols[
                                            i * MAX_DATA_COLS : (i + 1) * MAX_DATA_COLS
                                        ]
                                        table_df = final_df[actual_keys + chunk_cols]
                                        self._add_dataframe_as_table(
                                            doc, table_df, word_app
                                        )
                                        doc.Paragraphs.Add()  # Espaçamento entre tabelas divididas
                                else:
                                    self._add_dataframe_as_table(
                                        doc, final_df, word_app
                                    )
                        # Por enquanto, apenas printamos. No futuro, podemos adicionar a tabela ao Word.
                        except Exception as e:
                            print(f"Erro ao ler o arquivo Excel '{excel_path}': {e}")

                selection.TypeParagraph()

            # --- 3. SALVA E FECHA ---
            doc.Fields.Update()  # Força a atualização de todos os campos
            doc.Save()
            messagebox.showinfo(
                "Sucesso",
                f"Conteúdo adicionado e numerado com sucesso em:\n{self.docx_path}",
            )

        except Exception as e:
            messagebox.showerror(
                "Erro ao Adicionar ao Documento",
                f"Ocorreu um erro fatal ao interagir com o Word:\n{e}",
            )
        finally:
            if doc:
                doc.Close(SaveChanges=0)
            if word_app:
                word_app.Quit()

        self._define_magin_title(self.docx_path, 36, 36)
        self._centralize_image(self.docx_path)
