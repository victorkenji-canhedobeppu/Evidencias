# src/core/doc_creator_win32.py
# Versão alternativa usando win32com para controlar o MS Word diretamente.

import win32com.client as win32
from tkinter import filedialog, messagebox
import os


class DocxCreator:
    def __init__(self, template_path: str):
        self.template_path = os.path.abspath(
            template_path
        )  # win32com precisa do caminho absoluto

    def _set_content_control_text(self, doc, tag: str, value: str):
        """Encontra um Content Control pela sua Tag ou Título e define o texto."""
        try:
            # Tenta encontrar por Tag primeiro
            for cc in doc.ContentControls:
                if cc.Title.strip():
                    titulo = cc.Title.strip()
                    if titulo == tag:
                        cc.Range.Text = str(value)
                        print("str:", str(value))
                        return
        except Exception:
            print(
                f"Aviso: Não foi possível encontrar o Content Control com a tag/título '{tag}'."
            )

    def _ensure_heading_numbering(self, word_app, doc):
        """Liga 'Título 1' e 'Título 2' a um esquema 1-1.1-1.1.1."""
        wdOutlineNumberGallery = 3  # <-- 3, não 2!
        gallery = word_app.ListGalleries(wdOutlineNumberGallery)

        # Pega o 1.º modelo da galeria “1 Título 1, 1.1 Título 2…”
        list_template = gallery.ListTemplates(1)

        # Usa nomes em português *ou* em inglês, conforme a instalação
        for style_name, level in (
            ("Título 1", 1),
            ("Heading 1", 1),
            ("Título 2", 2),
            ("Heading 2", 2),
        ):
            try:
                style = doc.Styles(style_name)
                # argumentos **NOMEADOS** são essenciais no pywin32
                style.LinkToListTemplate(
                    ListTemplate=list_template, ListLevelNumber=level
                )
            except Exception:
                # ignora se o estilo não existe nesse idioma
                pass

    def _append_content_on_new_page(self, word_app, intro_text, title2_text):

        try:

            # Move o cursor para o final do documento
            selection = word_app.Selection
            selection.EndKey(Unit=6)  # wdStory

            # Adiciona uma quebra de página
            selection.InsertBreak(Type=7)  # wdPageBreak

            # Adiciona o Título 1 e o seu conteúdo
            if intro_text:
                selection.Style = "Título 1"
                selection.InsertAfter(
                    "INTRODUÇÃO"
                )  # Insere o texto sem apagar a numeração
                selection.Collapse(Direction=0)  # Move o cursor para o final do texto
                selection.TypeParagraph()  # Cria um novo parágrafo
                selection.Style = "Normal"
                selection.TypeText(Text=intro_text)
                selection.TypeParagraph()

            # Adiciona o Título 2 (AVANÇO FÍSICO) com numeração automática
            if title2_text:
                selection.TypeParagraph()
                selection.Style = "Título 1"
                selection.InsertAfter("AVANÇO FÍSICO DO PROJETO")
                selection.ParagraphFormat.PageBreakBefore = False
                selection.Collapse(Direction=0)
                selection.TypeParagraph()
                selection.Style = "Normal"
                selection.TypeText(Text=title2_text)
                selection.TypeParagraph()

        except Exception as e:
            print(e)

    def generate_document(self, main_data: dict, table_data: list):
        word_app = None
        try:
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False  # Não mostra o Word a abrir

            doc = word_app.Documents.Open(self.template_path)

            all_data_to_fill = main_data.copy()

            intro_text = all_data_to_fill.pop("INTRO_TEXT", "")
            title2_text = all_data_to_fill.pop("TITLE2_TEXT", "")

            # As chaves aqui (ex: "Revisão0") devem corresponder exatamente
            # aos Títulos dos Content Controls no seu template Word.
            all_data_to_fill[f"Revisão 0"] = table_data[0].get("revisao", "")
            all_data_to_fill[f"Data Revisão 0"] = table_data[0].get("data", "")
            all_data_to_fill[f"Descrição 0"] = table_data[0].get("descricao", "")

            # 3. Preenche todos os Content Controls num único loop
            for title, value in all_data_to_fill.items():
                if title not in ["INTRO_TEXT", "TITLE2_TEXT"]:
                    self._set_content_control_text(doc, title, value)

            self._ensure_heading_numbering(word_app, doc)

            self._append_content_on_new_page(word_app, intro_text, title2_text)

            save_path = filedialog.asksaveasfilename(
                title="Salvar Relatório Gerado",
                defaultextension=".docx",
                filetypes=[("Documento Word", "*.docx")],
            )
            if not save_path:
                doc.Close(SaveChanges=0)  # Fecha sem salvar
                return

            doc.SaveAs(os.path.abspath(save_path))
            doc.Close()
            messagebox.showinfo(
                "Sucesso", f"Documento salvo com sucesso em:\n{save_path}"
            )
            return save_path

        except FileNotFoundError:
            messagebox.showerror(
                "Erro", f"Template não encontrado em:\n{self.template_path}"
            )
            return None
        except Exception as e:
            messagebox.showerror(
                "Erro Win32",
                f"Ocorreu um erro ao interagir com o MS Word:\n{e}\n\nVerifique se o Word está instalado corretamente.",
            )
            return None
        finally:
            if word_app:
                word_app.Quit()
