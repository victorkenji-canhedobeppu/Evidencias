# src/core/folder_creator.py
# VERSÃO FINAL E DEFINITIVA: A lógica foi corrigida para garantir que as
# pastas de disciplina sejam sempre criadas, independentemente dos checkboxes.

import os
import pandas as pd
import re

from config.settings import (
    ANTT_DISCIPLINES_TYPES,
    ARTESP_FILE_GROUPED_TYPES,
)


class FolderCreator:
    def __init__(self, project_type: str, df: pd.DataFrame, file_path: str):
        self.project_type = project_type
        self.df = df
        self.source_file_path = file_path
        self.agency_code_column = "CÓDIGO AGENCIA"
        self.disciplines_map = self._get_disciplines_map()

    def _get_disciplines_map(self) -> dict:
        if self.project_type == "ARTESP":
            return ARTESP_FILE_GROUPED_TYPES
        elif self.project_type == "ANTT":
            return ANTT_DISCIPLINES_TYPES
        return {}

    def _sanitize_filename(self, filename: str) -> str:
        """Remove caracteres inválidos para nomes de diretório."""
        return re.sub(r'[\\/*?:"<>|]', "", filename)

    def _find_discipline(self, agency_code: str) -> str | None:
        if not agency_code or not isinstance(agency_code, str):
            return None
        agency_code_upper = agency_code.upper()
        for discipline, codes in self.disciplines_map.items():
            if any(code.upper() in agency_code_upper for code in codes):
                return discipline
        return None

    # --- FUNÇÃO ATUALIZADA COM A ESTRUTURA DE PASTAS CORRETA ---
    def create_folders_for_active_data(
        self, create_sondagens: bool, create_ensaios: bool
    ) -> tuple[bool, str, str | None]:

        try:
            # 1. Cria a pasta raiz (como antes)
            base_directory = os.path.dirname(self.source_file_path)
            excel_filename_raw = os.path.splitext(
                os.path.basename(self.source_file_path)
            )[0]

            sanitized_filename = self._sanitize_filename(excel_filename_raw)
            base_root_folder_name = f"{sanitized_filename}_Evidencias"

            root_folder_path = os.path.join(base_directory, base_root_folder_name)
            counter = 1
            while os.path.exists(root_folder_path):
                root_folder_path = os.path.join(
                    base_directory, f"{base_root_folder_name} ({counter})"
                )
                counter += 1

            final_root_folder_name = os.path.basename(root_folder_path)
            os.makedirs(root_folder_path, exist_ok=True)

            # 2. Cria as pastas fixas, se selecionadas
            if create_sondagens:
                os.makedirs(os.path.join(root_folder_path, "Sondagens"), exist_ok=True)

            if create_ensaios:
                os.makedirs(
                    os.path.join(root_folder_path, "Ensaios Especiais"), exist_ok=True
                )

        except Exception as e:
            return (
                False,
                f"Não foi possível criar a estrutura de pastas raiz.\nErro: {e}",
                None,
            )

        # --- LÓGICA DE CRIAÇÃO DAS PASTAS DE DISCIPLINA (SEMPRE EXECUTA) ---
        folders_created = set()
        if self.agency_code_column not in self.df.columns:
            # Se não houver a coluna, avisa, mas considera a operação um sucesso parcial
            return (
                True,
                f"Estrutura principal criada em '{final_root_folder_name}'. A coluna de disciplinas não foi encontrada.",
                root_folder_path,
            )

        if self.df.empty:
            return (
                True,
                f"Estrutura principal criada em '{final_root_folder_name}'. Não há dados para criar pastas de disciplina.",
                root_folder_path,
            )

        for index, row in self.df.iterrows():
            agency_code = row[self.agency_code_column]
            discipline = self._find_discipline(agency_code)
            if discipline:
                # Cria a pasta da disciplina diretamente dentro da pasta raiz
                folder_path = os.path.join(root_folder_path, discipline)
                try:
                    os.makedirs(folder_path, exist_ok=True)
                    folders_created.add(discipline)
                except OSError as e:
                    return (
                        False,
                        f"Não foi possível criar a pasta '{discipline}'.\nErro: {e}",
                        None,
                    )

        return (
            True,
            f"Estrutura de pastas criada com sucesso em:\n'{final_root_folder_name}'",
            root_folder_path,
        )
