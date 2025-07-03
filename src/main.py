# src/main.py

import sys
import os
from ui.app import App

# Adiciona o diretório 'src' ao path para permitir importações relativas
# Isso é útil se você executar o script de outro diretório
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


if __name__ == "__main__":
    app = App()
    app.mainloop()


# TODO Tabelas como subtitulo e.g 1.2.2.1
# TODO seleção de projet e.g Projeto de Geometria
# TODO Texto padrão para campos padrão
# TODO colocar listagem de arquivo emitidos no projeto
