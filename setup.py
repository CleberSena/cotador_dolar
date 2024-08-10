import sys
import os
from cx_Freeze import setup, Executable

# Definir o que deve ser incluído na pasta final
arquivos = ["teste.docx", "file.pdf", 'Grafico.png']
# Saida de arquivos
configuracao = Executable(
    script='cotacao.py',
    icon='icon.ico'
)
# Configurar o executável
setup(
    name='Gerenciador de Tarefas',
    version='1.0',
    description='Este programa verifica a cotação do dolar atual no Brasil',
    author='Cleber W. Sena',
    options={'build_exe':{ 
        'include_files' : arquivos,      
        'include_msvcr': True
    }},
    executables=[configuracao]
)