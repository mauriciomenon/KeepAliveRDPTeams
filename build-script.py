# pylint: disable=E0602,E0102,E1101
# build.py
import os
import sys
import subprocess
import shutil
from pathlib import Path


def build_exe(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"Erro: Arquivo {file_path} não encontrado!")
            sys.exit(1)

        print(f"Gerando executável para: {file_path}")

        # Instalar PyInstaller se necessário
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

        # Comando para gerar o executável
        cmd = [
            "pyinstaller",
            "--noconfirm",
            "--onefile",
            "--windowed",
            "--name=KeepAliveManager",
            "--clean",
            file_path,
        ]

        # Executar o comando
        subprocess.check_call(cmd)
        print("Executável criado com sucesso na pasta 'dist'!")

    except subprocess.CalledProcessError as e:
        print(f"Erro ao criar executável: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python build.py caminho/do/arquivo.py")
        sys.exit(1)

    build_exe(sys.argv[1])
