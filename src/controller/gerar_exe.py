import sys, os, subprocess, shutil
from src.controller.cmds_terminal import CmdsTerminal

# Classe responsável por gerar exe dos ciclos do compilador g-code
class GerarExe:
    def __init__(self, script_path: str, output_path: str, output_name: str, one_file=True, windowed=True, icon_path='assets/favicon.ico'):

        self.script_path = script_path
        self.output_path = output_path
        self.output_name = output_name
        self.icon_path = icon_path
        self.one_file = one_file
        self.windowed = windowed

        self.cmd = CmdsTerminal()
        self.validate()

    # Método responsável por validadr os campos de parâmetros da classe
    def validate(self):
        if not os.path.exists(self.script_path):
            raise FileNotFoundError(f'[ERRO] O script não foi encontrado: {self.script_path}')

        if self.icon_path and not os.path.exists(self.icon_path):
            raise FileNotFoundError(f'[ERRO] O icone não foi encontrado: {self.icon_path}')

    # Método responsável por remover arquivos e pastas que são geradas com .exe
    def clean_files(self):
        # Remove o arquivo .spec
        spec_file = f'{self.output_name}.spec'
        if os.path.exists(spec_file):
            os.remove(spec_file)

        # Remove a pasta build
        if os.path.exists('build'):
            shutil.rmtree('build')

    # Método responsável executar os comandos pyinstaller para gerar o .exe
    def compile_exe(self):
        command = [sys.executable, "-m", "PyInstaller", self.script_path, "--name", self.output_name]

        if self.one_file:
            command.append('--onefile')

        if self.windowed:
            command.append('--noconsole')

        if self.icon_path:
            command.extend(['--icon', self.icon_path])

        self.cmd.msg('\nGERANDO EXECUTÁVEL...', 4)

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        if result.returncode != 0:
            print(result.stderr)
            raise RuntimeError('[ERRO] Falha em gerar o executável.')

        self.cmd.msg('\nEXECUTÁVEL GERADO COM SUCESSO!', 5)