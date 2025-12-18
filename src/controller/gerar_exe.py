import os, subprocess
from src.controller.cmds_terminal import CmdsTerminal

# Classe responsável por gerar scripts de .py para .exe
class GerarExe:
    def __init__(self, script_path: str, output_name: str, one_file: bool=None, windowed: bool=None, icon_path: str=None):
        
        self.script_path = script_path
        self.output_name = output_name
        self.icon_path = icon_path
        self.one_file = one_file
        self.windowed = windowed

        self.cmd = CmdsTerminal()
        self.validate()

    # Método responsável por validar o caminho do script
    def validate(self):
        if not os.path.exists(self.script_path):
            raise FileNotFoundError(f'[ERRO] O script não foi encontrado: {self.script_path}')
        
        if self.icon_path and not os.path.exists(self.icon_path):
            raise FileNotFoundError(f'[ERRO] O icone não foi encotrado: {self.icon_path}')

    # Método responsável por compilar o script py
    def compile_exe(self):
        command = ['pyinstaller', self.script_path, "--name", self.output_name]

        if self.one_file:
            command.append('--onefile')

        if self.windowed:
            command.append('--noconsole')

        if self.icon_path:
            command.extend(['--icon', self.icon_path])

        self.cmd.msg('\nGERANDO EXECUTÁVEL...', 4)

        result = subprocess.run(command, 
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, 
                                text=True)

        if result.returncode != 0:
            raise RuntimeError('[ERRO] Falha em gerar o executável.')
        
        self.cmd.msg('\nEXECUTÁVEL GERADO SUCESSO!', 5)