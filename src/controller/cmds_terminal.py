import os
from time import sleep

class CmdsTerminal:

    # Método responsável por imprimir mensagens personalizadas
    def msg(self, text: str, delay: int=2):
        if not isinstance(text, str):
            raise TypeError('[ERRO] Aconteceu um erro no momento de passar o parâmetro de texto.')
        
        if not isinstance(delay, int):
            raise TypeError('[ERRO] Aconteceu um erro no momento de passar o parâmetro de delay.')
        
        os.system('cls')
        print(f'  {text}')
        sleep(delay)
        os.system('cls')