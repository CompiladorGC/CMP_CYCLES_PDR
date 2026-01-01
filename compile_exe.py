from src.controller.gerar_exe import GerarExe

if __name__ == '__main__':
    exe = GerarExe(
        script_path='src/model/cycles/cycle_desbaste.py',
        output_path='dist/cycle_desbaste/v1.0',
        output_name='ciclo_desbaste')   
    exe.compile_exe()
    exe.clean_files()