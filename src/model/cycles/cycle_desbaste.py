import openpyxl, os, textwrap
from src.controller.cmds_terminal import CmdsTerminal

# Classe responsável por gerar o ciclo de desbaste parametrizado
class CycleDesbaste:
    def __init__(self):
        self.file = openpyxl.load_workbook(r'C:\compilador g-code\ciclos\ciclos_padrões\ciclos_pdr.xlsx') # Caminho da pasta de trabalho excel
        self.plan_desb = self.file['CICLO DESBASTE']
        self.cmd = CmdsTerminal()
    
    # Método responsável por salvar o g-code do cycle  
    def save_code(self):
        try:
            arquivos = [self.file_main(), self.file_init(), self.file_configs(), self.file_cycle()]
            nomes = ['CMP_MAIN.MPF', 'CMP_INIT.SPF', 'CMP_CONFIGS.SPF', 'CMP_CYCLE.SPF']

            pasta = r'C:\compilador g-code\ciclos\ciclos_padrões' # Caminho da pasta responsável pelo o ciclo
            pasta_saida = os.path.join(pasta, 'arquivos g-code')
            os.makedirs(pasta_saida, exist_ok=True)

            for nome, arquivo in zip(nomes, arquivos):
                caminho = os.path.join(pasta_saida, nome)

                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(arquivo)
                    self.cmd.msg(f'\nSALVANDO ARQUIVO: {nome}', 1)

        except Exception as e:
            print(f'\n  [ERRO] ACONTECEU UM ERRO NO MOMENTO DE SALVAR OS ARQUIVOS G-CODE: {e}')
        else:
            self.cmd.msg('\nARQUIVOS SALVOS COM SUCESSO!', 2)

    # Método responsável pelo o arquivo CMP_MAIN.MPF
    def file_main(self):
        # Seção de Parâmetros Dimensionais
        diametro_inicial = self.plan_desb['I11'].value
        diametro_final = self.plan_desb['I12'].value
        espessura = self.plan_desb['I13'].value

        # Seção de Parâmetros de Corte
        passe = self.plan_desb['D16'].value

        rpm = self.plan_desb['D13'].value
        rpm_min = self.plan_desb['D14'].value
        rpm_max = self.plan_desb['D15'].value
        tipo_vc = self.plan_desb['D18'].value

        avanco_desb = self.plan_desb['D11'].value
        avanco_acab = self.plan_desb['D12'].value
        tipo_avanco = self.plan_desb['D19'].value

        # Seção de Parâmetros de Posicionamento
        x_pos = self.plan_desb['I17'].value
        z_pos = self.plan_desb['I18'].value
        aprox = self.plan_desb['I19'].value
        aproz = self.plan_desb['I20'].value

        self.main = textwrap.dedent(f'''
        DEF REAL DIAMETRO_INICIAL, DIAMETRO_FINAL, ESPESSURA, X_POS, Z_POS, APROX, APROZ;
        DEF REAL PASSE, AVANCO_ACAB, AVANCO_DESB, RPM, RPM_MIN, RPM_MAX;
        DEF STRING [30] TIPO_VC, TIPO_AVANCO , MODO_OP;
        DEF INT FERRAMENTA_DESB, FERRAMENTA_ACAB;

        GROUP_BEGIN(0,"SECTION - PARÂMETROS DIMENSIONAIS",0,0)
        DIAMETRO_INICIAL = {diametro_inicial};
        DIAMETRO_FINAL = {diametro_final};
        ESPESSURA = {espessura};
        MODO_OP = "";
        GROUP_END(0,0)

        GROUP_BEGIN(0,"SECTION - PARÂMETROS DE CORTE",0)
        FERRAMENTA_DESB = ;
        FERRAMENTA_ACAB = ;
        PASSE = {passe};

        RPM = {rpm};
        RPM_MIN = {rpm_min};
        RPM_MAX = {rpm_max};
        TIPO_VC = "{tipo_vc}";

        AVANCO_DESB = {avanco_desb};
        AVANCO_ACAB = {avanco_acab};
        TIPO_AVANCO = "{tipo_avanco}";
        GROUP_END(0,0)

        GROUP_BEGIN(0,"SECTION - PARÂMETROS DE POSICIONAMENTO",0)
        APROX = {aprox};
        APROZ = {aproz};
        X_POS = {x_pos};
        Z_POS = {z_pos};
        GROUP_END(0,0)

        CALL "/_N_SPF_DIR/_N_CICLOS_PDR_DIR/_N_CICLO_DESBASTE_DIR/_N_CMP_CONFIGS_SPF";
        M30;''')

        return self.main
    
    # Método responsável pelo o arquivo CMP_INIT.SPF
    def file_init(self):
        self.init = textwrap.dedent('''
        N10 DEF STRING [30] _RESULT_OP, _RESULT_VC, _RESULT_AVANCO;

        N20 _RESULT_AVANCO = TOUPPER(TIPO_AVANCO);
        N30 _RESULT_OP = TOUPPER(MODO_OP);
        N40 _RESULT_VC = TOUPPER(TIPO_VC);

        GROUP_BEGIN(0,"SECTION - SALTO",0,0)
        IF (_RESULT_OP=="D")
            MSG(" - CARREGANDO CICLO DE DESBASTE...");
            CALL "/_N_SPF_DIR/_N_CICLOS_PDR_DIR/_N_CICLO_DESBASTE_DIR/_N_CMP_CYCLE_SPF";    
        ENDIF

        IF (_RESULT_OP=="A")
            MSG(" - CARREGANDO CICLO DE ACABAMENTO...");
            CALL "/_N_SPF_DIR/_N_CICLOS_PDR_DIR/_N_CICLO_DESBASTE_DIR/_N_CMP_CYCLE_SPF" BLOCK INICIO_ACAB TO FIM_ACAB;
        ENDIF
        GROUP_END(0,0)

        RET;''')

        return self.init
        
    # Método responsável pelo o arquivo CMP_CONFIGS.SPF
    def file_configs(self):
        self.configs = textwrap.dedent('''
        N10 DEF STRING [30] _RESULT_AVANCO, _RESULT_VC;

        N20 _RESULT_AVANCO = TOUPPER(TIPO_AVANCO);
        N30 _RESULT_VC = TOUPPER(TIPO_VC);

        GROUP_BEGIN(0,"SECTION - CONDIÇÕES")
        IF (_RESULT_VC=="CONSTANTE")
            IF (_RESULT_AVANCO=="MM/MIN")
                G[15]=7; G961
            ENDIF
            IF (_RESULT_AVANCO=="MM/ROT")
                G[15]=4; G96
            ENDIF
        ENDIF

        IF (_RESULT_VC=="FIXA")
            IF (_RESULT_AVANCO=="MM/MIN")
                G[15]=8; G971
            ENDIF
            IF (_RESULT_AVANCO=="MM/ROT")
                G[15]=5; G97
            ENDIF
        ENDIF
        GROUP_END(0,0)

        MSG(" - CARREGANDO PARÂMETROS G-CODES");
                                       
        N40 G290;
        N50 G18 G40 G71 G90;

        N60 G0 G54 X=X_POS Z=Z_POS;

        N70 T=FERRAMENTA_DESB;

        N80 G25 S=RPM_MIN;
        N90 G26 S=RPM_MAX;
        N100 S=RPM;

        N110 DIAMON;
        N120 M3;

        MSG("");

        CALL "/_N_SPF_DIR/_N_CICLOS_PDR_DIR/_N_CICLO_DESBASTE_DIR/_N_CMP_INIT_SPF";
        N130 RET;''')

        return self.configs
        
    # Método responsável pelo o arquivo CMP_CYCLE.SPF
    def file_cycle(self):
        self.cycle = textwrap.dedent('''
        N10 G0 X=APROX Z=APROZ;
        N20 G0 X=DIAMETRO_INICIAL;

        MSG(" - CICLO DE DESBASTE EM ANDAMENTO...");
        CYCLE951(DIAMETRO_INICIAL,0,DIAMETRO_FINAL,ESPESSURA,0,0,1,PASSE,0,0,11,0,0,0,1,AVANCO_DESB,0,2,1111000)

        INICIO_ACAB:

        MSG("");
        N30 G0 X=X_POS Z=Z_POS;

        IF (FERRAMENTA_DESB <> FERRAMENTA_ACAB)
            T=FERRAMENTA_ACAB;
        ENDIF

        MSG(" - CICLO DE ACABAMENTO EM ANDAMENTO...");

        N40 G0 X=APROX Z=APROZ;
        N50 G0 X=DIAMETRO_INICIAL;

        CYCLE951(DIAMETRO_INICIAL,0,DIAMETRO_FINAL,ESPESSURA,0,0,1,1,0,0,21,0,0,0,1,AVANCO_ACAB,0,2,1111000)

        MSG("");
        N60 G0 X=X_POS Z=Z_POS;

        FIM_ACAB:
        N70 RET;''')

        return self.cycle