#======================================================================================
#                                Codigo feito por Prane
#                                        2025
#                                  Aproveitem Cornos
#======================================================================================

#--------------------------------------------------------------------------------------
#                                VERIFICAÇÃO DE DEPENDÊNCIAS
#                             IMPORTANTE: NÃO REMOVA ESSA PARTE
#
#             Antes de rodar o código, ele verifica se as bibliotecas necessárias 
#                       (numpy, pandas e xlsxwriter) estão instaladas.
#            Se alguma estiver faltando, ele oferece a opção de instalar automaticamente.
#              Depois de instalar tudo pode comentar essa parte, pra economizar tempo
#                    não precisa, mas é bom deixar pra quem for usar depois
#--------------------------------------------------------------------------------------

# import sys
# import subprocess
# import importlib.util

# def verificar_dependencias():
#     # Dicionário com o nome do módulo e o nome do pacote no pip
#     pacotes_necessarios = {
#         'numpy': 'numpy',
#         'pandas': 'pandas',
#         'xlsxwriter': 'xlsxwriter'
#     }
    
#     pacotes_faltando = []
    
#     # Checa se cada pacote está instalado no ambiente
#     for modulo, pacote in pacotes_necessarios.items():
#         if importlib.util.find_spec(modulo) is None:
#             pacotes_faltando.append(pacote)
            
#     # Se faltar alguma coisa, aciona o prompt
#     if pacotes_faltando:
#         print(f"\n[AVISO] Faltam as seguintes bibliotecas para rodar o código: {', '.join(pacotes_faltando)}")
#         resposta = input("Deseja instalar agora? [Y/N]: ").strip().upper()
        
#         if resposta == 'Y':
#             print("\nInstalando dependências... Aguarde um momento.")
#             try:
#                 # Roda o comando pip install via terminal pelo próprio script
#                 subprocess.check_call([sys.executable, '-m', 'pip', 'install', *pacotes_faltando])
#                 print("\n[SUCESSO] Tudo instalado! Bora rodar a malha.\n" + "-"*40 + "\n")
#             except subprocess.CalledProcessError:
#                 print("\n[ERRO] Deu ruim na instalação automática. Tente instalar manualmente pelo terminal.")
#                 sys.exit(1) # Para a execução se a instalação falhar
#         else:
#             print("\nOperação cancelada. O código será encerrado pois precisa dessas bibliotecas para funcionar.")
#             sys.exit(0) # Encerra o código pacificamente

# # Executa a verificação antes de tentar importar o numpy e o pandas
# verificar_dependencias()

#--------------------------------------------------------------------------------------
#
#                                LEIA AQUI ANTES DE USAR
#
# Este código é projetado para gerar uma malha aerodinâmica para um caralinho aeroelastico,
# mantendo a proporção entre as cordas e permitindo refinar o bordo de ataque e fuga 
# de forma independente. Ele gera uma tabela completa com as malhas normalizadas e as
# proporções, e salva os resultados em um arquivo Excel com 6 casas decimais.
#--------------------------------------------------------------------------------------
#
#                                     COMO USAR
#
# 0. Recomendo que instale a extensão "Excel Viewer" pra ver o excel direto pelo VScode
# 1. Defina as cordas do painel frontal e traseiro (corda_frente e corda_tras).
# 2. Defina o número total de painéis que deseja para a malha (num_paineis_corda).
# 3. Execute o código para gerar a tabela completa e salvar no arquivo Excel "malhaEH.xlsx".
# 4. A planilha Excel conterá as colunas "Painel Frente", "
#   Aileron" e "Painel Inteiro" com os valores formatados para 6 casas decimais.
#--------------------------------------------------------------------------------------

import numpy as np
import pandas as pd

# Função que cria os pontos da malha mantendo a proporção entre as cordas 
# e permitindo refinar o BF

def gerar_malha(corda, refinar_bordo_ataque, refinar_bordo_fuga, num_paineis):
    x = np.linspace(0, 1, num_paineis + 1)
    if refinar_bordo_ataque:
        x = np.sin(x * np.pi / 2)
    if refinar_bordo_fuga:
        x = 1 - np.sin((1 - x) * np.pi / 2)
    return (x - np.min(x)) / (np.max(x) - np.min(x))

# Função para gerar a tabela completa com as malhas normalizadas e as proporções

def gerar_tabela_completa(corda_frente, corda_tras, num_paineis_totais):

    proporcao_frente = corda_frente / (corda_frente + corda_tras)
    num_paineis_frente = max(1, int(round(num_paineis_totais * proporcao_frente)))
    num_paineis_tras = max(1, num_paineis_totais - num_paineis_frente)

    malha_frente_norm = gerar_malha(corda_frente, False, True, num_paineis_frente)
    malha_tras_norm = gerar_malha(corda_tras, False, True, num_paineis_tras)

    malha_frente_dim = malha_frente_norm * corda_frente
    malha_tras_dim = malha_tras_norm * corda_tras + corda_frente

    max_valor_verde_claro = np.max(malha_tras_dim)
    prop_frente = malha_frente_dim / max_valor_verde_claro
    prop_tras = malha_tras_dim / max_valor_verde_claro

    prop_combinada = np.concatenate((prop_frente, prop_tras))

    max_len = max(len(malha_frente_norm), len(malha_tras_norm), len(prop_combinada))

# Função para padronizar os arrays para o mesmo comprimento, preenchendo com NaN onde necessário

    def pad_array(arr, target_len):
        return np.pad(arr, (0, max(0, target_len - len(arr))), constant_values=np.nan)

    df = pd.DataFrame({
        "Painel Frente": pad_array(malha_frente_dim / corda_frente, max_len),
        "Aileron": pad_array((malha_tras_dim - corda_frente) / corda_tras, max_len),
        "Painel Inteiro": pad_array(prop_combinada, max_len)
    })
    
    return df

# Função para salvar o DataFrame completo no Excel com 6 casas decimais

def salvar_em_excel(df, nome_arquivo):
    """
    Salva o DataFrame completo no Excel com 6 casas decimais.
    """
    with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Malha Completa", index=False, float_format="%.6f")

        workbook = writer.book
        worksheet = writer.sheets["Malha Completa"]

        formato = workbook.add_format({'num_format': '0.000000'})

        for col_num, _ in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, None, formato)




#--------------------------------------------------------------------------------------
#
#                          CONFIGURAÇÕES INICIAIS
# 
#      é aqui que tu vai colocar os valores que precisa pra esse trem funcionar
#
#--------------------------------------------------------------------------------------

#caso referencial for painel esquerdo do avião, pegar o valor 3-4
#caso referencial for painel direito do avião, pegar o valor 1-2
corda_frente = 0.556558  # corda do painel DLM da frente (da seção que tem o aileron, não complica porra)

corda_tras = 0.303303    # auto explicativo, não vou nem explicar aqui 

num_paineis_corda = 40   # tem como calcular o valor ideal para os paineis
                         # mas fica uma merda, pq ele sempre subestima   

df_malha_completa = gerar_tabela_completa(corda_frente, corda_tras, num_paineis_corda)

salvar_em_excel(df_malha_completa, "malhanew.xlsx")

print("Planilha Excel gerada com sucesso!")