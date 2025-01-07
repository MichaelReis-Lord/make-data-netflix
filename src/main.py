# Facilitar a leitura de arquivos de dados
import pandas as pd
# Manipular diretórios no sistema operacional
import os
# Manipular diretorios e nome de arquivos em massa
import glob

# Caminho para ler os arquivos
folder_path = 'src\\data\\raw'

# Listar todos os arquivos de excel
excel_files = glob.glob( os.path.join(folder_path, '*.xlsx'))

# Caso não encontre nenhum arquivo excel
if not excel_files:
    print("Nenhum arquivo compátivel encontrado!")
else:
    
    # dataFrame = tabela na memória para guardar o conteúdos dos arquivos
    dfs = []

    # Caso encontre algum arquivo, tentar ler o arquivo
    for excel_file in excel_files:
       
    #    Função resnponsável por ler o arquivo
        try:
            # Ler arquivo excel
            df_temp = pd.read_excel(excel_file)

            # Pegar o nome do arquivo
            file_name = os.path.basename(excel_file)

            # Adicionar coluna com nome do arquivo de origem
            df_temp['Filename'] = file_name

            # Pesquisar pelo país de origem no nome do arquivo e adicionar uam coluna com pais no aruivo 
            if 'brasil' in file_name.lower():
                df_temp['Location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['Location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['Location'] = 'it'

            # Criar uma nova coluna chamada campanha
            df_temp['Campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
            
            # Guarda dados tratados dentro de df commum
            dfs.append(df_temp)

            print(df_temp)
        
        except Exception as e:
            print("Erro ao ler o arquivo! {excel_file} : {e}")

if dfs:

    # Concatena todas as tabelas salvas no dfs em uma única tabela
    result = pd.concat(dfs, ignore_index=True)

    # Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')

    # Cofigurou o motor de escrita
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Leva os dados do resultado a serem escritos no motor de excel configurado
    result.to_excel(writer, index=False)

    # Salva o arquivo de Excel
    writer._save()

# No caso de não encontrar arquivos
else:
    print('Não há nenhum dado a ser Salvo!')