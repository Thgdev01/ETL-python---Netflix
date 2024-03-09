import pandas as pd
import os
import glob

# caminho para ler o arquivo
folder_path = 'src\\data\\raw'

# lista todos os arquivos exel da pasta crua(raw)
excel_files = glob.glob(os.path.join(folder_path, '*xlsx'))

if not excel_files:
    print('Nenhum arquivo encontrado')
else:

    # dataframe = tabela na memória para guardar os conteúdos dos arquivos
    dfs = []

    for excel_file in excel_files:

        try:
             # ler o arquivo do excel
            df_temp = pd.read_excel(excel_file)

            # pegar o nome do arquivo
            file_name = os.path.basename(excel_file)

            df_temp['File Name'] = file_name

             # criando uma nova coluna chamada champaign
            if 'brasil' in file_name.lower():
                 df_temp['Location'] = 'From Brazil'
            elif 'france' in file_name.lower():
                df_temp['Location'] = 'From France'
            elif 'italian' in file_name.lower():
                df_temp['Location'] = 'From Italy'

            # criamos uma coluna chamada campaign
            
            df_temp['Campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            # guarda dados tratados dentro de uma dataframe

            dfs.append(df_temp)
            
            


        except Exception as e:
             print(f'Erro ao ler o arquivo {excel_file} : {e}')


        if dfs:

            # concatena todas as tabelas no dataframe(dfs) em uma única tabela
            result = pd.concat(dfs, ignore_index=True)

            # caminho de saída
            output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')

            # configurou o motor de escrita
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

            # leva os dados do resultado a serem escritos no motor do excel configurado
            result.to_excel(writer, index=False)

            # salva o arquivo no Excel
            writer._save()
        else:
            print('Nenhum dado para ser salvo')