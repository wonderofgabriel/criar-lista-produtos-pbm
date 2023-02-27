import os
import pandas as pd

eans = []
medicamentos = []
descontos = []
regras = []
fabricantes = []
programas = []
plataformas = []

# Importando a tabela de produtos
usuario = os.getlogin()
# Código do Bruno abaixo
arquivo = pd.read_excel(f'C:\\Users\\{usuario}\\Meu Drive\\PBM\\PRODUTOS PBM.xlsx', dtype=str)

# Código do Gabriel abaixo
# arquivo = pd.read_excel(f'C:\\Users\\{usuario}\\OneDrive\\Documentos\\PBM\\PRODUTOS PBM.xlsx', dtype=str)

# Nome dos laboratórios
nome_laboratorios = ['ACHE', 'ALCON', 'ASTRAZENECA', 'BAYER', 'BIOLAB', 'BOEHRINGER', 'BRACEPHARMA', 
'E.M.S', 'E.M.S', 'FQM', 'GSK', 'HYPERA', 'LILLY', 'NOVARTIS', 'PFIZER', 'U.SK', 'UCB BIOPHARMA', 
'UNITED MEDICAL', 'VIATRIS', 'ABBOTT', 'ALLERGAN', 'APSEN', 'CHIESI', 'MSD', 'MUNDIPHARMA', 'PERRIGO',
'SANOFI', 'SERVIER', 'ZODIAC']
nome_programas = ['CUIDADOS PELA VIDA', 'VALE MAIS VISAO', 'FAZBEM', 'BAYER PARA VOCE', 
'SAUDE EM EVOLUCAO', 'ABRACAR A VIDA','ABRACE-ME', 'PROGRAMA +OFTA', 'EMS SAUDE',
'CONSCIENCIA PELA VIDA', 'VIVER MAIS GSK', 'MANTECORP SAUDE', 'LILLY MELHOR PARA VOCE', 
'VALE MAIS SAUDE' ,'MAIS PFIZER', 'U.SK MAIS', 'COMPROMISSO SAUDE UCB', 'CAMINHANDO JUNTOS', 
'SE CUIDA','ACARE', 'VIVER MAIS ALLERGAN', 'SOU MAIS VIDA', 'ACESSAR','RECEITA DE VIDA', 
'PROGRAMA CUIDAR', 'PARA COMIGO','PROGRAMA VIVA', 'SEMPRE CUIDANDO', 'VIVER ZODIAC']

# Usuário digita os números dos laboratórios desejados
print('Digite o(s) número(s) do(s) laboratórios(s) que você deseja adicionar, separando-os por virgula.')
print()
contador = 0
for valor in range(29):
    if contador == 0:
        print('PORTAL DA DROGARIA:')
    elif contador == 19:
        print('\nFUNCIONAL:')
    else:
        pass
    print(f'{contador}. {nome_laboratorios[valor]} - {nome_programas[valor]}')
    contador += 1

selecao = input('\nDigite: ')
selecao = selecao.split(sep=',')
selecao_int = list()
for valor in selecao:
    selecao_int.append(int(valor))

for valor in selecao_int:
    produtos = arquivo.loc[arquivo['PROGRAMA'] == nome_programas[valor]]

    for x in produtos['EAN']:
        eans.append(x)

    for x in produtos['MEDICAMENTO']:
        medicamentos.append(x)

    for x in produtos['DESCONTO']:
        try:
            descontos.append(float(x) * 100)
        except ValueError:
            descontos.append('PREÇO SUGERIDO:')

    for x in produtos['REGRA']:
        regras.append(x)

    for x in produtos['FABRICANTE']:
        fabricantes.append(x)

    for x in produtos['PROGRAMA']:
        programas.append(x)

    for x in produtos['PLATAFORMA']:
        plataformas.append(x)

lista_produtos = {'EAN': eans,
                  'MEDICAMENTO': medicamentos,
                  'DESCONTO (%)': descontos,
                  'REGRA': regras,
                  'FABRICANTE': fabricantes,
                  'PROGRAMA': programas,
                  'PLATAFORMA': plataformas}

dataframe = pd.DataFrame.from_dict(lista_produtos)

cnpj = str(input('\nDigite o CNPJ da loja: '))

dataframe.to_excel(excel_writer=f'Produtos PBM {cnpj}.xlsx', sheet_name='Produtos', index=False)

print('Planilha gerada com sucesso.')
