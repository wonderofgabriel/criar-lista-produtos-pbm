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
arquivo = pd.read_excel(f'C:\\Users\\{usuario}\\Meu Drive\\PBM\\PRODUTOS PBM.xlsx', skiprows=3, dtype=str)
# arquivo = pd.read_excel(f'C:\\Users\\pbm\\OneDrive\\Documentos\\PBM\\PRODUTOS PBM.xlsx', skiprows=3, dtype=str)
# Nome dos laboratórios
laboratorios = ['ACHE', 'ALCON', 'ASTRAZENECA', 'BAYER', 'BIOLAB', 'BOEHRINGER',
                'E.M.S', 'FARMOQUIMICA', 'GERMED', 'GSK', 'HYPERA', 'LILLY', 'NOVARTIS', 'PFIZER',
                'U.SK', 'UCB BIOPHARMA', 'UNITED MEDICAL', 'VIATRIS', 'ABBOTT',
                'ALLERGAN', 'APSEN', 'CHIESI', 'MSD', 'MUNDIPHARMA',
                'PERRIGO', 'SANOFI', 'SERVIER', 'ZODIAC']

# Usuário digita os números dos laboratórios desejados
print('Digite o(s) número(s) do(s) laboratórios(s) que você deseja adicionar, separados por virgula.')
print()
contador = 0
for laboratorio in laboratorios:
    if contador == 0:
        print('PORTAL DA DROGARIA:')
    elif contador == 18:
        print('\nFUNCIONAL:')
    else:
        pass
    print(f'{contador}. {laboratorio}')
    contador += 1

laboratorios_selecionados = input('\nDigite: ')
laboratorios_selecionados = laboratorios_selecionados.split(sep=',')
laboratorios_selecionados_int = list()
for valor in laboratorios_selecionados:
    laboratorios_selecionados_int.append(int(valor))
# print(laboratorios_selecionados_int)

for valor in laboratorios_selecionados_int:
    produtos = arquivo.loc[arquivo['FABRICANTE'] == laboratorios[valor]]

    for x in produtos['EAN']:
        eans.append(x)

    for x in produtos['MEDICAMENTO']:
        medicamentos.append(x)

    for x in produtos['DESCONTO']:
        try:
            descontos.append(float(x) * 100)
        except ValueError:
            descontos.append('Consultar regra')

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
