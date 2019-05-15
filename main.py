import xlrd 
import csv

def tem_mesmas_colunas(paths):
    # Monta matriz de cabecalhos pra conferir se arquivos tem mesmos nomes de colunas
    cabecalho = []
    for i, file_path in enumerate(paths):
        wb = xlrd.open_workbook(file_path) 
        sheet = wb.sheet_by_index(0) 
        
        #Exemplo de acesso a dados
        #sheet.cell_value(0, 0) 
        
        cabecalho.append([])

        #Confere se Nomes de colunas são iguais em todos os arquivos
        for j in range(sheet.ncols): 
            valor = sheet.cell_value(0, j)
            cabecalho[i].append(valor)
            #print(i, j, '=', valor) 

    # Verifica se arquivos tem mesmos nomes de colunas
    is_igual = True # flag para verificação
    for i, e in enumerate(cabecalho[0]): # passa pelas colunas primeiro
        for j, ee in enumerate(cabecalho): # passa pelas linhas
            if e == cabecalho[j][i]:
                print('arquivo:', j, 'coluna:', i, 'temos', e, '===', cabecalho[j][i])
            else:
                is_igual= False
    if is_igual:
        print('\n\n>>> Todos os arquivos tem nome de colunas iguais: Continuando o processamento...\n\n')
        return True
    else:
        return False

def gera_matrix(paths):
    #Processamento para juntar arquivos
    print('Processamento para juntar arquivos iniciando...')
    temp = []
    for file_i, file_path in enumerate(paths):
        wb = xlrd.open_workbook(file_path) 
        sheet = wb.sheet_by_index(0) 
        #se primeiro arquivo então inclui cabeçalho
        if file_i==0:
            temp.append([]) # inclui linha
            for i in range(sheet.ncols): 
                temp[0].append(sheet.cell_value(0, i))
        #Insere linha a linha
        for i in range(sheet.nrows-1): # vai de 1 a n-1 
            linha = []
            temp.append(linha) # inclui linha
            for j in range(sheet.ncols): 
                linha.append(sheet.cell_value(i+1, j)) # vai de 1 a n-1
    return temp
    
def imprime_matriz(m):
    print('\nMatriz gerada:\n')
    for e in m:
        for ee in e:
            print('|', ee, end='')
        print('')

def gera_csv(m):
    print('\nGerando CSV...\n')
    with open('charles.csv', 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=';', quotechar=' ', quoting=csv.QUOTE_MINIMAL)
        for e in m:
            spamwriter.writerow(e)
        
def main(path):
    try:
        if tem_mesmas_colunas(paths):    
            m = gera_matrix(paths)
            #imprime_matriz(m)
            gera_csv(m)
        else:
            print('Arquivos não tem as mesmas colunas!')
    except:
        print('Exceção lançada...')

if __name__ == "__main__":
    
    paths = ['Despesas_2015_(Final).xlsx',
    'Despesas_Ano_2018_(Final).xlsx',
    'Despesas_2016_a_2017_(Final).xlsx',
    'Despesas_2019.xlsx',
    'Despesas GOV SC 2011 e 2012_(Final).xlsx',
    'Despesas GOV SC 2013 e 2014_(Final).xlsx' ]
    #paths = ['teste1.xlsx', 'teste2.xlsx', ]
    main(paths)