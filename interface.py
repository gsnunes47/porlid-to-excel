import time
from colarExcel import colarExcel
from mesclarExcel import mesclarExcel
from listarEmultiplicar import listarEmultiplicar

def printc(str):
    print()
    print(str)
    print()

printc("Bem vindo ao Porlid - Excel")
printc("Passe o caminho completo do porlid.txt que deseja converter.")
printc("Para copiar basta clicar com o botão direito em cima do arquivo e selecionar a opção 'copiar como caminho'")
printc("Alternativamente também pode se usar o atalho ctrl + shift + c")

file_path = str(input('Caminho: ')).replace('"', '')

printc("O arquivo final sera arquivado no mesmo diretório.")
printc("Por favor, digite o nome do arquivo final.")

output_name = str(input('Nome: ')) + '.xlsx'

path_sample = ''
split = file_path.split('\\')
split.pop()
for i in split:
    path_sample += i + '\\'

output_path = path_sample + output_name

colarExcel(file_path)

time.sleep(1.5)

colunas = mesclarExcel(path_sample + 'dados-convertidos.xlsx')

time.sleep(1.5)

printc("Iniciando a majoração das colunas.")

printc("Insira o valor do majorador e aperte enter.")

listarEmultiplicar(file_path=path_sample + 'dados-convertidos.xlsx', output_path=output_path, colunas=colunas)

printc("Conversão concluída e majoradores aplicados.")  
