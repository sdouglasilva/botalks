"""

Descrever os passos manuais e transformar isso em código

"""
#ler planilha e guardar informações sobre nome e telefone
#criar links personalizados do whatsapp e enviar mensagens para cada cliente com base nos dados da planilha

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(10)
workbook = openpyxl.load_workbook('Leads - TALKS.xlsx')
pagina_leads = workbook['repiqueTALKS']
for linha in pagina_leads.iter_rows(min_row=2):
  #nome e telefone
  nome = linha[0].value
  telefone = linha[2].value
  mensagem = f'Olá {nome}🛑 *ESSA É NOSSA ULTIMA TENTATIVA DE CONTATO*🛑, Estamos em período de repescagem de candidatos que não retornaram ao primeiro contato do projeto, sendo assim todos serão avaliados para distribuição de nossas últimas 13 bolsas, você deseja prosseguir? Lembrando que a bolsa para estudar inglês pela https://escolatalks.com.br/ é de 100% e o candidato é isento de todas suas matrículas, mensalidades, e taxas de certificado, ficando responsável financeiramente somente por seus materiais didáticos. Para mais informações acesse : https://www.projetobrasilmaisbilíngue.com.br'
link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
webbrowser.open(link_mensagem_whatsapp)
print(mensagem)
# print(nome)
# print(telefone)