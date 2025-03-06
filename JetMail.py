import os
import pandas as pd
import yagmail

df = pd.read_excel('MENSALISTAS-MAR-2025.xlsx')

yag = yagmail.SMTP('maanaimbpo@gmail.com', 'vxrp rgoh tlnv qwxp')

for index, row in df.iterrows():
    nome = row['RESPONSAVEL']
    destinatario = row['Contato']
    pdf_filename = f'pdfs/{nome}.pdf'

    # Verifica se o arquivo PDF existe no caminho especificado.
    if not os.path.exists(pdf_filename):
        print(f'O arquivo {pdf_filename} não foi encontrado. Pulando o envio para {destinatario}.')
        continue  # Pula para o próximo destinatário

    corpo = f'Olá {nome}, segue em anexo o seu arquivo PDF.'

    try:
        yag.send(destinatario, 'Seu arquivo PDF', [corpo, pdf_filename])
        print(f'E-mail enviado para {destinatario}')
    except Exception as e:
        print(f'Erro ao enviar para {destinatario}: {e}')
