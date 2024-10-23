import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrindo a Planilha

worrkbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')

# Esta pegando os dados apenas da parte da planilha Sheet1
Sheet_alunos = worrkbook_alunos['Sheet1']

# Ira pegar os dados a partir da linha 2
for i, linha in enumerate(Sheet_alunos.iter_rows(min_row=2)):
    # Pegando os Dados da planilha
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participante = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    # Transferir os dados para a Planilha
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    imagem = Image.open('./certificado_padrao.jpg')

    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1020, 827), nome_participante,
                  fill='black', font=fonte_nome)

    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)

    desenhar.text((1435, 1065), tipo_participante,
                  fill='black', font=fonte_geral)

    desenhar.text((1480, 1182), str(carga_horaria),
                  fill='black', font=fonte_geral)

    desenhar.text((750, 1770), str(data_inicio),
                  fill='black', font=fonte_data)

    desenhar.text((750, 1930), str(data_final), fill='black', font=fonte_data)

    desenhar.text((2220, 1930), str(data_emissao),
                  fill='black', font=fonte_data)

    imagem.save(f'./Certificados/{i}{nome_participante} certificado.png')
