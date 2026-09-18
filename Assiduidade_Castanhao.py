from tkinter import messagebox, Tk, Toplevel, ttk
import pandas as pd
import datetime
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import time
from ui_helpers import create_loading_dialog


class castanhao():
    try:
        def funcao_a_executar4(self):
            # Reutiliza a janela principal do aplicativo para o progresso.
            root = self.janela

            # Dimensões da janela de loading
            largura = 400
            altura = 100

            # Calcula posição central
            x = (root.winfo_screenwidth() // 2) - (largura // 2)
            y = (root.winfo_screenheight() // 2) - (altura // 2)

            loading, progresso = create_loading_dialog(root)
            # Use a função read_excel para ler o arquivo
            tabela = pd.read_excel('ASSIDUIDADE_CASTANHAO.xlsx')

            # cria os arrays
            NOME = []
            CTPS = []
            SERIE_CTPS = []
            CPF = []

            # Busca a data atual do sistema e converte no formato brasileiro
            data_atual = datetime.date.today()
            data = data_atual.strftime('%d/%m/%Y')

            # Iterar sobre as linhas
            for index, linha in tabela.iterrows():
                NOME.append(linha['Nome'])
                CTPS.append(linha['CTPS'])
                SERIE_CTPS.append(linha['Série CTPS'])
                CPF.append(linha['CPF'])
                # coluna_local = linha['Descrição do Local']

                # Conta a quantidade de linhas
                quant = len(NOME)
                valores = []

            progresso.configure(maximum=max(quant, 1))

            # Executa um loop gravando as informações no texto
            for dado in range(quant):
                arquivo = f"""          Castanhão Atacado Comercio e Distribuidora LTDA., inscrita no CNPJ/MF: 24.274.152/0001-40, empregador, através do presente Termo esclarece quanto ao pagamento em forma de crédito referente ao vale compra-assiduidade, nos termos e condições abaixo aduzidas:
             Considerando que o vale compra-assiduidade corresponde ao percentual de 3% sobre o valor do salário-base previsto na cláusula 7ª da CCT-2023/2024.\n
             Considerando que o Valor do Salário-Base previsto na Cláusula 28ª da CCT 2023/2024 é de R$1.882,00 x 3%, o valor do Vale Compra-Assiduidade é de R$ R$ 56,46 (Cinquenta e Seis Reais e Quarenta e Seis Centavos), na forma de crédito.
             Considerando que só terá direito ao respectivo recebimento do vale compra-assiduidade no mês, o funcionário que:
             1º Recebe salário até R$ 2.857,05 – Cláusula 27ª
             2º Não tenha faltado ao trabalho – sendo aceitas apenas as ausências decorrentes de Casamento, Falecimentos conforme previsto em Lei e na CCT, Falecimento de Sogro ou Sogra, Genro ou Nora e Licença Paternidade, as demais faltas mesmo que justificadas não farão jus ao recebimento do vale compra-assiduidade – Alínea “a” da Cláusula 27ª
             3º Não estiver afastado nos termos da lei, como auxílio doença, auxílio acidentário, auxílio maternidade ou gozando de férias. – Alínea “b” da Cláusula 27ª.\n
             Considerando que o valor do vale compra-assiduidade somente poderá ser utilizado para aquisição de produtos comercializados na própria empresa, conforme previsto na alínea “c” da cláusula 28ª da CCT 2023/2024.\n
             Considerando que o valor ofertado a título de vale compra-assiduidade não guarda qualquer relação com a remuneração ou salário in natura, não se valendo dos preceitos contidos no artigo 457 e artigo 458 da CLT.\n
             Eu {NOME[dado]}, CTPS: {CTPS[dado]}-{SERIE_CTPS[dado]}, declaro estar ciente e de pleno acordo com as condições que acima estão transcritas, que neste mês eu preencho os requisitos necessários ao recebimento do vale compra-assiduidade, e que neste ato, recebo através de crédito o valor de R$ 56,46 (Cinquenta e Seis Reais e Quarenta e Seis Centavos) estou ciente de que consumirei na aquisição de produtos comercializados pelo empregador Castanhão atacado comercio e distribuidora Ltda.
             São Paulo, {data}.\n
             ________________________________
             {NOME[dado]}
             CPF: {CPF[dado]}"""

                # Adicione a string ao array
                valores.append(arquivo)

                # Crie um novo documento do Word
                doc = Document()

                # Adicione um cabeçalho
                header = doc.sections[0].header
                paragraph = header.paragraphs[0]
                run = paragraph.add_run(
                    '      TERMO DE CONCESSÃO DO VALE COMPRA –ASSIDUIDADE – CLÁUSULA 27ª –\n                                      CONVENÇÃO COLETIVA 2023/2024\n\n')

                # Adiciona um texto em negrito
                font = run.font
                font.bold = True
                font.size = Pt(12)

                # Adicionando imagem ao rodapé
                section = doc.sections[0]
                footer = section.footer
                footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                footer.paragraphs[0].add_run().add_picture('Logo Castanhao.jpg', width=Inches(1.5), height=Inches(1.0))

                # Adicione as strings ao documento
                for texto in valores:
                    doc.add_paragraph(texto)
                    print('ok')
                    # Salve o documento
                    doc.save('ASSIDUIDADE_CASTANHAO.docx')
                    # Atualiza barra de progresso
                    progresso['value'] = dado + 1
                    loading.update()
                    time.sleep(0.1)

            loading.destroy()
            messagebox.showwarning(title="Assiduidades", message="Arquivo gerado com sucesso")

    except Exception as erro:
        messagebox.ERROR(f"{erro}")
