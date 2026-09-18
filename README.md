# Assiduidades | Castanha

Aplicativo desktop (Python + Tkinter) desenvolvido para o **Supermercado Castanha**, usado pelo setor de RH para gerar automaticamente documentos de assiduidade e benefícios a partir de planilhas Excel, produzindo arquivos Word (.docx) prontos para impressão e assinatura.

## Funcionalidades

O aplicativo abre uma janela principal com botões que disparam a geração de cada tipo de documento, lendo os dados de uma planilha `.xlsx` correspondente e preenchendo um modelo `.docx`:

- **Assiduidade** — gera os termos de assiduidade a partir de `ASSIDUIDADE.xlsx`.
- **Assiduidade Castanhão** — versão do documento para a unidade Castanhão, a partir de `ASSIDUIDADE_CASTANHAO.xlsx`.
- **Autorização de Trabalho** — gera a autorização de trabalho em dias específicos conforme o aditamento da convenção coletiva dos comerciários.
- **Vale Natal** / **Vale Natal Castanhão** — geram os documentos de vale/prêmio de Natal por unidade.
- **Vale** — geração do vale padrão (ex.: vale-domingo).

Cada geração exibe uma janela de progresso (`ui_helpers.py`) enquanto os documentos são criados.

## Estrutura do projeto

```
main.py                      # Janela principal (Tkinter) e ponto de entrada
Assiduidade.py                # Lógica de geração do termo de assiduidade
Assiduidade_Castanhao.py       # Variante para a unidade Castanhão
Autorizacao_Trabalho.py        # Geração da autorização de trabalho em dias específicos
Vale.py / Vale_Natal.py        # Geração dos vales
Vale_Natal_Castanhao.py        # Variante do vale de Natal para o Castanhão
ui_helpers.py                  # Componentes visuais reutilizados (janela de progresso)
requirements.txt               # Dependências Python
*.xlsx                         # Planilhas de origem dos dados (nomes, CPF, etc.)
*.docx                         # Modelos/documentos gerados
*.spec                         # Specs do PyInstaller para gerar o executável
```

## Tecnologias

- **Python 3** com **Tkinter** para a interface gráfica.
- **pandas** + **openpyxl** para leitura das planilhas Excel.
- **python-docx** para geração dos documentos Word.
- **Pillow** para os ícones/logos da interface.
- **PyInstaller** para empacotar o aplicativo como executável Windows (`.exe`).

## Como executar

1. Crie um ambiente virtual e instale as dependências:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
2. Coloque as planilhas `.xlsx` de origem (com os dados dos funcionários) na mesma pasta do script — elas **não** fazem parte deste repositório (veja a observação abaixo) e precisam ser obtidas separadamente com o RH.
3. Execute:
   ```bash
   python main.py
   ```

## Gerando o executável

O projeto já inclui specs do PyInstaller (`main.spec`, `Assiduidades-atualizado-v2.spec`, `Assiduidades-tk-corrigido.spec`). Para gerar o `.exe`:

```bash
pyinstaller main.spec
```

O executável e os arquivos de suporte são gerados nas pastas `build/` e `dist/` (ignoradas pelo Git).

## Observação

Este repositório contém apenas o **código-fonte** do aplicativo. As planilhas `.xlsx` e os documentos `.docx` gerados contêm dados internos de RH (nomes, CPF, valores de benefícios) de funcionários do Supermercado Castanha e **nunca são versionados** (estão no `.gitignore`) — o repositório é público, mas esses arquivos devem ser mantidos e trocados apenas por fora do Git.
