# WordHunter - Processamento de dados

## Descrição:

WordHunter é uma ferramenta desenvolvida em Python que permite processar múltiplos arquivos (PDF, DOCX e DOC) para buscar palavras específicas, como "comodato" e "contrato", de forma sensível a maiúsculas e minúsculas (case sensitive). Com uma interface gráfica simples construída com Tkinter, o usuário pode selecionar arquivos, definir regras de busca e visualizar os resultados diretamente no aplicativo. Ideal para quem precisa analisar documentos rapidamente, como em contextos jurídicos, administrativos ou de pesquisa.
O projeto foi criado para ser flexível: novas regras de busca podem ser adicionadas facilmente ao código, tornando-o expansível para diferentes necessidades.

## Funcionalidades:
- Suporte a múltiplos formatos: Lê arquivos PDF, DOCX e DOC.
- Busca case sensitive: Identifica palavras exatas, como "comodato" (diferente de "Comodato").
- Regras personalizáveis: Busca por palavras definidas pelo usuário (ex.: "comodato", "contrato") e padrões como datas (DD/MM/YYYY).
- Interface gráfica: Seleção de arquivos e exibição de resultados via Tkinter.
- Processamento em lote: Analisa vários arquivos de uma só vez.

## Como funciona:
O usuário insere as palavras ou regras a serem buscadas (ex.: "comodato, contrato") em um campo de texto.
Seleciona os arquivos a serem processados usando um seletor de arquivos.
O programa verifica cada arquivo e retorna se as palavras foram encontradas ou não, exibindo os resultados na interface.

## Requisitos:
- Python 3.x
- Bibliotecas Python:
- tkinter (geralmente já vem com o Python)
- pdfplumber (para leitura de PDFs)
- python-docx (para leitura de DOCX)
- pywin32 (para leitura de DOC no Windows)
- Sistema operacional: Testado no Windows (para suporte a DOC); adaptações podem ser necessárias para Linux/Mac.

## Limitações:
- A leitura de arquivos DOC requer o Microsoft Word instalado (via pywin32).
- A busca por palavras é literal e sensível a maiúsculas/minúsculas.

## Autor:
Criado por: *Lucas Bevenuto*
