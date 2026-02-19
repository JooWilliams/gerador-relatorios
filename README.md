# Gerador de Relatórios PDF

Este projeto automatiza a geração de relatórios em PDF para solicitações de autorização de sessões de terapia (Psicologia, Fonoaudiologia, etc.) a partir de uma planilha de controle de atendimentos.

## Funcionalidades

- Leitura de dados de uma planilha Excel (`atendimentos-cbmdf.xlsx`).
- Agrupamento de sessões por paciente e especialidade.
- Geração automática de PDFs formatados com os dados do paciente, convênio e quantidade de sessões solicitadas.
- Suporte a convênios específicos (CBMDF, PMDF).

## Estrutura do Projeto

- `gerador-relatorios-pdf.py`: Script principal de geração dos relatórios.
- `main.py`: Script de teste simples para geração de PDF.
- `atendimentos-cbmdf.xlsx`: Planilha de entrada com os dados dos atendimentos (não incluída no repositório por conter dados sensíveis).
- `relatorios/`: Pasta onde os PDFs gerados são salvos.

## Pré-requisitos

Certifique-se de ter o Python instalado e as bibliotecas necessárias:

```bash
pip install fpdf openpyxl
```

## Como Usar

1.  Certifique-se de que a planilha `atendimentos-cbmdf.xlsx` está na raiz do projeto e segue o formato esperado (colunas: Paciente, Tipo Atendimento, Plano, Data).
2.  Verifique se o arquivo de logo `logo-lavorato.png` está presente na raiz.
3.  Execute o script principal:

```bash
python gerador-relatorios-pdf.py
```

Os relatórios serão gerados automaticamente na pasta `relatorios/`.

## Configuração

As configurações de caminhos de arquivos e nomes de colunas podem ser ajustadas diretamente no início do arquivo `gerador-relatorios-pdf.py` na seção de CONFIGURAÇÕES.
