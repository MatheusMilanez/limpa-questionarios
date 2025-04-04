# Projeto de Tratamento de Dados - Arena do Saber

Este programa foi desenvolvido com o objetivo de auxiliar no tratamento e análise de dados de um projeto de Iniciação Científica financiado pela **FAPERJ** (Fundação de Amparo à Pesquisa do Estado do Rio de Janeiro).

O projeto foi realizado com **alunos reais de uma universidade**, onde foi conduzida uma pesquisa educacional utilizando o framework **Arena do Saber** — um sistema de quiz que apresenta questões de múltipla escolha com fins pedagógicos.

## Objetivo

Automatizar o processamento das respostas fornecidas pelos alunos em dois questionários distintos (Questionário 94 e 95), cada um com **10 questões de múltipla escolha**. O objetivo é comparar as respostas dos alunos com o gabarito oficial e gerar relatórios consolidados em **Excel** e **PDF**.

## Estrutura dos Arquivos

- Os arquivos `.xlsx` contendo as respostas dos alunos devem ser inseridos na pasta `input/`.
- Os resultados processados serão gerados na pasta `output/`, nos formatos:
  - `resultado_94.xlsx` e `resultado_94.pdf`
  - `resultado_95.xlsx` e `resultado_95.pdf`

## Como Usar

1. **Instale as dependências**:

   ```bash
   npm install xlsx pdfkit
   ```

2. **Insira o arquivo `.xlsx` com os dados brutos** na pasta `input/`. O arquivo deve conter duas abas: `94` e `95`.

3. **Execute o script**:

   ```bash
   node index.js
   ```

4. Os relatórios serão gerados automaticamente na pasta `output/`.

## Observações

- O questionário **94** possui uma estrutura onde cada linha representa uma resposta de uma questão, com as colunas `Email`, `Questão` e `Resposta`.
- O questionário **95** possui uma estrutura onde cada linha representa todas as respostas de um aluno, com as colunas `Email` e `answer`, sendo que `answer` é um array com as respostas das 10 questões.

## Créditos

Este programa foi gerado com o auxílio de **Inteligência Artificial**, com o intuito de resolver um problema específico relacionado ao tratamento de dados educacionais de um projeto científico.
