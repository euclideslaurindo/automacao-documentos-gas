# Automator de Documentos Institucionais (Google Apps Script)

Este repositório contém um script em JavaScript (Google Apps Script) criado para automatizar a geração de documentos, como contratos, termos de rescisão e aditivos. 

O script consome dados de uma base instável no Google Sheets (com milhares de registros de funcionários) e gera os arquivos finais diretamente no Google Docs, prontos para uso.

## O gargalo

O processo original era inteiramente manual: cruzar linhas de uma planilha e preencher arquivos do Word. Isso consumia cerca de 10 a 15 minutos por documento e gerava problemas frequentes com erros de digitação em campos críticos, como datas e salários.

Com a implementação da automação, o tempo caiu para **cerca de 5 segundos por documento**, poupando dezenas de horas de trabalho braçal.

## Notas de Implementação

Na hora de passar o problema para código, esbarrei em algumas instabilidades da base de dados que exigiram lógicas específicas:

- **Busca Bottom-up (Histórico de registros):** A base mantinha os registros antigos misturados com os novos. Uma busca normal pegaria contratos já encerrados. Tive que inverter o laço de busca (`for` varrendo a base de trás para frente) para garantir que o script capture sempre a atualização mais recente do funcionário.
- **Mapeamento dinâmico de colunas:** Como a planilha é alimentada por várias pessoas, as colunas mudam de lugar. Criei um mapeador que localiza os dados pelo nome do cabeçalho em tempo real, usando validação de strings para não confundir colunas como "Endereço" e "UF Endereço".
- **Sanitização de inputs:** A carga horária muitas vezes vinha suja ("200h/a"). O script usa Regex para limpar os caracteres, cruzar com a tabela interna, calcular o salário e repassar ao documento.
- **Extenso monetário customizado:** Criei a função `escreverPorExtenso()` do zero para calcular e escrever os valores financeiros por extenso direto no documento final, sem depender de bibliotecas externas.

## Como o script opera

O arquivo `.js` é anexado no Google Apps Script da planilha base. A função `onOpen()` cria um menu dropdown nativo no Google Sheets, onde o usuário apenas seleciona a linha desejada e clica no tipo de documento que precisa gerar.
