---
name: tabela-de-contratos
description: Use quando o advogado quiser levantar a mesma informação em vários contratos de uma pasta (por exemplo cláusula de não concorrência, vencimento, multa, foro) e receber uma planilha.
---

# Tabela de cláusulas de contratos

## Antes do plano, confirme com o advogado
- Qual pasta tem os contratos.
- Quais informações procurar em cada contrato.
- Se ele quer anonimização.

## Como procurar
- Leia todos os arquivos PDF e Word da pasta.
- Procure cada informação pelo sentido, não só pela expressão exata. Exemplo: "não concorrência" pode aparecer como "vedação de atuação", "exclusividade" ou "não competição". No plano, liste os termos que vai usar e pergunte se o advogado quer acrescentar algum.
- PDF digitalizado, sem texto: o programa marca "sem leitura", porque não consegue ler imagem. Depois de rodar o programa, leia você mesmo as páginas como imagem. Se conseguir ler, informe o trecho e a data na conversa e na coluna Observação, com a situação "Revisar: lido de imagem", nunca "Conferido", e avise que a leitura de imagem pode trocar números ou letras. Se a imagem estiver ilegível, mantenha "sem leitura". Nunca adivinhe o conteúdo.
- Prazo sem data certa (por exemplo "enquanto durar a relação comercial"): escreva "sem data certa" e marque para revisar.
- Cláusula que remete a anexo ou documento que não está na pasta: escreva "depende de documento ausente" e marque para revisar.

## O que entregar
Uma planilha contratos.xlsx com as colunas:
Arquivo | Trecho encontrado | Data | Fonte | Situação

- Trecho encontrado: copie o texto exatamente como está no contrato. Não resuma.
- Fonte: página, no caso de PDF; número da cláusula, no caso de Word, porque Word não tem página fixa.
- Situação: "Conferido" só quando o trecho e a fonte foram encontrados. Nos outros casos, "Revisar" ou "Sem leitura".
- Em uma aba separada, o resumo: quantos foram lidos, conferidos, para revisar e sem leitura.

Salve o programa em ferramentas/tabela_contratos.py, para o advogado rodar de novo quando chegarem contratos novos.
