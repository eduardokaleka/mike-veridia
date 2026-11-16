# Mike, by Veridia

**Você descreve. O Mike constrói. Você confere.**

O Mike é um assistente de programação para advogados que não sabem programar. O advogado descreve, em português, uma tarefa que se repete no trabalho dele. O Mike pergunta o objetivo jurídico, apresenta um plano sem código, espera a aprovação, constrói uma ferramenta, testa nos arquivos do advogado e entrega o resultado com a fonte de cada dado e a lista do que não conseguiu conferir.

O Mike não é um modelo de IA novo. Ele é uma configuração (regras, permissões e modelos de tarefa) em volta do Claude Code, o assistente de programação da Anthropic. O modelo é o mesmo; o que muda é como ele se comporta.

Projeto da APS de Ciência de Dados Aplicada ao Direito I, Insper, 2026.2.
Grupo Veridia: Eduardo Kaleka, Guilherme Martins, Guilherme Rodovalho, João Henrique Carvalho e Murilo Manoel.
Site do produto: https://mike-veridia.netlify.app

## Como funciona

1. **Você descreve a tarefa.** Do mesmo jeito que pediria a um estagiário. Se faltar informação, o Mike pergunta para que o resultado vai servir.
2. **O Mike apresenta um plano.** Em português, sem código: o que vai construir, o que vai ler e como vai testar. Nada é construído antes da sua aprovação.
3. **A ferramenta é construída e testada.** O Mike roda a ferramenta nos seus arquivos e corrige os erros que aparecerem.
4. **Você recebe a ferramenta e o relatório.** Cada dado vem com o trecho copiado do documento e a página de onde saiu. O que não foi possível conferir aparece numa lista separada.

## O que tem neste repositório

| Arquivo | O que é |
|---|---|
| `CLAUDE.md` | As regras de comportamento do Mike: como falar com o advogado, plano antes de agir, nunca inventar dado, diferença entre "não encontrado" e "ausente", sigilo, histórico e limites. |
| `.claude/settings.json` | As travas que o programa obriga a cumprir, independente do que o modelo decidir: o Mike sempre abre em modo de planejamento e não tem acesso à internet. |
| `.claude/skills/tabela-de-contratos/SKILL.md` | O modelo de tarefa "tabela de contratos": levantar a mesma cláusula em vários contratos, procurando pelo sentido e não só pela palavra exata. |
| `ferramentas/tabela_contratos.py` | A ferramenta que o Mike construiu no teste. Fica salva para rodar de novo quando chegarem contratos novos. |
| `Contratos de fornecedores/` | Os 10 contratos fictícios usados no teste. |
| `contratos.xlsx` | A planilha que o Mike entregou. |
| `historico.md` | O registro que o Mike escreveu: o que foi pedido, o plano aprovado e o que foi entregue. |

## Como usar

1. Instale o Claude Code e o Python no computador.
2. Baixe esta pasta e abra o Claude Code dentro dela.
3. Coloque os seus documentos numa subpasta e descreva a tarefa, por exemplo: *"Quero uma tabela com todos os contratos da pasta, com a cláusula de não concorrência e a data de vencimento de cada um. Se não tiver data certa, me avise."*

É preciso uma assinatura do Claude que inclua o Claude Code.

## O teste

Criamos 10 contratos fictícios de fornecimento com armadilhas planejadas: cláusulas com outros nomes ("vedação de atuação", "exclusividade", "não competir"), cláusula na página 3, data escrita por extenso, prazo sem data certa, remissão a um anexo que não está na pasta, um PDF escaneado sem texto e um contrato sem a cláusula.

**Resultado:** o Mike perguntou o objetivo, apresentou o plano, construiu a ferramenta e, ao conferir o próprio trabalho, achou e corrigiu um erro (a data por extenso do contrato 009 tinha saído como 31/01/2020 em vez de 31/01/2028). No fim, todas as linhas bateram com o gabarito, com trecho literal e fonte.

**Comparação com o Claude Code sem configuração:** demos o mesmo pedido ao Claude Code puro. Ele acertou as mesmas datas, e mais rápido. A diferença está no resto: ele não perguntou o objetivo, não mostrou plano, resumiu as cláusulas em vez de copiá-las, não indicou a página, não deixou arquivo nem ferramenta salva, e apresentou o dado do PDF escaneado como certo. O Mike não acerta mais. Ele entrega um resultado que o advogado consegue conferir e repetir.

**Teste com contratos novos:** depois, demos ao Mike 3 contratos que nem ele nem a ferramenta tinham visto, com armadilhas novas: uma cláusula de "não aliciamento" (parecida, mas que não é não concorrência), um aditivo que prorroga o prazo da cláusula e um prazo contado a partir do fim do contrato. A ferramenta salva, sozinha, errou 2 dos 3: não reconheceu a forma "abster-se-á" e não percebeu o aditivo. O Mike encontrou os dois erros ao comparar o resultado com o texto dos contratos, corrigiu a ferramenta e rodou de novo os 10 contratos antigos para garantir que nada tinha mudado. Resultado final: 3 de 3 corretos, com o aditivo marcado para revisão (as duas datas aparecem e quem decide qual vale é o advogado) e a cláusula de não aliciamento avisada, sem ser confundida com não concorrência.

## Limites conhecidos

- **Sigilo:** o texto dos documentos é processado nos servidores da Anthropic. O Mike não envia nada a terceiros sem autorização, e a anonimização reduz o que chega ao modelo, mas os documentos não ficam 100% no computador.
- **Pedidos de autorização:** o Mike pede confirmação muitas vezes. Isso protege, mas cansa, e quem se cansa passa a aprovar sem ler.
- **Ferramenta e modelo:** a ferramenta salva procura por palavras. Quem entende sinônimos e lê PDF escaneado é o modelo. Por isso a ferramenta, sozinha, marca o escaneado como "sem leitura", e contratos escritos de um jeito novo podem exigir que o Mike ajuste a ferramenta, como aconteceu no teste com contratos novos.
- **Teste pequeno:** 10 contratos fictícios, criados com IA. O site fala em centenas de contratos reais.
- **Uma ferramenta pronta:** das várias que o site lista, construímos a tabela de contratos, que é o exemplo principal.

## Próximos passos

- Ensinar a ferramenta a ler PDF escaneado sem internet (reconhecimento de texto).
- Reduzir os pedidos de autorização, liberando de antemão o que é só leitura.
- Criar os outros modelos do site: painel de contencioso, comparador de versões e cronologia dos fatos.

## Uso de IA

Este projeto foi feito com IA, como o professor permitiu. O código da ferramenta foi escrito pelo próprio Mike (Claude Code). O rascunho das regras do `CLAUDE.md` e os contratos de teste foram feitos com o Claude e revisados pelo grupo. O relatório de uso de IA detalha o que usamos, o que testamos e descartamos, e por quê.
