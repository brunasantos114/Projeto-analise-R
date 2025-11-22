# Projeto-analise-R

FASE 1

- NECESSIDADE DO CLIENTE
Neste projeto, o cliente necessitava de organizar um arquivo excel, em que todas as informações estavam contempladas em uma unica aba, tornando o o arquivo extenso e dificil de análisar, diante disso, me procurou para propor uma solução que tornasse este arquivo mais intuitivo e de fácil análise.

- CARACTERÍSICA DOS DADOS
As informações da tabela contemplam inconsistencias/erros de preenchimento de uma base de dados. Nela é possivel ver quais setores possuem erros e qual tipo de erro está presente. Os erros podem ser de QUANTIDADE, CALORIA, STATUS ou ATRASO. 

- SOLUÇÃO
Através do código em R, importei o arquivo em questão e fiz a transformação em algumas colunas, de modo que o arquivo ficasse mais compacto horizonatalmente, sem muita necessidade de rolagem. Em seguida, o arquivo foi dividido em abas, onde cada aba apresenta informações de um setor espefícico. Após essa divisão, dentro de cada aba houve o agrupamento por tipo de erro, de modo que permitisse analisar, quais erros eram mais recorrentes no preenchimento da base principal.

- RESULTADO PARA O CLIENTE
Um arquivo excel separado por abas, no qual é possivel analisar setor por setor individualmente;
Arquivo mais compacto e com menos necessidade de rolagem verticam e horizonal, possibilitando uma fácil compreensão do que está acontecendo naquele contexto;

Tecnologias

**Linguagem:** R
**Pacotes Principais:** `dplyr` (Transformação e Agrupamento), `tidyr` (Pivotamento), `readxl`/`writexl` (Manipulação de Excel)
**Ferramentas:** RStudio, Git/GitHub

FASE 2 - 

- NECESSIDADE DO CLIENTE
Após a construção de um controle de inconformidades mais estruturado, chegou a hora de alimentar esta tabela. Com esta nova estrurura definida, o "crescimento" da tabela ocorreria apenas na vertical. 
Para isso, foi criado um script de Atualização Controle de Respostas, no qual as pendências resolvidas pelos setores deveriam preencher o campo de retorno, data em que foi resolvida e possíveis observações de um determinado participante, visita e erro. Caso a pendência não estivesse nesta lista, a linha inteira deveria ser acrescentada no controle de erros.

Estas respostas vem de um outro documento excel (tambem construido por mim e que será abordado posteriormente). Este documento é gerado, entregue aos setores para que eles preenchem se os erros mencionados foram resolvidos e posteriormente devolvido ao profissional responsavel. Em seguida, o controle de erros é atualizado e este, servirá de filtro para o novo relatorio gerado, filtre os erros já corrigidos. 

Resumidamente o fluxo é: Codigo 1 (ATUALIZA CONTROLE DE ERROS) -> Código 2 (BUSCA OS NOVOS ERROS NA BASE PRINCIPAL; ORGANIZA-OS; FILTRA OS ERROS CORRIGIDOS) E IMPRIME ESTE EXCEL CONTEMPLANDO OS NOVOS ERROS JÁ COM O FILTRO DOS ERROS RESOLVIDOS APLICADOS -> ENVIA ESTE NOVO DOCUMENTO COM OS ERROS PENDENTES AO SETORES -> ELES RESPONDEM NOVAMENTE COM OS NOVOS ERROS CORRIGIDOS... e assim segue.

