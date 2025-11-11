# Projeto-analise-R
- NECESSIDADE DO CLIENTE
Neste projeto, o cliente necessitava de organizar um arquivo excel, em que todas as informações estavam contempladas em uma unica aba, tornando o o arquivo extenso e dificil de análisar, diante disso, me procurou para propor uma solução que tornasse este arquivo mais intuitivo e de fácil análise.

- CARACTERÍSICA DOS DADOS
As informações da tabela contemplam inconsistencias/erros de preenchimento de uma base de dados. Nela é possivel ver quais centros possuem erros e qual tipo de erro está presente. Os erros podem ser de QUANTIDADE, CALORIA, STATUS ou ATRASO. 

- SOLUÇÃO
Através do código em R, importei o arquivo em questão e fiz a transformação em algumas colunas, de modo que o arquivo ficasse mais compacto horizonatalmente, sem muita necessidade de rolagem. Em seguida, o arquivo foi dividido em abas, onde cada aba apresenta informações de um centro espefícico. Após essa divisão, dentro de cada aba houve o agrupamento por tipo de erro, de modo que permitisse analisar, quais erros eram mais recorrentes no preenchimento da base principal.

- RESULTADO PARA O CLIENTE
Um arquivo excel separado por abas, no qual é possivel analisar centro por centro individualmente;
Arquivo mais compacto e com menos necessidade de rolagem verticam e horizonal, possibilitando uma fácil compreensão do que está acontecendo naquele contexto;

Tecnologias

**Linguagem:** R
**Pacotes Principais:** `dplyr` (Transformação e Agrupamento), `tidyr` (Pivotamento), `readxl`/`writexl` (Manipulação de Excel)
**Ferramentas:** RStudio, Git/GitHub
