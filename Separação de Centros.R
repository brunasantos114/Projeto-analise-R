# --------------------------------------------
# PACOTES
# --------------------------------------------

library(openxlsx)
library(dplyr)

# --------------------------------------------
# ARQUIVOS / PATHS
# --------------------------------------------
ARQUIVO_CONTROLE <- "Controle_inconformidades.xlsx"

downloads_path <- file.path(Sys.getenv("USERPROFILE"), "Downloads")
p_ <- function(x) file.path(downloads_path, x)
`%||%` <- function(a,b) if (!is.null(a) && length(a)>0 && !is.na(a)) a else b

# --------------------------------------------
# LEITURA DO ARQUIVO
# --------------------------------------------
dados <- read.xlsx(p_(ARQUIVO_CONTROLE))
colnames(dados) <- trimws(colnames(dados))

# --------------------------------------------
# DEFINIÇÃO DA COLUNA DE CENTRO
# --------------------------------------------

coluna_centro <- "nome_centro"

if (!(coluna_centro %in% colnames(dados))) {
  coluna_centro <- "ncentro"
}

# --------------------------------------------
# LIMPA E TRATA OS NOMES DE CENTRO
# --------------------------------------------
dados[[coluna_centro]] <- dados[[coluna_centro]] %>%
  as.character() %>%
  trimws() %>%
  ifelse(is.na(.), "Centro_desconhecido", .)

centros <- unique(dados[[coluna_centro]])

# --------------------------------------------
# GERAÇÃO DO NOVO ARQUIVO
# --------------------------------------------
wb <- createWorkbook()

for (centro in centros) {
  dados_centro <- dados %>% filter(.data[[coluna_centro]] == centro)
  
  
  nome_aba <- gsub("[[:punct:][:space:]]+", "_", centro)
  nome_aba <- substr(nome_aba, 1, 31)
  nome_aba <- nome_aba %||% "Centro_sem_nome"
  
  addWorksheet(wb, nome_aba)
  writeData(wb, nome_aba, dados_centro)
}

ARQUIVO_SAIDA <- "Controle_inconformidades_por_centro.xlsx"
saveWorkbook(wb, p_(ARQUIVO_SAIDA), overwrite = TRUE)

cat("✅ Novo relatório salvo em:", p_(ARQUIVO_SAIDA), "\n")


