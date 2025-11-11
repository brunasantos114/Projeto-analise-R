# --------------------------------------------
# PACOTES
# --------------------------------------------
# install.packages("openxlsx")
# install.packages("dplyr")
# install.packages("lubridate")
# install.packages("stringi")
library(openxlsx)
library(dplyr)
library(lubridate)
library(stringi)

# --------------------------------------------
# ARQUIVOS / PATHS
# --------------------------------------------
ARQUIVO_CONTROLE <- "Controle_inconformidades.xlsx"

downloads_path <- file.path(Sys.getenv("USERPROFILE"), "Downloads")
p_ <- function(x) file.path(downloads_path, x)
`%||%` <- function(a,b) if (!is.null(a) && length(a)>0 && !is.na(a)) a else b

# --------------------------------------------
# LEITURA E PADRONIZAÇÃO
# --------------------------------------------
dados <- read.xlsx(p_(ARQUIVO_CONTROLE))
colnames(dados) <- colnames(dados) |>
  trimws() |>
  stri_trans_general("Latin-ASCII") |>
  tolower() |>
  gsub("[^a-z0-9]+", "_", x = _)

# --------------------------------------------
# IDENTIFICA COLUNA DE CENTRO
# --------------------------------------------
coluna_centro <- if ("nome_centro" %in% names(dados)) "nome_centro" else "ncentro"

# --------------------------------------------
# GARANTE EXISTÊNCIA DAS COLUNAS DE VERIFICAÇÃO
# --------------------------------------------
cols_verif <- c("verificacao_status", "verificacao_kcal", "verificacao_quantidade")
cols_just  <- c("justificativa_verificacao_status", "justificativa_verificacao_kcal", "justificativa_verificacao_quantidade")
for (c in c(cols_verif, cols_just)) {
  if (!(c %in% names(dados))) dados[[c]] <- NA
}

# --------------------------------------------
# CRIA NOVAS COLUNAS
# --------------------------------------------
dados <- dados %>%
  mutate(
    tipo_de_erro = case_when(
      !is.na(verificacao_status) & verificacao_status != "" ~ "Status",
      !is.na(verificacao_kcal) & verificacao_kcal != "" ~ "Kcal",
      !is.na(verificacao_quantidade) & verificacao_quantidade != "" ~ "Quantidade",
      TRUE ~ NA_character_
    ),
    descricao_erro = case_when(
      tipo_de_erro == "Status" ~ verificacao_status,
      tipo_de_erro == "Kcal" ~ verificacao_kcal,
      tipo_de_erro == "Quantidade" ~ verificacao_quantidade,
      TRUE ~ NA_character_
    ),
    resposta = case_when(
      tipo_de_erro == "Status" ~ justificativa_verificacao_status,
      tipo_de_erro == "Kcal" ~ justificativa_verificacao_kcal,
      tipo_de_erro == "Quantidade" ~ justificativa_verificacao_quantidade,
      TRUE ~ NA_character_
    )
  )

# --------------------------------------------
# CÁLCULO DE MESES EM ABERTO
# --------------------------------------------
dados <- dados %>%
  mutate(
    data = suppressWarnings(as.Date(data, origin = "1899-12-30")),
    data = if_else(is.na(data), as.Date(NA), data),
    qtd_meses_em_aberto = {
      data_ref <- floor_date(Sys.Date(), "month") - months(1)
      if_else(!is.na(data),
              interval(floor_date(data, "month"), data_ref) %/% months(1),
              NA_integer_)
    }
  )

# --------------------------------------------
# INCLUI ATRASADO COMO TIPO DE ERRO
# --------------------------------------------
dados <- dados %>%
  mutate(
    tipo_de_erro = if_else(tolower(status24h) == "atrasado", "Atrasado", tipo_de_erro)
  )

# --------------------------------------------
# ENCONTRA COLUNA DE DATA DE RESOLUÇÃO
# --------------------------------------------
col_data_resolucao <- grep("^data_de_resolucao$", names(dados), value = TRUE)
if (length(col_data_resolucao) == 0) {
  stop("❌ Nenhuma coluna correspondente a 'Data de resolução' encontrada mesmo após padronização.")
}

# --------------------------------------------
# SELECIONA E RENOMEIA COLUNAS
# --------------------------------------------
dados <- dados %>%
  mutate(Observacoes = NA_character_) %>%
  select(
    npac, data, visita, grupo, tipo_de_erro, status24h, diasemana, kcal,
    descricao_erro, resposta, all_of(col_data_resolucao), qtd_meses_em_aberto, Observacoes,
    all_of(coluna_centro)
  ) %>%
  rename(
    ID = npac,
    Data = data,
    `Tipo Visita` = visita,
    Grupo = grupo,
    `Tipo Erro` = tipo_de_erro,
    `Status 24h` = status24h,
    diasemana = diasemana,
    Calorias = kcal,
    `Descrição Erro` = descricao_erro,
    Resposta = resposta,
    `Data Resolução` = all_of(col_data_resolucao),
    `Qtd Meses em Aberto` = qtd_meses_em_aberto,
    Observações = Observacoes,
    Centro = all_of(coluna_centro)
  )

# --------------------------------------------
# GERAÇÃO DO NOVO ARQUIVO COM SEÇÕES
# --------------------------------------------
wb <- createWorkbook()
headerStyle <- createStyle(textDecoration = "bold", halign = "center", valign = "center", border = "Bottom")

centros <- unique(dados$Centro)
for (centro in centros) {
  dados_centro <- dados %>% filter(Centro == centro)
  
  nome_aba <- gsub("[[:punct:][:space:]]+", "_", centro)
  nome_aba <- substr(nome_aba, 1, 31)
  nome_aba <- nome_aba %||% "Centro_sem_nome"
  addWorksheet(wb, nome_aba)
  
  tipos <- c("Kcal", "Quantidade", "Status", "Atrasado")
  linha_atual <- 1
  
  for (tipo in tipos) {
    subset <- dados_centro %>% filter(`Tipo Erro` == tipo)
    if (nrow(subset) > 0) {
      # título
      writeData(wb, nome_aba, paste0("ERRO ", toupper(tipo)), startRow = linha_atual, colNames = FALSE)
      linha_atual <- linha_atual + 1
      
      # dados
      writeData(wb, nome_aba, subset %>% select(-Centro),
                startRow = linha_atual, colNames = TRUE, headerStyle = headerStyle)
      
      linha_atual <- linha_atual + nrow(subset) + 3
    }
  }
}

ARQUIVO_SAIDA <- "Controle_inconformidades_por_centro_formatado.xlsx"
saveWorkbook(wb, p_(ARQUIVO_SAIDA), overwrite = TRUE)

cat("✅ Relatório final salvo em:", p_(ARQUIVO_SAIDA), "\n")
