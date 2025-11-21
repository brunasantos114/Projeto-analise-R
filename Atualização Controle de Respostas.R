# ============================================
# ATUALIZAÇÃO DO CONTROLE VIA PASTA 09
# Versão: 19/11/25
# ============================================

library(openxlsx)
library(dplyr)
library(stringi)

downloads <- file.path(Sys.getenv("USERPROFILE"), "Downloads")
arquivo_controle <- file.path(downloads, "Controle_inconformidades_por_centro.xlsx")
pasta_respostas <- file.path(downloads, "09")

if (!file.exists(arquivo_controle)) stop("Arquivo de controle não encontrado.")
if (!dir.exists(pasta_respostas)) stop("Pasta '09' não encontrada.")

abas_controle <- getSheetNames(arquivo_controle)
wb_controle <- loadWorkbook(arquivo_controle)


norm <- function(x) {
  x <- as.character(x)
  x <- tolower(x)
  x <- stri_trans_general(x, "Latin-ASCII")
  x <- gsub("\\s+", "_", x)
  x <- gsub("\\.", "_", x)
  x <- trimws(x)
  x
}

cols_padrao <- c("id","tipo_visita","tipo_erro","resposta","data_resolucao","observacoes")

for (aba in abas_controle) {
  message("🔄 Processando aba: ", aba)
  
  
  df_ctrl <- read.xlsx(arquivo_controle, sheet = aba)
  if (nrow(df_ctrl) == 0) next
  
 
  arqs <- list.files(pasta_respostas, pattern = paste0("^", aba), full.names = TRUE)
  if(length(arqs) == 0) next
  
  df_resp <- read.xlsx(arqs[1])
  if(nrow(df_resp) == 0) next
  
 
  names(df_resp) <- norm(names(df_resp))
  names(df_ctrl) <- norm(names(df_ctrl))
  
  
  for(col in cols_padrao) {
    if(!(col %in% names(df_resp))) df_resp[[col]] <- NA
  }
  
  
  if("calorias" %in% names(df_resp)) df_resp$calorias <- suppressWarnings(as.numeric(df_resp$calorias))
  if("calorias" %in% names(df_ctrl)) df_ctrl$calorias <- suppressWarnings(as.numeric(df_ctrl$calorias))
  
  cols_texto <- c("id", "tipo_visita", "tipo_erro", "resposta", "observacoes", "grupo", "descricao_erro")
  for(cx in cols_texto) {
    if(cx %in% names(df_resp)) df_resp[[cx]] <- as.character(df_resp[[cx]])
    if(cx %in% names(df_ctrl)) df_ctrl[[cx]] <- as.character(df_ctrl[[cx]])
  }
  
  
  df_ctrl <- df_ctrl %>% mutate(
    id_norm = norm(id),
    tv_norm = norm(tipo_visita),
    te_norm = norm(tipo_erro),
    .join   = paste(id_norm, tv_norm, te_norm, sep="||")
  )
  
  df_resp <- df_resp %>% mutate(
    id_norm = norm(id),
    tv_norm = norm(tipo_visita),
    te_norm = norm(tipo_erro),
    .join   = paste(id_norm, tv_norm, te_norm, sep="||")
  )
  
 
  lookup_resp <- df_resp %>% group_by(.join) %>% slice(1) %>% ungroup() %>%
    select(.join, resposta, data_resolucao, observacoes)
  
  df_joined <- df_ctrl %>% left_join(lookup_resp, by=".join", suffix=c("",".resp"))
  
  atualizadas_count <- 0
  for(r in c("resposta","data_resolucao","observacoes")){
    col_resp <- paste0(r,".resp")
    if(col_resp %in% names(df_joined)){
      idx <- which(!is.na(df_joined[[col_resp]]) & nzchar(as.character(df_joined[[col_resp]])))
      if(length(idx)>0){
        df_joined[[r]][idx] <- df_joined[[col_resp]][idx]
        atualizadas_count <- atualizadas_count + length(idx)
      }
      df_joined[[col_resp]] <- NULL
    }
  }
  
  
  keys_ctrl <- unique(df_joined$.join)
  keys_resp <- unique(df_resp$.join)
  keys_novas <- setdiff(keys_resp, keys_ctrl)
  
  adicionadas_count <- 0
  df_novos_out <- NULL
  
  if(length(keys_novas) > 0){
    df_novos <- df_resp %>% filter(.join %in% keys_novas)
    cols_ctrl <- names(df_ctrl)[!names(df_ctrl) %in% c("id_norm","tv_norm","te_norm",".join")]
    
    colunas_faltantes <- setdiff(cols_ctrl, names(df_novos))
    if(length(colunas_faltantes) > 0) {
      for(cf in colunas_faltantes) { df_novos[[cf]] <- NA }
    }
    
    df_novos_out <- df_novos[, cols_ctrl, drop=FALSE]
    adicionadas_count <- nrow(df_novos_out)
  }
  
  
  df_final <- df_joined %>% select(-id_norm, -tv_norm, -te_norm, -.join)
  
  if(!is.null(df_novos_out)) {
    df_final <- bind_rows(df_final, df_novos_out)
  }
  
  
  
  if(all(c("tipo_erro", "id") %in% names(df_final))) {
    df_final <- df_final %>% 
      arrange(tipo_erro, id, tipo_visita)
  }
  
  
  writeData(wb_controle, sheet=aba, x=df_final, withFilter=FALSE)
  
  message(sprintf("   → %s: %d atualizadas | %d NOVAS |",
                  aba, atualizadas_count, adicionadas_count))
}

saveWorkbook(wb_controle, arquivo_controle, overwrite=TRUE)
message("✅ Processo finalizado. Controle atualizado ")