library(bizdays)
library(dplyr)
library(tidyr)
library(RDCOMClient)

bizdays::bizdays.options$set(calendario)

#Define calendario dias uteis e D-1
data_ref <- bizdays::offset(lubridate::today(), -1)

#Define o formato da data
data_ref <- format(data_ref, "%Y%m%d")

#Define caminho de base
BASE_PATH <- paste0("define o caminho na data", data_ref, "arquivox")
end <- paste0(BASE_PATH)

#Lista os arquivos
Files <- list.files(end)
Files

#Localiza o arquivo desejado
filename_Arquivo <- file.path(end, Files[which(stringr::str_detect(Files, "Nome do arquivo"))])

#Lê o arquivo
df <- read.csv(filename_Arquivo)

#Cria a coluna x e define quantidade de letras de uma coluna em especifico
df <- df%>% mutate(x = substr(df$colunaemespecifico, 1,3))

#Filtra e agrupa a coluna, baseada no objeto, após soma outra coluna
df2 <- filter(df, x == "Objetoprocurado") %>% group_by(x) %>% summarise(`Outra Coluna` = sum(`Outra coluna`))

#Condição para se houver mais de x na soma, encaminhar um email
if(df2$`Outra Coluna` >= 100) {
  df_html <- xtable::xtable(df2)
  df_html <- print(df_html, typo = "html", include.rownames = F, print.results = T)
  OutApp <- COMCreate ("Outlook.Application")
  outMail = OutApp $ CreateItem (0)
  outMail$GetInspector()
  Signature <- outMail [["HTMLbody"]]
  corpo <- paste0("Texto escolhido", sprintf(df_html))
  outMail [["To"]] = "email de destino"
  outMail [["subject"]] = "Assunto"
  outMail [["HTMLbody"]] = paste0("<p>", corpo, "</p>", df_html, Signature)
  cat("Mensagem enviada. \n\n")
  outMail$Display()
}