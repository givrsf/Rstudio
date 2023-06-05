#Chama as Bibliotecas
library(bizdays)
library(dplyr)
library(tidyr)
library(rJava)
library(RDCOMClient)

#Define a data em que o código será rodado
data_ref <- bizdays::offset(lubridate::today(), -1)

#Informa onde será salvo o output
DIR_DESTINO <- paste("Destino")

#Caminho da Base
BASE_PATH <- paste0("Caminho")

#Informa que o arquivo está no caminho de base, na data de referencia
end <- paste0(BASE_PATH, data_ref, "/")

#Seleciona os arquivos da base
Files <- list.files(end)
Files

#Informa qual arquivo da base será utilizado
Filename_ativo_objeto <- file.path((end, Files[which(stringr::str_detect(Files, "AtivoObjeto_"))]))

#Le o arquivo
df <- read.csv(Filename_ativo_objeto)

#Define quais colunas serão utilizadas
new_df <- data.frame(df$coluna1, df$coluna2)

#Transforma espaços em branco em NA
new_df[new_df == ""] <- NA

#Omite todos os NA`s do arquivo
df2 <- na.omit(new_df)

#Altera o nome das colunas
names(df2) <- c("novo nome", "novo nome")

#Cria um arquivo xlsx para salvar o DF
wb <- openxlsx::createWorkbook()

openxlsx::addWorksheet(wb, "Ativo", grindLines = TRUE)

#Dine estivo do arquivo
plan_style <- openxlsx::crateStyleborder = "TopBotttomLeftRight", textDecoration = "Bold", halign = "Center", valign = "Center", wrapText = TRUE, fgFill = "#4F81BD", fontColour = "#ffffff", fontSize = 9)

#Escreve arquivo
openxlsx::writeData(wb, "Ativo", df2, rowNames = FALSE,
                    colNames = TRUE, startCol = "1", startRow = 1,
                    headerStyle = plan_style, borders = "surrounding",
                    borderColour = "black")

df_style <- openxlsx:: createStyle(textDecoration = "bold", fgFill=rgb(166,166,166, maxColorValue = 255))

openxlsx::addStyle(wb, "nome", style = df_style, rows = 75, cols = c(5), stack = TRUE)

#Define o estilo
Destino <- paste0(DIR_DESTINO, paste0("Arquivo", format(data_ref, "%Y%m%d"), ".xlsx"))
opnexlsx::saveWorkbook(wb, Destino, overwrite = TRUE)
}
AnexoParametros(data_ref)

#Informa onde será salvo o output
NEW_DEST <- paste0("Caminho/", data_ref, "/")

NEW_BASE <- paste0("Caminho/", data_ref, "/base")

#Define data D-1
data_d1 <- bizdays::offset(lubridate::today(), -1)

#Define data D-2 
data_d2 <- bizdays::offset(lubridate::today(), -2)

#Caminho de base D-1 
caminho_d1 <- paste0("Caminho/", data_d1, "/d-1", format(data_d1, "%Y%m%d"), ".xlsx")

#Caminho de base D-2
caminho_d2 <- paste0("Caminho/", data_d2, "/d-2", format(data_d1, "%Y%m%d"), ".xlsx")

#Lê o arquivo da base D-1
atvb1 <- readxl:: read_xlsx(caminho_d1)

#Lê o arquivo da base D-2
atvb2 <- readxl:: read_xlsx(caminho_d2)

#Efetua a junção dos arquivos, ddefinindo colunas que não alteram
referencia <- full_join(atvb1, atvb2, by=c('Modelo', 'De', 'Nome'))

#Filtra o que está diferente entre as informações, baseado na coluna informada
dif <- filter(referencia, `Nome da coluna.x` != `Nome da Coluna.y`)

#Cria uma condição para valores diferentes das colunas x e y
varia <- referencia %>% mutate(Variação = ifelse(`Coluna.x` != `Coluna.y`, {
  
#Condição 1 - Encaminhar um e-mail se caso diferentes
  df_html <- xtable::xtable(dif)
  df_html <- print(df_html type = "html", include.rownames = TRUE, print.results = TRUE)
  OutApp <- COMCreate ('Outlook.Application')
  outMail = OutApp $ CreateItem (0)
  outMail$GetInspector()
  Signature <- outMail [['HTMLbody']]
  corpo <- paste0('Texto para o e-mail',
                  sprintf(df_html))
  outMail [['To']] = 'Email_de_destino'
  outMail [['subject']] = 'Assunto'
  outMail [['HTMLbody']] = paste0('<p>', corpo, '</p>', df_html, Signature)
  path_to_file <- arquivo para anexo
  outMail[['Attachments']]$Add(path_to_file)
  cat('Mensagem de e-mail enviada. \n\n')
  outMail$Display()
    }, "Segunda condição, criar coluna e deixar a mensagem manteve"))

#Criando arquivo excel
final <- openxlsx::createWorkbook()

openxlsx::addWorksheet( final, 'Plan1', gridLines = TRUE)

#Alterando o estilo
header_Style <- openxlsx::createStyle(border = "TopBottomLeftRight", textDecoration = "Bold", halign = "Center", balign = "Center", wrapText = TRUE, fgFill = "#4F81BD", fontColour = "#ffffff", fontSize = 9)

#Definindo o nome do arquivo 
opnexlsx::writeData(final, "Plan1", varia, rowNames = FALSE, colnames = TRUE, startCol = "1", startRow = 1, headerStylel = header_Style, borders = "surrounding", borderColour = "black")

new_style <- openxlsx::createStyle(textDecoration = "bold", fgFill=rgb(166,166,166, maxColorValue = 255))

openxlsx::addStyle(final, "Plan1", style = new_style, rows = 75, cols = c(8), stack = TRUE)

#Infomando onde o arquivo deve ser salvo
dest_final <-paste0 (NEW_DEST, paste0("Nome do arquivo", format(data_ref, "%Y%m%d"), ".xlsx"))

#Salvando o arquivo
opnexlsx::saveWorkbook(final, dest_final, overwrite = TRUE)