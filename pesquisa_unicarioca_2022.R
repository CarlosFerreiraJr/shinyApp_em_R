#===============================================================================
# Pós Graduação em Ciência de Dados
#
# Trabalho de Conclusão da Disciplina:
# 1. Análise Computacional e Qualitativa de Dados - Prof. Manuel Martins 
#
# Autores do Trabalho:
#  1. Allan Santana
#  2. Bruce
#  3. Carlos Ferreira
#  4. Clovis 
#  5. Jefferson Anjos
#-------------------------------------------------------------------------------
# Histórico de Alterações:
# Data        Autor             Descrição
# 09/07/2022  Carlos Ferreira   Versão Inicial
# 13/072022   Clóvis Souza      Inclusão da Estatística 7
#===============================================================================

# Remove os warnings do Console do R Studio
options(warn=-1)


###---------------------------------------
### Inclusão das bibliotecas e pacotes
###---------------------------------------
if(!require(tidyverse))
{
  install.packages("tidyverse");
  require(tidyverse)
}


library(readxl)    #install.packages("readxl")
library(ggplot2)   #install.packages("ggplot2")
library(dplyr)
library(plyr)
library(tibble)
library(writexl)   #install.packages("writexl")
library(outliers)  #install.packages("outliers")
library(shiny)
library(readr)

###---------------------------------------
### Variáveis Globais e Constantes
###---------------------------------------
diretorio <- "C:\\05. UniCarioca-Pós-Graduação em Ciência de Dados\\04. Análise Computacional e Qualitativa de Dados\\Trabalho\\"
nome_arquivo <- "pesquisa.xlsx"
gsPathFile <- paste(diretorio,nome_arquivo, sep="")
extensao_arq_excel <- ".xlsx"
extensao_arq_txt <- ".txt"

## Gera o gráfico de Barras - Tipo 1
cores = c("#0000FF", "#F4A460", "#00FF00", "#8A2BE2", 
          "#FF00FF", "#DC143C", "#FFFF00", "#A9A9A9",
          "#66CDAA")

nomeEstatistica <- c("estat01", "estat02", "estat03", "estat04", "estat05", "estat06","estat07", "sumarios")

estatisticas <- c("Quantidade de Alunos por Modalidade" =  paste(diretorio, nomeEstatistica[1], extensao_arq_excel, sep=""),
                  "Percentual de Alunos por Modalidade" =  paste(diretorio, nomeEstatistica[2], extensao_arq_excel, sep=""),
                  "Quantidade de Alunos por Turma/Modalidade" =  paste(diretorio, nomeEstatistica[3], extensao_arq_excel, sep=""),
                  "Quantidade de Alunos por Disciplina/Modalidade" =  paste(diretorio, nomeEstatistica[4], extensao_arq_excel, sep=""),
                  "Quantidade de Alunos por Dia da Semana/Modalidade" =  paste(diretorio, nomeEstatistica[5], extensao_arq_excel, sep=""),
                  "Quantidade de Alunos por Sexo/Modalidade" =  paste(diretorio, nomeEstatistica[6], extensao_arq_excel, sep=""),
                  "Quantidade de Alunos por Periodo" =  paste(diretorio, nomeEstatistica[7], extensao_arq_excel, sep="") ,
                  "Análise Exploratória" =  paste(diretorio, nomeEstatistica[8], extensao_arq_txt, sep="") 
                 )

Ctrs <- estatisticas

##----------------------------------------------------
##----------------------------------------------------
### PROGRAMA PRINCIPAL
##----------------------------------------------------
##----------------------------------------------------

# Limpa o console do R
cat("\014")

## Leitura do Arquivo de Dados
df_excel <- read_excel(gsPathFile, sheet=1, .name_repair = "minimal")

## Cria vetores com cada coluna do data frame
strTurma <- df_excel[[1]]
strDiaSemana <- substr(df_excel[[1]],3,3)
strTurno <- substr(df_excel[[1]],4,4)
strDisciplina  <- df_excel[[2]]
strPeriodo <- df_excel[[3]]
strSexo  <- df_excel[[4]]
strPesquisa  <- df_excel[[5]]

#------------------------------------------------------------------------------------
#  Estatística 01
#  Quantidade de Alunos por Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strPesquisa)
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "0"] <- "Não Responderam"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "H"] <- "Híbrido"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "P"] <- "Presencial"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "R"] <- "Remoto"
dfestat01 <- as.data.frame(strResultadoPesquisa)
names(dfestat01)[names(dfestat01) == "Var1"] <- "coluna1"
names(dfestat01)[names(dfestat01) == "Freq"] <- "coluna2"

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[1], extensao_arq_excel, sep="")
write_xlsx(dfestat01, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 02
#  Percentual de Alunos por Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strPesquisa)
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "0"] <- "Não Responderam"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "H"] <- "Híbrido"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "P"] <- "Presencial"
names(strResultadoPesquisa)[names(strResultadoPesquisa) == "R"] <- "Remoto"
dfestat02 <- as.data.frame(strResultadoPesquisa)
names(dfestat02)[names(dfestat02) == "Var1"] <- "coluna1"
names(dfestat02)[names(dfestat02) == "Freq"] <- "coluna2"

# Lista a quantidade de linhas de um data frame
qtd_alunos <- nrow(df_excel)
dfestat02[2] <- round((dfestat02[2]/qtd_alunos)*100,2)

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[2], extensao_arq_excel, sep="")
write_xlsx(dfestat02, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 03
#  Quantidade de Alunos por Turma/Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strTurma, strPesquisa)
dfestat03 <- as.data.frame(strResultadoPesquisa, stringsAsFactors = FALSE)
names(dfestat03)[names(dfestat03) == "strTurma"] <- "coluna1"
names(dfestat03)[names(dfestat03) == "Freq"] <- "coluna2"
names(dfestat03)[names(dfestat03) == "strPesquisa"] <- "coluna3"
dfestat03["coluna3"][dfestat03["coluna3"] == "0"] <- "Não Responderam"
dfestat03["coluna3"][dfestat03["coluna3"] == "H"] <- "Híbrido"
dfestat03["coluna3"][dfestat03["coluna3"] == "P"] <- "Presencial"
dfestat03["coluna3"][dfestat03["coluna3"] == "R"] <- "Remoto"

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[3], extensao_arq_excel, sep="")
write_xlsx(dfestat03, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 04
#  Quantidade de Alunos por Disciplina/Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strDisciplina, strPesquisa)
dfestat04 <- as.data.frame(strResultadoPesquisa, stringsAsFactors = FALSE)
names(dfestat04)[names(dfestat04) == "strDisciplina"] <- "coluna1"
names(dfestat04)[names(dfestat04) == "Freq"] <- "coluna2"
names(dfestat04)[names(dfestat04) == "strPesquisa"] <- "coluna3"
dfestat04["coluna3"][dfestat04["coluna3"] == "0"] <- "Não Responderam"
dfestat04["coluna3"][dfestat04["coluna3"] == "H"] <- "Híbrido"
dfestat04["coluna3"][dfestat04["coluna3"] == "P"] <- "Presencial"
dfestat04["coluna3"][dfestat04["coluna3"] == "R"] <- "Remoto"

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[4], extensao_arq_excel, sep="")
write_xlsx(dfestat04, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 05
#  Quantidade de Alunos por Dia da Semana/Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strDiaSemana, strPesquisa)
dfestat05 <- as.data.frame(strResultadoPesquisa, stringsAsFactors = FALSE)
names(dfestat05)[names(dfestat05) == "strDiaSemana"] <- "coluna1"
names(dfestat05)[names(dfestat05) == "Freq"] <- "coluna2"
names(dfestat05)[names(dfestat05) == "strPesquisa"] <- "coluna3"
dfestat05 <- arrange(dfestat05,coluna1)
dfestat05["coluna3"][dfestat05["coluna3"] == "0"] <- "Não Responderam"
dfestat05["coluna3"][dfestat05["coluna3"] == "H"] <- "Híbrido"
dfestat05["coluna3"][dfestat05["coluna3"] == "P"] <- "Presencial"
dfestat05["coluna3"][dfestat05["coluna3"] == "R"] <- "Remoto"
dfestat05["coluna1"][dfestat05["coluna1"] == "2"] <- "2-Segunda-feira"
dfestat05["coluna1"][dfestat05["coluna1"] == "3"] <- "3-Terça-feira"
dfestat05["coluna1"][dfestat05["coluna1"] == "4"] <- "4-Quarta-feira"
dfestat05["coluna1"][dfestat05["coluna1"] == "5"] <- "5-Quinta-feira"
dfestat05["coluna1"][dfestat05["coluna1"] == "6"] <- "6-Sexta-feira"

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[5], extensao_arq_excel, sep="")
write_xlsx(dfestat05, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 06
#  Quantidade de Alunos por Sexo/Modalidade
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strSexo, strPesquisa)
dfestat06 <- as.data.frame(strResultadoPesquisa, stringsAsFactors = FALSE)
names(dfestat06)[names(dfestat06) == "strSexo"] <- "coluna1"
names(dfestat06)[names(dfestat06) == "Freq"] <- "coluna2"
names(dfestat06)[names(dfestat06) == "strPesquisa"] <- "coluna3"
dfestat06 <- arrange(dfestat06,coluna1)
dfestat06["coluna3"][dfestat06["coluna3"] == "0"] <- "Não Responderam"
dfestat06["coluna3"][dfestat06["coluna3"] == "H"] <- "Híbrido"
dfestat06["coluna3"][dfestat06["coluna3"] == "P"] <- "Presencial"
dfestat06["coluna3"][dfestat06["coluna3"] == "R"] <- "Remoto"
dfestat06["coluna1"][dfestat06["coluna1"] == "0"] <- "Não Informado"
dfestat06["coluna1"][dfestat06["coluna1"] == "1"] <- "Masculino"
dfestat06["coluna1"][dfestat06["coluna1"] == "2"] <- "Feminino"

# Grava a Estatística 01 em Excel para utilização pelo App Web
arquivo_saida = paste(diretorio, nomeEstatistica[6], extensao_arq_excel, sep="")
write_xlsx(dfestat06, arquivo_saida)


#------------------------------------------------------------------------------------
#  Estatística 07
#  Por Periodo
#------------------------------------------------------------------------------------

strResultadoPesquisa <- table(strPeriodo, strPesquisa)

dfestat07 <- as.data.frame(strResultadoPesquisa, stringsAsFactors = FALSE)
names(dfestat07)[names(dfestat07) == "strPeriodo"] <- "coluna1"
names(dfestat07)[names(dfestat07) == "Freq"] <- "coluna2"
names(dfestat07)[names(dfestat07) == "strPesquisa"] <- "coluna3"
dfestat07 <- arrange(dfestat07,coluna1)
dfestat07["coluna3"][dfestat07["coluna3"] == "0"] <- "Não Responderam"
dfestat07["coluna3"][dfestat07["coluna3"] == "H"] <- "Híbrido"
dfestat07["coluna3"][dfestat07["coluna3"] == "P"] <- "Presencial"
dfestat07["coluna3"][dfestat07["coluna3"] == "R"] <- "Remoto"
dfestat07["coluna1"][dfestat07["coluna1"] == "1P"] <- "1 Periodo"
dfestat07["coluna1"][dfestat07["coluna1"] == "2P"] <- "2 Periodo"
dfestat07["coluna1"][dfestat07["coluna1"] == "3P"] <- "3 Periodo"
dfestat07["coluna1"][dfestat07["coluna1"] == "4P"] <- "4 Periodo"
dfestat07["coluna1"][dfestat07["coluna1"] == "5P"] <- "5 Periodo"

arquivo_saida = paste(diretorio, nomeEstatistica[7], extensao_arq_excel, sep="")
write_xlsx(dfestat07, arquivo_saida)


#------------------------------------------------------------------------------------
#  Análise Exploratória
#  Quantidade de Alunos por Modalidade
#------------------------------------------------------------------------------------

# Variáveis globais desta Estatística
str_valor = ""
strMsg = ""
vet_retorno = ""

# Função f_Maior - Obtêm o maior valor de um dataset
# O objeto p_dataframe deve possuir o seguinte formato:
# p_dataframe[linha, coluna], onde:
# linha = são as linhas do dataframe
# coluna = o dataset deve possuir 2 colunas. 
#     coluna 1 = Campo do tipo String contendo o que se deseja obter
#     coluna 2 = Campo do tipo numérico que será usado na comparação
# Retornos: maior, menor, media, desvio padrão, mediana
f_estatistica <- function(p_dataframe)
{
  maior_valor = 0
  menor_valor = 99999
  str_maior_valor = ""
  str_menor_valor = ""
  ret_msg = ""
  soma = 0 #Quantidade total de alunos
  media = 0
  mediana = 0
  desvio_padrao = 0
  qtd_turma = 0 
  tb_generico <- table(p_dataframe)
  df_generico = as.data.frame(tb_generico)
  
  for (linha in 1:nrow(df_generico)) {
    if (df_generico[linha, 2] > maior_valor) {
      str_maior_valor = df_generico[linha, 1]
      maior_valor = df_generico[linha, 2]
    }
    if (df_generico[linha, 2] < menor_valor) {
      str_menor_valor = df_generico[linha, 1]
      menor_valor = df_generico[linha, 2]
    }
    qtd_turma = qtd_turma + 1
    soma = soma + df_generico[linha, 2]
  }
  media = round(mean(df_generico[, 2]),2)
  mediana = round(median(df_generico[, 2]),2)
  desvio_padrao = round(sd(df_generico[, 2]),2)
  
  ret_msg = paste(maior_valor, str_maior_valor, 
                  menor_valor, str_menor_valor,
                  soma,
                  qtd_turma,
                  media,
                  mediana,
                  desvio_padrao,
                  sep = ",")
  return (ret_msg)
} # Fim da função f_Maior

############# Rotina principal da Análise Exploratória

# Define uma varável para identificar o arquivo texto
arq_sumario <- paste(diretorio, nomeEstatistica[8], extensao_arq_txt, sep="")

# Grava o Header do arquivo TXT
strMsg = "Análise exploratória"
cat(strMsg, file = arq_sumario, append = FALSE, sep = "\n")  

vet_retorno = unlist(strsplit(f_estatistica(strTurma), split = ","))

strMsg = paste("Total geral de Alunos :", vet_retorno[5], sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("A Turma [", vet_retorno[2] , "] possui a maior quantidade de alunos: ", vet_retorno[1], " alunos", sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("A Turma [", vet_retorno[4] , "] possui a menor quantidade de alunos: ", vet_retorno[3], " alunos", sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("Quantidade de turmas :", vet_retorno[6], sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("Média alunos por turma :", vet_retorno[7], sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("Mediana de alunos por turma :", vet_retorno[8], sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")
strMsg = paste("Desvio Padrão de alunos por turma :", vet_retorno[9], sep ="")
cat(strMsg, file = arq_sumario, append = TRUE, sep = "\n")


###-----------------------------------------------------
### # Define UI for application that draws a histogram
###-----------------------------------------------------
ui <- fluidPage(
  
  # Application title
  titlePanel("Pesquisa sobre a Modalidade de Ensino"),
  
  # Sidebar with dropdown
  sidebarLayout(
    sidebarPanel(
      selectInput(inputId = "selects", 
                  choices = Ctrs,
                  selected = "Quantidade de Alunos por Modalidade",
                  label = "Selecione a Estatística", multiple = FALSE)
    ),
    
    # Show a plot of the generated distribution
    mainPanel(
      uiOutput(outputId = "Plot")
    )
  )
)

###---------------------------------------------------
### Define server logic required to draw a histogram
###---------------------------------------------------
server <- function(input, output) {
  
  
  escolha <- reactive({
    arquivo = substr(input$selects, nchar(input$selects)-11, nchar(input$selects))
    if (arquivo == "sumarios.txt") {   
      FALSE }
    else {
      TRUE }
  })
  
  grafico = reactive({
    
    arquivo = substr(input$selects, nchar(input$selects)-11, nchar(input$selects))
    if (arquivo != "sumarios.txt")    
    { 
      if ((arquivo == "estat01.xlsx") || (arquivo == "estat02.xlsx"))   { 
        dados = read_excel(input$selects, sheet=1, .name_repair = "minimal", range = cell_cols("A:B"), skip=1) }
      else {
        dados = read_excel(input$selects, sheet=1, .name_repair = "minimal", range = cell_cols("A:C"), skip=1) }
    }  
  })
  
  
  plot1 <- reactive({
    
    arquivo = substr(input$selects, nchar(input$selects)-11, nchar(input$selects))
    if (arquivo == "estat01.xlsx")
    {
      varx = "Modalidade"  
      vary = "Qtd de Alunos"
      strMensagem = paste("Estatística: Quantidade de Alunos por Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 200
    } 
    if (arquivo == "estat02.xlsx")
    {
      varx = "Modalidade"  
      vary = "Percentual de Alunos"
      strMensagem = paste("Estatística: Percentual de Alunos por Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 100
    } 
    
      # this should be a complete plot image
      mydata = grafico()
      ggplot(mydata,
             aes(x = coluna1, y = coluna2)) +
        geom_col() +
        scale_y_continuous(limits = c(0, limite))+
        geom_text(aes(label = coluna2), 
                  vjust = -1) +
        labs(title = strMensagem, x=varx, y=vary) +
        theme_grey(base_size = 10) 
    
  })
  
  plot2 <- reactive({
    varx = ""  
    vary = ""
    strMensagem = ""
    limite = 100
    strLegenda = ""
    coluna1=""
    coluna2=""
    coluna3=""

    arquivo = substr(input$selects, nchar(input$selects)-11, nchar(input$selects))
    if (arquivo == "estat03.xlsx")
    {
      varx = "Turma"  
      vary = "Quantidade de Alunos"
      strMensagem = paste("Estatística: Quantidade de Alunos por Turma/Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 100
      strLegenda = "Modalidade"
    }
    
    if (arquivo == "estat04.xlsx")
    {
      varx = "Disciplina"  
      vary = "Quantidade de Alunos"
      strMensagem = paste("Estatística: Quantidade de Alunos por Disciplina/Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 100
      strLegenda = "Modalidade"
    }
    
    if (arquivo == "estat05.xlsx")
    {
      varx = "Dia da Semana"  
      vary = "Quantidade de Alunos"
      strMensagem = paste("Estatística: Quantidade de Alunos por Dia da Semana/Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 100
      strLegenda = "Modalidade"
    }
    
    if (arquivo == "estat06.xlsx")
    {
      varx = "Sexo"  
      vary = "Quantidade de Alunos"
      strMensagem = paste("Estatística: Quantidade de Alunos por Sexo/Modalidade - Total de Alunos contemplados na pesquisa: ", qtd_alunos, sep = "")
      limite = 100
      strLegenda = "Modalidade"
    }
    
    # this should be a complete plot image
    mydata = grafico()
    ggplot(mydata,
           aes(x = coluna1, y = coluna2, fill = coluna3)) +
      geom_bar(stat="identity", position = "dodge") + 
      geom_col(aes(fill = coluna3), position = "dodge") +
      geom_text(aes(label = coluna2), position = position_dodge(0.9), vjust = 0) +
      labs(fill = strLegenda, title = strMensagem, x=varx, y=vary) 
  })
  
  graphInput <- reactive({
    arquivo = substr(input$selects, nchar(input$selects)-11, nchar(input$selects))
      if ((arquivo == "estat01.xlsx") || (arquivo == "estat02.xlsx")) {   
        plot1() }
      else {
        plot2() }
    })
  
  plot3 <- reactive({
    arquivo = input$selects
    Dadostxt <- read.delim(arquivo, header = TRUE)
    Dadostxt
  })
  
  
  output$Plot <- renderUI({
    if (escolha()) {
      renderPlot({ graphInput()  })  
    }  
    else {
      renderPrint({ plot3() })  
    }
  })
}


# Run the application 
shinyApp(ui = ui, server = server)
