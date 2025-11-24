#* *********************************************************
#* Parâmetros iniciais
#* *********************************************************
#*
nomearq       = "/Users/jfrega/Documents/GitHub/CCAD-2024-2026-Rascunho/Dados/GRADE 2026.1 por extenso todas obrigatorias.xlsx"
senha         = "xpt0*!2025JrF-+"
sepcol        = FALSE
comentarios   = TRUE
travaPlanilha = FALSE
travaCopia    = FALSE 

colMN         = "lightgreen"
colNM         = "lightgreen"
colM2         = "lightblue"
colN2         = "lightblue"
colPeriodo    = "#F0F0A0"

#* *********************************************************
#* Função para colocar borda espessa em torno do conjunto
#* de células
#* *********************************************************
#*
OutsideBorders <-
  function(wb_,
           sheet_,
           rows_,
           cols_,
           border_col = "black",
           border_thickness = "medium") {
    left_col = min(cols_)
    right_col = max(cols_)
    top_row = min(rows_)
    bottom_row = max(rows_)
    
    sub_rows <- list(c(bottom_row:top_row),
                     c(bottom_row:top_row),
                     top_row,
                     bottom_row)
    
    sub_cols <- list(left_col,
                     right_col,
                     c(left_col:right_col),
                     c(left_col:right_col))
    
    directions <- list("Left", "Right", "Top", "Bottom")
    
    mapply(function(r_, c_, d) {
      temp_style <- createStyle(border = d,
                                borderColour = border_col,
                                borderStyle = border_thickness)
      addStyle(
        wb_,
        sheet_,
        style = temp_style,
        rows = r_,
        cols = c_,
        gridExpand = TRUE,
        stack = TRUE
      )
      
    }, sub_rows, sub_cols, directions)
  }

#* *********************************************************
#* Carrega bibliotecas necessárias
#* *********************************************************
#*
library(magrittr)
library(openxlsx)
library(sqldf)

#* *****************************************************
#* Leitura do arquivo de dados
#* *****************************************************
#*
read.xlsx(nomearq, sheet = 1, startRow = 2, cols = 1:11) -> df

#* *****************************************************
#* Pega todos os periodos disponiveis e monta as 
#* estruturas de dias da semana
#* 
colVazia = matrix(rep(NA, 7))
fullData = list()
df$Período = stringi::stri_trim_both(df$Período)
df$Período[which(df$Período=="EXT.")] = "DISCIPLINAS EXTENSIONISTAS"
unique(df$Período) -> periodo
diaSemana = c("SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO")
dfDiaSemana = data.frame(nDiaSemana = 1:6, diaSemana)

#* *****************************************************
#* Ordena o arquivo de dados por Período, Turno, 
#* Dia da Semana e Professor
#* 
"select a.* from df a, dfDiaSemana b where a.`Dia.da.Semana`=b.diaSemana order by a.Período, a.Turno, b.nDiaSemana, a.Professor " %>% sqldf -> df
strRemove = "<- NÃO OFERTAR"
which(substring(df$X8, 1, nchar(strRemove))!=strRemove | is.na(df$X8)) -> w
df = df[w,]

#* *****************************************************
#* Cria estilos de formatação para o EXCEL
#* (cabeçalho (linha de horário e dia da semana) e corpo)
#* 
headerStyle = createStyle(
  fontColour = "white",
  fgFill = "#000070",
  halign = "center",
  valign = "center",
  textDecoration = "Bold", 
  wrapText = TRUE
)

header2Style = createStyle(
  fontColour = "white",
  fgFill = "#2020C0",
  halign = "center",
  valign = "center",
  textDecoration = "Bold", 
  wrapText = TRUE
)

periodoStyle = createStyle(
  fontColour = "#3030AF",
  fgFill = colPeriodo,
  halign = "center",
  valign = "center",
  textDecoration = "Bold", 
  wrapText = TRUE, 
  fontSize = 16
)

sty = createStyle(
  fgFill = "white",
  halign = "center",
  valign = "center",
  wrapText = TRUE
)

colsepstyle = createStyle(
  fontColour = "#606060",
  fgFill = "#606060"
)

#* *****************************************************
#* Cria a pasta de trabalho
#* 
options("openxlsx.minWidth" = 0.2)
wb = createWorkbook()

#* *****************************************************
#* Para cada período
#*  
for(oPeriodo in periodo) {
  if(oPeriodo=="OUTROS CURSOS:") next
  # seleciona as linhas daquele período
  w = which(df$Período==oPeriodo)
  dfw = df[w,]
  # cria dataframe vazio para a planilha
  data = NULL
  # verifica o número de dias em cada dia da semana para determinar o máximo de colunas
  "select Turno, `Dia.da.Semana`, count(*) as n from dfw group by 1, 2 order by 2" %>% sqldf -> dfTemp
  dfTemp = rbind(dfTemp, data.frame(Turno = "PHANT", `Dia.da.Semana` = dfDiaSemana$diaSemana, n = 1))
  #* *****************************************************
  #* para cada um dos turnos
  #*
  nTurnos = 0
  for (oTurno in sort(unique(dfw$Turno))) {
    nTurnos = nTurnos + 1
    # Ajeita o formato dos horários
    if(oTurno=="MANHÃ") {
      turnoOficial = "07:30 às 11:30h"
    } else {
      if(oTurno=="NOITE") {
        turnoOficial = "19:00 às 23:00h"
      } else {
        if(oTurno=="11 ÀS 13H") {
          turnoOficial = "11:30 às 13:30h"
        } else {
          turnoOficial = oTurno
        }
      }
    }
    # Seleciona as linhas do turno específico
    dfww = dfw[which(dfw$Turno==oTurno),]
    if(sepcol) {
      data1 = colVazia
    } else {
      data1 = NULL
    }
    #* *****************************************************
    #* Para cada dia da semana
    #* 
    colsep = 1
    for(oDiaSem in diaSemana) {
      # Acho os elementos desse dia da semana e do turno
      dfwww = dfww[which(dfww$Dia.da.Semana==oDiaSem & dfww$Turno==oTurno),]
      data2 = NULL
      # se achei elementos
      if(nrow(dfwww)>0) {
        # faço o rbind (colunas para linhas)
        dfwww$Vagas[which(is.na(dfwww$Vagas))] = 55
        dfwww$Horas[which(is.na(dfwww$Horas))] = 60
        data2 = rbind(turnoOficial, 
                      dfwww$Dia.da.Semana, 
                      dfwww$`Disciplina.(código)`, 
                      dfwww$Turma, dfwww$Professor, 
                      dfwww$Área, 
                      paste0(
                        "(", 
                        dfwww$Vagas, 
                        " VAGAS - ", 
                        dfwww$Horas, 
                        " HORAS) ", 
                        ifelse(
                          comentarios, 
                          paste0(
                            ifelse(is.na(dfwww$X8), "", dfwww$X8), 
                            ifelse(is.na(dfwww$X9), "", dfwww$X9)
                          ), "")
                      )
        )
        # se não achei insere célula em branco para esse dia da semana
      } else {
        data2 = rbind(turnoOficial, oDiaSem, "", "", "", "", "")
      }
      data1 = cbind(data1, data2)
      # Vejo qual o máximo de colunas por turno 
      # para cada dia da semana
      w = which(dfTemp$Dia.da.Semana==oDiaSem)
      n = max(dfTemp$n[w])
      # se tem linhas para esse dia da semana e turno
      # completo com o máximo de colunas para esse dia
      # da semana
      if(!is.null(data2)) {
        if(n>ncol(data2)){
          for(k in 1:(n-ncol(data2))) { data1 = cbind(data1, NA) }
        }
      } else {
        if(n>0) {
          data2 = colVazia
          if(n>1) {
            for(k in 2:n) {
              data2 = cbind(data2, colVazia)
            }
          }
        }
      }
      #***********************************************
      #* Coloca coluna divisória entre dias da semana
      #* 
      if(sepcol) data1 = cbind(data1, rep(NA, 7))
      colsep = c(colsep, ncol(data1))
    }
    # # Monto o dataframe para salvar na planilha
    data = rbind(data, data1)
    data = rbind(data, NA)
  }
  # abro uma nova planilha para o período indicado
  sh = addWorksheet(wb, oPeriodo)
  
  # ajeito as alturas de linhas e escrevo o dataframe 
  if(!is.null(data)) removeRowHeights(wb, sh, 1:nrow(data)) # c(0:10*8+3, 0:10*8+7)) #, 56)
  
  # função para formatar as células com cores 
  fCorCelula = function(cod, col) {
    .sty = createStyle(fgFill = col, fontColour = "#802000", textDecoration = "Bold")
    l <- sapply(1:ncol(data), function(x) grep(cod, data[,x]))
    if(!is.null(l)) {
      for(i in 1:length(l)) {
        nelm = length(l[[i]])
        if(nelm>0){
          for(linha in 1:nelm) {
            addStyle(wb = wb, sheet = sh, style = .sty, rows = l[[i]][linha]+(-1:3)+1, cols = i, stack = TRUE, gridExpand = TRUE)
            # print(paste0("Achei ", cod, " na planilha ", sh, " na coluna ", i, " linha ", l[[i]][linha]))
          }
        }
      }
    }
  }
  
  
  # se a planilha tem dados
  if(!is.null(data)) {
    # aplica o estilo padrão
    addStyle(wb, sh, style = sty, rows = 1:(nrow(data)+1)+1, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)
    
    # coloca cores nas turmas não regulares
    fCorCelula("MN", colMN)
    fCorCelula("NM", colNM)
    fCorCelula("M2", colM2)
    fCorCelula("N2", colN2)
    
    # acerta a largura das colunas
    setColWidths(wb,sh,(1:ncol(data)),18)
    if(sepcol) {
      setColWidths(wb,sh,colsep, 0.05)
      addStyle(wb, sh, style = colsepstyle, 2:nrow(data)-1+1, colsep, gridExpand = TRUE, stack = TRUE)
    }
    
    # aplica o estilo para as linhas de cabecalho
    rows = ((1:nTurnos-1)*8+1)
    addStyle(wb, sh, style = headerStyle, rows = rows+1, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sh, style = header2Style, rows = rows+1+1, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)
    
    # cria as bordas espessas para cada grupo de células
    for (k in 1:ncol(data)) {
      for(kk in rows) OutsideBorders(wb, sh, kk-1+1:7+1, k)
    }
    
    # escreve o dataframe na planilha
    # Identifica o período
    if(oPeriodo %in% paste(1:8, "º", sep = "")) oPeriodo = paste0(oPeriodo, " PERÍODO")
    # Formata a primeira linha onde vai escrito o período
    setRowHeights(wb, sh, 1, 35)
    writeData(wb, sh, oPeriodo, startCol = 1, startRow = 1)
    addStyle(wb, sh, periodoStyle, rows = 1, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)
    mergeCells(wb, sh, cols = 1:ncol(data), rows = 1)
    OutsideBorders(wb, sh, 1, 1:ncol(data))
    # escreve o dataframe e coloca o default de impressão em paisagem
    # condensando em uma página
    writeData(wb, sh, data, colNames = FALSE, startRow = 2)
    pageSetup(wb, sh, orientation = "landscape", fitToWidth = TRUE, fitToHeight = TRUE)
    if(travaPlanilha) openxlsx::protectWorksheet(wb, sh, password = senha, lockSelectingUnlockedCells = travaCopia, lockSelectingLockedCells = travaCopia)
    # Guarda os dados em uma lista
    fullData[[oPeriodo]] = data
  }
}

#* *****************************************************
#* Cria a planilha de consistência dos docentes
#* 
df$Horas[which(is.na(df$Horas))] = 60
df$Vagas[which(is.na(df$Vagas))] = 55
"select Área, Professor, `Dia.da.Semana`, Turno, `Disciplina.(código)`, Horas, Vagas from df a, dfDiaSemana b where a.`Dia.da.Semana`=b.DiaSemana order by Professor, nDiaSemana, Turno" %>% sqldf -> ddf
sh = addWorksheet(wb, "CONSISTÊNCIA PROFESSORES")
setColWidths(wb,sh,1:ncol(ddf),"auto")
if(travaPlanilha) openxlsx::protectWorksheet(wb, sh, password = senha, lockSelectingUnlockedCells = travaCopia, lockSelectingLockedCells = travaCopia)
writeData(wb, sh, ddf)


#* *****************************************************
#* Insere uma planilha com os rateios de carga horária
#* por área
#* 
"select trim(Área) as Área, count(distinct(Professor)) as nProf, count(*) as nTurmas, sum(Horas) as totHoras, round(sum(Horas)/count(distinct(Professor)), 1) as horasPerCapita, round((count(*)+0.0)/count(distinct(Professor)),1) as turmasPerCapita from ddf group by 1 order by 1" %>% sqldf -> dfRateio
sh = addWorksheet(wb, "RATEIO DAS ÁREAS")
setColWidths(wb,sh,1:ncol(dfRateio),"auto")
if(travaPlanilha) openxlsx::protectWorksheet(wb, sh, password = senha, lockSelectingUnlockedCells = travaCopia, lockSelectingLockedCells = travaCopia)
writeData(wb, sh, dfRateio)

#* *****************************************************
#* Insere uma planilha com os rateios de carga horária
#* por docente
#* 
"select trim(Professor) as Professor, count(*) as nTurmas, sum(Horas) as totHoras, sum(Vagas) as TotVagas from ddf group by 1 order by 1" %>% sqldf -> dfRateio
sh = addWorksheet(wb, "CARGA POR PROFESSOR")
setColWidths(wb,sh,1:ncol(dfRateio),"auto")
if(travaPlanilha) openxlsx::protectWorksheet(wb, sh, password = senha, lockSelectingUnlockedCells = travaCopia, lockSelectingLockedCells = travaCopia)
writeData(wb, sh, dfRateio)

#* *****************************************************
#* Insere uma planilha com os dados de base para a 
#* geração
#* 
sh = addWorksheet(wb, "BASE PARA GERAÇÃO")
setColWidths(wb,sh,1:ncol(df),"auto")
if(travaPlanilha) openxlsx::protectWorksheet(wb, sh, password = senha, lockSelectingUnlockedCells = travaCopia, lockSelectingLockedCells = travaCopia)
writeData(wb, sh, df)

#* *****************************************************
#* Salva a pasta de trabalho e abre no EXCEL
#* 
if(travaCopia) openxlsx::protectWorkbook(wb, password = senha, lockStructure = TRUE)
fname = "/Users/jfrega/Documents/GitHub/CCAD-2024-2026-Rascunho/Dados/Grade Formato Blocos.xlsx"
saveWorkbook(wb, fname, overwrite = TRUE)
system(paste0("open '", fname, "'"))

#* *****************************************************
#* Copia a pasta de trabalho para o Onedrive
#* 
#* system(paste0('cp "', fname, '" "/Users/jfrega/Onedrive/CCAD-Arquivos/"'))
