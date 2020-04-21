
library(openxlsx)
library(Hmisc)
library(stringr)
library(zoo)
library(nplr)
library(WriteXLS)
library(readxl)

analyze = function(raw.data.files = "20191108 resazurin R.xlsx", 
                   info.data.files = "20191108 resazurin information.xlsx",
                   raw.data.files.name = "20191108 resazurin R",
                   pos.control.cutoff = 25000){
  #read raw data
  data = list()
  for (file.name in raw.data.files) {
    if (!is.na(file.name)) {
      nSheets = getSheetNames(file.name)
      for (j in 1:length(nSheets)) {
        table = read.xlsx(xlsxFile = file.name, sheet = j,  rows = 7:54, cols = 2:25, colNames = T)
        data = append(data, list(list(raw.data = table)))
      }
    }
  }
  n = length(data)
  information = read.xlsx(xlsxFile = info.data.files, colNames = T) 
  agent1 = information$Agent1 
  agent2 = information$Agent2
  agent1 = agent1[!is.na(agent1)]
  agent2 = agent2[!is.na(agent2)]
  plate = information$Plate.ID
  plate = plate[!is.na(plate)]
  cellline = information$Cell.line
  cellline = cellline[!is.na(cellline)]
  dose1 = information$Dose1
  dose2 = information$Dose2
  dose1 = dose1[!is.na(dose1)]
  dose2 = dose2[!is.na(dose2)]
  m = length(dose2)
  dir.create(file.path(paste("Chalice template ", raw.data.files.name, sep = "")))
  dir.create(file.path(paste("Plots data ", raw.data.files.name, sep = "")))
  
  ## analysis part
  ## extract number
  for ( k in 1:n){
    combination = sapply(dose1, function(x) eval(bquote(expression(.(agent2[k])~" + "~.(x)~mu*M~.(agent1[k])))))
    drugname = c(agent2[k],agent1[k],combination)                     
    platedata = data[[k]]$raw.data
    index_1 = seq(1,45,4)
    index_2 = seq(2,46,4)
    rawcount = platedata[index_1,]
    excount = platedata[index_2,]
    rawcount = sapply(rawcount, as.numeric)
    excount = sapply(excount, as.numeric)
    
    ## average positive control
    poscontrol = as.numeric(c(excount[-c(3:6),2], excount[,23]))
    ## cutoff 
    poscontrol = poscontrol[which(poscontrol>pos.control.cutoff)]
    poscontrol_average = mean(poscontrol)
    
    # percentage transfer
    percent = excount[, 3:22]/poscontrol_average*100
    
    ## drug concentration
    concentration = rep(25,m)
    for (i in 2:m){
      concentration[i] = concentration[i-1]*0.5
    }
    concentration_log = log10(concentration)
    
    ## plots
    agent2_1_index = seq(1,2*m,2)
    agent2_1 = percent[1,agent2_1_index]
    agent2_2_index = seq(2,2*m,2)
    agent2_2 = percent[1,agent2_2_index]
    agent2_3_index = seq(1,2*m,2)
    agent2_3 = percent[2,agent2_3_index]
    agent2_4_index = seq(2,2*m,2)
    agent2_4 = percent[2,agent2_4_index]
    
    agent1_1_index = seq(1,2*m,2)
    agent1_1 = percent[3,agent1_1_index]
    agent1_2_index = seq(2,2*m,2)
    agent1_2 = percent[3,agent1_2_index]
    agent1_3_index = seq(1,2*m,2)
    agent1_3 = percent[4,agent1_3_index]
    agent1_4_index = seq(2,2*m,2)
    agent1_4 = percent[4,agent1_4_index]
    
    C4_1_index = seq(1,2*m,2)
    C4_1 = percent[5,C4_1_index]
    C4_2_index = seq(2,2*m,2)
    C4_2 = percent[5,C4_2_index]
    C4_3_index = seq(1,2*m,2)
    C4_3 = percent[6,C4_3_index]
    C4_4_index = seq(2,2*m,2)
    C4_4 = percent[6,C4_4_index]
    
    C3_1_index = seq(1,2*m,2)
    C3_1 = percent[7,C3_1_index]
    C3_2_index = seq(2,2*m,2)
    C3_2 = percent[7,C3_2_index]
    C3_3_index = seq(1,2*m,2)
    C3_3 = percent[8,C3_3_index]
    C3_4_index = seq(2,2*m,2)
    C3_4 = percent[8,C3_4_index]
    
    C2_1_index = seq(1,2*m,2)
    C2_1 = percent[9,C2_1_index]
    C2_2_index = seq(2,2*m,2)
    C2_2 = percent[9,C2_2_index]
    C2_3_index = seq(1,2*m,2)
    C2_3 = percent[10,C2_3_index]
    C2_4_index = seq(2,2*m,2)
    C2_4 = percent[10,C2_4_index]
    
    C1_1_index = seq(1,2*m,2)
    C1_1 = percent[11,C1_1_index]
    C1_2_index = seq(2,2*m,2)
    C1_2 = percent[11,C1_2_index]
    C1_3_index = seq(1,2*m,2)
    C1_3 = percent[12,C1_3_index]
    C1_4_index = seq(2,2*m,2)
    C1_4 = percent[12,C1_4_index]
    
    data_plot = as.data.frame(cbind(concentration, concentration_log, 
                                    agent2_1, agent2_2, agent2_3, agent2_4,
                                    agent1_1, agent1_2, agent1_3, agent1_4, C1_1, C1_2, C1_3, C1_4,
                                    C2_1, C2_2, C2_3, C2_4, C3_1, C3_2, C3_3, C3_4, 
                                    C4_1, C4_2, C4_3, C4_4))
    agent2_mean = rowMeans(data_plot[,3:6], na.rm = T)
    agent2_sd = apply(data_plot[,3:6], 1, sd)
    agent2_CIL =  agent2_mean -  agent2_sd
    agent2_CIU =  agent2_mean +  agent2_sd
    agent1_mean = rowMeans(data_plot[,7:10], na.rm = T)
    agent1_sd = apply(data_plot[,7:10], 1, sd)
    agent1_CIL = agent1_mean - agent1_sd
    agent1_CIU = agent1_mean + agent1_sd
    C1_mean = rowMeans(data_plot[,11:14], na.rm = T)
    C1_sd = apply(data_plot[,11:14], 1, sd)
    C1_CIL = C1_mean - C1_sd
    C1_CIU = C1_mean + C1_sd
    C2_mean = rowMeans(data_plot[,15:18], na.rm = T)
    C2_sd = apply(data_plot[,15:18], 1, sd)
    C2_CIL = C2_mean - C2_sd
    C2_CIU = C2_mean + C2_sd
    C3_mean = rowMeans(data_plot[,18:22], na.rm = T)
    C3_sd = apply(data_plot[,18:22], 1, sd)
    C3_CIL = C3_mean - C3_sd
    C3_CIU = C3_mean + C3_sd
    C4_mean = rowMeans(data_plot[,23:26], na.rm = T)
    C4_sd = apply(data_plot[,23:26], 1, sd)
    C4_CIL = C4_mean - C4_sd
    C4_CIU = C4_mean + C4_sd
    
    data_plot_summary = as.data.frame(cbind(concentration, concentration_log, agent2_mean, agent1_mean, C1_mean, 
                                            C2_mean, C3_mean, C4_mean, agent2_CIL, agent1_CIL, C1_CIL,
                                            C2_CIL, C3_CIL, C4_CIL, agent2_CIU, agent1_CIU, C1_CIU,
                                            C2_CIU, C3_CIU, C4_CIU))
    ## prepare template for Combenefit
    template1 = as.data.frame(cbind(concentration, agent2_mean, agent1_mean, C1_mean, 
                                    C2_mean, C3_mean, C4_mean))
    template1[template1=="NaN"] = NA
    template2 = template1[order(template1$concentration),]
    template2 = t(template2)
    template3 = template2[-1,]
    colnames(template3) = template2[1,]
    doseB = c(0,dose1)
    agentB = NULL
    agentB[1] = 100
    if (agent1[k] == "Brigatinib"){
      for (i in 1:4){
        agentB[i+1] = template1$agent1_mean[6-i]
      }
    }else{
      for (i in 1:4){
        agentB[i+1] = template1$agent1_mean[7-i]-((template1$concentration[7-i]-doseB[i+1])/
                                                    (template1$concentration[7-i]-template1$concentration[8-i])*
                                                    (template1$agent1_mean[7-i]-template1$agent1_mean[8-i]))
      }
    }
    agentB = round(agentB, digits = 3)
    template4 = as.data.frame(cbind(doseB, agentB,template3[-2,]))
    template4 = as.data.frame(rbind(c(NA,0,round(template2[1,], digits = 2)), template4))
    template5 = template4
    template4[1,13] = "(=Agent 2)"
    template4[7,1] = "(=Agent 1)"
    template4[8,1] = "Agent 1"
    template4[8,2] = agent1[k]
    template4[9,1] = "Agent 2"
    template4[9,2] = agent2[k]
    template4[10,1] = "Unit"
    template4[10,2] = "uM"
    template4[11,1] = "Unit"
    template4[11,2] = "uM"
    template4[12,1] = "Title"
    template4[12,2] = paste(agent1[k], " vs. ", agent2[k], " in MODEL LOEWE")
    dir.create(file.path(paste("Combenefit batch ", raw.data.files.name, sep = ""), paste(agent2[k], ".", agent1[k], ".", "Plate", k, sep = "")), recursive = TRUE)
    WriteXLS(template4, 
             ExcelFileName = paste("Combenefit batch ", raw.data.files.name, "/", agent2[k], ".", agent1[k], ".", "Plate", k, "/", agent2[k], ".", agent1[k], ".", "Plate", k, ".xls", sep = ""),
             SheetNames = "Sheet 1", perl = "perl",
             verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
             row.names = F, col.names = F,
             AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
             na = "",
             FreezeRow = 0, FreezeCol = 0,
             envir = parent.frame())
    
    ### prepare template for Chalice score
    template6 = as.data.frame(rbind(template5[6,], template5[5,], template5[4,],
                                    template5[3,], template5[2,], template5[1,]))
    # slope = (c(TAE226_mean[-1], 0)-TAE226_mean)/(c(concentration[-1], 0)-concentration)
    # A = NULL
    # for ( j in 1:4){
    #   A[j] = TAE226_mean[3+j]+(doseB[6-j]-concentration[3+j])*slope[2+j]
    # }
    # A[5] = 100
    template6[1:5,2:12] = 100 - template6[1:5,2:12]
    template6[template6=="NaN"] = NA
    WriteXLS(template6, 
             ExcelFileName = paste("Chalice template ", raw.data.files.name, "/", "Chalice.", agent2[k], ".", agent1[k], ".", "Plate", k, ".", cellline[k], ".xls", sep = ""),
             SheetNames = "Sheet 1", perl = "perl",
             verbose = FALSE, Encoding = c("UTF-8", "latin1", "cp1252"),
             row.names = F, col.names = F,
             AdjWidth = FALSE, AutoFilter = FALSE, BoldHeaderRow = FALSE,
             na = "",
             FreezeRow = 0, FreezeCol = 0,
             envir = parent.frame())
    
    
    ### plot prepare
    data_plot_agent2 = as.data.frame(cbind(concentration_log, agent2_mean, agent2_CIL, agent2_CIU))
    data_plot_agent1 = as.data.frame(cbind(concentration_log, agent1_mean, agent1_CIL, agent1_CIU))
    data_plot_C1 = as.data.frame(cbind(concentration_log, C1_mean, C1_CIL, C1_CIU))
    data_plot_C2 = as.data.frame(cbind(concentration_log, C2_mean, C2_CIL, C2_CIU))
    data_plot_C3 = as.data.frame(cbind(concentration_log, C3_mean, C3_CIL, C3_CIU))
    data_plot_C4 = as.data.frame(cbind(concentration_log, C4_mean, C4_CIL, C4_CIU))
    
    data_plot_all = NULL
    for (i in seq(2,22,4)){
      table = data_plot[,c(1, 1+i, 2+i, 3+i, 4+i)]
      longdata = reshape(table, 
                         varying = 2:5, 
                         idvar="concentration", sep="_", timevar="order",
                         direction = "long")
      longdata = longdata[order(longdata$concentration),]
      longdata = longdata[,-2]
      colnames(longdata) = c("concentration", "res")
      longdata$res = longdata$res/100
      data_plot_all = as.data.frame(rbind(data_plot_all, longdata))
    }
    
    drug = c(rep(agent2[k], 40), rep(agent1[k], 40), rep("C1", 40),
             rep("C2", 40), rep("C3", 40), 
             rep("C4", 40))
    data_plot_all = as.data.frame(cbind(drug, data_plot_all))
    if (agent1[k] == "Brigatinib"){
      Brigatinib.concentration = rep(10,m)
      for (i in 2:m){
        Brigatinib.concentration[i] = Brigatinib.concentration[i-1]*0.5
      }
      Brigatinib.concentration2 = rep(Brigatinib.concentration,each=4)
      data_plot_all$concentration[41:80] = rev(Brigatinib.concentration2)
    }
    saveRDS(data_plot_all, 
            file = paste("Plots data ", raw.data.files.name, "/", agent2[k], ".", agent1[k], ".", "Plate", k, ".", cellline[k], ".data_plot_all.rds", sep = ""))
  }
} 


Score = function(raw.data.files.name = "20191108 resazurin R",
                 info.data.files = "20191108 resazurin information.xlsx") {
  information = read.xlsx(xlsxFile = info.data.files, colNames = T) 
  agent1 = information$Agent1
  agent1 = agent1[!is.na(agent1)]
  agent2 = information$Agent2
  agent2 = agent2[!is.na(agent2)]
  plate = information$Plate.ID
  plate = plate[!is.na(plate)]
  cellline = information$Cell.line
  cellline = cellline[!is.na(cellline)]
  dose1 = information$Dose1
  dose2 = information$Dose2
  dose1 = dose1[!is.na(dose1)]
  dose2 = dose2[!is.na(dose2)]
  m = length(plate)
  ## load data (multiple)
  excess = list()
  rawexcess = list()
  ### excess inhibition
  for ( k in 1:m){
    filename1 = paste(agent2[k], agent1[k], "Plate", sep = ".")
    path = paste("Combenefit batch ", raw.data.files.name, "/",  filename1, k, "/Analysis LOEWE/data/", "Mean_Loewe_SYN_ANT_", agent1[k],"  vs.  ", agent2[k],"  in MODEL LOEWE.txt", sep = "")
    table = read.table(path, header = FALSE, sep = "", dec = ".")
    table2 = data.frame(rbind(table[4,], table[3,], table[2,], table[1,]))
    excess = append(excess, list(list(raw.data = table2)))
    rawexcess = append(rawexcess, list(list(raw.data = table)))
  }
  ### data inhibition (input data)
  data = list()
  rawdata = list()
  for ( k in 1:m){
    filename2 = paste("Chalice.", agent2[k], ".", agent1[k], ".", "Plate", k, ".", cellline[k], sep = "")
    path = paste("Chalice template ", raw.data.files.name, "/", filename2, ".xls", sep = "")
    table = read_excel(path, col_names = FALSE)
    table2 = table[1:4,3:12]
    data = append(data, list(list(raw.data = table2)))
    rawdata = append(rawdata, list(list(raw.data = table)))
  }
  
  ### calculate chalice score
  V = NULL
  S = NULL
  for ( k in 1:m){
    E = excess[[k]]$raw.data
    D = data[[k]]$raw.data
    PE = E
    PE[PE<0] = 0
    PE[PE=="NaN"] = 0
    E[E=="NaN"] = 0
    V[k] = sum(apply(E,1,sum), na.rm = T)/100 ## inhibition volume
    S[k] = sum(apply(PE*D,1,sum), na.rm = T)*log(2)*log(2)/10000 ## Chalice score
  }
  
  results = data.frame(cbind(plate, V, S))
  colnames(results) = c("Plate ID", "Inhibition Volume", "Synergy Score")
  dir.create(file.path(paste("Summary results ", raw.data.files.name, sep = "")))
  write.xlsx(results, file = paste("Summary results ", raw.data.files.name, "/", raw.data.files.name, " Summary.xlsx", sep = ""), row.names = F)
  
  ### plots
  for ( k in 1:m){
    combination = sapply(dose1, function(x) eval(bquote(expression(.(agent2[k])~" + "~.(x)~mu*M~.(agent1[k])))))
    drugname = c(agent2[k],agent1[k],combination)                     
    data_plot_all = readRDS(paste("Plots data ", raw.data.files.name, "/", agent2[k], ".", agent1[k], ".", "Plate", k, ".", cellline[k], ".data_plot_all.rds", sep = ""))
    data_plot_all$drug <- factor(data_plot_all$drug, levels=c(agent2[k], agent1[k], "C1", "C2", "C3", "C4"))
    drugList <- split(data_plot_all, data_plot_all$drug)
    maxlimy = round(max(data_plot_all$res, na.rm = T), digits = 1)
    maxlimx = round(log10(max(data_plot_all$concentration)), digits = 0)+0.5
    minlimx = round(log10(min(data_plot_all$concentration)), digits = 0)-0.5
    
    Models <- lapply(drugList, function(tmp){
      nplr(tmp$concentration, tmp$res, silent = TRUE, npars=4)})
    # Visualizing
    pdf(paste("Summary results ", raw.data.files.name, "/", "Plate", k, ".", cellline[k], ".", "Plot.pdf", sep = ""))
    color = c("red", "orange", "purple", "blue", "cyan", "green")
    par(xpd = T, mar = par()$mar + c(3,0,0,7))
    overlay(Models, xlab = expression(paste(log,"[",mu,M,"]")), ylab = "Relative viability (%)", 
            xlim = c(minlimx, maxlimx), ylim = c(0, maxlimy),
            main = cellline[k], cex.main=1.5, lwd = 3, yaxt="n", xaxt="n", 
            Cols = color, pch = 19,showLegend = F)
    legend(maxlimx, maxlimy,
           drugname, lty = 1, pch = 19, 
           col = color, cex = 0.65)
    axis(2, at = pretty(seq(0, maxlimy, by = 0.1)), xpd = TRUE, 
         lab=pretty(seq(0, maxlimy, by = 0.1)*100), las = T)
    axis(1, at = pretty(seq(minlimx, maxlimx, by = 0.5)), xpd = TRUE, 
         lab= pretty(seq(minlimx, maxlimx, by = 0.5)), las = T)
    box(which = "plot", bty = "l")
    score = round(S[k], digits = 2)
    text(maxlimx+0.5, maxlimy*0.2, paste("Score:", score))
    par(mar=c(5, 4, 4, 2) + 0.1)
    dev.off()
  }
}  
