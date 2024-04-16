###############################################################################
#                       INSTALANDO E CARREGANDO PACOTES                       #
###############################################################################


#Cria um vetor com o nome dos pacotes a serem instalados
pacotes <- c("tidyverse",
             "ggplot2",
             "readxl",
             "qcc",
             "dplyr")

#Instala e carrega os pacotes
if(sum(as.numeric(!pacotes %in% installed.packages())) != 0){
  instalador <- pacotes[!pacotes %in% installed.packages()]
  for(i in 1:length(instalador)) {
    install.packages(instalador, dependencies = T)
    break()}
  sapply(pacotes, require, character = T) 
} else {
  sapply(pacotes, require, character = T) 
}

#Remove o vetor pacotes
rm(pacotes)


###############################################################################
#                              IMPORTANDO DADOS                               #
###############################################################################

#Importa base de dados para o novo objeto baseTO
baseTO <- read_xlsx("03 - baseTO-v1.xlsx") 


###############################################################################
#                                CALCULOS                                     #
###############################################################################


#Adiciona variável no data frame com calculo do percentual de turnover
#Dentro do conceito do Seis Sigma este calculo é equivalente ao:
# >> DPO = Defeitos por Oportunidade
baseTO["perTurnover"] <- (baseTO["qtdDesligamentos"] / baseTO["qtdTotalQuadro"])


#Adiciona variável no data frame com calculo do z-Bench

# --> nao funciona> 
#baseTO["zBench"] <- (qnorm(exp(-baseTO["perTurnover"]),mean=0, sd=1, lower.tail = TRUE, log.p = FALSE))

    #Workaround
    baseTO["probabilidade_EXP"] <- (exp(-baseTO["perTurnover"]))
    for (tmpProbabilidade in baseTO["probabilidade_EXP"]) {}
    baseTO["zBench"] <- qnorm(tmpProbabilidade,mean=0, sd=1, lower.tail = TRUE, log.p = FALSE)
    rm(tmpProbabilidade)
    

#Adiciona variável no data frame com calculo da Capacidade Sigma do processo
baseTO["capSigma"] <- (1.5 + baseTO["zBench"])


#Exibe metricas descritivas da base geral
summary(baseTO["qtdTotalQuadro"])
summary(baseTO["qtdDesligamentos"])
summary(baseTO["perTurnover"])


#Compararemos os mesmos meses de dois anos diferentes, para mitigar sazonalidades
#Compararemos 2 meses de 2012 com processo antigo e 2 meses de 2013 com processo novo
#Os meses comparados serão fevereiro e março.

#Visualizar base comparativa - Processo antigo
subset(baseTO, baseTO["processo"] == "antigo" & baseTO["baseComp"] == "sim")


#Visualizar base comparativa - Processo novo
subset(baseTO, baseTO["processo"] == "novo" & baseTO["baseComp"] == "sim")

#Cria bases dos processo antigo e processo novo
baseTOProcessoAntigo <- subset(baseTO, baseTO["processo"] == "antigo" & baseTO["baseComp"] == "sim")
baseTOProcessoNovo <- subset(baseTO, baseTO["processo"] == "novo" & baseTO["baseComp"] == "sim")


#Exibe metricas descritivas da base processo antigo
summary(baseTOProcessoAntigo["qtdTotalQuadro"])
summary(baseTOProcessoAntigo["qtdDesligamentos"])
summary(baseTOProcessoAntigo["perTurnover"])


#Exibe metricas descritivas da base processo novo
summary(baseTOProcessoNovo["qtdTotalQuadro"])
summary(baseTOProcessoNovo["qtdDesligamentos"])
summary(baseTOProcessoNovo["perTurnover"])


#Calcula indicadores do processo antigo, para o periodo a ser comparado
perTurnoverAntigo <- (sum(baseTOProcessoAntigo["qtdDesligamentos"]) / sum(baseTOProcessoAntigo["qtdTotalQuadro"]))
zBenchAntigo <- (qnorm(exp(-perTurnoverAntigo),mean=0, sd=1, lower.tail = TRUE, log.p = FALSE))
capSigmaAntigo <- (1.5 + zBenchAntigo)

#Calcula indicadores do processo novo, para o periodo a ser comparado
perTurnoverNovo <- (sum(baseTOProcessoNovo["qtdDesligamentos"]) / sum(baseTOProcessoNovo["qtdTotalQuadro"]))
zBenchNovo<- (qnorm(exp(-perTurnoverNovo),mean=0, sd=1, lower.tail = TRUE, log.p = FALSE))
capSigmaNovo <- (1.5 + zBenchNovo)


#Realiza teste de Qui-Quadrado, para avaliar se as alterações no processo tem influêncua no turnover
#H0: x e y são independentes
#H1: x e y não são independentes
#x = Processo (Antigo / Novo) 
#y = TurnOver (Desligado / Não desligado)


#Visualiza os dados para o teste Qui-Quadrado
sum(baseTOProcessoAntigo["qtdDesligamentos"]) 
sum(baseTOProcessoAntigo["qtdTotalQuadro"]) - sum(baseTOProcessoAntigo["qtdDesligamentos"])

sum(baseTOProcessoNovo["qtdDesligamentos"]) 
sum(baseTOProcessoNovo["qtdTotalQuadro"]) - sum(baseTOProcessoNovo["qtdDesligamentos"])


#Cria tabela com dados a serem utilizados no teste de Qui-Quadrado
baseQuiQuadrado <- data.frame (desligados = c(sum(baseTOProcessoAntigo["qtdDesligamentos"]) , sum(baseTOProcessoNovo["qtdDesligamentos"]) ),
                               NaoDesligados = c(sum(baseTOProcessoAntigo["qtdTotalQuadro"]) - sum(baseTOProcessoAntigo["qtdDesligamentos"]) , sum(baseTOProcessoNovo["qtdTotalQuadro"]) - sum(baseTOProcessoNovo["qtdDesligamentos"])))

#Executa o teste de Qui-Quadrado

chisq.test(baseQuiQuadrado)


###############################################################################
#                                GRÁFICOS                                     #
###############################################################################
#Cria gráficos para a variáveis de Turnover, Quantidade total de funcionários e
#quantidade de funcionários desligados

#Adiciona uma colula com percentual de turnover formatado, para gerar uma melhor visualização. 
baseTO["labelPerTurnover"] <- (baseTO["perTurnover"] * 100)


#Gera gráfico histórico com percentual de turnover, por mês
ggplot(data = baseTO) +
  geom_line(aes(x = dtMesAno, y = labelPerTurnover)) +
  geom_point(aes(x = dtMesAno, y = labelPerTurnover, size = 3, color= "orange3")) +
  geom_text(aes(x = dtMesAno, y = labelPerTurnover, label= format(labelPerTurnover, digits= 3), vjust = 2, angle = 45, fontface="bold")) +
  ylim(1,7) + 
  labs(x= "Meses",
       y= "% Turnover",
       title="Histórico do percentual de turnover (por mês)")+
  theme_light()


#Gera gráfico histórico com a quantidade total de funcionários, por mês
ggplot(data = baseTO) +
  geom_col(aes(x = dtMesAno, y = qtdTotalQuadro), fill = "lightsteelblue3") +
  geom_text(aes(x = dtMesAno, y = qtdTotalQuadro, label= qtdTotalQuadro, vjust = -1, fontface="bold")) + 
  ylim(0,2100) + 
  labs(x= "Meses",
       y= "Quantidade total funcionários",
       title="Histórico do total de funcionários (por mês)")+
  theme_light()


#Gera gráfico histórico com a quantidade de desligamentos, por mês
ggplot(data = baseTO) +
  geom_col(aes(x = dtMesAno, y = qtdDesligamentos), fill = "skyblue4") +
  geom_text(aes(x = dtMesAno, y = qtdDesligamentos, label= qtdDesligamentos, vjust = -1, fontface="bold")) + 
  ylim(0,105) + 
  labs(x= "Meses",
       y= "Quantidade desligamentos",
       title="Histórico dos desligamentos de funcionários (por mês)") +
  theme_light()

