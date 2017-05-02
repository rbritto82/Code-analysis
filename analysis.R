library(openxlsx)
library(tidyverse)
library(lubridate)
library(stringr)
library(Hmisc)
library(ggrepel)
library(scales)

#reading file
pcData <- cbind(
                as_tibble(read.xlsx("./TEDD Ericsson Karlskrona learning data_v4.3.xlsx", sheet=5, cols = c(1, 2, 4, 6))),
                as_tibble(read.xlsx("./TEDD Ericsson Karlskrona learning data_v4.3.xlsx", sheet=9, cols = c(3, 4),  detectDates = TRUE)),
                as_tibble(read.xlsx("./TEDD Ericsson Karlskrona learning data_v4.3.xlsx", sheet=6, cols = c(6))),
                as_tibble(read.xlsx("./TEDD Ericsson Karlskrona learning data_v4.3.xlsx", sheet=10, cols = c(3:14)))
                )

#Removing observations that should not be in the analysis and renaming the maturity values
pcData <- pcData[-c(9, 18, 28, 29, 32, 33, 36, 41, 42, 44, 45, 47:55),]
pcData$refinedTeamMaturity[pcData$refinedTeamMaturity == "Immature teams"] <- "Immature"
pcData$refinedTeamMaturity[pcData$refinedTeamMaturity == "Mature teams"] <- "Mature"

pcData <- pcData[pcData$location2 != "Distributed", ]

#Boxplots-----------------------------------------------------------------------------------------
boxplot <- ggplot(pcData, aes(x = refinedTeamMaturity, y = nLocRatio)) + 
  geom_boxplot() + 
  scale_y_continuous(name = "Normalized Ratio Submitted/pushed LOC",
                     breaks = seq(min(pcData$nLocRatio), max(pcData$nLocRatio), 0.01), 
                     limits=c(min(pcData$nLocRatio), max(pcData$nLocRatio))) + 
  scale_x_discrete(name = "Maturity") +
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + 
  theme(axis.text.x = element_text(size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15))
ggsave("./Normalized_LOC_Ratio.pdf", boxplot)

boxplot <- ggplot(pcData, aes(x = refinedTeamMaturity, y = nLocDuplicated)) + 
  geom_boxplot() + 
  scale_y_continuous(name = "Normalized Duplicated LOC",
                     breaks = seq(min(pcData$nLocDuplicated), max(pcData$nLocDuplicated), 20), 
                     limits=c(min(pcData$nLocDuplicated), max(pcData$nLocDuplicated))) + 
  scale_x_discrete(name = "Maturity") +
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + 
  theme(axis.text.x = element_text(size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15))
ggsave("./Normalized_Duplicated_LOC.pdf", boxplot)

boxplot <- ggplot(pcData, aes(x = refinedTeamMaturity, y = nCyclomatic)) + 
  geom_boxplot() + 
  scale_y_continuous(name = "Normalized Cyclomatic Complexity",
                     breaks = seq(min(pcData$nCyclomatic), max(pcData$nCyclomatic), 1.5), 
                     limits=c(min(pcData$nCyclomatic), max(pcData$nCyclomatic))) + 
  scale_x_discrete(name = "Maturity") +
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + 
  theme(axis.text.x = element_text(size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15))
ggsave("./Normalized_Cyclomatic_Complexity.pdf", boxplot)

boxplot <- ggplot(pcData, aes(x = refinedTeamMaturity, y = nTechnicalDebt)) + 
  geom_boxplot() + 
  scale_y_continuous(name = "Normalized Technical Debt",
                     breaks = seq(min(pcData$nTechnicalDebt), max(pcData$nTechnicalDebt), 10), 
                     limits=c(min(pcData$nTechnicalDebt), max(pcData$nTechnicalDebt))) + 
  scale_x_discrete(name = "Maturity") +
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + 
  theme(axis.text.x = element_text(size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15))
ggsave("./Normalized_Technical_Debt.pdf", boxplot)

#Wilcoxon------------------------------------------------------------------------------------------
#Creating workbook
wb <- createWorkbook("Code analysis.xlsx")
addWorksheet(wb, "Summary Norm. Ratio")
addWorksheet(wb, "P-value Norm. Ratio")
addWorksheet(wb, "Summary Norm. Duplicated")
addWorksheet(wb, "P-value Norm. Duplicated")
addWorksheet(wb, "Summary Norm. Cyclomatic")
addWorksheet(wb, "P-value Norm. Cyclomatic")
addWorksheet(wb,  "Summary Norm. Debt")
addWorksheet(wb,  "P-value Norm. Debt")

#Summary
analysisData <- group_by(pcData, refinedTeamMaturity)
analysisDataSummary <- summarise(analysisData, observations = length(refinedTeamMaturity), 
                                            mean = mean(nLocRatio,  na.rm = TRUE),
                                            median = median(nLocRatio,  na.rm = TRUE),
                                            min = min(nLocRatio,  na.rm = TRUE),
                                            max = max(nLocRatio,  na.rm = TRUE),
                                            standardDeviation = sd(nLocRatio,  na.rm = TRUE))
#Hypotehesis test
wilcox <- wilcox.test(nLocRatio ~ refinedTeamMaturity, data = pcData, alternative = "two.sided")
#Save results
writeData(wb, "Summary Norm. Ratio", analysisDataSummary, withFilter = TRUE)
writeData(wb, "P-value Norm. Ratio", wilcox$p.value, withFilter = TRUE)

#Summary
analysisDataSummary <- summarise(analysisData, observations = length(refinedTeamMaturity), 
                                 mean = mean(nLocDuplicated,  na.rm = TRUE),
                                 median = median(nLocDuplicated,  na.rm = TRUE),
                                 min = min(nLocDuplicated,  na.rm = TRUE),
                                 max = max(nLocDuplicated,  na.rm = TRUE),
                                 standardDeviation = sd(nLocDuplicated,  na.rm = TRUE))
#Hypotehesis test
wilcox <- wilcox.test(nLocDuplicated ~ refinedTeamMaturity, data = pcData, alternative = "two.sided")
#Save results
writeData(wb, "Summary Norm. Duplicated", analysisDataSummary, withFilter = TRUE)
writeData(wb, "P-value Norm. Duplicated", wilcox$p.value, withFilter = TRUE)

#Summary
analysisDataSummary <- summarise(analysisData, observations = length(refinedTeamMaturity), 
                                 mean = mean(nCyclomatic,  na.rm = TRUE),
                                 median = median(nCyclomatic,  na.rm = TRUE),
                                 min = min(nCyclomatic,  na.rm = TRUE),
                                 max = max(nCyclomatic,  na.rm = TRUE),
                                 standardDeviation = sd(nCyclomatic,  na.rm = TRUE))
#Hypotehesis test
wilcox <- wilcox.test(nCyclomatic ~ refinedTeamMaturity, data = pcData, alternative = "two.sided")
#Save results
writeData(wb, "Summary Norm. Cyclomatic", analysisDataSummary, withFilter = TRUE)
writeData(wb, "P-value Norm. Cyclomatic", wilcox$p.value, withFilter = TRUE)

#Summary
analysisDataSummary <- summarise(analysisData, observations = length(refinedTeamMaturity), 
                                 mean = mean(nTechnicalDebt,  na.rm = TRUE),
                                 median = median(nTechnicalDebt,  na.rm = TRUE),
                                 min = min(nTechnicalDebt,  na.rm = TRUE),
                                 max = max(nTechnicalDebt,  na.rm = TRUE),
                                 standardDeviation = sd(nTechnicalDebt,  na.rm = TRUE))
#Hypotehesis test
wilcox <- wilcox.test(nTechnicalDebt ~ refinedTeamMaturity, data = pcData, alternative = "two.sided")
#Save results
writeData(wb, "Summary Norm. Debt", analysisDataSummary, withFilter = TRUE)
writeData(wb, "P-value Norm. Debt", wilcox$p.value, withFilter = TRUE)

#Save data
saveWorkbook(wb, file = "./Code analysis.xlsx", overwrite = TRUE)




#Evolution plots----------------------------------------------------------------------------------------
ggplot(pcData, aes(x = end, y = nLocRatio, label = ID)) + 
  geom_point() + 
  stat_smooth(method = "loess")+
  labs(x="End date", y="Normalized LOC Ratio") + #change the labels
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + #set legend theme
  theme(axis.text.x = element_text(angle=90, size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15)) +
  scale_x_date(labels = date_format("%y/%m"), date_breaks = "1 month", limits = c(min(pcData$end), max(pcData$end))) +
  geom_label_repel(aes(fill = location2, size = 5)) +
  scale_fill_manual(values = c("orange", "deepskyblue", "yellow", "red"))
ggsave("./LOWESS_nRatio_LOC.pdf")

ggplot(pcData, aes(x = end, y = nLocDuplicated, label = ID)) + 
  geom_point() + 
  stat_smooth(method = "loess")+
  labs(x="End date", y="Normalized Duplicated LOC") + #change the labels
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + #set legend theme
  theme(axis.text.x = element_text(angle=90, size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15)) +
  scale_x_date(labels = date_format("%y/%m"), date_breaks = "1 month", limits = c(min(pcData$end), max(pcData$end))) +
  geom_label_repel(aes(fill = location2, size = 5)) +
  scale_fill_manual(values = c("orange", "deepskyblue", "yellow", "red"))
ggsave("./LOWESS_nDuplicated_LOC.pdf")

ggplot(pcData, aes(x = end, y = nCyclomatic, label = ID)) + 
  geom_point() + 
  stat_smooth(method = "loess")+
  labs(x="End date", y="Normalized Cyclomatic Complexity") + #change the labels
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + #set legend theme
  theme(axis.text.x = element_text(angle=90, size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15)) +
  scale_x_date(labels = date_format("%y/%m"), date_breaks = "1 month", limits = c(min(pcData$end), max(pcData$end))) +
  geom_label_repel(aes(fill = location2, size = 5)) +
  scale_fill_manual(values = c("orange", "deepskyblue", "yellow", "red"))
ggsave("./LOWESS_nCyclomatic.pdf")

ggplot(pcData, aes(x = end, y = nTechnicalDebt, label = ID)) + 
  geom_point() + 
  stat_smooth(method = "loess")+
  labs(x="End date", y="Normalized Technical Debt") + #change the labels
  theme(legend.position="bottom", legend.direction="horizontal",legend.title = element_blank()) + #set legend theme
  theme(axis.text.x = element_text(angle=90, size = 15),
        axis.text.y = element_text(size = 15),
        axis.title.x = element_text(size = 15),
        axis.title.y = element_text(size = 15)) +
  scale_x_date(labels = date_format("%y/%m"), date_breaks = "1 month", limits = c(min(pcData$end), max(pcData$end))) +
  geom_label_repel(aes(fill = location2, size = 5)) +
  scale_fill_manual(values = c("orange", "deepskyblue", "yellow", "red"))
ggsave("./LOWESS_nDebt.pdf")

#Correlations--------------------------------------------------------------------------------------------

correlations <- rcorr(as.matrix(pcData[,-c(1:6)]), type = "spearman")

addWorksheet(wb,  "Correlation values")
writeData(wb, "Correlation values", correlations$r, withFilter = TRUE)
addWorksheet(wb,  "Correlation P-values")
writeData(wb, "Correlation P-values", correlations$P, withFilter = TRUE)
saveWorkbook(wb, file = "./Code analysis.xlsx", overwrite = TRUE)
