require(openxlsx);
require(xlsx);
require(reshape2);
require(chron);
require(Cairo);
setwd("/Users/richard/Google Drive/dupont/test");
#setwd("L:/Richard/test");
plant_codes <- c("FCEFF","FCNT2","FCPT2","FCPT6","FCPT7","FCSOL","FCUTI","FCUTW","FCNIT","FCPT5","FCMIS","FCNT1","FCHYD");
#plant_codes <- c("FCEFF","FCNT2","FCPT2","FCPT5","FCPT6","FCPT7","FCSOL","FCUTI","FCUTW","FCNIT");
two_weeks_ago <- as.Date(Sys.time()) - 15;
results_raw_raw <- read.csv(file="last_two_weeks.csv",header=TRUE,sep=",");
results_raw <- results_raw_raw[!(results_raw_raw$Date.result.authorised == ""),];two_weeks_ago_tag = as.Date(as.POSIXlt(results_raw$Login.date,format='%d-%b-%Y %H:%M:%S')) > two_weeks_ago
results <- na.omit(results_raw[two_weeks_ago_tag,]);
no_data <- read.table(file="L:/Richard/no_data.csv",header=TRUE,sep=";");
levels(results$Sample.Template.Id) -> templates;
results_wb <- createWorkbook();
#setwd("L:/Richard/test");
sample_plot <- function(data1,data2,ytitle,mtitle) {
 plot(data1,data2,main=mtitle,xlab="Time Sample Was Entered",ylab=ytitle,col="steelblue",type="n");
 lines(data1,data2,pch=2); }

sample_plot_upper_lim <- function(data1,data2,ytitle,mtitle,upper_lim) {
 plot(data1,data2,main=mtitle,xlab="Time Sample Was Entered",ylab=ytitle,col="steelblue",type="n");
 lines(data1,data2,pch=2);
 abline(h=upper_lim,col=2) }

sample_plot_upper_and_lower_lim <- function(data1,data2,ytitle,mtitle,upper_lim,lower_lim) {
 plot(data1,data2,main=mtitle,xlab="Time Sample Was Entered",ylab=ytitle,col="steelblue",type="n");
 lines(data1,data2,pch=2); 
 abline(h=upper_lim,col=2)
 abline(h=lower_lim,col=2)	
}

#Wash Water Plot
WW = results_raw$Description == "#2 Nitration Unit MNB Primary Wash Water"
wash_water <- results_raw[WW,];
#wash_water <- results_raw[results_raw$Description == "#2 Nitration Unit MNB Primary Wash Water"];
wash_water_times <- wash_water$Date.result.authorised;
wash_water_results <- wash_water$Result.value;
if (length(wash_water_results) == 0) {print("No results to graph.");} else {
Cairo(file="Wash_water_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(wash_water_times,format='%d-%b-%Y %H:%M:%S'),wash_water_results,"NAOH in Wash Water","Wash Water % Caustic");
dev.off(); }

bz_sample_plot <- function(data1,data2,ytitle) {
 plot(data1,data2,main="Wash Water % Caustic",xlab="Time Sample Was Entered",ylab=ytitle,col="steelblue",xaxt="n");
 axis(side = 1, at = data1, labels = TRUE);
 lines(data1,data2,pch=2); }
#Benzene Stripper - Benzene Plot
BZ = results$Description == "#2 Nitration Unit MNB Benzene Stripper Bottom"
Bz_stripper <- results[BZ,];
BZB = Bz_stripper$Component == "Benzene"
Bz_stripper_benzene <- Bz_stripper[BZB,];
Bz_stripper_benzene_times <- Bz_stripper_benzene$Date.result.authorised;
Bz_stripper_benzene_results <- Bz_stripper_benzene$Result.value;
if (length(Bz_stripper_benzene_results) == 0) {print("no results");} else {
Cairo(file="Bz_stripper_benzene_results.png",bg="white",type="png",dpi=72);
sample_plot_upper_lim(as.POSIXlt(Bz_stripper_benzene_times,format='%d-%b-%Y %H:%M:%S'),Bz_stripper_benzene_results,"Benzene (ppm) in Benzene Stripper","Benzene slippage in Bz Stripper",10);
dev.off(); }

CC_3_ind = results$Description == "Effluent System #3 Carbon Column Exit"
CC_3 <- results[CC_3_ind,];
CC_3_an = CC_3$Component == "Aniline"
CC_3_aniline <- CC_3[CC_3_an,];
CC_3_aniline_times <- CC_3_aniline$Date.result.authorised;
CC_3_aniline_results <- CC_3_aniline$Result.value;
if (length(CC_3_aniline_results) == 0) { print("No results to graph"); } else {
Cairo(file="CC_3_aniline_results.png",bg="white",type="png",dpi=72);
sample_plot_upper_lim(as.POSIXlt(CC_3_aniline_times,format='%d-%b-%Y %H:%M:%S'),CC_3_aniline_results,"Aniline (ppm) in Carbon Column #3","Aniline in Carbon Column #3",150);
dev.off();
}

#747 - Benzene Plot
b747 = results$Description == "#2 Nitration Unit MNB C747 Organics"
b747_results <- results[b747,];
b747_c = b747_results$Component == "Benzene"
b747_benzene <- b747_results[b747_c,];
b747_benzene_times <- b747_benzene$Date.result.authorised;
b747_benzene_results <- b747_benzene$Result.value;
if (length(b747_benzene_results) == 0) { print("No results to graph"); } else {
Cairo(file="b747_benzene_results.png",bg="white",type="png",dpi=72);
sample_plot_upper_lim(as.POSIXlt(b747_benzene_times,format='%d-%b-%Y %H:%M:%S'),b747_benzene_results,"Benzene (ppm) in 747","Benzene slippage in 747",9);
dev.off();
}

#ANOH - MNB Plot
ANOH = results$Description == "#2 Plant Aniline Column Overhead"
ANOH_results <- results[ANOH,];
ANOH_MNB = ANOH_results$Component == "Nitrobenzene"
ANOH_MNB_results <- ANOH_results[ANOH_MNB,];
ANOH_MNB_times <- ANOH_MNB_results$Date.result.authorised;
ANOH_MNB_values <- ANOH_MNB_results$Result.value;
if (length(ANOH_MNB_values) == 0) { print("No results to graph"); } else {
Cairo(file="ANOH_MNB_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(ANOH_MNB_times,format='%d-%b-%Y %H:%M:%S'),ANOH_MNB_values,"MNB (ppm) in #2 Aniline Overhead","MNB levels in #2 Aniline Overhead");
dev.off();
}

#ANOH - phenol Plot
ANOH = results$Description == "#2 Plant Aniline Column Overhead"
ANOH_results <- results[ANOH,];
ANOH_phenol = ANOH_results$Component == "Phenol"
ANOH_phenol_results <- ANOH_results[ANOH_phenol,];
ANOH_phenol_times <- ANOH_phenol_results$Date.result.authorised;
ANOH_phenol_values <- ANOH_phenol_results$Result.value;
if (length(ANOH_phenol_values) == 0) { print("No results to graph"); } else {
Cairo(file="ANOH_phenol_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(ANOH_phenol_times,format='%d-%b-%Y %H:%M:%S'),ANOH_phenol_values,"Phenol (ppm) in #2 Aniline Overhead","Phenol levels in #2 Aniline Overhead");
dev.off();
}

#R1900 - DPA Plot
R1900 = results$Description == "NDPA Reaction Crude from R-1900"
R1900_results <- results[R1900,];
R1900_DPA = R1900_results$Component == "Diphenylamine"
R1900_DPA_results <- R1900_results[R1900_DPA,];
R1900_DPA_times <- R1900_DPA_results$Date.result.authorised;
R1900_DPA_values <- R1900_DPA_results$Result.value;
if (length(R1900_DPA_values) == 0) { print("No results to graph"); } else {
Cairo(file="R1900_DPA_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(R1900_DPA_times,format='%d-%b-%Y %H:%M:%S'),R1900_DPA_values,"Diphenylamine in R1900","DPA levels in R1900");
dev.off();
}

#R1900 - Tri-Alk's Plot
R1900 = results$Description == "NDPA Reaction Crude from R-1900"
R1900_results <- results[R1900,];
R1900_trialks = R1900_results$Component == "Tri-alkyl DPA"
R1900_trialks_results <- R1900_results[R1900_trialks,];
R1900_trialks_times <- R1900_trialks_results$Date.result.authorised;
R1900_trialks_values <- R1900_trialks_results$Result.value;
if (length(R1900_trialks_values) == 0) { print("No results to graph"); } else {
Cairo(file="R1900_trialks_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(R1900_trialks_times,format='%d-%b-%Y %H:%M:%S'),R1900_trialks_values,"Tri-alkyl DPA in R1900","Tri-alkyl DPA levels in R1900");
dev.off();
}

#R1910 - Nonene Plot
R1910 = results$Description == "NDPA Nonene Strip - Strip Product from R-1910"
R1910_results <- results[R1910,];
R1910_nonene = R1910_results$Component == "Nonene"
R1910_nonene_results <- R1910_results[R1910_nonene,];
R1910_nonene_times <- R1910_nonene_results$Date.result.authorised;
R1910_nonene_values <- R1910_nonene_results$Result.value;
if (length(R1910_nonene_values) == 0) { print("No results to graph"); } else {
Cairo(file="R1910_nonene_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(R1910_nonene_times,format='%d-%b-%Y %H:%M:%S'),R1910_nonene_values,"Nonene in R1910","Nonene levels in R1910");
dev.off();
}

#R1910 - Tri-alks Plot
R1910 = results$Description == "NDPA Nonene Strip - Strip Product from R-1910"
R1910_results <- results[R1910,];
R1910_trialks = R1910_results$Component == "Tri-alkyl DPA"
R1910_trialks_results <- R1910_results[R1910_trialks,];
R1910_trialks_times <- R1910_trialks_results$Date.result.authorised;
R1910_trialks_values <- R1910_trialks_results$Result.value;
if (length(R1910_trialks_values) == 0) { print("No results to graph"); } else {
Cairo(file="R1910_trialks_results.png",bg="white",type="png",dpi=72);
sample_plot(as.POSIXlt(R1910_trialks_times,format='%d-%b-%Y %H:%M:%S'),R1910_trialks_values,"Tri-alkyl DPA in R1910","Tri-alkyl DPA levels in R1910");
dev.off();
}

conditional_format <- function(analyte,spec_limit,digits) {
  for (l in 1:length(analyte)) {   analyte[l] <- round(analyte[l],digits) } 
  for (l in 1:length(analyte)) {   if(analyte[l] == "NaN") { analyte[l] <- -1; } }
  sprintf("<%d",spec_limit) -> field_value;
  for (l in 1:length(analyte)) { if(as.numeric(analyte[l]) < spec_limit && as.numeric(analyte[l]) >= 0) { analyte[l] <- field_value} } 
  for (l in 1:length(analyte)) { if(analyte[l] == -1) { analyte[l] <- " "; } } 
  return(analyte);
}

for (i in 1:13) {
 plant <- grepl(plant_codes[i],results$Template.id);
 plant_results <- results[plant,];
 plant_code <- plant_codes[i];
 levels(factor(plant_results$Description)) -> descriptions;
 length(descriptions) -> num_descriptions;
 if (num_descriptions == 0) { descriptions <- c("None","None"); num_descriptions <- 1; }
 else {
 print(descriptions);
 print(num_descriptions);
 for (j in 1:num_descriptions) { 
#  if (descriptions
  L = descriptions[j] == plant_results$Description;
  plant_results[L,] -> results_template;
  print(results_template);
  description_name <- descriptions[j];
  length(results_template$Component.name) -> k; 
#  sample_date <- as.POSIXlt(results_template$Date.result.authorised[i],format='%d-%b-%Y %H:%M:%S');
  if (length(plant_results$Description) == 0) { print("no data"); results_template_agg <- no_data; }
  else {  results_template_agg <- dcast(results_template,Description + Login.date ~ Component.name, value.var = "Result.value", fun.aggregate = mean, na.rm=TRUE); }
  if (grepl("Unscheduled",descriptions[j]) == TRUE) { results_template_agg <- dcast(results_template,Description + Login.date + Sample.name ~ Component.name, value.var = "Result.value"); }
  if (grepl("4-ADP",descriptions[j]) == TRUE) { results_template_agg <- dcast(results_template,Description + Login.date + Sample.name ~ Component.name, value.var = "Result.value"); }
  print(results_template_agg);
  dim(results_template_agg)[2] -> ncols; 
  if (description_name == "#2 Plant Aniline Column Overhead Decanter") { description_name <- "Overhead Decanter"; print("yep"); }
  if (description_name == "Calibration & Instrument Check of LC3 Using 2 & 4 NP") { description_name <- "LC3 2 & 4 NP Standard"; print("yep2"); }
  if (description_name == "Solvent Extraction Extracted Water Bottom") { description_name <- "Extracted Water Stripper Bottoms"; print("yep3"); }
  if (description_name == "#6 Plant FORAFAC 1157N Step-1 PSCN Reactor Feed") { description_name <- "1157N Step1 PSCN Rx Feed" }
  if (description_name == "#6 Plant FORAFAC 1157N Step-4 Water Addition") { description_name <- "1157N Step4 Water Addition" }
  if (description_name == "#6 Plant Treated Water Tank (TK1448)") { description_name <- "#6 Plant (1448) Treater Water" }
  if (description_name == "Instrument Check and Calibration for GC4B using FC113_Std") { description_name <- "GC4B Standard" }
  if (description_name == "Instrument MS #2 QC Check using ppb PFOA") { description_name <- "MS #2 ppb PFOA Standard" }
  if (description_name == "#6 Plant FORAFAC 1157N Step-1 PSCN Crude Product Wash") { description_name <- "1157N Step1 PSCN Crude Wash" }
  if (description_name == "#6 Plant FORAFAC 1157N Step-1 PSCN End of Reaction") { description_name <- "1157N Step1 PSCN End of Rx" }
  if (description_name == "Instrument MS #1 QC Check using ppb PFOA") { description_name <- "MS #1 ppb PFOA Standard" }
  if (description_name == "#6 Plant FORAFAC 1157N Step-2 PSCL Crude Product Wash") { description_name <- "1157N Step2 Crude Prod. Wash" }
  if (description_name == "#6 Plant FORAFAC 1157N Step-3 F1145N Crude Product Wash") { description_name <- "1157N Step3 F1145N Crude Prod. Wash" }
  if (description_name == "Instrument MS #1 QC Check using ppm C6-I, C8-I and 8-2-8 ester") { description_name <- "MS #1 ppm C6-I, C8-I & 8-2-8 ester" }
  if (description_name == "Instrument MS #2 QC Check using ppm C8-I and 8-2-8 ester") { description_name <- "MS #2 ppm C8-I & 8-2-8 ester" }
  if (description_name == "Instrument MS #1 QC Check using ppm PFOA") { description_name <- "MS #1 ppm PFOA Standard" }
  if (description_name == "Instrument MS #2 QC Check using ppm PFOA") { description_name <- "MS #2 ppm PFOA Standard" }
  if (description_name == "Plant #7 NDPA Reclaimed Nonene from TK-1954") { description_name <- "Reclaimed Nonene from TK-1954" }
  if (description_name == "#1 Nitration Unit MNT Incoming Toluene Truck") { description_name <- "Incoming Toluene Truck" }
  if (description_name == "#6 Plant C62-I Step-1 PSCN Crude Product Wash") { description_name <- "C62-I Step1 Crude Prod. Wash" }
  if (description_name == "#6 Plant C62-I Step-2 PSCN Crude Product Wash") { description_name <- "C62-I Step2 Crude Prod. Wash" }
  if (description_name == "#6 Plant C62-I Step-2 PSCL Crude Product Wash") { description_name <- "C62-I Step2 PSCL Crude Wash" }
  if (description_name == "#2 Plant Aniline Recycle Gas - Low Concentration Hydrogen") { description_name <- "Low Conc. Hydrogen" }
  if (description_name == "Effluent Discharge Water - West Tank Farm") { description_name <- "West Tank Farm Discharge" }

#  if (description_name == "Nitric Plant 23rd Tray") { print(results_template_agg$Result.value); for(l in 1:length(results_template_agg$Chloride)) { if(results_template_agg$Chloride[l] < 25) { results_template_agg$Chloride[l] <- "< 25.0 ppm" } } }
  if (description_name == "Nitric Plant 23rd Tray") { results_template_agg$Chloride <- conditional_format(results_template_agg$Chloride,25,1) }
  if (description_name == "NDPA Filtration Dried Cake from R-1930") { print(results_template_agg$Result.value); for(l in 1:length(results_template_agg$"Flash Point")) { if(results_template_agg$"Flash Point"[l] >= 200) { results_template_agg$"Flash Point"[l] <- ">200" } } }
  if (description_name == "#6 Plant FORAFAC C1470 Step-4 After Product Filtration") { description_name <- "C1470 S4 After Filtration" }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$"2 Nitrophenol" <- conditional_format(results_template_agg$"2 Nitrophenol",30,0) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$"4 Nitrophenol" <- conditional_format(results_template_agg$"4 Nitrophenol",30,0) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$Aniline <- conditional_format(results_template_agg$Aniline,1,1) }
#  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$Benzene <- conditional_format(results_template_agg$Benzene,10,0) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$TOC <- conditional_format(results_template_agg$TOC,1,1) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$Phenol <- conditional_format(results_template_agg$Phenol,1,1) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$"Ortho Toluidine" <- conditional_format(results_template_agg$"Ortho Toluidine",1,1) }
  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$"p-Toluidine" <- conditional_format(results_template_agg$"p-Toluidine",1,1) }
#  if (description_name == "Effluent System #1 Carbon Column Exit") { results_template_agg$Toluene <- conditional_format(results_template_agg$Toluene,10,0) }

  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$"4 Nitrophenol" <- conditional_format(results_template_agg$"4 Nitrophenol",30,0)} 
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$"2 Nitrophenol" <- conditional_format(results_template_agg$"2 Nitrophenol",30,0)} 
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$Aniline <- conditional_format(results_template_agg$Aniline,1,1) }
#  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$Benzene <- conditional_format(results_template_agg$Benzene,10,0) }
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$TOC <- conditional_format(results_template_agg$TOC,1,1) }
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$Phenol <- conditional_format(results_template_agg$Phenol,1,1) }
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$"Ortho Toluidine" <- conditional_format(results_template_agg$"Ortho Toluidine",1,1) }
  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$"p-Toluidine" <- conditional_format(results_template_agg$"p-Toluidine",1,1) }
#  if (description_name == "Effluent System #2 Carbon Column Exit") { results_template_agg$Toluene <- conditional_format(results_template_agg$Toluene,10,0) }

  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$"4 Nitrophenol" <- conditional_format(results_template_agg$"4 Nitrophenol",30,0)} 
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$"2 Nitrophenol" <- conditional_format(results_template_agg$"2 Nitrophenol",30,0)} 
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$Aniline <- conditional_format(results_template_agg$Aniline,1,1) }
#  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$Benzene <- conditional_format(results_template_agg$Benzene,10,0) }
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$TOC <- conditional_format(results_template_agg$TOC,1,1) }
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$Phenol <- conditional_format(results_template_agg$Phenol,1,1) }
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$"Ortho Toluidine" <- conditional_format(results_template_agg$"Ortho Toluidine",1,1) }
  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$"p-Toluidine" <- conditional_format(results_template_agg$"p-Toluidine",1,1) }
#  if (description_name == "Effluent System #3 Carbon Column Exit") { results_template_agg$Toluene <- conditional_format(results_template_agg$Toluene,10,0) }

  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$"4 Nitrophenol" <- conditional_format(results_template_agg$"4 Nitrophenol",30,0)} 
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$"2 Nitrophenol" <- conditional_format(results_template_agg$"2 Nitrophenol",30,0)} 
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$Aniline <- conditional_format(results_template_agg$Aniline,1,1) }
#  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$Benzene <- conditional_format(results_template_agg$Benzene,10,0) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$TOC <- conditional_format(results_template_agg$TOC,1,1) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$Phenol <- conditional_format(results_template_agg$Phenol,1,1) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$"Ortho Toluidine" <- conditional_format(results_template_agg$"Ortho Toluidine",1,1) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$"p-Toluidine" <- conditional_format(results_template_agg$"p-Toluidine",1,1) }
#  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$Toluene <- conditional_format(results_template_agg$Toluene,10,0) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$pH <- conditional_format(results_template_agg$pH,0,1) }
  if (description_name == "Effluent System pH Adjustment Tank") { results_template_agg$"Hydrogen peroxide" <- conditional_format(results_template_agg$"Hydrogen peroxide",99,0) }

  print(description_name);

  results_template_sheet <- createSheet(wb=results_wb,sheetName = description_name);
  print(results_template_agg[order(as.POSIXct(results_template_agg$Login.date,format='%d-%b-%Y %H:%M:%S')),]);
  if (grepl("Unscheduled",descriptions[j]) == TRUE || grepl("4-ADP",descriptions[j]) == TRUE) { createFreezePane(results_template_sheet,2,4); }
#  if (grepl("4-ADP",descriptions[j]) == TRUE) { createFreezePane(results_template_sheet,2,4); }
  else {  createFreezePane(results_template_sheet,2,3); }
  addDataFrame(x=results_template_agg[order(as.POSIXct(results_template_agg$Login.date,format='%d-%b-%Y %H:%M:%S')),],sheet=results_template_sheet,row.names=FALSE);
   for (k in 1:ncols) {
    autoSizeColumn(results_template_sheet,k) }
  if (description_name == "#2 Nitration Unit MNB Primary Wash Water") { addPicture("Wash_water_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "#2 Nitration Unit MNB Benzene Stripper Bottom") { addPicture("Bz_stripper_benzene_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "#2 Nitration Unit MNB C747 Organics") { addPicture("b747_benzene_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "#2 Plant Aniline Column Overhead") { addPicture("ANOH_MNB_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "#2 Plant Aniline Column Overhead") { addPicture("ANOH_phenol_results.png",results_template_sheet,startRow=27,startColumn=ncols+1); }
  if (description_name == "NDPA Reaction Crude from R-1900") { addPicture("R1900_DPA_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "NDPA Reaction Crude from R-1900") { addPicture("R1900_trialks_results.png",results_template_sheet,startRow=27,startColumn=ncols+1); }
  if (description_name == "NDPA Nonene Strip - Strip Product from R-1910") { addPicture("R1910_nonene_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
  if (description_name == "NDPA Nonene Strip - Strip Product from R-1910") { addPicture("R1910_trialks_results.png",results_template_sheet,startRow=27,startColumn=ncols+1); }
  if (description_name == "Effluent System #3 Carbon Column Exit") { addPicture("CC_3_aniline_results.png",results_template_sheet,startRow=2,startColumn=ncols+1); }
 }
}
}
setwd("S:/Plant Data/Results over last two weeks");
#setwd("L:/Richard/test");
saveWorkbook(results_wb,"last_two_weeks_aggregated_tmp.xlsx");