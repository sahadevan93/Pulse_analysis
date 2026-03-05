#=====================================
#SCRIPT PROTOCOL: Automated Pulse Experiment Analysis
# ==============================================================================
# DATA STRUCTURE REQUIREMENT: 
# This script requires the raw Keysight BenchVue .xlsx export in its original 
# columnar architecture as defined in Table S1 of the manuscript. 
#
# CRITICAL COLUMN MAPPING:
# - Col D: Timestamp 
# - Col G/H: Initial 10-min Baseline (Set Current/ Get Potential)
# - Col I/J: Loop Operating Phase (Set Current/Get Potential)
# - Col K/L: Loop Pulse Phase (Set Current/Get Potential)
#
#===============================================
# ==============================================================================
# BATCH PROCESSING LOGIC:
# 1. Place all raw .xlsx BenchVue files into a single dedicated folder.
# 2. The script will prompt for this folder directory via a pop-up window.
# 3. Output: An 'Analysis_files' folder will be created automatically.
# 4. Results: Individual analysis files are generated for each raw input, 
#    plus a consolidated 'Analysis_summary.csv/xlsx' merging all datasets.
# ==============================================================================
options(java.parameters = "-Xmx25000m")
library(dplyr)
library(xlsx)
library(svDialogs)
require(tcltk)


##Important functions for analysis

convert_time <- function(time) {
  posix_time <- as.POSIXct(time, format = "%Y-%m-%d %H:%M:%S", tz = "GMT")
  
  # Total seconds since Unix epoch
  total_seconds <- as.numeric(posix_time)
  
  # Total minutes since Unix epoch
  total_minutes <- total_seconds / 60
  
  # Total hours since Unix epoch
  total_hours <- total_seconds / 3600
  
  return(c(total_hours, total_minutes, total_seconds))
}

identify_numeric_blocks <- function(vec) {
  # 1. Determine which elements are NOT NA (i.e., numeric or non-missing)
  is_numeric_block <- !is.na(vec)
  
  # 2. Use Run Length Encoding to find consecutive runs of TRUE/FALSE
  r <- rle(is_numeric_block)
  
  # 3. Initialize variables to store results
  block_starts <- numeric()
  block_ends <- numeric()
  current_pos <- 0 # Keeps track of the current position in the original vector
  
  # 4. Iterate through the runs to find the "TRUE" blocks (numeric blocks)
  for (i in seq_along(r$lengths)) {
    # Update current position to the start of the current run
    current_pos <- current_pos + 1 # Start of this run
    
    # If the current run is a block of TRUEs (non-NA values)
    if (r$values[i]) {
      start_index <- current_pos
      end_index <- current_pos + r$lengths[i] - 1
      block_starts <- c(block_starts, start_index)
      block_ends <- c(block_ends, end_index)
    }
    
    # Advance current_pos for the next iteration (move past the current block)
    current_pos <- current_pos + r$lengths[i] - 1
  }
  
  if (length(block_starts) == 0) {
    message("No numeric blocks found.")
    return(NULL)
  } else {
    return(data.frame(start = block_starts, end = block_ends))
  }
}

#### Getting raw files and initial user inputs.

print("Select the raw analysis folder using the pop-up window that just appeared outside your R studio")
setwd(choose.dir(caption = "Select the folder which has all your raw files ending with .xlsx"))

filepath<-getwd()
xlsx_files <- list.files(pattern = "\\.xlsx$", full.names = FALSE)
dir.create(paste0(filepath,"/Analysis_files"))


form_data <- list(
  "Imax:NUM" = 3.87,  # Default value 0 for the first numeric input
  "EH2:NUM" = 0.09  # Default value 0 for the second numeric input
)

result <- dlg_form(
  form = form_data,
  title = "Enter Imax and EH2 values for analysis",
  message = paste0("You have ", length(xlsx_files)," raw files in this folder. Please enter the Imax and EH2 values for analysing these files.")
)$res

Imax<-suppressWarnings(as.numeric(result$Imax))
EH2<-suppressWarnings(as.numeric(result$EH2))

analysis_summary<-matrix(nrow = 37)
pb <- txtProgressBar(min = 0, max = length(xlsx_files), initial = 0, style = 3, label = "Analysis progress")

for (f in 1:length(xlsx_files)) {
  setTxtProgressBar(pb, f)
  print( paste0("Analysing the raw file : ", xlsx_files[f]))
  filename<-xlsx_files[f]

  raw<-read.xlsx(file = filename,sheetName = "Sheet1")

  
  
  analysis<-matrix(nrow = length(raw$Time), ncol = 23, dimnames = list(NULL,c(
    "Time (sec)",	"I_No Pulse",	"E_No pulse",	"Time_base", "I_base",	"E_base",	"Time_pulse","I_pulse",
    "E_pulse","Time_Combined",	"I_Combined",	"E_combined",	"T_lowest_Base",	"Lowest E_base","T_Highest_Base","Highest E_base","Avg.E_Base",
    "T_Pulse",	"Pulse_highest point",	"E_Pulse_Average",
    "E*t_base",	"E*t pulse",	"Eaverage(base+pulse)"
    
  )))
  
  
  
  
  ###Populating the time column
  start_time<-convert_time(raw$Time[1])[3]
  analysis[1,1]<-0
  
  for (i in 2:length(raw$Time)) {
    tstamp<-convert_time(raw$Time[i])[3]
    #print(tstamp)
    analysis[i,1]<-tstamp-start_time
  }
  
  
  
  
  ##copying column G and H into analysis B and C
  
  analysis[,2]<-raw[,7]
  analysis[,3]<-raw[,8]
  
  
  ##Populating Time_base in column D in analysis
  
  raw_filtered_base<-raw %>% filter(!is.na(.[[9]]))
  start_time_base<-convert_time(raw_filtered_base$Time[1])[3]
  analysis[1,4]<-0
  for (i in 2:length(raw_filtered_base$Time)) {
    tstamp<-convert_time(raw_filtered_base$Time[i])[3]
    #print(tstamp)
    analysis[i,4]<-tstamp-start_time_base
  }
  
  ##copying column G and H to analysis E and F
  nas_to_add <- rep(NA, length(analysis[,5])-length(raw_filtered_base[,8]))
  analysis[,5]<-c(raw_filtered_base[,9],nas_to_add)
  analysis[,6]<-c(raw_filtered_base[,10],nas_to_add)
  
  ##Populating Time_pulse in column G in analysis
  
  raw_filtered_Pulse<-raw %>% filter(!is.na(.[[11]]))
  #start_time_Pulse<-convert_time(raw_filtered_Pulse$Time[1])[3]
  #analysis[1,7]<-0
  for (i in 1:length(raw_filtered_Pulse$Time)) {
    tstamp<-convert_time(raw_filtered_Pulse$Time[i])[3]
    #print(tstamp)
    analysis[i,7]<-tstamp-start_time_base
  }
  
  ##copying column G and H to analysis H and I
  nas_to_add <- rep(NA, length(analysis[,7])-length(raw_filtered_Pulse[,11]))
  analysis[,8]<-c(raw_filtered_Pulse[,11],nas_to_add)
  analysis[,9]<-c(raw_filtered_Pulse[,12],nas_to_add)
  
  
  
  # Population Combined values of Time, I and E in columns J,K,L in the analysis
  raw_filtered_Combined<-raw %>% filter(is.na(.[[7]]))
  raw_filtered_Combined<-raw_filtered_Combined[-1,]
  
  start_time_combined<-convert_time(raw_filtered_Combined$Time[1])[3]
  analysis[1,10]<-0
  
  for (i in 2:length(raw_filtered_Combined$Time)) {
    tstamp<-convert_time(raw_filtered_Combined$Time[i])[3]
    #print(tstamp)
    analysis[i,10]<-tstamp-start_time_base
  }
  
  # Correcting the assignment for column 9 from column 11
  raw_filtered_Combined[is.na(raw_filtered_Combined[, 9]), 9] <- raw_filtered_Combined[is.na(raw_filtered_Combined[, 9]), 11]
  
  # And similarly for column 10 from column 12
  raw_filtered_Combined[is.na(raw_filtered_Combined[, 10]), 10] <- raw_filtered_Combined[is.na(raw_filtered_Combined[, 10]), 12]
  
  nas_to_add <- rep(NA, length(analysis[,11])-length(raw_filtered_Combined[,9]))
  
  analysis[,11]<-c(raw_filtered_Combined[,9],nas_to_add)
  analysis[,12]<-c(raw_filtered_Combined[,10],nas_to_add)
  
  
  ## Populating M, N, O, ,P and Q in the analysis file
 
  basedata<-raw[,10]
  basedata_block<-identify_numeric_blocks(basedata)
  pb2 <- txtProgressBar(min = 0, max = length(row.names(basedata_block)), initial = 0, style = 3, label = "Basedata analysis")
  for (i in 1:length(row.names(basedata_block))){
    if((as.numeric(basedata_block[i,][2])-as.numeric(basedata_block[i,][1]))>1){
      setTxtProgressBar(pb2, i)
      block_indices<-seq((as.numeric(basedata_block[i,][1])+1),as.numeric(basedata_block[i,][2]),1 ) ###Ignored the first value of each block
      #print(block_indices)
      block_values<-c()
      time_values<-c()
      for (j in block_indices) {
        block_values<-append(block_values,basedata[j])
        time_values<-append(time_values,(convert_time(raw$Time[j])[3]-start_time_base))
      }
      block<-data.frame(Time = time_values, Values = block_values)
      #print(block)
      avg_value<-mean(block$Values)
      max_value<-which(max(block$Values)==block$Values)[1]
      min_value<-which(min(block$Values)==block$Values)[1]
      analysis[i,13]<-block[min_value,1]
      analysis[i,14]<-block[min_value,2]
      analysis[i,15]<-block[max_value,1]
      analysis[i,16]<-block[max_value,2]
      analysis[i,17]<-avg_value
    }
    
  }
  close(pb2)
  
  #filling column R, S and T
 
  pulsedata<-raw[,12]
  pulsedata_block<-identify_numeric_blocks(pulsedata)
  pb3 <- txtProgressBar(min = 0, max = length(row.names(pulsedata_block)), initial = 0, style = 3, label = "Pulsedata analysis")
  for (i in 1:length(row.names(pulsedata_block))){
    setTxtProgressBar(pb3, i)
    if(((as.numeric(pulsedata_block[i,][2]))-(as.numeric(pulsedata_block[i,][1])))>1){
      block_indices<-seq((as.numeric(pulsedata_block[i,][1])+1),as.numeric(pulsedata_block[i,][2]),1 ) ###Ignored the first value of each block
    }else{
      block_indices<-c(as.numeric(pulsedata_block[i,][1]),as.numeric(pulsedata_block[i,][2]))
    }
    
    #print(block_indices)
    block_values<-c()
    time_values<-c()
    for (j in block_indices) {
      block_values<-append(block_values,pulsedata[j])
      time_values<-append(time_values,(convert_time(raw$Time[j])[3]-start_time_base))
    }
    block<-data.frame(Time = time_values, Values = block_values)
    #print(block)
    avg_value<-mean(block$Values)
    max_value<-which(max(block$Values)==block$Values)[1]
    analysis[i,18]<-block[max_value,1]
    analysis[i,19]<-block[max_value,2]
    analysis[i,20]<-avg_value
    
  }
  close(pb3)
  
  ###### Collect parameters for calculating outputs
  
  
  I_nopulse<- unique(analysis[,2])
  I_nopulse<-I_nopulse[!is.na(I_nopulse)]
  I_pulseapplied<-unique(analysis[,8])
  I_pulseapplied<-I_pulseapplied[!is.na(I_pulseapplied)]
  basedata_block2<-identify_numeric_blocks(raw[,9])
  Pulse_interval<-round((convert_time(raw$Time[as.numeric(basedata_block2[1,][2])])[3])-(convert_time(raw$Time[as.numeric(basedata_block2[1,][1])])[3]))
  pulsedata_block2<-identify_numeric_blocks(raw[,11])
  Pulse_width<-round((convert_time(raw$Time[as.numeric(pulsedata_block2[1,][2])])[3])-(convert_time(raw$Time[as.numeric(pulsedata_block2[1,][1])])[3]))
  Total_duration<-analysis[,10]
  Total_duration<-Total_duration[!is.na(Total_duration)]
  Total_duration<-round(max(Total_duration))
  
  ncycles<-floor(Total_duration/(Pulse_interval+Pulse_width))
  T_fullcycle<-ncycles*(Pulse_interval+Pulse_width)
  Remaining_T<-Total_duration-T_fullcycle
  
  ncycle_crossvalid_base<-length((analysis[,13][!is.na(analysis[,13])]))
  ncycle_crossvalid_pulse<-length((analysis[,18][!is.na(analysis[,18])]))
  
  charge_percycle<-(I_nopulse*Pulse_interval)+(I_pulseapplied*Pulse_width)
  Total_charge<-charge_percycle*ncycles
  
  I_combined<-analysis[,11][!is.na(analysis[,11])]
  Remaining_Charge<-I_combined[length(I_combined)]*Remaining_T
  Combined_Total_charge<-Total_charge+Remaining_Charge
  
  deltaH=285000
  ne=2
  Faraday=96485
  
  ###Calculating Analysis Results
  
  #Eaverage(base+pulse)
  for (i in 1:length(analysis[,17])) {
    analysis[i,23]<-(((analysis[i,17])*Pulse_interval)+(analysis[i,20]*Pulse_width))/(Pulse_interval+Pulse_width)
    
  }
  
  E_avg.pulse<-mean(analysis[,23][!is.na(analysis[,23])])
  E_avg.nopulse<-mean(analysis[,3][!is.na(analysis[,3])])
  I_nopulse<-mean(analysis[,2][!is.na(analysis[,2])])
  I_avg.pulse<-Combined_Total_charge/Total_duration
  E_nopulse_tend<-analysis[,3][!is.na(analysis[,3])]
  E_nopulse_tend<-E_nopulse_tend[length(E_nopulse_tend)]
  Lowest_Ebase<-analysis[,14]
  Lowest_Ebase<-Lowest_Ebase[!is.na(Lowest_Ebase)]
  E_1st_pulse<-as.numeric(Lowest_Ebase[2])
  E_last_pulse<-as.numeric(Lowest_Ebase[length(Lowest_Ebase)])
  immediate_recovery_nopulse<-((E_nopulse_tend-E_1st_pulse)*100)/E_nopulse_tend
  immediate_recovery_H2<-((E_nopulse_tend-E_1st_pulse)/(E_nopulse_tend-EH2))*100
  Steady_recovery_nopulse <-((E_nopulse_tend-E_last_pulse)/E_nopulse_tend)*100
  Steady_recovery_H2<-((E_nopulse_tend-E_last_pulse)/(E_nopulse_tend-EH2))*100
  SE_nopulse<-(I_nopulse/Imax)*100
  SE_pulse<-(I_avg.pulse/Imax)*100
  EE1_NoPulse<-((1-((ne*Faraday*E_avg.nopulse)/deltaH)))*100
  EE1_Pulse<-((1-((ne*Faraday*E_avg.pulse)/deltaH)))*100
  EE2_NoPulse<-(SE_nopulse*EE1_NoPulse)/100
  EE2_Pulse<-(SE_pulse*EE1_Pulse)/100
  PC_nopulse<-I_nopulse*E_avg.nopulse
  PC_pulse<-I_avg.pulse*E_avg.nopulse
  
  
  
  ###Output files
  
  
  result <- matrix(
    nrow = 37,
    ncol = 1,
    dimnames = list(
      c( # This list defines the row names (now 38)
        "Imax (A)",
        "I_nopulse (A)", # Keeping this one with units
        "I_Pulse applied(A)",
        "Pulse interval (s)",
        "Pulse width (s)",
        "Total duration (sec)",
        "No: of cycles",
        "Time for full cycles (seconds)",
        "Remaining time (seconds)",
        "No:of cycle _cross validation_base",
        "No:of cycle _cross validation_Pulse",
        "delta H (J/mol)",
        "n",
        "F",
        "E_average, pulse",
        "E_average, no-pulse",
        "Charge per cycle (Coulombs)",
        "Total charge from full cycles (C)",
        "Charge from remaining time (C)",
        "Total charge (Coulombs)",
        "I_average, pulse",
        "E_no pulse_tend",
        "E_1st pulse",
        "E_last pulse",
        "E_H2",
        "immediate recovery_nopulse",
        "immediate recovery_H2",
        "Steady recovery_no pulse",
        "Steady recovery_H2",
        "SE_no pulse",
        "SE_pulse",
        "EE(1)_No Pulse",
        "EE(1)_Pulse",
        "EE(2) _no pulse",
        "EE(2)_Pulse",
        "PC_no pulse",
        "PC_ Pulse"
      ),
      c(filename) # This list defines the column name (just one here)
    )
  )
  result[,1]<-c(Imax,I_nopulse,I_pulseapplied,Pulse_interval,Pulse_width,Total_duration,ncycles,
                T_fullcycle,Remaining_T, ncycle_crossvalid_base,ncycle_crossvalid_pulse,deltaH,ne,Faraday,
                E_avg.pulse,E_avg.nopulse,charge_percycle,Total_charge,Remaining_Charge,Combined_Total_charge,
                I_avg.pulse,E_nopulse_tend,E_1st_pulse,E_last_pulse,EH2,immediate_recovery_nopulse,immediate_recovery_H2,
                Steady_recovery_nopulse,Steady_recovery_H2,SE_nopulse,SE_pulse,EE1_NoPulse,EE1_Pulse,EE2_NoPulse, EE2_Pulse,
                PC_nopulse,PC_pulse)
  
  Output_filename<-paste0(filepath,"/Analysis_files/","Analyzed_",filename)
  write.xlsx(analysis, file=Output_filename,append = T, showNA = F,sheetName = "Analysis",row.names = F)
  write.xlsx(result, file=Output_filename,append = T, showNA = F,sheetName = "Results",row.names = T)
  
  analysis_summary<-cbind(result, analysis_summary)
  write.csv(analysis_summary,paste0(filepath,"/Analysis_files/","Analysis_summary.csv"))
}
write.xlsx(analysis_summary, file=paste0(filepath,"/Analysis_files/","Analysis_summary.xlsx"),append = T, showNA = F,sheetName = "Results",row.names = T)
close(pb)
print(("******************************ANALYSIS COMPLETE****************************************************************"))
