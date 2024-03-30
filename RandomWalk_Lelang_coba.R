###############################
# Random Walk Seksi AH Dit. PS
###############################
library(readxl) # karena csv valuesnya berubah - kita pake excel nya - package ini lebih mudah
library(openxlsx) # tuk print dalam bentu xlsx 
library(tidyr) # untuk pivot dan filter
library(dplyr) # tuk sort data

# cleaning
rm(list=ls())

outdet<-"C:/Users/fatchul.lailin/OneDrive - Kemenkeu/LELANG & RENCANA LELANG/2024/12 Lelang 19 Maret 2024/IBPA/Randomwalk R"
tempat_IBPA<-"C:/Users/fatchul.lailin/OneDrive - Kemenkeu/LELANG & RENCANA LELANG/2024/12 Lelang 19 Maret 2024"

date <- "18032024"  # replace it with your actual date, cukup update tanggal di sini saja

setwd(outdet)

#input data
MainData <- read.xlsx("20240205_20240318_Series_Historical_GB.xlsx")

#input IBPA tuk dpt nama seri
# input data IBPA t-1
setwd(tempat_IBPA) # tempat ngambil file IBPA, jadi satu saja dengan PLTE
ibpa_file <- paste("IBPA_t-1_", date, ".xlsx", sep = "") # Construct file names with the date
IBPA <- read_excel(ibpa_file)# Read Excel files

# Subset MainData for the current series
seri <- "PBS032"

# preparation
series_name <-colnames(IBPA) #bikin daftar seri tuk looping per seri  
yield_rw <- data.frame(SERI = character(), yield_rw = numeric()) # Initialize an empty data frame to store the WAY values for each series

# Loop over each series
for (seri in series_name) {

new_data <- subset(MainData, SERIES == seri)


# Select specific columns
column_select <- c("DATE", "FAIR.YIELD")
new_data <- new_data[, column_select]

# calculate Randomwalk
  LV <- tail(new_data$FAIR.YIELD, n = 1) #mengambil nilai terakhir 
  LV26 <- tail(new_data$FAIR.YIELD, n = 26) #mengambil 26 nilai terakhir, karena untuk memperhitungkan perbedaan 25 hari terakhir membutuhkan 1 hari tambahan.
  ln25_mean <- mean(diff(log(LV26))*100) #menghitung rata-rata perbedaan harian selama 25 hari terakhir 
  randomwalk <- (round((LV+LV*ln25_mean/100), digits=2)) #menghitung randomwalk (value terakhir + value terakhir*rata2 perbedaan 25 hari terakhir)

  # Add WAY value to the data frame
  yield_rw <- rbind(yield_rw, data.frame(SERI = seri, yield_rw = randomwalk))
}
  


# Print to csv
write.xlsx(yield_rw, "exp_yield.xlsx", rowNames = FALSE)

##################### selamat mencoba ##################################################
