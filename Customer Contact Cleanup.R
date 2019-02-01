# Cleaning up a spreadsheet of client names and addresses. The spreadsheet was generated from
# Quickbooks and contained duplicate entries, multiple emails on the same row and useless title information
# The goal is to generate a single list of emails that can be imported into MailChimp for marketing. 

#Load necessary programs to read the xlsx file and split columns
  library(xlsx)
  library(splitstackshape)

#create the initial dataframe from excel and delete the first 4 rows and first column
  dat <- read.xlsx("name_of_file.xlsx", 1, startRow = 5, colIndex = c(2:4), header=TRUE)

#Create new dataframe that includes new columns for additional email (JF)
  dat1 <- cSplit(dat, "Email", sep = ", ", type.convert = FALSE)
  dat1 <- cSplit(dat, "Email", sep = ",", type.convert = FALSE)
  
#First, break each email column out into its own data frame (JF)
  dat_e1 <- subset(dat1, select = -c(Email_2, Email_3))
  dat_e2 <- subset(dat1, select = -c(Email_1, Email_3))
  dat_e3 <- subset(dat1, select = -c(Email_1, Email_2))

#Rename columns to match
  dat_names <- c("Customer", "Name", "Email")
  names(dat_e1) <- dat_names
  names(dat_e2) <- dat_names
  names(dat_e3) <- dat_names
  
#Bind new data frames into single data frame (JF)
  dat_e4 <- as.data.frame(rbind(dat_e1, dat_e2, dat_e3))

#Order dataframe based on Customer & delete rows without an email
  dat_e4 <- dat_e4[order(dat_e4$Customer),]
  dat_e4 <- dat_e4[complete.cases(dat_e4),]
  
#Remove duplicate rows with duplicate email addresses
  dat_clean <- dat_e4[!duplicated(dat_e4$Email),]

#Clean up data no longer needed
  rm(dat_e1, dat_e2, dat_e3, dat, dat1, dat_names, dat_e4)
  
#Save the new dataframe as a .csv file to your working directory
  write.xlsx(dat_clean, file = "new_file_name.xlsx", sheetName = "Email List") 
  