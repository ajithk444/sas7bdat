# Import the Haven package to read sas files
# Import the writexl package to write to excel files
# Import the log4r package for logging.
library(haven)
library(writexl)
library(log4r)

# Create a new logger object with create.logger().
logger <- create.logger()

# Set the logger's file output: currently only allows flat files.
logfile(logger) <- file.path('C:/Users/KumarAji/Documents/Visual Studio 2015/Projects/R/R/Logs/Log.log')

# Set the current level of the logger.
level(logger) <- "INFO"

# Set directory path
sasDirectory <- trimws("C:/Users/KumarAji/Documents/Visual Studio 2015/Projects/R/R/Sas/")
xlsxDirectory <- trimws("C:/Users/KumarAji/Documents/Visual Studio 2015/Projects/R/R/Xlsx/")

# Get all the sas filenames in a list
sasFileNames <- list.files(path = sasDirectory, pattern = "*.sas7bdat")


# Loop and Convert each sas files to xlsx files 
for (i in 1:length(sasFileNames)) {
    # Set Sas input filepath 
    sasFilePath <- paste(sasDirectory, sasFileNames[i], sep = "")
    # Set xlsx file output path
    xlsxFilePath <- paste(xlsxDirectory, sasFileNames[i], sep = "")
    # Get output filename
    xlsFileName <- sasFileNames[i]

    print(xlsFileName)
    info(logger, xlsFileName)

    info(logger, paste("Reading sas7bdat file from :", sasFilePath, sep = ""))
    # Read sas file to a dataframe
    sasDataset <- read_sas(sasFilePath, catalog_file = NULL, encoding = NULL, cols_only = NULL)
    info(logger, paste("Completed reading sas7bdat file from :", sasFilePath, sep = ""))

    info(logger, paste("Trying to write to xlsx file :", paste(xlsxDirectory, ".xlsx", sep = ""), sep = ""))
    # Write dataframe to excel file
    write_xlsx(sasDataset, path = paste(xlsxFilePath, ".xlsx", sep = ""), col_names = TRUE)
    info(logger, paste("Xlsx file created successfully :", paste(xlsxFilePath, ".xlsx", sep = ""), sep = ""))
}
info(logger, paste(i, " - Files were processed successfully.", sep = ""))
print(paste(i, " - Files were processed successfully.", sep = ""))