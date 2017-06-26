#' A function to read in csv file and return a confirmation
#' 
#' @param DataPath The folder where the file is stored in 
#' @param filename The name of the file
#' @return A data frame in R, and a printed sentense report whether this file can be successfully loaded.
#' @seealso read.csv and read.csv2
#' @export 
#' @examples
#' x<-ReadCSVFile(/Users/newfile, newfile)

ReadCSVFile <- function(DataPath, filename) {
  
  # Function: 
  # Check for existance of *filename* in *DataPath*.
  # Read file into local object *responsetable* and return it.
  ## Reads blank and negative response values "", "-1" and "-2" as missing 
  ## values and recorded as "NA" in the dataframe.
  
  # Construct filename
  
  datfilename <- paste(DataPath, filename, ".csv", sep = "", collapse = "")
  
  # Check whether file exists
  
  if (file.access(datfilename, mod = 4) != 0) {
    print(paste("ERROR: Cannot open input file:", datfilename, collapse = " ") )
  } 
  else {
    # File does exist
    print(paste("Found file:", datfilename, collapse = " ") )
    
    # Read data table
    responsetable <- read.csv(file = datfilename,
                              #colClasses = "character",
                              header = TRUE, na.strings = c("", "-1", "-2"), 
                              row.names = NULL, flush = TRUE, fill = TRUE,check.names = FALSE )      
    return(responsetable)
  }  
}
