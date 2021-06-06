
#-----------------------------------------------------------
#fetch_report_format
#'@export
#'@title Fetch The Specified Report Formatting Information
#'@description Returns a list the default font format for the report element
#'
#'@param cfg ubiquity system object    
#'@param rptname report name initialized with \code{system_report_init}
#'@param element report element to fetch: for Word reports it can be 
#' "default" (default),       "Normal",        "Code",  "TOC",          
#' "Heading_1",  "Heading_2", "Heading_3", "Table", "Table_Labels",
#' "Table_Caption", "Figure",  and  "Figure_Caption"
#'
#'@return list of current parameter gauesses
fetch_report_format <- function(cfg, rptname="default", element="Default"){

default_format = NULL
isgood = TRUE

if(cfg[["reporting"]][["enabled"]]){
  if(element %in% names(cfg[["reporting"]][["reports"]][[rptname]][["meta"]][["md_def"]])){
    default_format = cfg[["reporting"]][["reports"]][[rptname]][["meta"]][["md_def"]][[element]]
  } else if(!(rptname %in% names(cfg[["reporting"]][["reports"]]))){
    isgood = FALSE
    vp(cfg, paste("Error: The report name >", rptname,"< not found", sep=""))
  } else if(!(element %in% names(cfg[["reporting"]][["reports"]][[rptname]][["meta"]][["md_def"]]))){
    isgood = FALSE
    vp(cfg, paste("Error: The report element >", element,"< not found", sep=""))
  } 
} else {
  isgood = FALSE
  vp(cfg, "Error: Reporting not enabled")
}

# If we failed to get the report element formatting for the rptname we try to
# pull a default ubiquity format:
if(!isgood){
  # First we try the word format
  if(element %in% names(cfg[["reporting"]][["meta_docx"]][["md_def"]])){
    default_format = cfg[["reporting"]][["meta_docx"]][["md_def"]][[element]]
    vp(cfg, paste("Returning default format for Word element >", element,"< for ubiquity default document", sep=""))
  }else if(element %in% names(cfg[["reporting"]][["meta_pptx"]][["md_def"]])){
  # Then we try the powerpoint format
    default_format = cfg[["reporting"]][["meta_pptx"]][["md_def"]][[element]]
    vp(cfg, paste("Returning default format for PowerPoint element >", element,"< for ubiquity default document", sep=""))

  } else {
    vp(cfg, paste("Unable to find formatting for element >", element,"< in either Word or Powerpoint default ubiquity documents", sep=""))
    vp(cfg, "Returning NULL")
  }
  vp(cfg, "fetch_report_format()")
}


  res = list(isgood         = isgood,
             default_format = default_format)
res}

