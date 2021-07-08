#'@export
#'@title Fetch The Specified Report Formatting Information
#'@description Returns a list of the default font format for the report element
#'
#'@param obnd onbrand report object
#'@param format_name Name of report format to fetch; this is defined in the md_def
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned list.
#'section for the given report type (\code{"default"})
#'
#'@return list containing the following elements
#'\itemize{
#'\item{isgood}: Boolean variable indicating success or failure
#'\item{msgs}: Vector of messages
#'\item{format_details}: List containing the format details for the specified
#'format_name
#'}
#'@examples
#' obnd = read_template(
#'        template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'         mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#' fr = fetch_report_format(obnd)
fetch_report_format <- function(obnd, format_name="default", verbose=TRUE){

  isgood         = TRUE
  msgs           = c()
  format_details = NULL

  if(!obnd[["isgood"]]){
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  if(obnd[["rpttype"]] == "PowerPoint"){
    meta_section = "rpptx"
  } else if(obnd[["rpttype"]] == "Word"){
    meta_section = "rdocx"
  }

  # This should contain all the markdown defaults for the current report type
  md_def = obnd[["meta"]][[meta_section]][["md_def"]]

  if(format_name %in% names(md_def)){
    format_details = md_def[[format_name]]
  } else {
    msgs = c(msgs, paste0("The format_name >", format_name, "< was not found in the "))
    msgs = c(msgs, paste0(obnd[["rpttype"]], " section >", meta_section, "< of the mapping file:"))
    msgs = c(msgs, paste0("  ", obnd[["mapping"]]))
  }

  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, "fetch_report_format()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }


  res = list(isgood         = isgood,
             msgs           =  msgs,
             format_detials = format_details)
res}

