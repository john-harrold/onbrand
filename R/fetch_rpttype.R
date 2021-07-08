#'@export
#'@title Determines Type of Report Template
#'@description  Based on the file extension for a template
#'
#'@param template Name of PowerPoint or Word file
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned list.
#'
#'@return List with the following elements
#' \itemize{
#' \item{rpttype}: Either Word, PowerPoint or Unknown
#' \item{rptext}: Either docx, pptx, or Unknown
#' \item{rptobj}: Either rdocx, rpptx, or Unknown
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{msgs}: Vector of messages
#'}
#'@examples
#' rpttype = fetch_rpttype(template=
#'   file.path(system.file(package="onbrand"), "templates", "report.pptx"))
#'
fetch_rpttype = function(template= NULL, verbose=TRUE){

  isgood = TRUE
  msgs = c()
  rpttype = "Unknown"
  rptext  = "Unknown"
  rptobj  = "Unknown"

  if(is.null(template)){
    isgood = FALSE
    msgs = c(msgs, "You must supply a file name")
  } else {
    if(grepl(pattern="pptx$", template)){
      rpttype = "PowerPoint"
      rptext  = "pptx"
      rptobj  = "rpptx"
    }else if(grepl(pattern="docx$", template)){
      rpttype = "Word"
      rptext  = "docx"
      rptobj  = "rdocx"
    } else {
      msgs = c(msgs, "Only pptx and docx file extensions are allowed.")
      isgood  = FALSE
    }
  }

  if(!isgood){
    msgs = c(msgs, "fetch_rpttype()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }
  res = list(isgood  = isgood,
             rpttype = rpttype,
             rptext  = rptext,
             rptobj  = rptobj,
             msgs    = msgs)
res}
