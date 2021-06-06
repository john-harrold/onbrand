#'@export
#'@title Determines Type of Report Template
#'@description  Based on the file extension for a template
#'
#'@param template Name of PowerPoint or Word file
#'
#'@return List with the following elements
#' \itemize{
#' \item{rpttype} Either Word, PowerPoint or Unknown
#' \item{isgood} Boolean variable indicating success or failure
#' \item{msgs} Vector of messages
#'}
#'@examples
#' fetch_rpttype(template= file.path(system.file(package="onbrand"), "templates", "report.pptx"))
#'
fetch_rpttype = function(template= file.path(system.file(package="onbrand"), "templates", "report.pptx")){

  isgood = TRUE
  msgs = c()

  if(grepl(pattern="pptx$", template)){
    rpttype = "PowerPoint"
  }else if(grepl(pattern="docx$", template)){
    rpttype = "Word"
  } else {
    rpttype = "Unknown"
    msgs = c(msgs, "Only pptx and docx file extensions are allowed.")
    isgood  = FALSE
  }

  res = list(isgood  = isgood,
             rpttype = rpttype,
             msgs    = msgs)
res}
