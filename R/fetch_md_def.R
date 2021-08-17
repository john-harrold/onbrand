#'@export
#'@title Fetch Markdown Default Format from onbrand Object
#'@description  Used to extract the formatting elements for a given style from
#'an onbrand object.
#'
#'@param obnd onbrand report object
#'@param style name of style in md_def for the report type in obnd to fetch (\code{"default"})
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@return list with the following elements
#' \itemize{
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{md_def}: List with the default format for the specified style
#' \item{msgs}: Vector of messages
#'}
#'@examples
#' obnd = read_template(
#'    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' obnd = fetch_md_def(obnd, style="default")
#' md_def =  obnd[["md_def"]]
fetch_md_def = function(obnd, style = "default", verbose=TRUE){

  isgood = TRUE
  msgs   = c()
  md_def = NULL

  if(obnd[[isgood]]){
    # Figuring out if we have a rpptx or rdocx object:
    meta_section =  obnd[["rptobj"]]
    if(style %in% names(obnd[["meta"]][[meta_section]][["md_def"]])){
      md_def = obnd[["meta"]][[meta_section]][["md_def"]][[style]]
    } else {
      isgood = FALSE
      msgs = c(msgs, paste0("The specified style >", style,"< was not found."))
      msgs = c(msgs, paste0("The following styles are avaialble:"))
      msgs = c(msgs, paste0("  ",paste(names(obnd[["meta"]][[meta_section]][["md_def"]]), collapse=", ")))
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, "onbrand::fetch_md_def()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  res = list(isgood = isgood,
             msgs   = msgs,
             md_def = md_def)

res}
