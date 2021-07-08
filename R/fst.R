#'@export
#'@title Fetch Word Style
#'@description  Retrieves the style name in Word for a specified onbrand style
#'name.
#'
#'@param obnd onbrand report object
#'@param osn onbrand Word style name to fetch
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned list.
#'
#'@return List with the following elements
#' \itemize{
#' \item{wsn}:    Word style name that corresponds to the specified onbrand style name (\code{osn})
#' \item{dff}:    Default font format for that style (the corresponding \code{md_def} section of the yaml file for that style)
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{msgs}:   Vector of messages
#'}
#'@examples
#'# Creating an onbrand object:
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'# Pulling out the placeholder information:
#'st = fst(obnd, "Heading_3")
#'
#'
fst = function(obnd,
               osn      = NULL,
               verbose  = TRUE){

  isgood = TRUE
  msgs   = c()
  wsn    = NULL
  dff    = NULL

  #----------------------------
  # First we make sure the onbrand object is good. Next, we walk through the
  # onbrand object and make sure we have: meta data, powerpoint data, the
  # template exists and lastly if the specified placeholder name exists
  if(obnd[["isgood"]]){
    if("meta" %in% names(obnd)){
      if("rdocx" %in% names(obnd[["meta"]])){
        if(osn %in% names(obnd[["meta"]][["rdocx"]][["styles"]])){
         wsn = obnd[["meta"]][["rdocx"]][["styles"]][[osn]]
         dff =  obnd[["meta"]][["rdocx"]][["md_def"]][[osn]]
        } else {
          isgood = FALSE
          msgs = c(msgs, paste0("The specified onbrand style name>", osn, "< was not found"))
          msgs = c(msgs, paste0("in the mapping file."))
          msgs = c(msgs, paste0("The following named styles are available:"))
          msgs = c(msgs, paste0("  ", paste(names(obnd[["meta"]][["rdocx"]][["styles"]]), collapse=", ")))
        }
      } else {
        isgood = FALSE
        msgs = c(msgs, "No Word mapping information found in onbrand object")
      }
    } else {
      isgood = FALSE
      msgs = c(msgs, "No mapping information found in onbrand object")
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }
  #----------------------------
  # If errors were encountered we make sure that the state of the reporting
  # object is set to false
  if(!isgood){
    msgs = c(msgs, "fst()")
  }
  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  res = list(wsn     = wsn,
             dff     = dff,
             isgood  = isgood,
             msgs    = msgs)

res}

