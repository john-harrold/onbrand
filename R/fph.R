#'@export
#'@title Fetch PowerPoint Placeholder
#'@description  Retrieves the placeholder name in PowerPoint for a specified
#'layout element.
#'
#'@param obnd onbrand report object
#'@param template Name of slide template (name from templates in yaml mapping file)
#'@param pn Placehodler name to fetch
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned list.
#'
#'@return List with the following elements
#' \itemize{
#' \item{ph}:      Placeholder label or \code{NULL} if failure
#' \item{type}:    Placeholder content type in PowerPoint or \code{NULL} if failure
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{msgs}: Vector of messages
#'}
#'@examples
#'# Creating an onbrand object:
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'# Pulling out the placeholder information:
#'ph = fph(obnd, "two_content_header_text", "content_left_header")
#'
#'
fph = function(obnd,
               template = NULL,
               pn       = NULL,
               verbose  = TRUE){

  isgood = TRUE
  msgs   = c()
  pl     = NULL
  type   = NULL


  #----------------------------
  # First we make sure the onbrand object is good. Next, we walk through the
  # onbrand object and make sure we have: meta data, powerpoint data, the
  # template exists and lastly if the specified placeholder name exists
  if(obnd[["isgood"]]){
    if("meta" %in% names(obnd)){
      if("rpptx" %in% names(obnd[["meta"]])){
        if(template %in% names(obnd[["meta"]][["rpptx"]][["templates"]])){
          if(pn %in% names(obnd[["meta"]][["rpptx"]][["templates"]][[template]])){
            # If we've made it this far then it's all good and we pull out hte
            # placeholder lable and content type
            pl   = obnd[["meta"]][["rpptx"]][["templates"]][[template]][[pn]][["ph_label"]]
            type = obnd[["meta"]][["rpptx"]][["templates"]][[template]][[pn]][["content_type"]]
          } else {
            isgood = FALSE
            msgs = c(msgs, paste0("The placeholder name >", pn, "< was not found"))
            msgs = c(msgs, paste0("in the layout template >", template, "<"))
            msgs = c(msgs, paste0("The following named placeholder names are available:"))
            msgs = c(msgs, paste0("  ", paste(names(obnd[["meta"]][["rpptx"]][["templates"]][[template]]), collapse=", ")))
          }
        } else {
          isgood = FALSE
          msgs = c(msgs, paste0("The slide template >", template, "< was not found"))
          msgs = c(msgs, paste0("in the mapping file."))
          msgs = c(msgs, paste0("The following named layout templates are available:"))
          msgs = c(msgs, paste0("  ", paste(names(obnd[["meta"]][["rpptx"]][["templates"]]), collapse=", ")))
        }
      } else {
        isgood = FALSE
        msgs = c(msgs, "No PowerPoint mapping information found in onbrand object")
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
    msgs = c(msgs, "onbrand::fph()")
  }
  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  res = list(pl      = pl,
             type    = type,
             isgood  = isgood,
             msgs    = msgs)

res}
