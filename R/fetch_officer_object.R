#'@export
#'@title Extracts Officer Object From Onbrand Report Object
#'@description If you need modify the onbrand report object directly with
#'officer functions you can use this function to extract the report object
#'from the onbrand object.
#'
#'@param obnd onbrand report object
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@return List with the following elements
#' \itemize{
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{rpt}: Officer object
#' \item{msgs}: Vector of messages
#'}
#'@examples
#'
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#' rpt = fetch_officer_object(obnd)$rpt
#'
#'@seealso \code{\link{set_officer_object}}
fetch_officer_object  = function (obnd, verbose=TRUE){

  isgood = TRUE
  msgs = c()
  rpt = NULL

  if(obnd[["isgood"]]){
    rpt = obnd[["rpt"]]
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, "fetch_officer_object()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  res = list(isgood = isgood,
             rpt    = rpt,
             msgs   = msgs)


res}
