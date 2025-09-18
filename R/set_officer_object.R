#'@export
#'@title Places Officer Object Into Onbrand Report Object
#'@description After modifying a report object manually, you can return it to
#'the onbrand object using this function.
#'
#'@param obnd    onbrand report object
#'@param rpt     officer object        
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@return onbrand object with the report replaced
#'
#'@examples
#'
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'# pulling out the report
#'rpt = fetch_officer_object(obnd)$rpt
#'
#'# Modifications would be made here with officer directly
#'
#'# Replacing the report into the onbrand object
#'obnd = set_officer_object(obnd, rpt)
#'
#'@seealso \code{\link{fetch_officer_object}}
set_officer_object  = function (obnd, rpt = NULL, verbose=TRUE){

  isgood = TRUE
  msgs = c()

  if(obnd[["isgood"]]){
    rpt = obnd[["rpt"]]
    if(class(rpt) == class(obnd[["rpt"]])){
      obnd[["rpt"]] = rpt
    } else {
      isgood = FALSE
      msgs = c(msgs, paste0("The report type specified >", class(rpt), "< does not match"))
      msgs = c(msgs, paste0("the report type found in the officer object >", class(obnd[["rpt"]]), "<."))
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, "onbrand::set_officer_object()")
  }
  
  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  # Setting the messages returned
  if(!is.null(msgs) & obnd[["isgood"]]){
    obnd[[msgs]] = msgs
  }
  # Setting the obnd object as false
  if(!isgood & obnd[["isgood"]]){
    obnd[["isgood"]] = FALSE
  }

obnd}
