#'@export
#'@title Save Onbrand Report to a File
#'@description Saves report in onbrand object to the specified file.
#'
#'@param obnd onbrand report object
#'@param output_file File name to save the report.
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@return List with the following elements
#' \itemize{
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{msgs}:   Vector of messages
#'}
#'@examples
#'
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#' save_report(obnd, tempfile(fileext = ".pptx"))
#'
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#' save_report(obnd, tempfile(fileext = ".docx"))
#'
save_report  = function (obnd,
                         output_file = NULL,
                         verbose     = TRUE){

  msgs          = c()
  isgood        = TRUE
  rpttype       = "Unknown"


  if(obnd[["isgood"]]){
     rpttype = obnd[["rpttype"]]

  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  if(is.null(output_file)){
    isgood = FALSE
    msgs   = c(msgs, "You need to specify an output_file name.")
  } else {
    fr = fetch_rpttype(output_file, verbose=verbose)
    # Appending any messages from fetch_rpttype if verbose was turned off
    if(!fr[["isgood"]] & !verbose){
      msgs = c(msgs, fr[["msgs"]])
    }
    # Making sure the file extensions match
    if(fr[["rptext"]] != obnd[["rptext"]]){
      isgood = FALSE
      msgs = c(msgs, paste0("The output_file type >", fr[["rpttype"]], "< does not match the report " ))
      msgs = c(msgs, paste0("found in the supplied onbrand object >", obnd[["rpttype"]], "<." ))
      msgs = c(msgs, paste0("You must supply an output file with a >.", obnd[["rptext"]], "< file extension." ))
    }
  }


  # If we pass all the checks above we save the file
  if(isgood){
    #-------
    # Applying any placeholders
    if(rpttype == "Word"){
      if("placeholders" %in% names(obnd)){
        # Looping through each placeholder
        for(phn in names(obnd[["placeholders"]])){
          # Here we pull out the value (phv) and locatio (phl) of each
          # placeholder:
          pht = paste("===",phn,"===", sep="")
          phv = obnd[["placeholders"]][[phn]][["value"]]
          phl = obnd[["placeholders"]][[phn]][["location"]]
          if(phl == "body"){
            obnd[["rpt"]] = officer::body_replace_all_text(
                 old_value      = pht,
                 new_value      = phv ,
                 fixed          = TRUE,
                 only_at_cursor = FALSE,
                 warn           = FALSE,
                 x              = obnd[["rpt"]]
                 )
          }
          if(phl == "header"){
            obnd[["rpt"]] = officer::headers_replace_all_text(
                 old_value      = pht,
                 new_value      = phv ,
                 fixed          = TRUE,
                 only_at_cursor = FALSE,
                 warn           = FALSE,
                 x              = obnd[["rpt"]]
                 )
          }
          if(phl == "footer"){
            obnd[["rpt"]] = officer::footers_replace_all_text(
                 old_value      = pht,
                 new_value      = phv ,
                 fixed          = TRUE,
                 only_at_cursor = FALSE,
                 warn           = FALSE,
                 x              = obnd[["rpt"]]
                 )
          }
        }
      }

      # Checking for post processing code
      if("post_processing" %in% names(obnd[["meta"]][["rdocx"]])){
        if(!is.null(obnd[["meta"]][["rdocx"]][["post_processing"]])){
          # pulling out the officer object
          rpt = fetch_officer_object(obnd=obnd, verbose=FALSE)
          tcres =
            tryCatch(
              {
               # Evaulating the post processing code
                eval(parse(text=obnd[["meta"]][["rdocx"]][["post_processing"]]))
              list(isgood=TRUE, rpt = rpt)},
             error = function(e) {
              list(isgood=FALSE, error=e)})
      
      
          # putting the object back, but only if the
          # code executed successfully
          if(tcres[["isgood"]]){
            obnd = set_officer_object(obnd=obnd, rpt=tcres[["rpt"]], verbose=FALSE)
          } else {
            #if the user defined code failed, then we generate an error
            isgood = FALSE
            msgs = c(msgs, "Failed to evaluate the following user defined code:")
            msgs = c(msgs, obnd[["meta"]][["rdocx"]][["post_processing"]])
            msgs = c(msgs, "The following error was encountered")
            msgs = c(msgs, tcres$e$message)
          }
        }
      }
    }

    if(rpttype == "PowerPoint"){
      # Checking for post processing code
      if("post_processing" %in% names(obnd[["meta"]][["rpptx"]])){
        if(!is.null(obnd[["meta"]][["rpptx"]][["post_processing"]])){
          # pulling out the officer object
          rpt = fetch_officer_object(obnd=obnd, verbose=FALSE)
          tcres =
            tryCatch(
              {
               # Evaulating the post processing code
                eval(parse(text=obnd[["meta"]][["rpptx"]][["post_processing"]]))
              list(isgood=TRUE, rpt = rpt)},
             error = function(e) {
              list(isgood=FALSE, error=e)})
          # putting the object back, but only if the
          # code executed successfully
          if(tcres[["isgood"]]){
            obnd = set_officer_object(obnd=obnd, rpt=tcres[["rpt"]], verbose=FALSE)
          } else {
            #if the user defined code failed, then we generate an error
            isgood = FALSE
            msgs = c(msgs, "Failed to evaluate the following user defined code:")
            msgs = c(msgs, obnd[["meta"]][["rpptx"]][["post_processing"]])
            msgs = c(msgs, "The following error was encountered")
            msgs = c(msgs, tcres$e$message)
          }
        }
      }
    }


    #-------
    print(obnd[["rpt"]], output_file)
  }

  # If errors were encountered we make sure that the state of the reporting
  # object is set to false
  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, "onbrand::save_report()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  # Stopping
  if(!isgood){
    stop()
  }


  res = list(isgood = isgood,
             msgs   = msgs)
res}
