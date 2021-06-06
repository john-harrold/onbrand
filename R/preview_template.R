#'@export
#'@title Generate Report Previewing the Locations From Mapping File
#'@description  Takes an onbrand object with a loaded template and populates
#'the template with the elements from the mapping file.
#'
#'@param obnd onbrand report object
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object. 
#'
#'@return onbrand object with template previews added and any messages passed
#'along
#'@examples
#' obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'                       mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' obnd = preview_template(obnd)
#' # JMH add Word example
preview_template = function(obnd, verbose=TRUE){

  isgood  = TRUE
  msgs    = c()
  rpttype = NULL

  if(obnd[["isgood"]]){
      rpttype = obnd[["rpttype"]]
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  # Here we loop through each layout template
  if(isgood){
    if(rpttype == "PowerPoint"){
      templates = names(obnd[["meta"]][["rpptx"]][["templates"]])
      for(template in templates){
        # This will hold the content elements for the slide
        elements = list()
        # these are all the placeholder elements in the current template
        phs = names(obnd[["meta"]][["rpptx"]][["templates"]][[template]])
        # Now we build out the element list
        for(ph in phs){
          
          tmp_text = paste0(ph, ":", obnd[["meta"]][["rpptx"]][["templates"]][[template]][[ph]][["ph_label"]])

          # If we're processing a title then we add the template as well
          if(grepl("^title", ignore.case=TRUE, obnd[["meta"]][["rpptx"]][["templates"]][[template]][[ph]][["ph_label"]])){
            tmp_text = paste0(template, "=", tmp_text)
          }

          if(obnd[["meta"]][["rpptx"]][["templates"]][[template]][[ph]][["content_type"]] == "text"){
            elements[[ph]] = list(content = tmp_text,
                                  type     = "text")
          } else {
            elements[[ph]] = list(content = c("1", tmp_text),
                                  type     = "list")
          }
        }
        # Now adding a slide with each element on it
        obnd = report_add_slide(obnd,
          template = template,
          elements = elements)
      }
    } else if(rpttype == "Word"){
      # JMH add word stuff here
    }
  }
    

  #-------
  # If errors were encountered we make sure that the state of the reporting
  # object is set to false
  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, paste0("mapping file: ", obnd[["mapping"]]))
    msgs = c(msgs, "report_add_slide()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  # Passing the messages along in the onbrand object
  obnd[["msgs"]] = msgs
  
obnd}

