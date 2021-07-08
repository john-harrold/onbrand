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
#' obnd = read_template(
#'    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' obnd = preview_template(obnd)
#'
#' obnd = read_template(
#'    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' obnd = preview_template(obnd)
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
          # Skipping NULL (missing) placeholders
          if(!is.null(obnd[["meta"]][["rpptx"]][["templates"]][[template]][[ph]][["ph_label"]])){
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
        }
        # Now adding a slide with each element on it
        obnd = report_add_slide(obnd,
          template = template,
          elements = elements,
          verbose  = verbose)

        msgs = c(msgs, obnd[["msgs"]])
      }
    } else if(rpttype == "Word"){
     # Walking through each style and adding content for that style

     # Pulling a summary of the layouts out of the officer object
     lay_sum = officer::styles_info(obnd[["rpt"]])

     # pulling out the template styles specified in the mapping file
     template_styles = obnd[["meta"]][["rdocx"]][["styles"]]

     # Walking through and adding content for each style to annotate
     # the user specified style name/actual style used.
     for(user_style in names(template_styles)){
       # Style name used in the document:
       docx_style = template_styles[[user_style]]
       # type of style using Word docx terminology
       docx_type   = dplyr::filter(lay_sum, .data[["style_name"]] == docx_style)[["style_type"]]

       tmp_text = paste0(user_style, " (onbrand name):", docx_style, " (Word name)")
       tab_example = data.frame( Number = c(1,2,3,4),
                                 Text   = "Here")

       # Depending on the docx type we add either text or tabular content
       if(docx_type %in% c("paragraph", "charcter")){
         obnd = report_add_doc_content(obnd,
                    type     = "text",
                    content  = list(text=tmp_text, style=user_style),
                    verbose  = verbose)
         msgs = c(msgs, obnd[["msgs"]])
       }
       if(docx_type %in% c("table")){
        obnd[["rpt"]] = officer::body_add_par(  x=obnd[["rpt"]], value=tmp_text)
        obnd[["rpt"]] = officer::body_add_table(x=obnd[["rpt"]], value=tab_example, style = docx_style)
       }
      }
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

