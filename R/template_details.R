#'@export
#'@title Show Template Details for `onbrand` Object
#'@description  Takes an onbrand object with a loaded template and displays
#'details about the template. For PowerPoint this contains the template names
#'and elements present for that template. For Word it will contain defined
#'text and table styles. This information can be displayed in the console,
#'returned as text or formatted for use in RMarkdown documentation.
#'
#'@param obnd onbrand report object
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned results
#'object.
#'
#'@return list with the following elements:
#' \itemize{
#'  \item{rpttype}:      Type of report (either PowerPoint or Word)
#'  \item{msgs}:         Vector of messages with details or any errors that were encountered
#'  \item{txt}:          Vector of template details in text format
#'  \item{df}:           Vector of template details in a dataframe
#'  \item{ft}:           Vector of template details in flextable format
#'  \item{isgood}:       Boolean variable indicating the current state of the object
#' }
#'@examples
#' obnd = read_template(
#'    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' details = template_details(obnd)
#'
#' obnd = read_template(
#'    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#' details = template_details(obnd)
template_details = function(obnd, verbose=TRUE){

  isgood  = TRUE
  msgs    = c()
  rpttype = NULL
  res     = list()
  d_txt   = c()
  d_df    = NULL     # Details in data frame format
  d_ft    = NULL     # Details in flextable format

  if(obnd[["isgood"]]){
      rpttype = obnd[["rpttype"]]
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  # Here we loop through each layout template
  if(isgood){
    d_txt = c(d_txt, paste0("Mapping:     ", obnd[["mapping"]]))
    d_txt = c(d_txt, paste0("Report Type: ", rpttype))
    if(rpttype == "PowerPoint"){
      templates = names(obnd[["meta"]][["rpptx"]][["templates"]])
      for(template in templates){
        # This will hold the content elements for the slide
        elements = list()
        # these are all the placeholder elements in the current template
        phs = names(obnd[["meta"]][["rpptx"]][["templates"]][[template]])
        # Now we build out the element list
        d_txt = c(d_txt, paste0(template, " (master/template)"))
        for(ph in phs){
          content_type = obnd[["meta"]][["rpptx"]][["templates"]][[template]][[ph]][["content_type"]]
          # Text details
          d_txt = c(d_txt, paste0("  > ",ph, " (", content_type, ")"))
          # Dataframe details:
          d_df = rbind(d_df,
             data.frame(template=template,
                        ph      = ph,
                        ct      = content_type))
        }

        msgs = c(msgs, obnd[["msgs"]])
      }

      # Creating the flextable
      header = list(template = c("Master/Template", "Name"),
                    ph       = c("onbrand",         "Placeholder"),
                    ct       = c("Content",         "Type"))

      d_ft = flextable::flextable(d_df)
      d_ft = flextable::delete_part(d_ft, part   =        "header")
      d_ft = flextable::add_header(d_ft,  values = as.list(header))
      d_ft = flextable::theme_box(d_ft)
      d_ft = flextable::merge_v(d_ft, j=~template)
    } else if(rpttype == "Word"){
      # pulling out the template styles specified in the mapping file
      template_styles = obnd[["meta"]][["rdocx"]][["styles"]]

      # Pulling a summary of the layouts out of the officer object
      lay_sum = officer::styles_info(obnd[["rpt"]])

      # Walking through each style:
      for(user_style in names(template_styles)){
#       # Style name used in the document:
        docx_style = template_styles[[user_style]]
#       # type of style using Word docx terminology
        docx_type   = dplyr::filter(lay_sum, .data[["style_name"]] == docx_style)[["style_type"]]
        d_df = rbind(d_df,
          data.frame(onbrand = user_style,
                     Word    = docx_style,
                     type    = docx_type ))
      }


      # Sorting by type then onbrand name
      d_df = dplyr::arrange(d_df, .data[["type"]], .data[["onbrand"]])

      # Now creating the text result
      d_txt = c(d_txt, "  onbrand style (word style, style type)")
      d_txt = c(d_txt, "  --------------------------------------")
      for(ridx in 1:nrow(d_df)){
        d_txt = c(d_txt,
          paste0("  ", d_df[ridx,][["onbrand"]], " (", d_df[ridx,][["Word"]],  ", ",d_df[ridx,][["type"]], ")")
          )
      }

      # Creating the flextable
      header = list(onbrand  = c("onbrand", "Style"),
                    Word     = c("Word",    "Style"),
                    type     = c("Style",   "Type"))

      d_ft = flextable::flextable(d_df)
      d_ft = flextable::delete_part(d_ft, part   =        "header")
      d_ft = flextable::add_header(d_ft,  values = as.list(header))
      d_ft = flextable::theme_box(d_ft)
      d_ft = flextable::autofit(d_ft)
    }
  }

  #-------
  if(!isgood){
    msgs = c(msgs, paste0("mapping file: ", obnd[["mapping"]]))
    msgs = c(msgs, "onbrand::template_details()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }
  if(verbose & !is.null(d_txt)){
    message(paste(d_txt, collapse="\n"))
  }

  res = list(
          msgs    = msgs,
          rpttype = rpttype,
          txt     = d_txt,
          df      = d_df,
          ft      = d_ft,
          isgood  = isgood
        )

return(res)}

