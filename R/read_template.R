#'@export
#'@title Read Word or PowerPoint Templates
#'@description  Takes a given template file/yaml mapping file combination, reads in
#'that information, checks to make sure the mapping information is correct and
#'then returns an onbrand object.
#'
#'@param template Name of PowerPoint or Word file to annotate (defaults to included PowerPoint template)
#'@param mapping Name of yaml file with configuration information
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object. 
#'
#'@return onbrand object which is a list with the following elements:
#' \itemize{
#' \item{isgood} Boolean variable indicating the current state of the object
#' \item{rpt} Officer object containing the initialized report
#' \item{rpttype} Type of report (either PowerPoint or Word)
#' \item{meta} Metadata read in from the yaml file
#' \item{mapping} Mapping yaml file
#' \item{msgs} Vector of messages indicating any errors that were encountered
#'}
#'@examples
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
read_template = function(template    = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                         mapping     = file.path(system.file(package="onbrand"), "templates", "report.yaml"),
                         verbose     = TRUE){

  rpt    = NULL
  meta   = NULL

  fr = fetch_rpttype(template=template)
  isgood  = fr[["isgood"]]
  msgs    = fr[["msgs"]]
  rpttype = fr[["rpttype"]]


  # these are required styles and their defaults
  required_styles = list(default = c(
                  "    default:",
                  "      color:                 black",
                  "      font.size:             12",
                  "      bold:                  TRUE",
                  "      italic:                FALSE",
                  "      underlined:            FALSE",
                  "      font.family:           Helvetica",
                  "      vertical.align:        baseline",
                  "      shading.color:         transparent"),
                          Table_Labels =  c(
                  "    Table_Labels:",
                  "      color:                 black",
                  "      font.size:             12",
                  "      bold:                  TRUE",
                  "      italic:                FALSE",
                  "      underlined:            FALSE",
                  "      font.family:           Helvetica",
                  "      vertical.align:        baseline",
                  "      shading.color:         transparent"))
  

  # Reading in the yaml file:
  if(file.exists(mapping)){
    meta = yaml::read_yaml(mapping)

    if(rpttype == "PowerPoint"){
      if(!("rpptx" %in% names(meta))){
        isgood = FALSE
        msgs = c(msgs, "The template is for PowerPoint but the mapping file does not contain an rpptx section.")
      }
    }
    if(rpttype == "Word"){
      if(!("rdocx" %in% names(meta))){
        isgood = FALSE
        msgs = c(msgs, "The template is for Word but the mapping file does not contain an rdocx section.")
      }
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, paste0("Template mapping file >", mapping, "< not found"))
  }


  # If everything checks out we read in the template file and compare it to
  # the mapping information
  if(isgood){
    if(rpttype == "PowerPoint"){
      rpt =officer::read_pptx(template)

      # pulling out the layout summary
      lay_sum = officer::layout_summary(rpt)

      # Checking to see if any layouts specified in the mapping file are
      # missing
      lay_meta = names(meta[["rpptx"]][["templates"]])
      if(any(!(lay_meta %in% lay_sum[["layout"]]))){
        isgood = FALSE
        msgs = c(msgs, "The following template layouts were specified in the mapping file but were not found in the supplied template:")
        msgs = c(msgs, paste(lay_meta[!(lay_meta %in% lay_sum[["layout"]])], collapse=", "))
      }

      # For each layout that was found we check supplied placeholders
      for(lay_found in lay_meta[(lay_meta %in% lay_sum[["layout"]])]){

        lp = officer::layout_properties(x= rpt, layout=lay_found, master=meta[["rpptx"]][["master"]])
        # All of the placeholders for the current layout
        phs = names(meta[['rpptx']][["templates"]][[lay_found]])
        for(ph in phs){
          phele = meta[['rpptx']][["templates"]][[lay_found]][[ph]]

          tmplp = dplyr::filter(lp, .data[["ph_label"]] == phele[["ph_label"]])

          if(!(nrow(tmplp) == 1)){
            isgood = FALSE
            msgs = c(msgs, "The following placeholder was not found:")
            msgs = c(msgs, paste0("  Layout:      ", lay_found))
            msgs = c(msgs, paste0("  Placeholder: ", ph))
          }
        }
      }
      # now we check the md_def for required elements
      if("md_def" %in% names(meta[['rpptx']])){

        # If we're missing any required styles we flag that here:
        if(!all(names(required_styles) %in% names(meta[["rpptx"]][["md_def"]]))){
          isgood = FALSE
          msgs = c(msgs, "The following required style(s) were not found:")
          msgs = c(msgs, paste0("  ", paste(names(required_styles)[!(names(required_styles) %in% names(meta[["rpptx"]][["md_def"]]))], collapse=", ")))
          msgs = c(msgs, "You can use the following:")
          for(required_style in names(required_styles)[!(names(required_styles) %in% names(meta[["rpptx"]][["md_def"]]))]){
            msgs = c(msgs, required_styles[[required_style]])
          }
        }
      } else {
        isgood = FALSE
        msgs = c(msgs, "You must have an md_def defined with values for at least the following styles:")
        msgs = c(msgs, paste0("  ",          paste(names(required_styles), collapse=", ")))
        msgs = c(msgs, "You can use the following:")
        msgs = c(msgs, "  md_def:")
        for(required_style in required_styles){
          msgs = c(msgs, required_style)
        }
      }
    } else {
      rpt     =  officer::read_docx(template)
      lay_sum = officer::styles_info(rpt)

      meta_styles = as.vector(unlist(meta[["rdocx"]][["styles"]]))


      # If we don't find all of the styles specified in the mapping file
      # then we flag that
      if(!all(meta_styles %in% lay_sum[["style_name"]])){
        isgood = FALSE
        msgs = c(msgs, "The following styles were specified in the mapping file but were not found in the supplied template:")
        msgs = c(msgs, paste( meta_styles[!(meta_styles %in% lay_sum[["style_name"]])] , collapse=", "))

        }

      # Now we're checking all of the meta_styles to ensure there is a md_def
      # entry as well
      if(!all(names(meta[["rdocx"]][["styles"]])  %in% names(meta[["rdocx"]][["md_def"]]))){
        isgood = FALSE
        msgs = c(msgs, "The following styles were specified but no md_def entries were found")
        msgs = c(msgs, paste(
           names(meta[["rdocx"]][["styles"]])[!(names(meta[["rdocx"]][["styles"]])  %in% names(meta[["rdocx"]][["md_def"]]))],
           collapse=", "))
        }
      }
    }

  if(!isgood){
    msgs = c(msgs, "read_template()")
  }
  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  res = list(isgood = isgood,
             rpt    = rpt,
             rpttype= rpttype,
             mapping= mapping,
             msgs   = msgs,
             meta   = meta)
res}
