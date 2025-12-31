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
#' \item{isgood}:       Boolean variable indicating the current state of the object
#' \item{rpt}:          Officer object containing the initialized report
#' \item{rpttype}:      Type of report (either PowerPoint or Word)
#' \item{key_table}:    Empty (NULL) mapping table for tracking cross referencing (Word only)
#' \item{placeholders}: Empty list to hold placeholder substitution text (Word only)
#' \item{meta}:         Metadata read in from the yaml file
#' \item{mapping}:      Mapping yaml file
#' \item{msgs}:         Vector of messages indicating any errors that were encountered
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
  rptobj  = fr[["rptobj"]]
  rptext  = fr[["rptext"]]

  # For Word reports this defines both the default order and the allowed
  # commands
  def_tab_order  = c("table", "notes", "caption")
  def_fig_order  = c("figure", "notes", "caption")
  def_tab_seq_id = "Table"
  def_fig_seq_id = "Figure"
  def_tab_num    = 'list(officer::run_autonum(pre_label = "", seq_id = Caption_Seq_Id, post_label="", start_at=Caption_Start_At))'
  def_fig_num    = 'list(officer::run_autonum(pre_label = "", seq_id = Caption_Seq_Id, post_label="", start_at=Caption_Start_At))'

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

      # JMH filter first by Master to make sure the correct master was specified:
      # Include Master name in the message below

      # Checking to see if any layouts specified in the mapping file are
      # missing
      lay_meta = names(meta[["rpptx"]][["templates"]])
      if(any(!(lay_meta %in% lay_sum[["layout"]]), na.rm=TRUE)){
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

          # This template may not have each label defined, we skip the NULL
          # values
          if(!is.null(phele[["ph_label"]])){
            tmplp = dplyr::filter(lp, .data[["ph_label"]] == phele[["ph_label"]])
            if(!(nrow(tmplp) == 1)){
              isgood = FALSE
              msgs = c(msgs, "The following placeholder was not found:")
              msgs = c(msgs, paste0("  Layout:      ", lay_found))
              msgs = c(msgs, paste0("  Placeholder: ", ph))
            }
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

      # Checking to make sure that document defaults have been specified
      #                   Default sytle name   allowed style types in word document
      req_doc_defs = list("Text"               = c("paragraph", "character"),
                          "Table"              = c("table"),
                          "Table_Caption"      = c("paragraph", "character"),
                          "Figure_Caption"     = c("paragraph", "character"),
                          "Notes"              = c("paragraph", "character"))
      if(isgood){
        # First we make sure that the expected defaults were specified:
        if(all(names(req_doc_defs) %in% names(meta[["rdocx"]][["doc_def"]]) )){
          # Now we make sure those specified defaults are actual styles:
          def_styles = as.vector(unlist(meta[["rdocx"]][["doc_def"]]))
          if(all(def_styles %in% names(meta[["rdocx"]][["styles"]]))){
            # Checking the default styles to make sure they are the correct type:
            # Here we define the styles locally in terms of the user specified names
            for(def_style in names(req_doc_defs)){
              Word_style      = meta[["rdocx"]][["styles"]][[meta[["rdocx"]][["doc_def"]][[def_style]]]]
              Word_style_type = dplyr::filter(lay_sum, .data[["style_name"]] == Word_style)[["style_type"]]
              allowed_style_types = req_doc_defs[[def_style]]
              # if(Word_style == "Table Caption"){
              #     browser()
              # }
              # If the word style type is not in the allowed types we flag it
              if(!(Word_style_type %in% allowed_style_types)){
                isgood = FALSE
                msgs = c(msgs, "The requred document style default (doc_def) is the wrong type.")
                msgs = c(msgs, paste0("  default:      ", def_style))
                msgs = c(msgs, paste0("  style:        ", Word_style))
                msgs = c(msgs, paste0("  style type:   ", Word_style_type))
                msgs = c(msgs, paste0("  allowed types ", paste(allowed_style_types, collapse=", ")))
              }
            }
          } else {
            isgood = FALSE
            msgs = c(msgs, "Default user styles in doc_def are present that have not been defined in styles.")
            msgs = c(msgs, "Please check the following values specified in doc_def")
            msgs = c(msgs, paste0("  ", paste(def_styles[!(def_styles %in% names(meta[["rdocx"]][["styles"]]))], collapse=", ")))
          }
        } else {
          isgood = FALSE
          defs_missing = names(req_doc_defs)[ !(names(req_doc_defs) %in% names(meta[["rdocx"]][["doc_def"]]))]
          msgs = c(msgs, "In doc_def you must specify default styles to be used.")
          msgs = c(msgs, "The following default styles were not specified:")
          msgs = c(msgs, paste0("  ", paste0(defs_missing, collapse=", ")))
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


          if(!("separator" %in% names(meta[["rdocx"]][["formatting"]]))){
            isgood = FALSE
            msgs = c(msgs, "Unable to find the separator formatting field")
            msgs = c(msgs, 'In the US or Canada this should probably be "," while in Europe it should probably be ";".')
          } else {
            if(!(meta[["rdocx"]][["formatting"]][["separator"]] %in% c(",", ";"))){
              isgood = FALSE
              msgs = c(msgs, paste0("The separator formatting field provided >", meta[["rdocx"]][["formatting"]][["separator"]],"<"))
              msgs = c(msgs, 'is invalid. It should be either ";" or "," ')

            }
          }

          # Checking for depreciated things
          if("Table_Caption_Location" %in% names(meta[["rdocx"]][["formatting"]])){
            msgs = c(msgs, "Warning: Table_Caption_Location")
            msgs = c(msgs, "  -> This formatting option has been depreciated use Table_Order now.")
          }
          if("Figure_Caption_Location" %in% names(meta[["rdocx"]][["formatting"]])){
            msgs = c(msgs, "Warning: Figure_Caption_Location")
            msgs = c(msgs, "  -> This formatting option has been depreciated use Figure_Order now.")
          }
          
          # Now we're checking optional things 
          # Figure and table order
          if(!("Table_Order" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Table_Order"]] = def_tab_order
            msgs = c(msgs, "Table_Order not specified using defaults:")
            msgs = c(msgs, paste0("  -> ", paste(def_tab_order, collapse=", ")))
          }
          if(!("Figure_Order" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Figure_Order"]] = def_fig_order
            msgs = c(msgs, "Figure_Order not specified using defaults:")
            msgs = c(msgs, paste0("  -> ", paste(def_fig_order, collapse=", ")))
          }

          # Table and figure seq_ids 
          if(!("Table_Seq_Id" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Table_Seq_Id"]] = def_tab_seq_id
            msgs = c(msgs, "Table_Seq_Id not specified using defaults:")
            msgs = c(msgs, paste0('  -> "', def_tab_seq_id, '"'))
          }
          if(!("Figure_Seq_Id" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Figure_Seq_Id"]] = def_fig_seq_id
            msgs = c(msgs, "Figure_Seq_Id not specified using defaults:")
            msgs = c(msgs, paste0('  -> "', def_fig_seq_id, '"'))
          }
          #Figure and table numbers
          if(!("Table_Number" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Table_Number"]] = def_tab_num   
            msgs = c(msgs, "Table_Number not specified using defaults:")
            msgs = c(msgs, paste0('  -> "', def_tab_num, '"'))
          }
          if(!("Figure_Number" %in% names(meta[["rdocx"]][["formatting"]]))){
            meta[["rdocx"]][["formatting"]][["Figure_Number"]] = def_fig_num   
            msgs = c(msgs, "Figure_Number not specified using defaults:")
            msgs = c(msgs, paste0('  -> "', def_fig_num, '"'))
          }

          # Checking the Table_Order and Figure_Order to make sure they only
          # have the correct fields
          if(!all((meta[["rdocx"]][["formatting"]][["Table_Order"]] %in% def_tab_order))){
            isgood = FALSE
            msgs = c(msgs, "Table_Order has the following unknown element(s):")
            msgs = c(msgs, paste0("  -> ", paste(meta[["rdocx"]][["formatting"]][["Table_Order"]][!(meta[["rdocx"]][["formatting"]][["Table_Order"]] %in% def_tab_order)], collapse=", ")))
            msgs = c(msgs, "Only the following are allowed:")
            msgs = c(msgs, paste0("  -> ", paste0(def_tab_order, collapse=", ")))
          }
          if(!all((meta[["rdocx"]][["formatting"]][["Figure_Order"]] %in% def_fig_order))){
            isgood = FALSE
            msgs = c(msgs, "Figure_Order has the following unknown element(s):")
            msgs = c(msgs, paste0("  -> ", paste(meta[["rdocx"]][["formatting"]][["Figure_Order"]][!(meta[["rdocx"]][["formatting"]][["Figure_Order"]] %in% def_fig_order)], collapse=", ")))
            msgs = c(msgs, "Only the following are allowed:")
            msgs = c(msgs, paste0("  -> ", paste0(def_fig_order, collapse=", ")))
          }
        }
      }
    }

  if(!isgood){
    msgs = c(msgs, "onbrand::read_template()")
  }
  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  if(!isgood){
    stop("Unable to read the template.")
  }

  res = list(isgood        = isgood,
             rpt           = rpt,
             rpttype       = rpttype,
             rptext        = rptext,
             rptobj        = rptobj,
             key_table     = NULL,
             placeholders  = list(),
             mapping       = mapping,
             msgs          = msgs,
             meta          = meta)
res}
