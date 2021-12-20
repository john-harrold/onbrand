#'@export
#'@title Add Content to Body of a Word Document Report
#'@description Appends content to the body of a Word document
#'
#'@param obnd           onbrand report object
#'@param type           Type of content to add
#'@param content        Content to add
#'@param fig_start_at   Indicates that you want to restart figure numbering at
#'the specified value (e.g. 1) after adding this content a value of \code{NULL} (default)
#'will ignore this option.
#'@param tab_start_at   Indicates that you want to restart figure numbering at
#'the specified value (e.g. 1) after adding this content a value of \code{NULL} (default)
#'will ignore this option.
#'@param verbose        Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@details
#' For each \code{content} types listed below the different \code{content} outlined is expected. Text
#' can be specified in different formats: \code{"text"} indicates plain text,
#' \code{"fpar"} is formatted text defined by the \code{fpar} command from the
#' \code{officer} package, \code{"ftext"} is a list of formatted text defined
#' by the \code{ftext} command, and \code{"md"} is text formatted in markdown
#' format (\code{?md_to_officer} for markdown details).
#'
#' \itemize{
#'  \item \code{"break"} page break, content is (\code{NULL}) and a page break will be inserted here
#'  \item \code{"ph"} adds placeholder text substitution
#'   \itemize{
#'      \item \code{"name"} placeholder name (value in body of text surrounded by three equal signs, e.g. if you have \code{"===MYPH==="}. in the document the name is just \code{"MYPH"})
#'      \item \code{"value"} value to be substituted into the placeholder (\code{"my text"})
#'      \item \code{"location"} document location where the placeholder will be located (either \code{"header"}, \code{"footer"}, or \code{"body"})
#'    }
#'  \item \code{"toc"} generates the table of contents, and content is a list
#'  that must contain __one__ of the following. 
#'   \itemize{
#'      \item \code{"level"} number indicating the depth of the contents to display (default: \code{3})
#'      \item \code{"style"} string containing the onbrand style name to use to build the TOC
#'    }
#'  \item \code{"section"} formats the current document section
#'   \itemize{
#'      \item \code{"section_type"} type of section to apply, either \code{"columns"},  \code{"continuous"}, \code{"landscape"}, \code{"portrait"}, \code{"columns"},  or  \code{"columns_landscape"}
#'      \item \code{"width"}        override the default page width with this value in inches (\code{NULL})
#'      \item \code{"height"}       override the default page height with this value in inches (\code{NULL})
#'      \item \code{"widths"}       column widths in inches, number of columns set by number of values (\code{NULL})
#'      \item \code{"space"}        space in inches between columns (\code{NULL})
#'      \item \code{"sep"}          Boolean value controlling line separating columns (\code{FALSE})
#'    }
#'  \item \code{"text"} content is a list containing a paragraph of text with the following elements
#'   \itemize{
#'      \item \code{"text"} string containing the text content either a string or the output of \code{"fpar"} for formatted text.
#'      \item \code{"style"} string containing the style to use (default \code{NULL} will use the \code{doc_def}, \code{Text} style)
#'      \item \code{"format"} string containing the format, either \code{"text"}, \code{"fpar"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'    }
#'  \item \code{"imagefile"} content is a list containing information describing an image file with the following elements
#'   \itemize{
#'      \item \code{image} string containing path to image file
#'      \item \code{caption} caption of the image (\code{NULL})
#'      \item \code{caption_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{notes} notes to add under the image  (\code{NULL})
#'      \item \code{notes_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{key} unique key for cross referencing e.g. "FIG_DATA" (\code{NULL})
#'      \item \code{height} height of the image (\code{NULL})
#'      \item \code{width} width of the image (\code{NULL})
#'    }
#'  \item \code{"ggplot"} content is a list containing a ggplot object, (eg. p = ggplot() + ....) with the following elements
#'   \itemize{
#'      \item \code{image} ggplot object
#'      \item \code{caption} caption of the image (\code{NULL})
#'      \item \code{caption_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{notes} notes to add under the image  (\code{NULL})
#'      \item \code{notes_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{key} unique key for cross referencing e.g. "FIG_DATA" (\code{NULL})
#'      \item \code{height} height of the image (\code{NULL})
#'      \item \code{width} width of the image (\code{NULL})
#'    }
#'  \item \code{"table"} content is a list containing the table content and other options with the following elements:
#'   \itemize{
#'      \item \code{table} data frame containing the tabular data
#'      \item \code{"style"} string containing the style to use (default \code{NULL} will use the \code{doc_def}, \code{Table} style)
#'      \item \code{caption} caption of the table (\code{NULL})
#'      \item \code{caption_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{notes} notes to add under the image  (\code{NULL})
#'      \item \code{notes_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{key} unique key for cross referencing e.g. "TAB_DATA" (\code{NULL})
#'      \item \code{header} Boolean variable to control displaying the header (\code{TRUE})
#'      \item \code{first_row} Boolean variable to indicate that the first row contains header information (\code{TRUE})
#'    }
#'  \item \code{"flextable"} content is a list containing flextable content and other options with the following elements (defaults in parenthesis):
#'   \itemize{
#'      \item \code{table} data frame containing the tabular data
#'      \item \code{caption} caption of the table (\code{NULL})
#'      \item \code{caption_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{notes} notes to add under the image  (\code{NULL})
#'      \item \code{notes_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{key} unique key for cross referencing e.g. "TAB_DATA" (\code{NULL})
#'      \item \code{header_top}, \code{header_middle}, \code{header_bottom} (\code{NULL}) a list with the same names as the data frame names containing the tabular data and values with the header text to show in the table
#'      \item \code{header_format} string containing the format, either \code{"text"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{merge_header} (\code{TRUE}) Set to true to combine column headers with the same information
#'      \item \code{table_body_alignment}, table_header_alignment ("center") Controls alignment
#'      \item \code{table_autofit} (\code{TRUE}) Automatically fit content, or specify the cell width and height with \code{cwidth} (\code{0.75}) and \code{cheight} (\code{0.25})
#'      \item \code{table_theme} (\code{"theme_vanilla"}) Table theme
#'    }
#'  \item \code{"flextable_object"} content is a list specifying the a user defined flextable object with the following elements:
#'   \itemize{
#'      \item \code{ft} flextable object
#'      \item \code{caption} caption of the table (\code{NULL})
#'      \item \code{caption_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{notes} notes to add under the image  (\code{NULL})
#'      \item \code{notes_format} string containing the format, either \code{"text"}, \code{"ftext"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{key} unique key for cross referencing e.g. "TAB_DATA" (\code{NULL})
#'    }
#'}
#'@return onbrand object with the content added to the body or isgood set
#'to FALSE with any messages in the msgs field. The isgood value is a Boolean variable
#'indicating the current state of the object.
#'@examples
#'
#'# Read  Word template into an onbrand object
#'obnd = read_template(
#'  template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'  mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'# The examples below use the following packages
#'library(ggplot2)
#'library(flextable)
#'library(officer)
#'
#'# Adding text
#' obnd = report_add_doc_content(obnd,
#'   type     = "text",
#'   content  = list(text="Text with no style specified will use the doc_def text format."))
#'
#'
#'# Text formatted with fpar
#'fpartext = fpar(
#'ftext("Formatted text can be created using the ", prop=NULL),
#'ftext("fpar ", prop=fp_text(color="green")),
#'ftext("command from the officer package.", prop=NULL))
#'
#'obnd = report_add_doc_content(obnd,
#'  type     = "text",
#'  content  = list(text   = fpartext,
#'                  format = "fpar",
#'                  style  = "Normal"))
#'
#'# Text formatted with markdown
#'mdtext = "Formatted text can be created using
#'**<color:green>markdown</color>** formatting"
#'obnd = report_add_doc_content(obnd,
#'  type     = "text",
#'  content  = list(text   = mdtext,
#'                  format = "md",
#'                  style  = "Normal"))
#'
#'
#'# Adding figures
#'p = ggplot() + annotate("text", x=0, y=0, label = "picture example")
#'imgfile = tempfile(pattern="image", fileext=".png")
#'ggsave(filename=imgfile, plot=p, height=5.15, width=9, units="in")
#'
#'# From an image file:
#'obnd = report_add_doc_content(obnd,
#'type     = "imagefile",
#'content  = list(image   = imgfile,
#'                caption = "This is an example of an image from a file."))
#'
#'# From a ggplot object
#'obnd = report_add_doc_content(obnd,
#'type     = "imagefile",
#'content  = list(image   = imgfile,
#'                caption = "This is an example of an image from a file."))
#'
#'
#'#Adding tables
#'tdf =    data.frame(Parameters = c("Length", "Width", "Height"),
#'                 Values     = 1:3,
#'                 Units      = c("m", "m", "m") )
#'
#'# Word table
#'tab_cont = list(table   = tdf,
#'                caption = "Word Table.")
#'obnd = report_add_doc_content(obnd,
#'  type     = "table",
#'  content  = tab_cont)
#'
#'# onbrand flextable abstraction:
#'tab_cont = list(table   = tdf,
#'                caption = "Word Table.")
#'obnd = report_add_doc_content(obnd,
#'type     = "table",
#'content  = tab_cont)
#'
#'# flextable object
#'tab_fto = flextable(tdf)
#'obnd = report_add_doc_content(obnd,
#'  type     = "flextable_object",
#'  content  = list(ft=tab_fto,
#'                  caption  = "Flextable object created by the user."))
#'
#'# Saving the report output
#'save_report(obnd, tempfile(fileext = ".docx"))
#'
report_add_doc_content = function(obnd,  
                                  type         = NULL, 
                                  content      = NULL, 
                                  fig_start_at = NULL,
                                  tab_start_at = NULL,
                                  verbose=TRUE){

  isgood        = TRUE
  ref_key       = NULL
  msgs          = c()
  style         = NULL
  caption       = NULL
  adornment_order = NULL

  # Boolean variable used to indicate if the reference key (used with figures
  # and tables) has been used before. If not normal captioning will be used.
  # Otherwise the previous figure or table number will be recycled.
  ref_key_used = FALSE

  if(type == "break"){
   content = list()
  }

  #-----
  # This is a set of basic tests on user input
  if(obnd[["isgood"]]){
    if(obnd[["rpttype"]] == "Word"){
      # If the onbrand object is good and we have a Word document we check the
      # userspecified style information against the

      # pulling out the template styles specified in the mapping file
      template_styles = obnd[["meta"]][["rdocx"]][["styles"]]

      # Capturing figure and table number restarts
      if(!is.null(tab_start_at)){
         obnd[["meta"]][["rdocx"]][["start_at"]][["tab_num"]] = tab_start_at
      }
      if(!is.null(fig_start_at)){
         obnd[["meta"]][["rdocx"]][["start_at"]][["fig_num"]] = fig_start_at
      }
    } else {

      # This funciton only works with Word reports
      isgood = FALSE
      msgs = c(msgs, paste0("The povided onbrand object is for rpttype >", obnd[["rpttype"]], "<"))
      msgs = c(msgs, "and should be Word")
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }

  #-----
  # If the basic tests are passed we start digging into the user specified
  # content:
  if(isgood){
    # All allowed content types are:
    allowed_types = c("break", "text", "toc", "section",  "imagefile", "ggplot", "ph", "table", "flextable", "flextable_object")

    # Making sure both the content and content type were defined
    if(is.null(content) | is.null(type)){
      isgood = FALSE
      msgs = c(msgs, "Either the content or the type was not specified")
    } else {
      # Sorting out the style information:
      # These two types can contain style informaiton
      if(type %in% c("text", "table")){
        docx_text_types  = c("paragraph", "character")
        docx_table_types = c("table")

        # This figures out the style to use:
        # Either the defaults if none was specified
        if(is.null(content[["style"]])){
          if(type %in% c("text")){
            style = obnd[["meta"]][["rdocx"]][["doc_def"]][["Text"]]
          } else if(type == "table"){
            style = obnd[["meta"]][["rdocx"]][["doc_def"]][["Table"]]
          }
        } else {
          style = content[["style"]]
        }

        # We want to check to make sure that the user-specified style is defined
        if((style %in% names(template_styles))){
          # Now we make sure the style is appropriate for the content type
          # Document layout summary
          lay_sum = officer::styles_info(obnd[["rpt"]])
          # Style name used in the doc template:
          docx_style = template_styles[[style]]
          # # type of style using Word docx terminology
          docx_type   = dplyr::filter(lay_sum, .data[["style_name"]] == docx_style)[["style_type"]]

          # Now we compare the docx_type against those allowed for the current
          # content type:
          if(!((type %in% c("text", "toc")  & docx_type %in% docx_text_types) |
               (type == "table"             & docx_type %in% docx_table_types))  ){
              isgood = FALSE
              msgs = c(msgs, "The supplied onbrand style cannot be applied to this content type.")
              msgs = c(msgs, paste0("  type:               ", type))
              msgs = c(msgs, paste0("  onbrand style:      ", style))
              msgs = c(msgs, paste0("  docx style:         ", docx_style))
              msgs = c(msgs, paste0("  docx type:          ", docx_type))

              if(type == "text"){
                allowed_docx_types = docx_text_types
              } else if(type =="table"){
                allowed_docx_types = docx_table_types
              }
              msgs = c(msgs, paste0("  allowed docx types: ", paste(allowed_docx_types, collapse=", ")))
          }
        } else {
          isgood = FALSE
          msgs = c(msgs, paste0("The user specified style >", style, "< was not found in the template mapping file."))
          msgs = c(msgs, paste0("The available styles are:"))
          msgs = c(msgs, paste0("  ", paste(names(template_styles), collapse=", ")))

        }
      }
      # Checking the content type
      if(!(type %in% allowed_types)){
        msgs = c(msgs, paste("The content type >", type, "< is not supported",sep=""))
        isgood = FALSE
      } else{

        # Here we perform tests against content based on the content type:
        if(type == "toc"){
          # Making sure the style exists
          if("style" %in% names(content)){
            if(!(content[["style"]] %in% names(template_styles))){
              isgood = FALSE
              msgs = c(msgs, paste0("The user specified style >", style, "< was not found in the template mapping file."))
              msgs = c(msgs, paste0("The available styles are:"))
              msgs = c(msgs, paste0("  ", paste(names(template_styles), collapse=", ")))
            }
          }
        }
        if(type == "text"){
          # Placeholder in case any other tests need to be run, but most of
          # the tests necessary for text are up with the docx_type vs
          # content type comparisons above.
        }
        if(type == "section"){
          allowed_section_types =  c("columns", "continuous",  "landscape", "portrait", "columns", "columns_landscape")
          # checking the section type
          if("section_type" %in% names(content)){
            if(content[["section_type"]] %in% allowed_section_types){
              # Placeholder in case any other tests need to be run for sections.
            } else {
              isgood = FALSE
              msgs = c(msgs, paste0("The section type >", content[["section_type"]], "< was not recognized."))
              msgs = c(msgs, "Allowed section types include:")
              msgs = c(msgs, paste0("  ",  paste(allowed_section_types, collapse=", ")))
            }
          } else {
            isgood = FALSE
            msgs = c(msgs, "When adding a section, a section type must be specified.")
            msgs = c(msgs, "Allowed section types include:")
            msgs = c(msgs, paste0("  ",  paste(allowed_section_types, collapse=", ")))
          }
        }
        if(type == "ph"){
          # allowed placeholder locations:
          ph_locations = c("body", "header", "footer")

          if(all(c("name", "value", "location") %in% names(content))){
            if(!content[["location"]] %in% ph_locations){
              msgs = c(msgs, paste0("The specified location >", content[["location"]], "< is invalid."))
              msgs = c(msgs, paste0("The location should be one of", paste(ph_locations, collapse = ", ")))

            }
          } else {
            isgood = FALSE
            msgs = c(msgs, "Placeholders requires you to specify the name, value and location")
          }
        }
        if(type == "imagefile"){
          if(!file.exists(content[["image"]])){
            msgs = c(msgs, paste("The imagefile >", content[["image"]], "< does not exist", sep=""))
            isgood = FALSE
          }
        }
        if(type == "ggplot"){
          if(!ggplot2::is.ggplot(content[["image"]])){
            msgs = c(msgs, paste("The image data found in >content$image< is not a ggplot object",sep=""))
            isgood = FALSE
          }
        }
        if(type %in% c("table", "flextable")){
          if(!is.data.frame(content[["table"]])){
            msgs = c(msgs, paste("The tabular information found in >content$table< is not a data.frame object",sep=""))
            isgood = FALSE
          }
        }
        if(type == "flextable_object"){
          # Making sure the caption defaults to NULL if it's not defined
          if(!("caption" %in% names(content))){
            content[["caption"]] = NULL
          }
          if(!("ft" %in% names(content))){
            msgs = c(msgs, paste("the flextable object >content$ft< was not found",sep=""))
            isgood = FALSE
          }
        }
      }
      # Checking reference keys. These only make sense if there is a caption
      # otherwise there is no number to reference:
      if("key" %in% names(content) & !is.null(content[["caption"]])){
        ref_key = content[["key"]]
        # If the reference key is in the key table we set
        # the ref_key_used to true
        if(ref_key %in% obnd[["key_table"]][["internal_key"]]){
          ref_key_used = TRUE
        } else {
          # Otherwise we the reference to the key table
          obnd[["key_table"]] =
            rbind(
              obnd[["key_table"]],
              data.frame(user_key     = content[["key"]],
                         internal_key = ref_key,
                         seq_text     = paste0(' REF ',  ref_key, ' \\h '),
                         ref_text     = paste0("<REF:", content[["key"]], ">")))
        }
      }
    }
  }

 # If all the checks have passed we add the content
 if(isgood){
   Caption_Format  = "text"
   Notes_Format    = "text"
   if("caption_format" %in% names(content)){
     Caption_Format = content[["caption_format"]]
   }
   if("notes_format" %in% names(content)){
     Notes_Format = content[["notes_format"]]
   }

   Notes_Style          = obnd[["meta"]][["rdocx"]][["doc_def"]][["Notes"]]
   Notes_Style_docx     = obnd[["meta"]][["rdocx"]][["styles"]][[Notes_Style]]
   Notes_Style_md_def   = fetch_md_def(obnd, Notes_Style)$md_def

   #------
   # Figure options
   if(type == "ggplot" | type == "imagefile"){
     Figure_Width =  obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Width"]]
     if("width" %in% names(content)){
       Figure_Width = content[["width"]]
     }

     Figure_Height =  obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Height"]]
     if("height" %in% names(content)){
       Figure_Height = content[["height"]]
     }
     Caption_Style        = obnd[["meta"]][["rdocx"]][["doc_def"]][["Figure_Caption"]]
     Caption_Label_Pre    = obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Caption_Label_Pre"]]
     Caption_Label_Post   = obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Caption_Label_Post"]]
     Caption_Style_docx   = obnd[["meta"]][["rdocx"]][["styles"]][[Caption_Style]]
     Caption_Seq_Id       = obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Seq_Id"]]
     Caption_Number       = obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Number"]]
     Caption_Style_md_def = fetch_md_def(obnd, Caption_Style)$md_def

     # Getting the order of the figure/table, caption, notes 
     adornment_order = obnd[["meta"]][["rdocx"]][["formatting"]][["Figure_Order"]]

     if(is.null(obnd[["meta"]][["rdocx"]][["start_at"]][["fig_num"]])){
       Caption_Start_At = NULL
     } else {
       Caption_Start_At     = obnd[["meta"]][["rdocx"]][["start_at"]][["fig_num"]]
       obnd[["meta"]][["rdocx"]][["start_at"]][["fig_num"]] = NULL
     }
   }

   #-------
   # Table options
   if(type == "table" | type == "flextable" | type=="flextable_object"){
     Caption_Style        = obnd[["meta"]][["rdocx"]][["doc_def"]][["Table_Caption"]]
     Caption_Label_Pre    = obnd[["meta"]][["rdocx"]][["formatting"]][["Table_Caption_Label_Pre"]]
     Caption_Label_Post   = obnd[["meta"]][["rdocx"]][["formatting"]][["Table_Caption_Label_Post"]]
     Caption_Seq_Id       = obnd[["meta"]][["rdocx"]][["formatting"]][["Table_Seq_Id"]]
     Caption_Number       = obnd[["meta"]][["rdocx"]][["formatting"]][["Table_Number"]]
     Caption_Style_docx   = obnd[["meta"]][["rdocx"]][["styles"]][[Caption_Style]]
     Caption_Style_md_def = fetch_md_def(obnd, Caption_Style)$md_def

     # Getting the order of the figure/table, caption, notes 
     adornment_order = obnd[["meta"]][["rdocx"]][["formatting"]][["Table_Order"]]

     if(is.null(obnd[["meta"]][["rdocx"]][["start_at"]][["tab_num"]])){
       Caption_Start_At = NULL
     } else {
       Caption_Start_At     = obnd[["meta"]][["rdocx"]][["start_at"]][["tab_num"]]
       obnd[["meta"]][["rdocx"]][["start_at"]][["tab_num"]] = NULL
     }
   }

   if(type == "table"){
     header    = TRUE
     if('header' %in% names(content)){
       header = content[["header"]]
     }
     first_row = TRUE
     if('first_row' %in% names(content)){
       first_row = content[["first_row"]]
     }
   }
   #-------
   if(type == "flextable"){
     # These are the default flextable options
     # and they can be over written by specifying
     # the same fields of the content list
     header_top              = NULL
     header_middle           = NULL
     header_bottom           = NULL
     header_format           = NULL
     merge_header            = TRUE
     table_body_alignment    ="center"
     table_header_alignment  ="center"
     table_autofit           = TRUE
     table_theme             ="theme_vanilla"
     cwidth                  = 0.75
     cheight                 = 0.25

     ftops = c("header_top",             "header_middle",    "header_bottom",
               "header_format",
               "merge_header",           "table_theme",      "table_body_alignment",
               "table_header_alignment", "table_autofit",    "cwidth",
               "cheight")


     # Defining the user specified flextable options:
     for(ftop in ftops){
       if(!is.null(content[[ftop]])){
         eval(parse(text=sprintf('%s = content[[ftop]]', ftop)))
       }
     }

     # Creating the table
     ft = flextable::regulartable(content[["table"]],  cwidth = cwidth, cheight=cheight)
     #-------
     # Defining the default formatting for tables. We default to NULL and
     # if it's currently defined in the meta for this report template then
     # we use that

     # determining the header depth:
     if(!is.null(names(header_bottom))){
       num_headers = 3
     }else if(!is.null(names(header_middle))){
       num_headers = 2
     }else if(!is.null(names(header_top))){
       num_headers = 1
     } else {
       num_headers = 0
     }
     #-------
     # Processing user defined headers. These will get stuck in header_list
     header_list    = list()
     if(num_headers > 0){
       for(cname in names(content[["table"]])){
         # creating an empty header by default
         header_list[[cname]] = rep("", times=num_headers)
         if(cname %in% names(header_top)){
            header_list[[cname]][1] = header_top[[cname]] }
         if(cname %in% names(header_middle)){
            header_list[[cname]][2] = header_middle[[cname]] }
         if(cname %in% names(header_bottom)){
            header_list[[cname]][3] = header_bottom[[cname]] }
       }
       ft = flextable::delete_part(ft, part   = "header")
       ft = flextable::add_header(ft,  values = header_list)

     }
     #-------
     # Processing markdown
     if(!is.null(header_format)){
       if(header_format == "md"){

         # Pulling out the default format for the Table element
         default_format_table =  obnd[["meta"]][["rdocx"]][["md_def"]][["Table_Labels"]]

         for(cname in names(content[["table"]])){
           # For each column name we run the header text through the markdown
           # conversion to produce the as_paragraph output:
           ft = flextable::compose(ft,
                             j     = cname,
                             part  = "header",
                             value = md_to_oo(strs= header_list[[cname]], default_format=default_format_table)$oo)
         }
       }
     }
     #-------


     # Setting the theme
     if(!is.null(table_theme)){
       eval(parse(text=paste("ft = flextable::", table_theme, "(ft)", sep=""))) }

     # Merging headers
     if(merge_header){
       ft = flextable::merge_h(ft, part="header") }

     if(table_autofit){
       ft = flextable::autofit(ft) }

     # Applying the aligment
     ft = flextable::align(ft, align=table_header_alignment, part="header")
     ft = flextable::align(ft, align=table_body_alignment,   part="body"  )


   }
   #-------
   if(type == "flextable_object"){
     ft = content[["ft"]]
   }

   # For figures and tables they can have adornments (captions and notes)
   # attached to them and the order can be arbitrary. So we loop through the 
   adornment_types = c("flextable", "flextable_object",
                       "ggplot",    "imagefile", "table") 
   if(type %in% adornment_types){
     for(adornment_ele in adornment_order){
       #-------
       # Adding the image/table
       if(type == "ggplot" & adornment_ele == "figure"){
         obnd[["rpt"]] = officer::body_add_gg(obnd[["rpt"]], value=content[["image"]], width = Figure_Width, height = Figure_Height)
       }
       if(type == "imagefile" & adornment_ele == "figure"){
         obnd[["rpt"]] = officer::body_add_img(obnd[["rpt"]], src=content[["image"]], width = Figure_Width, height = Figure_Height)
       }
       
       if(type == "table" & adornment_ele == "table"){
        obnd[["rpt"]] = officer::body_add_table(obnd[["rpt"]], value=content[["table"]], header=header, first_row=first_row, docx_style)
       }
       
       if((type == "flextable" | type=="flextable_object") & adornment_ele == "table"){
         obnd[["rpt"]] = flextable::body_add_flextable(x = obnd[["rpt"]], value = ft)
       }
       
       #------
       if(!is.null(content[["notes"]]) & adornment_ele == "notes"){
         if(Notes_Format == "text"){
           obnd[["rpt"]] = officer::body_add_par(obnd[["rpt"]], content[["notes"]], style=Notes_Style)
         } else if(Notes_Format == "ftext"){
           fpargs = paste(paste('content[["notes"]][[', 1:length(content[["notes"]]), ']]'), collapse=",")
           fp_cmd  = paste0("officer::fpar(",fpargs,")")
           fp_res = eval(parse(text=fp_cmd))
           obnd[["rpt"]] = officer::body_add_fpar(obnd[["rpt"]], fp_res, style=Notes_Style)
         } else if(Notes_Format == "md"){
           mdout = md_to_officer(content[["notes"]], default_format=Notes_Style_md_def)
           ftall = c()
           for(pgraph in mdout){
             ftall = c(ftall, pgraph[["ftext_cmd"]], 'officer::ftext(" ")')
           }
           fp_cmd  = paste0("officer::fpar(",paste(ftall, collapse = ", "), ")")
           fp_notes = eval(parse(text=fp_cmd))
       
           obnd[["rpt"]] = officer::body_add_fpar(obnd[["rpt"]], fp_notes,  style=Notes_Style_docx)
         }
       }
       #------
       # Caption
       if(!is.null(content[["caption"]]) & adornment_ele  == "caption"){
         # Creating the caption block based on the caption format:
         # this will define the object 'caption' that can be inserted
         # below based on whether it should be below or above the table/figure
         # JMH abstract this part out to yaml file

         if(ref_key_used){
           caption_runs   = NULL
           caption_rb_res = NULL
           run_num        = NULL
         } else {
           caption_runs   = eval(parse(text=Caption_Number))
           caption_rb_res = officer::run_bookmark(ref_key, caption_runs)
           run_num = officer::run_autonum(seq_id     = "fig",
                                          pre_label  = Caption_Label_Pre,
                                          post_label = Caption_Label_Post,
                                          bkm        = ref_key)
         }
         
         if(Caption_Format == "text"){
           if(ref_key_used){
             caption = officer::fpar(Caption_Label_Pre,
                                     officer::run_reference(ref_key),
                                     Caption_Label_Post,
                                     content[["caption"]])
           } else {
             caption = officer::fpar(Caption_Label_Pre,
                                     caption_rb_res,
                                     Caption_Label_Post,
                                     content[["caption"]])
           }
         } else if(Caption_Format == "ftext"){
           # Prepending the table/figure label and number
           if(ref_key_used){
             fpargs = paste("Caption_Label_Pre,", 
                            "officer::run_reference(ref_key),",
                            "Caption_Label_Post,")
           } else {
             #fpargs = paste("autonum=run_num,", fpargs)
             fpargs = paste("Caption_Label_Pre,",
                            "caption_rb_res,",
                            "Caption_Label_Post,")
           }
           fpargs = paste(fpargs, paste(paste('content[["caption"]][[', 1:length(content[["caption"]]), ']]'), collapse=","))
           fp_cmd  = paste0("officer::fpar(",fpargs,")")
           caption = eval(parse(text=fp_cmd))
         } else if(Caption_Format == "md"){
           mdout = md_to_officer(content[["caption"]], default_format=Caption_Style_md_def)
           # Prepending the table/figure label and number
           if(ref_key_used){
             ftall  = c("Caption_Label_Pre", 
                        "officer::run_reference(ref_key)",
                        "Caption_Label_Post")
           } else {
             ftall = c("Caption_Label_Pre",
                       "caption_rb_res",
                       "Caption_Label_Post")
           }
           for(pgraph in mdout){
             ftall = c(ftall, pgraph[["ftext_cmd"]], 'officer::ftext(" ")')
           }
           fp_cmd  = paste0("officer::fpar(",paste(ftall, collapse = ", "), ")")
           caption = eval(parse(text=fp_cmd))
         }

         # Now that the caption has been defined we add it here:
         obnd[["rpt"]] = officer::body_add_fpar(obnd[["rpt"]], caption, style=Caption_Style_docx)
       }
     }
   }
   #------
   if(type == "text"){
     # Pulling out the markdown defaults for this style:
     md_defaults = obnd[["meta"]][["rdocx"]][["md_def"]][[style]]

     # Figuring out the formatting
     if("format" %in% names(content)){
       Text_Format = content[["format"]]
     } else {
       # defaulting to text format
       Text_Format = "text"
     }
     if(Text_Format == "text"){
       obnd[["rpt"]] =  officer::body_add_par(obnd[["rpt"]], value=content[["text"]], style=docx_style)
     } else if(Text_Format == "fpar"){
       obnd[["rpt"]] = officer::body_add_fpar(obnd[["rpt"]], value=content[["text"]], style=docx_style)
     } else if(Text_Format == "md"){
       mdout = md_to_officer(str=content[["text"]], default_format = md_defaults)
       for(pgraph in mdout){
         obnd[["rpt"]] = officer::body_add_fpar(obnd[["rpt"]], value=eval(parse(text=pgraph[["fpar_cmd"]])), style=docx_style)
       }
     }
   }

   # Setting the placeholder
   if(type == "ph"){
    obnd[["placeholders"]][[content[["name"]]]] =
            list(location = content[["location"]],
                 value    = content[["value"]])
   }


   # Adding sections
   if(type == "section"){
     # Different section types allow different arguments, here we just
     # explicitly define them:
     allowed_args = list(
        continuous           = c(),
        landscape            = c("width",  "height"),
        portrait             = c("width",  "height"),
        columns              = c("widths", "space", "sep"),
        columns_landscape    = c("widths", "space", "sep", "width", "height"))

     # Creating function arguments
     fcnargs = c('x=obnd[["rpt"]]')

     if("sep" %in% allowed_args[[content[["section_type"]]]]){
       if(!is.null(content[["sep"]])){
          fcnargs = c(fcnargs, paste0("sep=",content[["sep"]])) }
     }
     if("w" %in% allowed_args[[content[["section_type"]]]]){
      if(!is.null(content[["width"]])){
         fcnargs = c(fcnargs, paste0("w=",content[["width"]])) }
     }
     if("h" %in% allowed_args[[content[["section_type"]]]]){
       if(!is.null(content[["height"]])){
          fcnargs = c(fcnargs, paste0("h=",content[["height"]])) }
     }
     if("widths" %in% allowed_args[[content[["section_type"]]]]){
       if(!is.null(content[["widths"]])){
          fcnargs = c(fcnargs, paste0("widths=c(",toString(content[["widths"]]), ")")) }
     }

     fcn = paste0("officer::body_end_section_", content[["section_type"]])

     fcncall = paste0('obnd[["rpt"]] = ', fcn, "(", paste(fcnargs, collapse = ", "), ")")

     eval(parse(text=fcncall))
   }

   # Adding the table of contents
   if(type == "toc"){
     # Default to no style and a level of 3
     TOC_Separator  = obnd[["meta"]][["rdocx"]][["formatting"]][["separator"]]
     TOC_Style_docx = NULL
     level = 3
     # here level has been provided 
     if("level" %in% names(content)){
       level = content[["level"]]
     } 

     # A style has been provided:
     if("style" %in% names(content)){
       TOC_Style       = content[["style"]]
       TOC_Style_docx  = obnd[["meta"]][["rdocx"]][["styles"]][[TOC_Style]]
     }
     obnd[["rpt"]] = officer::body_add_toc(obnd[["rpt"]], style=TOC_Style_docx, level=level, separator=TOC_Separator)
   }

   if(type == "break"){
       obnd[["rpt"]] = officer::body_add_break(obnd[["rpt"]])
   }
  }


  if(!isgood){
    msgs = c(msgs, "Unable to add content to document, see above for details")
    msgs = c(msgs, paste0("mapping file: ", obnd[["mapping"]]))
    msgs = c(msgs, "onbrand::report_add_doc_content() ")
    obnd[["isgood"]] = isgood
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  # Adding any messages to the report object
  obnd[["msgs"]] = msgs

  if(!isgood){
    stop("Unable to add content to the Word document.")
  }

obnd}
