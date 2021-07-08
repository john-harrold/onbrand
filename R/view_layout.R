#'@importFrom ggplot2 is.ggplot
#'@importFrom flextable add_header add_footer align as_chunk as_paragraph autofit body_add_flextable delete_part merge_h
#'@importFrom flextable regulartable set_header_labels theme_alafoli theme_box theme_tron_legacy
#'@importFrom flextable theme_vanilla theme_booktabs theme_tron theme_vader theme_zebra
#'@importFrom magrittr "%>%"
#'@importFrom officer add_slide annotate_base body_add_break block_caption body_add_caption body_add_fpar body_add_par body_add_gg body_add_img
#'@importFrom officer body_add_table body_add_toc body_bookmark body_end_section_continuous
#'@importFrom officer body_end_section_landscape body_end_section_portrait body_replace_all_text external_img
#'@importFrom officer footers_replace_all_text fpar ftext headers_replace_all_text layout_properties layout_summary ph_location_type
#'@importFrom officer ph_location_label ph_with read_pptx read_docx shortcuts  slip_in_seqfield slip_in_text
#'@importFrom officer run_autonum styles_info unordered_list
#'@importFrom stringr str_locate_all
#'@importFrom dplyr filter
#'@importFrom yaml read_yaml
#'@importFrom rlang .data

#'@export
#'@title Generate Annotated Layout for Report Templates
#'@description Elements of slide masters are identified by placeholder labels.
#' As PowerPoint masters are created the labels
#' can be difficult to predict. Word documents are identified by style names.
#' This function will create a layout file identifying all of the elements of
#' each slide master for a PowerPoint template or each paragraph and table
#' style for a Word template.
#'
#'@param template Name of PowerPoint or Word file to annotate (defaults to included PoerPoint template)
#'@param output_file name of file to place the annotated layout information, set to \code{NULL} and it will generate a file named layout with the appropriate extension
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'
#'@return List with the following elements
#' \itemize{
#' \item{isgood}: Boolean variable indicating success or failure
#' \item{rpt}: Officer with the annotated layout
#' \item{msgs}: Vector of messages
#'}
#'@examples
#'lpptx = view_layout(
#'      template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'      output_file   = file.path(tempdir(), "layout.pptx"))
#'
#'ldocx = view_layout(
#'      template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
#'      output_file   = file.path(tempdir(), "layout.docx"))
view_layout = function(template    = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                       output_file = NULL,
                       verbose     = TRUE){
  rpt    = NULL

  fr = fetch_rpttype(template=template, verbose=verbose)
  isgood  = fr[["isgood"]]
  msgs    = fr[["msgs"]]
  rpttype = fr[["rpttype"]]


  # Making sure the output_file is defined
  if(isgood){
    if(is.null(output_file)){
      if(rpttype  == "PowerPoint"){
        output_file = "layout.pptx"
      } else if (rpttype  == "Word"){
        output_file = "layout.docx"}
    }
  }

  if(isgood){

    # Dumping PowerPoint layout
    if(rpttype  == "PowerPoint"){
      # Flag for detecting placeholder repeats
      ph_repeats = FALSE

      # Getting the annotated report
      #rpt = officer::annotate_base(path=template, output_file=NULL)
      rpt <- read_pptx(path=template)

      # Removing any slides that are present in the file
      while(length(rpt)>0){
        rpt <- officer::remove_slide(rpt, 1)
      }

      # Pulling out all of the layouts stored in the template
      lay_sum <- layout_summary(rpt)

      # Looping through each layout
      for(lidx in seq_len(nrow(lay_sum))){
        # Pulling out the layout properties
        layout <- lay_sum[lidx, 1]
        master <- lay_sum[lidx, 2]
        lp <- layout_properties ( x = rpt, layout = layout, master = master)

        # Adding a slide for the current layout
        rpt   <- add_slide(x=rpt, layout = layout, master = master)
        size  <- officer::slide_size(rpt)
        fpar_ <- officer::fpar(sprintf('layout ="%s", master = "%s"', layout, master),
                      fp_t = officer::fp_text(color = "orange", font.size = 20),
                      fp_p = officer::fp_par(text.align = "right", padding = 5)
        )
        rpt <- ph_with(x = rpt, value = fpar_, ph_label = "layout_ph",
                       location = officer::ph_location(left = 0, top = -0.5, width = size$width, height = 1,
                                              bg = "transparent", newlabel = "layout_ph"))

        # Blank slides have nothing
        if(length(lp[,1] > 0)){
          # Now we go through each placholder
          for(pidx in seq_len(nrow(lp))){
            textstr <- paste("type=", lp$type[pidx], ", index=", lp$id[pidx], ", ph_label=",lp$ph_label[pidx])

            if(nrow(lp[lp$ph_label == lp$ph_label[pidx],]) == 1){
              rpt <- officer::ph_with(x=rpt,  value = textstr, location=officer::ph_location_label(ph_label=lp$ph_label[pidx]))
            } else {
              ph_repeats = TRUE
              msgs       = c(msgs, paste0("In layout >", layout, "<, the placeholder >",lp$ph_label[pidx],"< is used more than once."))
            }
          }
        }
      }

      # If placehoder repeats were detected we adda  general message
      if(ph_repeats){
         msgs  = c(msgs, paste0("In one or more slides a placeholder name was repeated."))
         msgs  = c(msgs, paste0("This placeholder will be unavailable for reporting."))
      }
    }

    # Dumping Word layout
    if(rpttype  == "Word"){
      rpt = officer::read_docx(template)

      tab_example = data.frame( Number = c(1,2,3,4),
                                Text   = "Here")

      list_exmaple = c(1, "Top level",
                       1, "Also top level",
                       2, "first sub bullet",
                       2, "second sub bullet",
                       1, "Third top level")

      # Pulling out the different styles
      lay_sum = officer::styles_info(rpt)
      disp_styles = c("paragraph", "character", "table")

      for(style_type in disp_styles){
        rpt = officer::body_add_par(x=rpt, value="")
        rpt = officer::body_add_par(x=rpt, value=paste("STYLES: ", style_type))

        tmp_lay_sum =  lay_sum[lay_sum$style_type == style_type, ]
        for(lidx in 1:length(tmp_lay_sum[,1])){
          #style_type   = tmp_lay_sum[lidx, ]$style_type
          style_id     = tmp_lay_sum[lidx, ]$style_id
          style_name   = tmp_lay_sum[lidx, ]$style_name

          # Paragraph styles
          if(style_type %in% c("paragraph", "charcter")){
            rpt = officer::body_add_par(x=rpt, value=paste("style_name: ", style_name), style=style_name)
          }

          # Table styles
          if(style_type %in% c("table")){
            rpt = officer::body_add_par(x=rpt, value=paste("style_name: ", style_name))
            rpt = officer::body_add_table(x=rpt, value=tab_example, style = style_name)
          }
        }
      }
    }
    print(rpt, output_file)
    msgs = c(msgs, "--------------------------------")
    msgs = c(msgs, "Generating annotated layout for a report template")
    msgs = c(msgs, paste0("Template:         ", template))
    msgs = c(msgs, paste0("Annotated layout: ", output_file))
    msgs = c(msgs, "--------------------------------")
    }


  if(!isgood){
    msgs = c(msgs, "view_layout()")
    msgs = c(msgs, "Layout not generated.")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }
  res = list(isgood = isgood,
             rpt    = rpt,
             msgs   = msgs)

  return(res)}
