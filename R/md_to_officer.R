#'@export
#'@title Parse Markdown for Officer
#'@description Parses text in Markdown format and returns fpar and as_paragraph command strings to be used with Officer
#'@param str     string containing Markdown can contain the following elements:
#' \itemize{
#'  \item paragraph:   two or more new lines creates a paragraph
#'  \item bold:        can be either \code{"*text in bold*"} or \code{"_text in bold_"}
#'  \item italics:     can be either \code{"**text in italics**"} or \code{"__text in italics__"}
#'  \item subscript:   \code{"Normal~subscript~"} 
#'  \item superscript: \code{"Normal^superscript^"} 
#'  \item color:       \code{"<color:red>red text</color>"} 
#'  \item shade:       \code{"<shade:#33ff33>shading</shade>"} 
#'  \item font family: \code{"<ff:symbol>symbol</ff>"} 
#'}
#'@param default_format  list containing the default format for elements not defined with markdown default values. 
#' \preformatted{
#'    default_format = list( 
#'       color          = "black",
#'       font.size      = 12,
#'       bold           = FALSE,
#'       italic         = FALSE,
#'       underlined     = FALSE,
#'       font.family    = "Cambria (Body)",
#'       vertical.align = "baseline",
#'       shading.color  = "transparent")
#' }
#'@return list with parsed paragraph elements ubiquity system object with the
#' content added to the body, each paragraph can be found in a numbered list
#' element (e.g. \code{pgraph_1}, \code{pgraph_2}, etc) each with the following
#' elements:
#' \itemize{
#'  \item \code{locs} Dataframe showing the locations of markdown elements in the current paragraph
#'  \item \code{pele} These are the individual parsed paragraph elements
#'  \item \code{fpar_cmd} String containing the fpar_cmd that can be run using
#'  \code{eval} to return the output of \code{fpar}. For example: 
#' \preformatted{
#'   myfpar = eval(parse(text=pgparse$pgraph_1$fpar_cmd))
#'  }
#'  \item \code{as_paragraph_cmd} String containing the as_paragraph_cmd that can be run using
#' \preformatted{
#'   myas_para = eval(parse(text=pgparse$pgraph_1$as_paragraph_cmd))
#'  }
#'}
md_to_officer = function(str,
     default_format = list( 
        color          = "black",
        font.size      = 12,
        bold           = FALSE,
        italic         = FALSE,
        underlined     = FALSE,
        font.family    = "Cambria (Body)",
        vertical.align = "baseline",
        shading.color  = "transparent")){



# First we find paragraphs:
pgraphs = unlist(base::strsplit(str, split="(\r\n|\r|\n){2,}"))


md_info = data.frame(
  md_name = c( "subscript",         "superscript",     "bold_us",  "bold_st",     "italic",            "color",                     "shading_color",                     "font_family"             ),
  pattern = c( "~.+?~",             "\\^.+?\\^",       "_.+?_",    "\\*.+?\\*",   "\\%\\%.+?\\%\\%",   "<color:\\S+?>.+?</color>",  "<shade:\\S+?>.+?</shade>",          "<ff:\\S+?>.+?</ff>"      ), 
  start   = c( "~",                 "\\^",             "_",        "\\*",         "\\%\\%",            "<color:\\S+?>",             "<shade:\\S+?>",                     "<ff:\\S+?>"              ),
  end     = c( "~",                 "\\^",             "_",        "\\*",         "\\%\\%",            "</color>",                  "</shade>",                          "</ff>"                   ),
  prop    = c( "vertical.align",    "vertical.align",  "bold",     "bold",        "italic",            "color",                     "shading_color",                     "font.family"             ))


pos_start = c()
pos_stop  = c()

# This is for chunks of text with no formatting. 
no_props_str = paste('officer::fp_text(bold = ', default_format[["bold"]],',',
                         'font.size = ',         default_format[["font.size"]], ',', 
                         'italic = ',            default_format[["italic"]], ',', 
                         'underlined = ',        default_format[["underlined"]], ',', 
                         'color = "',            default_format[["color"]], '",', 
                         'shading.color = "',    default_format[["shading.color"]], '",', 
                         'vertical.align = "',   default_format[["vertical.align"]], '",', 
                         'font.family = "',      default_format[["font.family"]],'")', sep = "")

no_props     = officer::fp_text(
        color          = default_format[["color"]], 
        font.size      = default_format[["font.size"]], 
        bold           = default_format[["bold"]], 
        italic         = default_format[["italic"]], 
        underlined     = default_format[["underlined"]], 
        font.family    = default_format[["font.family"]], 
        vertical.align = default_format[["vertical.align"]],
        shading.color  = default_format[["shading.color"]]) 


# Saving the parsed paragraphs
pgraphs_parse = list()

  # Now we walk through each paragraph
  pgraph_idx = 1
  for(pgraph_raw in pgraphs){

    # Removing all of the carriage returns in the paragraph:
    pgraph = gsub(pattern="(\r\n|\r|\n)", replacement=" ", pgraph_raw)
  
    # Storing the locations of the markdown in the string
    locs      = NULL

    # Visual id of md elements to debug finding stuff 
    md_visual = c()

    # Converting the ** to %% to make it easier to distinguish between bold and
    # italics
    pgraph = gsub(pgraph, pattern="\\*\\*", replacement="%%")
    pgraph = gsub(pgraph, pattern="__", replacement="%%")

    # Finding the locations of the markdown elements
    for(md_idx  in 1:nrow(md_info)){
      pattern = as.character(md_info[md_idx, ]$pattern)
      md_name = as.character(md_info[md_idx, ]$md_name)
      tmplocs = stringr::str_locate_all(pgraph, pattern)[[1]]
      if(nrow(tmplocs) > 0){
        tmplocs = as.data.frame(tmplocs)
        tmplocs$md_name = md_name
        locs = rbind(locs, tmplocs)
      }
    }

    
    # if locs is NULL then no markdown elements were found in the current 
    # current paragraph so we just raap that up
    if(is.null(locs)){
      pele     = list()
      pele$p_1 = list(text      = pgraph,
                      props     = c(no_props_str),
                      props_cmd = paste("prop=", no_props_str, sep=""))
    } else {
      # If locs isn't NULL we start working trough the markdown elements:

      # We begin by grouping nested markdown elements
      locs =locs[order(locs$start), ]
      locs$group = 1
      if(nrow(locs) > 1){
        for(loc_idx in 2:nrow(locs)){
          # If the current md element starts before the last one stops
          # they are grouped together, otherwise they become part of a new group
          if(locs[loc_idx,]$start < locs[loc_idx-1,]$end){
            locs[loc_idx,]$group = locs[loc_idx-1,]$group
          } else {
            locs[loc_idx,]$group = locs[loc_idx-1,]$group + 1
          }
        }
      }
      
      # Pulling out the separate paragraph elements
      pele     = list()
      pele_idx = 1
      # Processing each group
      for(group in unique(locs$group)){
        # pulling out the markdown elements for that group
        gr_md   = locs[locs$group == group, ]
      
        #----------
        # if we're dealing with the first group and it starts after the first
        # character then we add that first chunk of text
        if(group == 1 & gr_md[1, ]$start > 1){
          #pele_tmp = list(text = pgraph)
          pele[[paste('p_', pele_idx, sep="")]] = 
               list(text      = substr(pgraph, start=1, stop=(gr_md$start-1)),
                    props     = c(no_props_str),
                    props_cmd = paste("prop=", no_props_str, sep=""))
          pele_idx = pele_idx + 1
        }
        #----------
        # If we're dealing with a group after the first we need to pull out any
        # text between groups
        if(group > 1){
          # Previous group:
          gr_md_prev   = locs[locs$group == group-1, ]
      
          # If there is more than 1 character difference between the last md
          # element from the previous group and the first of the current group
          # then we need to add that text
          if(gr_md[1, ]$start-gr_md_prev[1, ]$end > 1){
            pele[[paste('p_', pele_idx, sep="")]] = 
                       list(text      = substr(pgraph, 
                                               start =(gr_md_prev[1, ]$end + 1),
                                               stop  =(gr_md[1, ]$start - 1)),
                            props     = c(no_props_str),
                            props_cmd = paste("prop=", no_props_str, sep=""))
            pele_idx = pele_idx + 1
          }
        }
        #----------
        # Processing the markdown for a group
        # First we pull out the text from the inner most markdown element
        md_text = substr(pgraph, 
                         start =gr_md[nrow(gr_md), ]$start,
                         stop  =gr_md[nrow(gr_md), ]$end)

        # now we strip off the beginning and ending of the markdown 
        md_name  = gr_md[nrow(gr_md), ]$md_name
      
        # patterns to strip off the beginning and end
        md_start = paste("^", as.character(md_info[md_info$md_name == md_name, ]$start), sep="")
        md_end   = paste(as.character(md_info[md_info$md_name == md_name, ]$end), "$", sep="")
      
        # Stripping those patterns off
        md_text = sub(md_text, pattern=md_start, replacement="")
        md_text = sub(md_text, pattern=md_end, replacement="")

        if(group == 4){
        }
      
        # Now we save the text:
        pele[[paste('p_', pele_idx, sep="")]] = 
                   list(text      = md_text,
                        props     = no_props,
                        props_cmd = no_props)
        
        # Making a copy of the format, these will be subtracted below as the
        # user defined formats are found:
        tmp_def_props = default_format
      
        tmp_props = c()
        # Next we add the properties associated with the markdown
        for(md_name in (gr_md$md_name)){
          md_start = as.character(md_info[md_info$md_name == md_name, ]$start)
          md_end   = as.character(md_info[md_info$md_name == md_name, ]$end)
          md_prop  = as.character(md_info[md_info$md_name == md_name, ]$prop)
      
          # Setting properties based on the type of markdown selected
          if(md_name == "bold_st" | md_name == "bold_us"){
            tmp_props = c(tmp_props, "bold = TRUE")
            # subtracting the default value
            tmp_def_props[["bold"]] = NULL
          }
      
          if(md_name == "italic"){
            tmp_props = c(tmp_props, "italic = TRUE")
            # subtracting the default value
            tmp_def_props[["italic"]] = NULL
          }
      
          if(md_name == "superscript"){
            tmp_props = c(tmp_props, 'vertical.align = "superscript"')
            # subtracting the default value
            tmp_def_props[["vertical.align"]] = NULL
          }
      
          if(md_name == "subscript"){
            tmp_props = c(tmp_props, 'vertical.align = "subscript"')
            # subtracting the default value
            tmp_def_props[["vertical.align"]] = NULL
          }
      
          if(md_name == "color"){
            # pulling out the color markdown text. It uses the first entry so
            # the outer most. There shouldn't be more than one. 
            md_text = substr(pgraph, 
                             start = gr_md[gr_md$md_name == "color", ]$start[1],
                             stop  = gr_md[gr_md$md_name == "color", ]$end[1])

            #extracting the color
            color = stringr::str_extract(md_text, md_start)
            color= gsub(color, pattern="<color:", replacement="") 
            color= gsub(color, pattern=">", replacement="") 

            tmp_props = c(tmp_props, paste('color = "', color, '"', sep=""))
            # subtracting the default value
            tmp_def_props[["color"]] = NULL
          }
          if(md_name == "shading_color"){
            # pulling out the color markdown text. It uses the first entry so
            # the outer most. There shouldn't be more than one. 
            md_text = substr(pgraph, 
                             start = gr_md[gr_md$md_name == "shading_color", ]$start[1],
                             stop  = gr_md[gr_md$md_name == "shading_color", ]$end[1])

            #extracting the color
            color = stringr::str_extract(md_text, md_start)
            color= gsub(color, pattern="<shade:", replacement="") 
            color= gsub(color, pattern=">", replacement="") 

            tmp_props = c(tmp_props, paste('shading.color = "', color, '"', sep=""))
            # subtracting the default value
            tmp_def_props[["shading.color"]] = NULL
          }

          if(md_name == "font_family"){
            md_text = substr(pgraph, 
                             start = gr_md[gr_md$md_name == "font_family", ]$start[1],
                             stop  = gr_md[gr_md$md_name == "font_family", ]$end[1])

            #extracting the font family
            ff = stringr::str_extract(md_text, md_start)
            ff = gsub(ff, pattern="<ff:", replacement="") 
            ff = gsub(ff, pattern=">", replacement="") 
            tmp_props = c(tmp_props, paste('font.family = "', ff, '"', sep=""))
            # subtracting the default value
            tmp_def_props[["font.family"]] = NULL
          }
        }

        # Now we add in the default properties that are left over
        if("color"            %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'color = "',          tmp_def_props[["color"]],         '"',  sep=""))}
        if("font.size"        %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'font.size = ',       tmp_def_props[["font.size"]],           sep=""))}
        if("bold"             %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'bold = ',            tmp_def_props[["bold"]],                sep=""))}
        if("italic"           %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'italic = ',          tmp_def_props[["talic"]],               sep=""))}
        if("underlined"       %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'underlined = ',      tmp_def_props[["underlined"]],          sep=""))}
        if("font.family"      %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'font.family = "',    tmp_def_props[["font.family"]],   '"',  sep=""))}
        if("vertical.align"   %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'vertical.align = "', tmp_def_props[["vertical.align"]], '"', sep=""))}
        if("shading.color"    %in% names(tmp_def_props)){ tmp_props = c(tmp_props, paste( 'shading.color = "',  tmp_def_props[["shading.color"]], '"',  sep=""))}
      
        pele[[paste('p_', pele_idx, sep="")]]$props     = tmp_props
        pele[[paste('p_', pele_idx, sep="")]]$props_cmd = paste("prop=officer::fp_text(", paste(tmp_props, collapse=", "), ")", sep="")
      
        pele_idx = pele_idx + 1
      
        #----------
        # If we're at the last group and it doesn't go to the end we add the
        # last part as well
        if(group == max(unique(locs$group))){
          # First we get the last piece of text:
          text_end = substr(pgraph, start=(gr_md[1, ]$end+1), stop=nchar(pgraph))
          # If that string isn't empty we add a paragraph element for it
          if(text_end != ""){
            pele[[paste('p_', pele_idx, sep="")]] = 
                       list(text      = text_end,
                            props     = c(no_props_str),
                            props_cmd = paste("prop=", no_props_str, sep=""))
            pele_idx = pele_idx + 1
          }
        }
        #----------
      }
      
      for(loc_idx in 1:nrow(locs)){
        tmpstr = paste(rep(" ", nchar(pgraph)), collapse="")
        tmpstr = paste(tmpstr, ":",  locs[loc_idx, ]$md_name, sep="")
        substr(tmpstr, locs[loc_idx, ]$start, locs[loc_idx, ]$start)  = "|"
        substr(tmpstr, locs[loc_idx, ]$end  , locs[loc_idx, ]$end  )  = "|"
        md_visual = c(md_visual, pgraph, tmpstr)
      }
    }


  fpar_cmd = ""
  for(tmpele in pele){
    if(fpar_cmd != ""){
     fpar_cmd = paste(fpar_cmd, ',\n') }
    fpar_cmd = paste(fpar_cmd, 'ftext("', tmpele$text, '", ', tmpele$props_cmd, ')', sep="")
  }
  fpar_cmd = paste("fpar(", fpar_cmd, ")", sep="")

  as_paragraph_cmd = ""
  for(tmpele in pele){
    if(as_paragraph_cmd != ""){
     as_paragraph_cmd = paste(as_paragraph_cmd, ',\n') }
    as_paragraph_cmd = paste(as_paragraph_cmd, 'flextable::as_chunk("', tmpele$text, '", ', tmpele$props_cmd, ')', sep="")
  }
  as_paragraph_cmd = paste("flextable::as_paragraph(", as_paragraph_cmd, ")", sep="")



  pgraphs_parse[[paste("pgraph_", pgraph_idx, sep="")]]$pele              = pele
  pgraphs_parse[[paste("pgraph_", pgraph_idx, sep="")]]$locs              = locs
  pgraphs_parse[[paste("pgraph_", pgraph_idx, sep="")]]$md_visual         = md_visual
  pgraphs_parse[[paste("pgraph_", pgraph_idx, sep="")]]$fpar_cmd          = fpar_cmd
  pgraphs_parse[[paste("pgraph_", pgraph_idx, sep="")]]$as_paragraph_cmd  = as_paragraph_cmd

  pgraph_idx = pgraph_idx + 1
  }
pgraphs_parse}
