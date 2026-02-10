#'@export
#'@title Parse Markdown for Officer
#'@description Parses text in Markdown format and returns fpar and as_paragraph command strings to be used with officer
#'
#'@param str     string containing Markdown can contain the following elements:
#' \itemize{
#'  \item paragraph:   two or more new lines creates a paragraph
#'  \item bold:        can be either \code{"**text in bold**"} or \code{"__text in bold__"}
#'  \item italics:     can be either \code{"*text in italics*"} or \code{"_text in italics_"}
#'  \item subscript:   \code{"Normal~subscript~"}
#'  \item superscript: \code{"Normal^superscript^"}
#'  \item color:       \code{"<color:red>red text</color>"}
#'  \item shade:       \code{"<shade:#33ff33>shading</shade>"}
#'  \item font family: \code{"<ff:symbol>symbol</ff>"}
#'  \item reference:   \code{"<ref:key>"} Where \code{"key"} is the value  assigned when adding a table or figure
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
#'@return list with parsed paragraph elements with the content added to the body,
#' each paragraph can be found in a numbered list element (e.g. \code{pgraph_1},
#' \code{pgraph_2}, etc) each with the following elements:
#' \itemize{
#'  \item{locs}: Dataframe showing the locations of markdown elements in the current paragraph
#'  \item{pele}: These are the individual parsed paragraph elements
#'  \item{ftext_cmd}: String containing the ftext commands.
#'  \item{fpar_cmd}: String containing the fpar commands that can be run using
#'  \code{eval} to return the output of \code{fpar}. For example:
#' \preformatted{
#'   myfpar = eval(parse(text=pgparse$pgraph_1$fpar_cmd))
#'  }
#'  \item{as_paragraph_cmd}: String containing the as_paragraph_cmd that can be run using
#' \preformatted{
#'   myas_para = eval(parse(text=pgparse$pgraph_1$as_paragraph_cmd))
#'  }
#'}
#'@examples
#'res              = md_to_officer("Be **bold**!")
#'fpar_obj         = eval(parse(text=res$pgraph_1$fpar_cmd))
#'as_paragraph_obj = eval(parse(text=res$pgraph_1$as_paragraph_cmd))
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

  #----------------------------
  # Splitting input into paragraphs (2+ newlines)
  pgraphs = unlist(base::strsplit(str, split="(\r\n|\r|\n){2,}"))

  #----------------------------
  # Markdown patterns (order matters: more specific first)
  md_patterns = list(
    list(name = "bold_italic", regex = "\\*\\*\\*(.+?)\\*\\*\\*",
         start = "***", end = "***", props = c("bold", "italic")),
    list(name = "bold", regex = "\\*\\*(.+?)\\*\\*",
         start = "**", end = "**", props = c("bold")),
    list(name = "bold", regex = "__(.+?)__",
         start = "__", end = "__", props = c("bold")),
    list(name = "italic_st", regex = "(?<!\\*)\\*([^*]+?)\\*(?!\\*)",
         start = "*", end = "*", props = c("italic")),
    list(name = "italic_us", regex = "_([^_]+?)_",
         start = "_", end = "_", props = c("italic")),
    list(name = "subscript", regex = "~(.+?)~",
         start = "~", end = "~", props = c("vertical.align"), value = "subscript"),
    list(name = "superscript", regex = "\\^(.+?)\\^",
         start = "^", end = "^", props = c("vertical.align"), value = "superscript"),
    list(name = "color", regex = "<color:(\\S+?)>(.+?)</color>",
         start_regex = "<color:(\\S+?)>", end_regex = "</color>", props = c("color"), extract_value = TRUE),
    list(name = "shading_color", regex = "<shade:(\\S+?)>(.+?)</shade>",
         start_regex = "<shade:(\\S+?)>", end_regex = "</shade>", props = c("shading.color"), extract_value = TRUE),
    list(name = "font_family", regex = "<ff:(\\S+?)>(.+?)</ff>",
         start_regex = "<ff:(\\S+?)>", end_regex = "</ff>", props = c("font.family"), extract_value = TRUE),
    list(name = "reference", regex = "<ref:(.+?)>",
         start = "<ref:", end = ">", is_reference = TRUE)
  )

  #----------------------------
  # Helper to build the fp_text properties string. The order of properties
  # matters for test compatibility with the original implementation.
  build_props = function(prop_list) {
    parts = c()
    if(!is.null(prop_list$bold))           parts = c(parts, paste0('bold = ', prop_list$bold))
    if(!is.null(prop_list$font.size))      parts = c(parts, paste0('font.size = ', prop_list$font.size))
    if(!is.null(prop_list$italic))         parts = c(parts, paste0('italic = ', prop_list$italic))
    if(!is.null(prop_list$underlined))     parts = c(parts, paste0('underlined = ', prop_list$underlined))
    if(!is.null(prop_list$color))          parts = c(parts, paste0('color = "', prop_list$color, '"'))
    if(!is.null(prop_list$shading.color))  parts = c(parts, paste0('shading.color = "', prop_list$shading.color, '"'))
    if(!is.null(prop_list$vertical.align)) parts = c(parts, paste0('vertical.align = "', prop_list$vertical.align, '"'))
    if(!is.null(prop_list$font.family))    parts = c(parts, paste0('font.family = "', prop_list$font.family, '"'))
    paste(parts, collapse = ",")
  }

  #----------------------------
  # Helper to recursively parse text for markdown patterns. This walks through
  # the text finding markdown elements and builds a list of segments with
  # their formatting properties.
  parse_text = function(text, current_props = default_format) {
    res = list()

    if(!is.null(text) && nchar(text) > 0) {
      # Finding all markdown matches in this text
      all_matches = list()
      for(pattern in md_patterns) {
        matches = stringr::str_locate_all(text, pattern$regex)[[1]]
        if(nrow(matches) > 0) {
          for(i in 1:nrow(matches)) {
            match_text = substr(text, matches[i,1], matches[i,2])
            all_matches[[length(all_matches) + 1]] = list(
              start   = matches[i,1],
              end     = matches[i,2],
              pattern = pattern,
              text    = match_text
            )
          }
        }
      }

      if(length(all_matches) == 0) {
        #-----
        # No matches found, this is plain text
        res = list(list(
          text      = text,
          props     = current_props,
          props_str = build_props(current_props),
          md_name   = "none"
        ))
      } else {
        #-----
        # Sorting by start position and processing the first match
        all_matches = all_matches[order(sapply(all_matches, function(m) m$start))]
        first_match = all_matches[[1]]

        # Adding text before match (if any)
        if(first_match$start > 1) {
          before_text = substr(text, 1, first_match$start - 1)
          res[[length(res) + 1]] = list(
            text      = before_text,
            props     = current_props,
            props_str = build_props(current_props),
            md_name   = "no_md"
          )
        }

        # Extracting content and value from match
        pattern = first_match$pattern
        is_ref  = !is.null(pattern$is_reference) && pattern$is_reference

        if(is_ref) {
          # References produce run_reference commands
          ref_match = stringr::str_match(first_match$text, "<ref:(.+?)>")
          ref_key   = ref_match[2]
          res[[length(res) + 1]] = list(
            text         = paste0('officer::run_reference("', ref_key, '")'),
            is_reference = TRUE
          )
        } else {
          if(!is.null(pattern$extract_value) && pattern$extract_value) {
            # Patterns with embedded values (color, shading, font family)
            if(pattern$name == "color") {
              value_match  = stringr::str_match(first_match$text, "<color:(\\S+?)>(.+?)</color>")
              prop_value   = value_match[2]
              inner_content = value_match[3]
            } else if(pattern$name == "shading_color") {
              value_match  = stringr::str_match(first_match$text, "<shade:(\\S+?)>(.+?)</shade>")
              prop_value   = value_match[2]
              inner_content = value_match[3]
            } else if(pattern$name == "font_family") {
              value_match  = stringr::str_match(first_match$text, "<ff:(\\S+?)>(.+?)</ff>")
              prop_value   = value_match[2]
              inner_content = value_match[3]
            }
          } else {
            # Simple markers (bold, italic, sub/superscript)
            inner_content = stringr::str_match(first_match$text, pattern$regex)[2]
            prop_value    = if(!is.null(pattern$value)) pattern$value else TRUE
          }

          # Building new properties by applying the current pattern
          new_props = current_props
          for(prop in pattern$props) {
            if(prop == "bold")                new_props$bold           = TRUE
            else if(prop == "italic")         new_props$italic         = TRUE
            else if(prop == "vertical.align") new_props$vertical.align = prop_value
            else if(prop == "color")          new_props$color          = prop_value
            else if(prop == "shading.color")  new_props$shading.color  = prop_value
            else if(prop == "font.family")    new_props$font.family    = prop_value
          }

          # Recursively parsing inner content with updated properties
          inner_results = parse_text(inner_content, new_props)
          for(i in seq_along(inner_results)) {
            if(is.null(inner_results[[i]]$md_name) || inner_results[[i]]$md_name == "none" || inner_results[[i]]$md_name == "no_md") {
              inner_results[[i]]$md_name = pattern$name
            }
          }
          res = c(res, inner_results)
        }

        # Processing remaining text after the current match
        if(first_match$end < nchar(text)) {
          remaining = substr(text, first_match$end + 1, nchar(text))
          res = c(res, parse_text(remaining, current_props))
        }
      }
    }

  res
  }

  #----------------------------
  # Processing each paragraph: parsing markdown, building pele structures,
  # and generating ftext/fpar/as_paragraph command strings
  res = list()
  for(pgraph_idx in seq_along(pgraphs)) {
    pgraph_raw = pgraphs[pgraph_idx]

    # Removing carriage returns within paragraphs
    pgraph = gsub(pattern="(\r\n|\r|\n)", replacement=" ", pgraph_raw)

    # Parsing the paragraph into segments
    segments = parse_text(pgraph)

    #-----
    # Building pele structure
    pele = list()
    for(i in seq_along(segments)) {
      seg = segments[[i]]
      if(!is.null(seg$is_reference) && seg$is_reference) {
        pele[[paste0('p_', i)]] = list(
          text      = seg$text,
          md_name   = "reference",
          props     = c(),
          props_cmd = ""
        )
      } else {
        props_cmd = paste0("prop=officer::fp_text(", seg$props_str, ")")
        md_name   = if(!is.null(seg$md_name)) seg$md_name else "formatted"
        pele[[paste0('p_', i)]] = list(
          text      = seg$text,
          md_name   = md_name,
          props     = seg$props_str,
          props_cmd = props_cmd
        )
      }
    }

    #-----
    # Building ftext command string
    ftext_parts = c()
    for(seg in segments) {
      if(!is.null(seg$is_reference) && seg$is_reference) {
        ftext_parts = c(ftext_parts, seg$text)
      } else {
        ftext_parts = c(ftext_parts, paste0('officer::ftext("', seg$text, '", prop=officer::fp_text(', seg$props_str, '))'))
      }
    }
    ftext_cmd = paste(ftext_parts, collapse = " ,\n")

    #-----
    # Building fpar and as_paragraph command strings
    fpar_cmd = paste0("officer::fpar(", ftext_cmd, ")")

    as_para_parts = c()
    for(seg in segments) {
      if(!is.null(seg$is_reference) && seg$is_reference) {
        as_para_parts = c(as_para_parts, seg$text)
      } else {
        as_para_parts = c(as_para_parts, paste0('flextable::as_chunk("', seg$text, '", prop=officer::fp_text(', seg$props_str, '))'))
      }
    }
    as_paragraph_cmd = paste0("flextable::as_paragraph(", paste(as_para_parts, collapse = " ,\n"), ")")

    #-----
    # Storing results for this paragraph
    res[[paste0("pgraph_", pgraph_idx)]] = list(
      pele             = pele,
      locs             = NULL,
      ftext_cmd        = ftext_cmd,
      fpar_cmd         = fpar_cmd,
      as_paragraph_cmd = as_paragraph_cmd
    )
  }

res}
