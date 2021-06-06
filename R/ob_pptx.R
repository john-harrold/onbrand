#'@export
#'@title Add Slide and Content
#'@description Creates a report slide and populates the content 
#'
#'@param obnd onbrand report object
#'@param template Name of slide template to use (name from templates in yaml mapping file)
#'@param elements Content and type for each placeholder you wish to fill for
#'this slide: This is a list with names set to palceholders for the specified
#'tempalte. Each placeholder is a list and should have a content element and a
#'type element (see Details below). 
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object. 
#'
#'@return onbrand report object with either the contenet added or isgood set
#'to FALSE with any messages in the msgs field.
#'
#'@details
#'For example consider the mapping information for the slide
#'template \code{title_slide} with the two place holders \code{title} and
#'\code{subtitle}.
#' 
#'\preformatted{
#'rpptx:
#'  master: Office Theme
#'  templates:
#'    title_slide:
#'      title:
#'        type:         ctrTitle
#'        index:        1
#'        ph_label:     Title 1
#'        content_type: text
#'      subtitle:
#'        type:         subTitle
#'        index:        1
#'        ph_label:     Subtitle 2
#'        content_type: text
#' }
#' 
#'This shows how to populate a title slide with text:
#' 
#'\preformatted{
#'obnd = report_add_slide(obnd,
#'  template = "title_slide",
#'  elements = list(
#'     title     = list( content      = "Slide Title",
#'                       type         = "text"),
#'     subtitle  = list( content      = "Subtitle",
#'                       type         = "text")))
#'}
#'
#'See the function \code{\link{report_add_pptx_ph_content}} for a list of
#'allowed values for \code{type}. Note that if mapping defines the
#'\code{content_type} as \code{text}, you cannot use a \code{list} type.
#'Similarly, if the \code{content_type} is defined as \code{list}, you 
#'cannot use a \code{text} type.
#'
#'@seealso \code{\link{report_add_pptx_ph_content}}
#'
#'@examples
#'obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'obnd = report_add_slide(obnd,
#'  template = "content_text", 
#'  elements = list(
#'     title         = list( content      = "Text Example",
#'                           type         = "text"),
#'     sub_title     = list( content      = "Adding a slide with a block of text",
#'                           type         = "text"),
#'     content_body  = list( content      = "A block of text",
#'                           type         = "text")))
#'
report_add_slide = function (obnd,
                             template = NULL,
                             elements = NULL,
                             verbose  = TRUE){
  msgs          = c()
  isgood        = TRUE
  general_types = c("imagefile", "ggplot", "table", "flextable", "flextable_object")

  #-----
  # This is a set of basic tests on user input
  if(obnd[["isgood"]]){
    if(obnd[["rpttype"]] != "PowerPoint"){
      isgood = FALSE
      msgs = c(msgs, paste0("The povided onbrand object is for rpttype >", obnd[["rpttype"]], "<"))
      msgs = c(msgs, "and should be PowerPoint")
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }
  if(is.null(template)){
    isgood = FALSE
    msgs = c(msgs, "No slide template was provided")
  }
  if(is.null(elements)){
    isgood = FALSE
    msgs = c(msgs, "No slide elements were provided")
  }
  #-----
  # If we pass the basic tests then we start walking through the user supplied
  # information to make sure it makes sense in the context of the current
  # templates
  if(isgood){
    if(template %in% names(obnd[["meta"]][["rpptx"]][["templates"]])){
      # this holds the template details for the current layout:
      td = obnd[["meta"]][["rpptx"]][["templates"]][[template]]
      # If the template exists we test each placeholder next
      if(all(names(elements) %in% names(td))){
        # now we check the elements that are specified as text/list 
        for(phname in names(elements)){
          # Making sure the content types are correct
          allowed_types = c(td[[phname]][["content_type"]], general_types)
          if(!(elements[[phname]][["type"]] %in% allowed_types)){
            isgood = FALSE
            msgs = c(msgs, paste0("The content type >", elements[[phname]][["type"]], "< is not allowed with"))
            msgs = c(msgs, paste0("the layout template >", template, "<"))
            msgs = c(msgs, paste0("The allowed values of type are:"))
            msgs = c(msgs, paste0("  ", paste(allowed_types, collapse=", ")))
          }
        }
      } else {
        isgood = FALSE
        msgs = c(msgs, paste0("The following placeholder elements were specified but"))
        msgs = c(msgs, paste0("not found in the mapping file: "))
        msgs = c(msgs, paste0("  ",
          paste(names(elements)[!(names(elements) %in% names(td))], collapse =", ")))

      }
    }  else {
      isgood = FALSE
      msgs = c(msgs, paste0("The slide template >", template, "< was not found"))
      msgs = c(msgs, paste0("In the mapping file."))
    }
  }


  #-------
  # If all that above checks out then we add the reporting elements
  if(isgood){

    # Pulling the report out of the obnd object to make working with things
    # easier
    tmprpt = obnd[["rpt"]]

    # First we add the new slide according to the specified template
    tmprpt = officer::add_slide(x      = tmprpt, 
                       layout = template,
                       master = obnd[["meta"]][["rpptx"]][["master"]])
      
    # Now we walk through each specified placeholder and add it
    for(phname in names(elements)){
      tmprpt = report_add_pptx_ph_content(rpt = tmprpt,
                 content      = elements[[phname]][["content"]], 
                 content_type = elements[[phname]][["type"]], 
                 ph_label     = td[[phname]][["ph_label"]])
    }
    # putting the report back in the obnd object
    obnd[["rpt"]] = tmprpt
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

  # Adding any messages to the report object   
  obnd[["msgs"]] = msgs


obnd}

#'@export
#'@title Populate Placeholder In Officer Report
#'@description Places content in a PowerPoint placeholder for a given Officer document.
#'
#'@param rpt officer pptx object
#'@param content_type string indicating the content type 
#'@param content   content
#'@param type      placeholder type (\code{"body"})
#'@param ph_label  placeholder location (text)
#'
#'@return officer pptx object with the content added
#'
#'@details
#'
#' For each content type listed below the following content is expected:
#'
#' \itemize{
#'  \item \code{"text"} text string of information
#'  \item \code{"list"} vector of paired values (indent level and text), eg.  c(1, "Main Bullet", 2 "Sub Bullet")
#'  \item \code{"imagefile"} string containing path to image file
#'  \item \code{"ggplot"} ggplot object, eg. p = ggplot() + ....
#'  \item \code{"table"} list containing the table content and other options with the following elements (defaults in parenthesis):
#'   \itemize{
#'      \item \code{table} Data frame containing the tabular data
#'      \item \code{header} Boolean variable to control displaying the header (\code{TRUE})
#'      \item \code{first_row} Boolean variable to indicate that the first row contains header information (\code{TRUE})
#'    }
#'  \item \code{"flextable"} list containing flextable content and other options with the following elements (defaults in parenthesis):
#'   \itemize{
#'      \item \code{table} Data frame containing the tabular data
#'      \item \code{header_top}, \code{header_middle}, \code{header_bottom} (\code{NULL}) a list with the same names as the data frame names containing the tabular data and values with the header text to show in the table
#'      \item \code{header_format} string containing the format, either \code{"text"}, or \code{"md"} (default \code{NULL} assumes \code{"text"} format)
#'      \item \code{merge_header} (\code{TRUE}) Set to true to combine column headers with the same information
#'      \item \code{table_body_alignment}, table_header_alignment ("center") Controls alignment
#'      \item \code{table_autofit} (\code{TRUE}) Automatically fit content, or specify the cell width and height with \code{cwidth} (\code{0.75}) and \code{cheight} (\code{0.25})
#'      \item \code{table_theme} (\code{"theme_vanilla"}) Table theme
#'    }
#'  \item \code{"flextable_object"} user defined flextable object 
#'  }
#'
#'@seealso \code{\link{view_layout}}
report_add_pptx_ph_content = function(rpt,content_type, content, ph_label){

    if(content_type == "text"){
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=content) 
    }
    else if(content_type == "list"){
      mcontent = matrix(data = content, ncol=2, byrow=TRUE)
      
      # constructing the elements for unordered_list below
      level_list  = c()
      str_list    = c()
      for(lidx in 1:length(mcontent[,1])){
        level_list  = c(level_list,  as.numeric(mcontent[lidx, 1]))
        str_list    = c(str_list, mcontent[lidx, 2])
      }

      # packing the list pieces into the ul object
      ul = officer::unordered_list(level_list = level_list, str_list = str_list)
      # adding it to the report
      rpt = officer::ph_with(x  = rpt,  
                             location  = officer::ph_location_label(ph_label=ph_label),
                             value     = ul) 
    }
    else if(content_type == "imagefile"){
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=officer::external_img(src=content)) 
    }
    else if(content_type == "ggplot"){
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=content) 
    }
    else if(content_type == "table"){
      if('header' %in% names(content)){
        header = content[["header"]]
      } else {header = TRUE}
      if('first_row' %in% names(content)){
        first_row = content[["first_row"]]
      } else {first_row = TRUE}
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=content[["table"]], header=header, first_row=first_row) 
    }
    else if(content_type == "flextable"){
      # These are the default table options
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
     #-------
     # JMH fix this fetch_report_format thing
     #    # Pulling out the default format for the Table element
     #    default_format_table = system_fetch_report_format(cfg, rptname=rptname, element="Table_Labels")
     #
     #    for(cname in names(content[["table"]])){
     #      # For each column name we run the header text through the markdown
     #      # conversion to produce the as_paragraph output:
     #      ft = flextable::compose(ft,
     #                        j     = cname,                                                    
     #                        part  = "header",                                                          
     #                        value = md_to_oo(strs= header_list[[cname]], default_format=default_format_table)[["oo"]])
     #    }
     # /JMH
     #-------
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

    rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=ft) 

    } 
    else if(content_type == "flextable_object"){
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=content) 
    }
return(rpt)}
