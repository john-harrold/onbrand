#'@export
#'@title Populate Placeholder In Officer Report
#'@description Places content in a PowerPoint placeholder for a given officer document.
#'
#'@param obnd onbrand report object
#'@param content_type string indicating the content type 
#'@param content   content
#'@param ph_label  placeholder location (text)
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object. 
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
add_pptx_ph_content = function(obnd,content_type, content, ph_label, verbose=TRUE){

    # Pulling the report out to make accessing it easier
    rpt = obnd[["rpt"]]

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
          # Pulling out the default format for the Table element
          ff_res               = fetch_report_format(obnd, format_name="Table_Labels", verbose=verbose)

          # the actual format details:
          default_format_table = ff_res[["format_details"]]
      
          for(cname in names(content[["table"]])){
            # For each column name we run the header text through the markdown
            # conversion to produce the as_paragraph output:
            ft = flextable::compose(ft,
                              j     = cname,                                                    
                              part  = "header",                                                          
                              value = md_to_oo(strs= header_list[[cname]], default_format=default_format_table)[["oo"]])
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

    rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=ft) 

    } 
    else if(content_type == "flextable_object"){
      rpt = officer::ph_with(x=rpt,  location=officer::ph_location_label(ph_label=ph_label), value=content) 
    }

    # Pulling the report back in the obnd object
    obnd[["rpt"]] = rpt

return(obnd)}
