#'@export
#'@title Render Markdown in flextable Object
#'@description Takes a flextable object and renders any markdown in the
#'specified part.
#'@param ft                        Flextable object.
#'@param obnd                      Optional onbrand object used to format markdown. The default \code{NULL} value will use default formatting.
#'@param part                      Part of the table can be one of \code{"all"}, \code{"body"} (default), \code{"header"}, or \code{"footer"}.
#'@param prows                     Optional rows of the part to process, ignored when \code{part = "all"}. Set to \code{NULL} (default) to process all rows.
#'@param pcols                     Optional columns of the part to process, ignored when \code{part = "all"}. Set to \code{NULL} (default) to process all columns.
#'@return flextable with markdown applied
#'@examples
#' library(onbrand)
#' library(flextable)
#'
#' df = data.frame(
#'  A = c("e^x^",      "text"),
#'  B = c("sin(x~y~)", "**<ff:symbol>S</ff>**~x~"))
#' 
#' ft = flextable(df) |>
#'      delete_part(part="header") |>
#'      add_header(values = 
#'        list(A= "*Italics*", 
#'             B= "**Bold**") )    |>
#'      theme_vanilla()            |>
#'      ft_apply_md(part="all")
#'
#' ft
ft_apply_md = function(ft, obnd=NULL, part = "body", prows = NULL, pcols = NULL){

  # This defines the defatul format for the header:
  if(is.null(obnd)){
    dft = NULL
  } else {
    dft      = fetch_md_def(obnd, style="Table_Labels")$md_def
  }

  if(part == "all"){
    parts = c("body", "header", "footer")
  } else {
    parts = part
  }

  isgood = TRUE 

  for(tmppart in parts){
    part_data = ft[[tmppart]]$data

    # If the part has no data then we skip it. This can happen for example if
    # all is selected and there is no footer present
    if(nrow(part_data) > 0){

      if(is.null(prows) | part == "all" ){
        prows = 1:nrow(part_data)
      } else {
        if(!all(prows %in% 1:nrow(part_data))){
          message(paste0("The following rows were specified but are beyond the bounds of the ", tmppart, ": "))
          message(paste0("  >", paste(prows[!(prows %in% 1:nrow(part_data))], collapse=", ")))
          isgood = FALSE
        }
      }
      if(is.null(pcols) | part == "all"){
        pcols = 1:ncol(part_data)
      } else {
        if(!all(pcols %in% 1:ncol(part_data))){
          message(paste0("The following columns were specified but are beyond the bounds of the ", tmppart, ": "))
          message(paste0("  >", paste( pcols[!(pcols %in% 1:ncol(part_data))], collapse=", ")))
          isgood = FALSE
        }
      }
      if(!isgood){
        stop("ft_apply_md() failed see above for details")
      }

      # Now we walk through each element of the current table part and apply
      # markdown to that element:
      for(hcol in pcols){
        for(hrow in prows){
          # To have markdown there has to be at least 3 characters and we only
          # apply it to character data
          if(is.character(part_data[hrow,hcol])){
          if(nchar(part_data[hrow,hcol])>2){
            # Here we apply markdown using either the default format from
            # onbrand or the value specified in the onbrand object
            if(is.null(dft)){
              new_value = md_to_oo(as.character(part_data[hrow, hcol]))$oo
            } else {
              new_value = md_to_oo(as.character(part_data[hrow, hcol], dft))$oo
            }
            ft = flextable::compose(ft,
              i = hrow, j = hcol,
              part = tmppart,
              value = new_value
            )
          }
          }
        }
      }
    }
  }



ft}

