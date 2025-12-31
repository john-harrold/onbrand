#'@export
#'@title Spread Large Table Over Smaller Tables
#'@description Takes a large table and spreads it over smaller tables to
#'paginate it. It will preserve common row  information on the left and
#'separate columns according to maximum specifications. The final tables will
#'have widths less than or equal to both max_col and max_width, and heights
#'less than or equal to both max_row and max_height.
#'
#'@param table_body                Data frame with the body of the large table.
#'@param row_common                Data frame with the common rows.
#'@param table_body_head           Data frame or matrix with headers for the table body.
#'@param row_common_head           Data frame or matrix with headers for the common rows.
#'@param header_format             Format of the header either \code{"text"} (default) or \code{"md"} for markdown.
#'@param obnd                      Optional onbrand object used to format markdown. The default \code{NULL} value will use default formatting.
#'@param max_row                   Maximum number of rows in output tables (A value of \code{NULL} will set max_row to the number of rows in the table).
#'@param max_col                   Maximum number of columns in output tables (A value of \code{NULL} will set max_col to number of columns in the table).
#'@param max_width                 Maximum width of the final table in inches (A value of \code{NULL} will use 100 inches).
#'@param max_height                Maximum height of the final table in inches (A value of \code{NULL} will use 100 inches).
#'@param table_alignment           Character string specifying the alignment #'of the table (body and headers). Can be \code{"center"} (default), \code{"left"}, \code{"right"}, or \code{"justify"}
#'@param inner_border              Border object for inner border lines defined using \code{officer::fp_border()}
#'@param outer_border              Border object for outer border lines defined using \code{officer::fp_border()}
#'@param set_header_inner_border_v Boolean value to enable or disable inner vertical borders for headers
#'@param set_header_inner_border_h Boolean value to enable or disable inner horizontal borders for headers
#'@param set_header_outer_border   Boolean value to enable or disable outer border for headers
#'@param set_body_inner_border_v   Boolean value to enable or disable inner vertical borders for the body
#'@param set_body_inner_border_h   Boolean value to enable or disable inner horizontal borders for the body
#'@param set_body_outer_border     Boolean value to enable or disable outer border borders for the body
#'@param notes_detect              Vector of strings to detect in output tables (example \code{c("NC", "BLQ")}).
#'@seealso \code{\link{build_span} for the relationship of inputs.}
#'@details
#' The way the data frames relate to each other are mapped out below. The
#' dimensions of the different data frames are identified below (nrow x ncol)
#'
#'
#'\preformatted{
#'  |-------------------------------------| ---
#'  |                 |                   |  ^
#'  |                 |                   |  |
#'  | row_common_head |  table_body_head  |  | m
#'  |      m x n      |       m x c       |  |
#'  |                 |                   |  v
#'  |-------------------------------------| ---
#'  |                 |                   |  ^
#'  |                 |                   |  |
#'  |    row_common   |    table_body     |  | r
#'  |      r x n      |      r x c        |  |
#'  |                 |                   |  |
#'  |                 |                   |  v
#'  |-------------------------------------| ---
#'
#'  |<--------------->|<----------------->|
#'          n                   c
#'}
#'@return list with the following elements
#' \itemize{
#'   \item{isgood:}    Boolean indicating the exit status of the function.
#'   \item{one_body:}  Full table with no headers.
#'   \item{one_table:} Full table with headers.
#'   \item{msgs:}      Vector of text messages describing any errors that were found.
#'   \item{tables:}    Named list of tables. Each list element is of the output.
#'   format from \code{build_span()}.
#'}
#'@example inst/test_scripts/span_table.R
span_table = function(table_body                = NULL,
                      row_common                = NULL,
                      table_body_head           = NULL,
                      row_common_head           = NULL ,
                      header_format             = "text",
                      obnd                      = NULL,
                      max_row                   = 20,
                      max_col                   = 10,
                      max_height                = 7,
                      max_width                 = 6.5,
                      table_alignment           = "center",
                      inner_border              = officer::fp_border(color="black", width = .3),
                      outer_border              = officer::fp_border(color="black", width = 2),
                      set_header_inner_border_v = TRUE,
                      set_header_inner_border_h = TRUE,
                      set_header_outer_border   = TRUE ,
                      set_body_inner_border_v   = TRUE,
                      set_body_inner_border_h   = FALSE,
                      set_body_outer_border     = TRUE ,
                      notes_detect              = NULL){

  isgood = TRUE
  msgs   = c()

  # Checking user input:
  if(is.null(table_body)){
      isgood = FALSE
      msgs = c(msgs, "table_body must be specified.")
  }
  if(is.null(row_common)){
      isgood = FALSE
      msgs = c(msgs, "row_common must be specified.")
  }

  if(isgood){
    if(is.null(max_row)){
      max_row = nrow(table_body) }
    if(is.null(max_col)){
      max_col = ncol(table_body) }
    if(is.null(max_height)){
      max_height = 100 }
    if(is.null(max_width)){
      max_width = 100 }
    if(!is.null(table_body_head)){
      if(ncol(table_body_head) !=
         ncol(table_body)){
        isgood = FALSE
        msgs = c(msgs, "table_body and table_body_head have a different number of columns.")
      }
    }

    if(!is.null(row_common_head)){
      if(ncol(row_common_head) !=
         ncol(row_common)){
        isgood = FALSE
        msgs = c(msgs, "row_common and row_common_head have a different number of columns.")
      }
    }

    if(nrow(row_common) != nrow(table_body)){
        isgood = FALSE
        msgs = c(msgs, "row_common and table_body have a different number of rows.")
    }

    if(!is.null(row_common_head) & is.null(table_body_head)){
          isgood = FALSE
          msgs = c(msgs, "You must specify both or neither:")
          msgs = c(msgs, "  row_common_head: specified")
          msgs = c(msgs, "  table_body_head: not specified")
    }

    if(is.null(row_common_head) & !is.null(table_body_head)){
          isgood = FALSE
          msgs = c(msgs, "You must specify both or neither:")
          msgs = c(msgs, "  row_common_head: not specified")
          msgs = c(msgs, "  table_body_head: specified")
    }

    if(!is.null(row_common_head) & !is.null(table_body_head)){
      if(nrow(row_common_head) != nrow(table_body_head)){
          isgood = FALSE
          msgs = c(msgs, "row_common_head and table_body_head must have the same number of rows.")
      }
    }
  }

  # JMH Add checks for max_height and max_width here:
  if(isgood){
    # Building the table with 1 row and 1 column to see if the row_common is
    # too wide or the headers are too tall.
    bsr = build_span(
      table_body                = table_body,
      row_common                = row_common,
      table_body_head           = table_body_head,
      row_common_head           = row_common_head,
      header_format             = header_format,
      obnd                      = obnd,
      row_sel                   = c(1:1),
      col_sel                   = c(1:1),
      table_alignment           = table_alignment,
      inner_border              = inner_border,
      outer_border              = outer_border,
      set_header_inner_border_v = set_header_inner_border_v,
      set_header_inner_border_h = set_header_inner_border_h,
      set_header_outer_border   = set_header_outer_border,
      set_body_inner_border_v   = set_body_inner_border_v,
      set_body_inner_border_h   = set_body_inner_border_h,
      set_body_outer_border     = set_body_outer_border,
      notes_detect              = notes_detect)

    tmp_dim = flextable::flextable_dim(bsr[["ft"]])
    if(max_width - tmp_dim$widths < 0){
      isgood = FALSE
      msgs = c(msgs, paste0("row_common is too wide and there isn't enough room for data."))
      msgs = c(msgs, paste0("  Too many grouping variables can cause this." ))
    }
    if(max_height - tmp_dim$heights < 0){
      isgood = FALSE
      msgs = c(msgs, paste0("table_body_head/row_common_head is too tall and there isn't enough room for data."))
      msgs = c(msgs, paste0("  Too many header rows can cause this." ))
    }
  }

  # This means that everything checks out and we're good to go.
  tables    = list()
  one_table = NULL
  one_body  = NULL
  one_header= NULL
  if(isgood){
    tb_idx = 1

    row_offset = 0
    while(row_offset < nrow(table_body)){
      row_sel_start = row_offset+1

      if(max_row+row_offset <= nrow(table_body)){
        row_sel_stop  = max_row+row_offset
      } else {
        row_sel_stop  = nrow(table_body)
      }

      row_sel = c(row_sel_start:row_sel_stop)

      col_offset = 0
      while(col_offset < ncol(table_body)){

        # We start 1 after the offset
        col_sel_start = col_offset+1

        # This is the intervals before the last column
        if((max_col+col_offset) <=( ncol(table_body) - ncol(row_common))){
          col_sel_stop  = max_col+col_offset - ncol(row_common)
        } else {
          # And this goes to the last column
          col_sel_stop  = ncol(table_body)
        }

        col_sel = unique(c(col_sel_start:col_sel_stop))

        tb_key = paste0("Table ", tb_idx)
        #cat(paste0("row: ", toString(c(min(row_sel), max(row_sel))), ", col: ", toString(c(min(col_sel),max(col_sel))), "\n"))

        # First we build the table based on row_sel and col_sel from max
        # values specified:
        bsr = build_span(
          table_body                = table_body,
          row_common                = row_common,
          table_body_head           = table_body_head,
          row_common_head           = row_common_head,
          header_format             = header_format,
          obnd                      = obnd,
          row_sel                   = row_sel,
          col_sel                   = col_sel,
          table_alignment           = table_alignment,
          inner_border              = inner_border,
          outer_border              = outer_border,
          set_header_inner_border_v = set_header_inner_border_v,
          set_header_inner_border_h = set_header_inner_border_h,
          set_header_outer_border   = set_header_outer_border,
          set_body_inner_border_v   = set_body_inner_border_v,
          set_body_inner_border_h   = set_body_inner_border_h,
          set_body_outer_border     = set_body_outer_border,
          notes_detect              = notes_detect)

        resize = FALSE
        # Now we test to make sure it falls within the max width and max
        # height values:
        ftdims = flextable::flextable_dim(bsr[["ft"]])

        #----------------------------------------
        # Trimming columns to fit within max_width
        if(ftdims[["widths"]] > max_width){
          resize = TRUE

          tmp_dim = dim(bsr[["ft"]])

          # This finds the number of columns to remove off the bottom:
          ntrim = 0
          while(sum(tmp_dim$widths) > max_width){
            ntrim = ntrim + 1
            tmp_dim$widths = utils::head(tmp_dim$widths, -1)
          }

          # This trims off those elements
          col_sel = col_sel[1:(length(col_sel) - ntrim)]
        }
        #----------------------------------------
        # Trimming rows to fit within max_height
        if(ftdims[["heights"]] > max_height){
          resize = TRUE
          tmp_dim = dim(bsr[["ft"]])

          # This finds the number of rows to remove off the bottom:
          ntrim = 0
          while(sum(tmp_dim$heights) > max_height){
            ntrim = ntrim + 1
            tmp_dim$heights = utils::head(tmp_dim$heights, -1)
          }

          # This trims off those elements
          row_sel = row_sel[1:(length(row_sel) - ntrim)]
        }
        #----------------------------------------
        # This triggers rebuilding of the table to fit within the specified
        # dimensions
        if(resize){
          bsr = build_span(
            table_body                = table_body,
            row_common                = row_common,
            table_body_head           = table_body_head,
            row_common_head           = row_common_head,
            header_format             = header_format,
            obnd                      = obnd,
            row_sel                   = row_sel,
            col_sel                   = col_sel,
            table_alignment           = table_alignment,
            inner_border              = inner_border,
            outer_border              = outer_border,
            set_header_inner_border_v = set_header_inner_border_v,
            set_header_inner_border_h = set_header_inner_border_h,
            set_header_outer_border   = set_header_outer_border,
            set_body_inner_border_v   = set_body_inner_border_v,
            set_body_inner_border_h   = set_body_inner_border_h,
            set_body_outer_border     = set_body_outer_border,
            notes_detect              = notes_detect)
        }
        # Storing the information about the current table.
        tables[[tb_key]] = bsr
        tables[[tb_key]][["tb_idx"]] = tb_idx

        tb_idx = 1+tb_idx
        col_offset = max(col_sel)
      }
      row_offset = max(row_sel)
    }


    #creating the single table as well:
    one_body   = cbind(row_common, table_body)
    one_header = cbind(row_common_head, table_body_head)
    if(!is.null(one_header)){
      one_header = as.data.frame(one_header)
      colnames(one_header) = names(one_body)
    }
    one_table  = rbind(one_header, one_body)
  }

  # If an error was encountered above we want to pack that into a tabular
  # output. This way it can cascade to an end user
  if(!isgood){
    df = data.frame("span_table_Failed" = msgs)
    ft = flextable::flextable(df)
    tables = list(Error = list(
              df = df,
              ft = ft,
              tb_idx = 1))
  }



  res = list(isgood     = isgood,
             one_table  = one_table,
             one_header = one_header,
             one_body   = one_body ,
             tables     = tables,
             msgs       = msgs)

res}



#'@title Construct Table Span From Components
#'@description Takes a large table, common rows, and header information and
#'constructs a table that is a subset of those components using supplied
#'ranges of rows and columns.
#'@param table_body                Data frame with the body of the large table.
#'@param row_common                Data frame with the common rows.
#'@param table_body_head           Data frame or matrix with headers for the table body.
#'@param row_common_head           Data frame or matrix with headers for the common rows.
#'@param header_format             Format of the header either \code{"text"} (default) or \code{"md"} for markdown.
#'@param obnd                      Optional onbrand object used to format markdown. The default \code{NULL} value will use default formatting.
#'@param col_sel                   Indices of columns to build to the table with.
#'@param row_sel                   Indices of rows to build to the table with.
#'@param table_alignment           Character string specifying the alignment #'of the table (body and headers). Can be \code{"center"} (default), \code{"left"}, \code{"right"}, or \code{"justify"}
#'@param inner_border              Border object for inner border lines defined using \code{officer::fp_border()}
#'@param outer_border              Border object for outer border lines defined using \code{officer::fp_border()}
#'@param set_header_inner_border_v Boolean value to enable or disable inner vertical borders for headers
#'@param set_header_inner_border_h Boolean value to enable or disable inner horizontal borders for headers
#'@param set_header_outer_border   Boolean value to enable or disable outer border for headers
#'@param set_body_inner_border_v   Boolean value to enable or disable inner vertical borders for the body
#'@param set_body_inner_border_h   Boolean value to enable or disable inner horizontal borders for the body
#'@param set_body_outer_border     Boolean value to enable or disable outer border borders for the body
#'@param notes_detect              Vector of strings to detect in output tables (example \code{c("NC", "BLQ")}).
#'@details
#' The way the data frames relate to each other are mapped out below. The
#' dimensions of the different data frames are identified below (nrow x ncol)
#'
#'
#' \preformatted{
#'                             col_sel
#'                       |<--------------->|
#'
#' |--------------------------------------------| ---
#' |                 |   .                 .    |  ^
#' |                 |   .                 .    |  |
#' | row_common_head |   . table_body_head .    |  | m
#' |      m x n      |   .      m x c      .    |  |
#' |                 |   .                 .    |  v
#' |--------------------------------------------| ---
#' |                 |   .                 .    |  ^
#' |                 |   .                 .    |  |
#' |    row_common   |   .   table_body    .    |  |
#' |      r x n      |   .     r x c       .    |  |
#' |                 |   .                 .    |  |
#' |.................|..........................|  |     -
#' |                 |   ./  /  /  /  /  / .    |  |     ^
#' |                 |   .  /  /  /  /  /  .    |  | r   |
#' |                 |   . /  /  /  /  /  /.    |  |     | row_sel
#' |                 |   ./  /  /  /  /  / .    |  |     |
#' |                 |   .  /  /  /  /  /  .    |  |     v
#' |.................|...../../../../../../.... |  |     -
#' |                 |   .                 .    |  |
#' |                 |   .                 .    |  v
#' |--------------------------------------------| ---
#'
#' |<--------------->|<------------------------>|
#'         n                    c
#' }
#'@return list with the following elements
#' \itemize{
#'   \item{df:}     Data frame with the built table.
#'   \item{ft:}     The data frame as a flextable object.
#'   \item{notes:}  Note placeholders found in the table.
#'}
#'@example inst/test_scripts/span_table.R
build_span  = function(table_body                = NULL,
                       row_common                = NULL,
                       table_body_head           = NULL,
                       row_common_head           = NULL ,
                       header_format             = "text",
                       obnd                      = NULL,
                       row_sel                   = NULL,
                       col_sel                   = NULL,
                       table_alignment           = "center",
                       inner_border              = officer::fp_border(color="black", width = .3),
                       outer_border              = officer::fp_border(color="black", width = 2),
                       set_header_inner_border_v = TRUE,
                       set_header_inner_border_h = TRUE,
                       set_header_outer_border   = TRUE ,
                       set_body_inner_border_v   = TRUE,
                       set_body_inner_border_h   = FALSE,
                       set_body_outer_border     = TRUE ,
                       notes_detect              = NULL){

  # Constructing the data frame with the chunk of the table:
  df = cbind(row_common[row_sel, ],
             table_body[row_sel, col_sel])

  # Finding notes in the current df
  notes = c()
  if(!is.null(notes_detect)){
    for(note in notes_detect){
      if(any(note == df, na.rm=TRUE)){
        notes = c(notes, note)
      }
    }
  }

  # Building the flextable and remove default formatting
  ft = flextable::flextable(df) |>
       flextable::border_remove()

  if(!is.null(row_common_head) &
     !is.null(table_body_head)){

    # This clears out the default headers for the table
    ft = flextable::delete_part(ft, part="header")


    # This should contain the portion of the body header
    # for the current data frame
    tmp_bh = table_body_head[, col_sel]

    # Merging the common header with the portion for the current
    # data frame:
    ft_head = cbind(row_common_head, tmp_bh)

    for(hidx in nrow(ft_head):1){

      cv  = c()
      ccw = c()

      tmp_hrow = unlist(as.vector(ft_head[hidx, ]))

      # Here we're collapsing rows that look like this:
      # A, A, A, B, B, B, B, A, A, C, C, Q into
      # cv  = ("A", "B", "A", "C", "Q")
      # ccw = ( 3 ,  4 ,  2 ,  2,   1 )

      tot_cctr   = 1
      tmp_cv     = tmpcv = tmp_hrow[1]
      tmp_ctr    = 0
      while(tot_cctr <= length(tmp_hrow)){
        if(tmp_cv != tmp_hrow[tot_cctr]){
          cv      = c(cv, tmp_cv)
          ccw     = c(ccw, tmp_ctr)
          tmp_cv  = tmp_hrow[tot_cctr]
          tmp_ctr = 0
        }

        tmp_ctr   = tmp_ctr  + 1

        if(tot_cctr == length(tmp_hrow)){
          cv      = c(cv, tmp_cv)
          ccw     = c(ccw, tmp_ctr)
        }

        tot_cctr  = tot_cctr + 1
      }

      # This adds the header row
      ft = flextable::add_header_row(ft, values=cv, colwidths=ccw)

    }

    # Applying the markdown formatting to the headers if requested:
    if(header_format == "md"){
      ft = ft_apply_md(ft, obnd, part="header")
    }
  }



  if(set_body_outer_border){
    ft = flextable::border_outer(ft, border=outer_border, part="all") }

  # Here we add all the specified borders:
  if(set_header_inner_border_v){
    ft = flextable::border_inner_v(ft, border=inner_border, part="header") }

  if(set_header_inner_border_h){
    ft = flextable::border_inner_h(ft, border=inner_border, part="header") }


  if(set_body_inner_border_v){
    ft = flextable::border_inner_v(ft, border=inner_border, part="body") }

  if(set_body_inner_border_h){
    ft = flextable::border_inner_h(ft, border=inner_border, part="body") }

  # Applying alignment to the entire table
  ft = flextable::align(   ft, part="all", align=table_alignment)

  # Combining header elements
  ft = flextable::merge_v(ft, part="header")
  ft = flextable::merge_h(ft, part="header")

  # This fixes odd border issues with merged cells
  ft <- flextable::fix_border_issues(x = ft)

  res  = list(df    = df,
              ft    = ft,
              notes = notes)
res}
