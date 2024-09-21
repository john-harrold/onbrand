if(interactive()){

res =
span_table(table_body      = lg_tbl_body,
           row_common      = lg_tbl_row_common,
           table_body_head = lg_tbl_body_head,
           row_common_head = lg_tbl_row_common_head,
           max_row         = 20,
           max_col         = 10,
           notes_detect    = c("BQL", "NC"))


# Notes detected in the first table:
res[["tables"]][["Table 1"]][["notes"]]

# First table as a data frame:
res[["tables"]][["Table 1"]][["df"]]

# First table as a flextable:
res[["tables"]][["Table 1"]][["ft"]]

}

