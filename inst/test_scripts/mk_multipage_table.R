# 
# This script is used in the build process for packaging onbrand.
# 

trow = c(1:51)
tcol = c(1:63)
cht  = c(rep("A", 20),
         rep("B", 20),
         rep("C", 11))


table_body = NULL
for(cn in tcol){
  if(is.null(table_body)){
    tmp_cmd =  paste0("table_body = data.frame(C_",
                      cn,"= paste0(trow, ',', cn))")
  } else {
    tmp_cmd =  paste0("table_body = cbind(table_body, data.frame(C_",
                      cn,"= paste0(trow, ',', cn)))")
  }
  eval(parse(text=tmp_cmd))
}


table_body[1,]    = "BQL"
table_body[5,8]   = "NC"
table_body[20,]   = "BQL"
table_body[25,2]  = "NC"

row_common = data.frame(
    ID = trow,
    CH = cht)


row_common_head = data.frame(
    ID  = c("ID",     "ID", "ID"),
    CH  = c("Cohort", "Cohort", "Cohort"))



table_body_head = NULL

cidx = 1
for(cn in names(table_body)){

  units = "units"
  range = "range"

  if(cidx < 4){
    range = "R A"
  } else if(cidx < 12 ){
    range = "R B"
  } else if(cidx < 18 ){
    range = "R C"
  } else if(cidx < 28 ){
    range = "R D"
  } else if(cidx < 35 ){
    range = "R E"
  } else if(cidx < 48 ){
    range = "R F"
  } else if(cidx < 55 ){
    range = "R G"
  } else if(cidx < 60 ){
    range = "R H"
  } else {
    range = "R I"
  }

  if(cidx < 4){
    units = "U A"
  } else if(cidx < 8  ){
    units = "U B"
  } else if(cidx < 14 ){
    units = "U Q"
  } else if(cidx < 18 ){
    units = "U C"
  } else if(cidx < 28 ){
    units = "U D"
  } else if(cidx < 35 ){
    units = "U E"
  } else if(cidx < 48 ){
    units = "U F"
  } else if(cidx < 55 ){
    units = "U G"
  } else if(cidx < 60 ){
    units = "U H"
  } else {
    units = "U I"
  }

  if(is.null(table_body_head)){
    tmp_cmd =  paste0("table_body_head = data.frame(",
                      cn,'= c("', cn, '", units, range))')
  } else {
    tmp_cmd =  paste0("table_body_head = cbind(table_body_head, data.frame(",
                      cn,'= c("', cn, '", units, range)))')
  }

  eval(parse(text=tmp_cmd))
  cidx = cidx + 1
}


lg_tbl_body             = table_body
lg_tbl_row_common       = row_common
lg_tbl_body_head        = table_body_head
lg_tbl_row_common_head  = row_common_head

repo_root = here::here()
save(lg_tbl_body           , file = file.path(repo_root,"data/lg_tbl_body.rda"),            compress = "bzip2", version = 2, ascii = FALSE)
save(lg_tbl_row_common     , file = file.path(repo_root,"data/lg_tbl_row_common.rda"),      compress = "bzip2", version = 2, ascii = FALSE)
save(lg_tbl_body_head      , file = file.path(repo_root,"data/lg_tbl_body_head.rda"),       compress = "bzip2", version = 2, ascii = FALSE)
save(lg_tbl_row_common_head, file = file.path(repo_root,"data/lg_tbl_row_common_head.rda"), compress = "bzip2", version = 2, ascii = FALSE)

