#'@export
#'@title Parse Markdown into Officer as_paragraph Result
#'@description Used to take small markdown chunks and return the as_paragraph
#' results. This function will take the markdown specified in str, calls
#' md_to_officer, evals the as_paragraph field from the first paragraph
#' returned, evals that result and returns the object from the as_paragraph
#' command.
#'@param strs    vector of strings containing Markdown can contain the following elements:
#'@param default_format  list containing the default format for elements not defined with markdown default values (format the same as \code{\link{md_to_officer}}, default is \code{NULL})
#'@return list with the following elements
#' \itemize{
#'  \item{isgood}:    Boolean value indicating the result of the function call
#'  \item{msgs}:      sequence of strings contianing a description of any problems
#'  \item{as_par_cmd}:as_paragraph generated code from md_to_officer
#'  \item{oo}:        as_paragraph officer object resulting from running the as_par_cmd code
#'}
#'@examples
#'res = md_to_oo("Be **bold**")
md_to_oo     = function(strs,default_format=NULL){

  isgood     = TRUE
  as_par_cmd = NULL
  oo         = NULL
  msgs       = c()


  str = strs
  for(str in strs){
    # IFf str is empty we need something to hold it's place this way the
    # length and order of oo will match strs
    if(str == ""){
      str = " "
    }

    if(is.null(default_format)){
      mdres = md_to_officer(str)
    } else {
      mdres = md_to_officer(str, default_format)
    }

    # Checking to make sure we got what we needed from the md_to_officer command
    # above
    isgood_mdres = FALSE
    tmpoo = NULL
    if("pgraph_1" %in% names(mdres)){
      if("as_paragraph_cmd" %in% names(mdres[["pgraph_1"]])){
        isgood_mdres = TRUE
        as_par_cmd =  paste("tmpoo =", mdres[["pgraph_1"]][["as_paragraph_cmd"]])
        eval(parse(text=as_par_cmd))
      }
    }

    # If either of the fields above are missing then something failed
    if(!isgood_mdres){
      isgood = FALSE
      msgs = c(msgs, "md_to_officer call failed")
    }

    if(!isgood){
       msgs = c(msgs, "onbrand::md_to_oo()", "Unable to evaluate markdown, see above for details")
    }

    # Appending the temp officer object to the vector of officer objects
    oo = c(oo, tmpoo)
  }


  res = list(isgood     = isgood,
             msgs       = msgs,
             as_par_cmd = as_par_cmd,
             oo         = oo)


res}

