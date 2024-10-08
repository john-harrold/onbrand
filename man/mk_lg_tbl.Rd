% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/mk_lg_tbl.R
\name{mk_lg_tbl}
\alias{mk_lg_tbl}
\title{Creates Large Table Data for Testing}
\usage{
mk_lg_tbl()
}
\value{
list with the following elements
\itemize{
  \item{lg_tbl_body:}              Table body to be spread across multiple pages.
  \item{lg_tbl_row_common:}        Common rows of data to be found horizontially.
  \item{lg_tbl_body_head:}         Header data for the body portion.
  \item{lg_tbl_row_common_head:}   Header data for the common_row portion.
}
}
\description{
Generates a large table for testing multi-page table reports.
It is used for testing \code{span_table()}.
}
\examples{
res =  mk_lg_tbl()
}
