% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/ft_apply_md.R
\name{ft_apply_md}
\alias{ft_apply_md}
\title{Render Markdown in flextable Object}
\usage{
ft_apply_md(ft, obnd = NULL, part = "body", prows = NULL, pcols = NULL)
}
\arguments{
\item{ft}{Flextable object.}

\item{obnd}{Optional onbrand object used to format markdown. The default \code{NULL} value will use default formatting.}

\item{part}{Part of the table can be one of \code{"all"}, \code{"body"} (default), \code{"header"}, or \code{"footer"}.}

\item{prows}{Optional rows of the part to process, ignored when \code{part = "all"}. Set to \code{NULL} (default) to process all rows.}

\item{pcols}{Optional columns of the part to process, ignored when \code{part = "all"}. Set to \code{NULL} (default) to process all columns.}
}
\value{
flextable with markdown applied
}
\description{
Takes a flextable object and renders any markdown in the
specified part.
}
\examples{
library(onbrand)
library(flextable)

df = data.frame(
 A = c("e^x^",      "text"),
 B = c("sin(x~y~)", "**<ff:symbol>S</ff>**~x~"))

ft = flextable(df) |>
     delete_part(part="header") |>
     add_header(values = 
       list(A= "*Italics*", 
            B= "**Bold**") )    |>
     theme_vanilla()            |>
     ft_apply_md(part="all")

ft
}
