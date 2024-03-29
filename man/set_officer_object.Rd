% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/set_officer_object.R
\name{set_officer_object}
\alias{set_officer_object}
\title{Places Officer Object Into Onbrand Report Object}
\usage{
set_officer_object(obnd, rpt = NULL, verbose = TRUE)
}
\arguments{
\item{obnd}{onbrand report object}

\item{rpt}{officer object}

\item{verbose}{Boolean variable when set to TRUE (default) messages will be
displayed on the terminal; Messages will be included in the returned onbrand
object.}
}
\value{
onbrand object with the report replaced
}
\description{
After modifying a report object manually, you can return it to
the onbrand object using this function.
}
\examples{

obnd = read_template(
      template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
      mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

# pulling out the report
rpt = fetch_officer_object(obnd)$rpt

# Modifications would be made here with officer directly

# Replacing the report into the onbrand object
obnd = set_officer_object(obnd, rpt)

}
\seealso{
\code{\link{fetch_officer_object}}
}
