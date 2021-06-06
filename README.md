
<!-- README.md is generated from README.Rmd. Please edit that file -->

# onbrand

<!-- badges: start -->

[![R-CMD-check](https://github.com/john-harrold/onbrand/workflows/R-CMD-check/badge.svg)](https://github.com/john-harrold/onbrand/actions)
<!-- badges: end -->

The `officer` package provides extensive methods for accessing,
creating, and modifying Word and PowerPoint documents. These methods
require accessing placeholder and style information that is document
specific. In order to switch between document templates, it is necessary
to change these references within the reporting code. The purpose of
`onbrand` is to provide an abstraction layer where template details can
be mapped to human-readable names. Scripts can be written using these
human-readable names, and different templates can be used by providing
the mapping information in a template-specific yaml file.

## Installation

You can install the released version of onbrand from
[CRAN](https://CRAN.R-project.org) with:

``` r
install.packages("onbrand")
```

And the development version from [GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("john-harrold/onbrand")
```
