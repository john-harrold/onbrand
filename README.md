
<!-- README.md is generated from README.Rmd. Please edit that file -->

# onbrand <img src="man/figures/logo.png" align="right" height="138.5" />

<!-- badges: start -->

[![R-CMD-check](https://github.com/john-harrold/onbrand/workflows/R-CMD-check/badge.svg)](https://github.com/john-harrold/onbrand/actions)
[![CRAN
checks](https://badges.cranchecks.info/worst/onbrand.svg)](https://cran.r-project.org/web/checks/check_results_onbrand.html)
[![version](https://www.r-pkg.org/badges/version/onbrand)](https://CRAN.R-project.org/package=onbrand)
![cranlogs](https://cranlogs.r-pkg.org/badges/onbrand)
![Active](https://www.repostatus.org/badges/latest/active.svg)
<!-- badges: end -->

The `officer` package provides extensive methods for accessing,
creating, and modifying both Word and PowerPoint documents. These
methods require obtaining document specific placeholder and style
information. In order to switch between document templates, it is
necessary to change these references within the reporting code. The
purpose of `onbrand` is to provide an abstraction layer where template
details are mapped to human-readable names.

These human-readable names combined with the mapping information - in a
template-specific yaml file - provides a systematic method to script
support for different Word of PowerPoint templates. Which means, the
same workflow will support multiple outputs. Which makes your life
easier and, thus, makes the world a little better place.

## Installation

You can install the released version of `onbrand` from
[CRAN](https://cran.r-project.org/package=onbrand) with:

``` r
install.packages("onbrand")
```

And the development version from
[GitHub](https://github.com/john-harrold/onbrand) with:

``` r
# Installing devtools if it's not already installed
if(system.file(package="devtools") == ""){
  install.packages("devtools") 
}
devtools::install_github("john-harrold/onbrand", dependencies=TRUE)
```

## Getting Started

Browse through the [documentation](https://onbrand.ubiquity.tools/) and
check out the vignettes:

1.  [Custom
    Templates](https://onbrand.ubiquity.tools/articles/Custom_Office_Templates.html)
2.  [Templated
    Workflows](https://onbrand.ubiquity.tools/articles/Creating_Templated_Office_Workflows.html)

These vignettes contain everything you need to walk through the basics.
