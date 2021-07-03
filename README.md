
<!-- README.md is generated from README.Rmd. Please edit that file -->

# onbrand <img src="inst/graphics/onbrand_logo.png" align="right" height="138.5" />

<!-- badges: start -->

[![R-CMD-check](https://github.com/john-harrold/onbrand/workflows/R-CMD-check/badge.svg)](https://github.com/john-harrold/onbrand/actions)
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
[CRAN](https://CRAN.R-project.org) with:

``` r
install.packages("onbrand")
```

And the development version from [GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("john-harrold/onbrand")
```

## Getting Started

`onbrand` is packaged with two Vignettes:

1.  Custom\_Office\_Templates  
2.  Creating\_Templated\_Office\_Workflows

These vignettes contain everything you need to walk through the basics.
We recommend starting with the first vignette:  
*insert code chunk*
