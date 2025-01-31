library(flextable)
library(onbrand)

tmpdf = data.frame(
 Cola = rep("A^x^", 10),
 Colb = exp(rnorm(10))+10,
 Colc = rep("C^x^", 10),
 Cold = exp(rnorm(10))+10,
 Cole = rep("E^x^", 10),
 Colf = exp(rnorm(10))+10,
 Colg = rep("G^x^", 10),
 Colh = rep("H^x^", 10))

obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

ft = flextable::flextable(tmpdf)

test_that("Apply MD beyond flextable dimensions",{
expect_error({
  ft_bad_body  =
    ft_apply_md(ft    = ft,     obnd  = obnd,
                part  = "body", prows = c(8:15),
                pcols = c(4:10))
        })
})


test_that("Apply MD to flextable body",{
expect_no_failure({
  ft_all_body  =
    ft_apply_md(ft    = ft,     obnd  = obnd,
                part  = "body")
  })
 })

test_that("Apply MD to flextable header",{
expect_no_failure({
  ft_all_header  =
    ft_apply_md(ft    = ft,     obnd  = obnd,
                part  = "header")
  })
 })

test_that("Apply MD portion of flextable body",{
expect_no_failure({
  ft_part_body =
    ft_apply_md(ft    = ft,
                obnd  = obnd,
                part  ="body",
                prows = c(2:6),
                pcols = c(4:8))
  })
 })

