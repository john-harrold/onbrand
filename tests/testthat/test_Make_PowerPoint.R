
#---
# Creating objects for the tests below
# Figures
p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "picture example")
imgfile = tempfile(pattern="image", fileext=".png")
ggplot2::ggsave(filename=imgfile, plot=p, height=5.15, width=9, units="in")


# Bulleted list
bl = c("1", "This is first level bullet",
       "2", "sub-bullet",
       "3", "sub-sub-bullet",
       "1", "Another first level bullet")

# Tables
tdf =    data.frame(Parameters = c("Length", "Width", "Height"),
                    Values     = 1:3,
                    Units      = c("m", "m", "m") )
# Used for PowerPoint Tables
tab_cont = list(table = tdf)


# Used for flextables
tab_ft = list(table         = tdf,
              header_top    = list(Parameters = "Name",
                                   Values     = "Value",
                                   Units      = "Units"),
              cwidth        = 0.8,
              table_autofit = TRUE,
              table_theme   = "theme_zebra")

# Used for flextable_objects
tab_fto = flextable::flextable(tdf)

# Loading the default PowerPoint template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

#--------------
test_that("PowerPoint: Adding Text",{
tmp_obnd = report_add_slide(obnd,
  template = "title_slide",
  elements = list(
     title     = list( content         = "Placeholder Types",
                       type            = "text"),
     sub_title = list( content         = "Testing each type of placeholder",
                       type            = "text")))
expect_true(tmp_obnd[["isgood"]]) })
#--------------
test_that("PowerPoint: Adding ggplot Objects",{
tmp_obnd = report_add_slide(obnd,
  template = "content_text",
  elements = list(
     title         = list( content      = "Figures",
                           type         = "text"),
     sub_title     = list( content      = "Adding ggplot objects",
                           type         = "text"),
     content_body  = list( content      = p,
                           type         = "ggplot")))
expect_true(tmp_obnd[["isgood"]]) })

#--------------
test_that("PowerPoint: Adding Figures From Files",{
tmp_obnd = report_add_slide(obnd,
  template = "content_text",
  elements = list(
     title         = list( content      = "Figures",
                           type         = "text"),
     sub_title     = list( content      = "Adding an image from a file",
                           type         = "text"),
     content_body  = list( content      = imgfile,
                           type         = "imagefile")))
expect_true(tmp_obnd[["isgood"]]) })

#--------------
test_that("PowerPoint: Adding Lists",{

tmp_obnd = report_add_slide(obnd,
  template = "content_list",
  elements = list(
     title         = list( content      = "Lists",
                           type         = "text"),
     sub_title     = list( content      = "Creating a list",
                           type         = "text"),
     content_body  = list( content      = bl,
                           type         = "list")))
expect_true(tmp_obnd[["isgood"]]) })

#--------------
test_that("PowerPoint: Adding PowerPoint Tables",{
tmp_obnd = report_add_slide(obnd,
  template = "content_text",
  elements = list(
     title         = list( content      = "Tables",
                           type         = "text"),
     sub_title     = list( content      = "Creating PowerPoint Table",
                           type         = "text"),
     content_body  = list( content      = tab_cont,
                           type         = "table")))

expect_true(tmp_obnd[["isgood"]]) })
#--------------
test_that("PowerPoint: Adding flextable Objects",{
tmp_obnd = report_add_slide(obnd,
  template = "content_text",
  elements = list(
     title         = list( content      = "Tables",
                           type         = "text"),
     sub_title     = list( content      = "Creating PowerPoint Table",
                           type         = "text"),
     content_body  = list( content      = tab_fto,
                           type         = "flextable_object")))
expect_true(tmp_obnd[["isgood"]]) })
#--------------
test_that("PowerPoint: Adding flextable Using onbrand Interface",{
tmp_obnd = report_add_slide(obnd,
  template = "content_text",
  elements = list(
     title         = list( content      = "Tables",
                           type         = "text"),
     sub_title     = list( content      = "Creating PowerPoint Table",
                           type         = "text"),
     content_body  = list( content      = tab_ft,
                           type         = "flextable")))
expect_true(tmp_obnd[["isgood"]]) })

test_that("PowerPoint: Adding flextable Using onbrand Interface",{
 tmp_obnd = report_add_slide(obnd,
                          template = "content_text",
                          elements = list(
                            title         = list( content      = "Combining elements, user defined locations, and formatting",
                                                  type         = "text")),
                          user_location = list(
                            txt_example   =
                              list( content        = officer::fpar(officer::ftext("This is formatted text", officer::fp_text(color="green", font.size=24))),
                                    type           = "text",
                                    start          = c(0,  .25),
                                    stop           = c(.25,.35)),
                            large_figure  =
                              list( content        = p,
                                    type           = "ggplot",
                                    start          = c(.25,.25),
                                    stop           = c(.99,.99)),
                            flextable_obj =
                              list( content        = tab_fto,
                                    type           = "flextable_object",
                                    start          = c(0,.75),
                                    stop           = c(.25,.95)),
                            small_figure  =
                              list( content        = p,
                                    type           = "ggplot",
                                    start          = c(0,  .35),
                                    stop           = c(.25,.74))
                          )
  )
expect_true(tmp_obnd[["isgood"]]) })
