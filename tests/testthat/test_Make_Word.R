#---
# Creating objects for the tests below
# Figures
p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "picture example")
imgfile = tempfile(pattern="image", fileext=".png")
ggplot2::ggsave(filename=imgfile, plot=p, height=5.15, width=9, units="in")

docxsavefile = tempfile(fileext=".docx")


# Bulleted list
bl = c("1", "This is first level bullet",
       "2", "sub-bullet",
       "3", "sub-sub-bullet",
       "1", "Another first level bullet")

# Tables
tdf =    data.frame(Parameters = c("Length", "Width", "Height"),
                    Values     = 1:3,
                    Units      = c("m", "m", "m") )
# Used for Word Tables
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


# fpar text
fpartext = officer::fpar(
officer::ftext("Formatted text can be created using the ", prop=NULL),
officer::ftext("fpar ", prop=officer::fp_text(color="green")),
officer::ftext("command from the officer package.", prop=NULL))
#--------------
# Running tests
#--------------
test_that("Word: Adding Plain Text",{

# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text="Text with no style specified will use the doc_def text format."))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text  ="First level header",
                           style = "Heading_1"))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text  ="Second level header",
                           style = "Heading_2"))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text  ="Third level header",
                           style = "Heading_3"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })

#--------------
test_that("Word: Adding fpar Text",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text   = fpartext,
                           format = "fpar",
                           style  = "Normal"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
test_that("Word: Adding Markdown Text",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_doc_content(obnd,
           type     = "text",
           content  = list(text   = "Formatted text can be created using **<color:green>markdown</color>** formatting",
                           format = "md",
                           style  = "Normal"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })

#--------------
test_that("Word: Creatign Placeholders",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_doc_content(obnd,
           type     = "ph",
           content  = list(name     = "BODYPHEXAMPLE",
                           value    = "placeholder text in the body",
                           location = "body"))

obnd = report_add_doc_content(obnd,
           type     = "ph",
           content  = list(name     = "FOOTERLEFT",
                           value    = "Footer Placeholder",
                           location = "footer"))

obnd = report_add_doc_content(obnd,
           type     = "ph",
           content  = list(name     = "HEADERLEFT",
                           value    = "Header Placeholder",
                           location = "header"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
test_that("Word: Creating Tables",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))


obnd = report_add_doc_content(obnd,
           type     = "flextable",
           content  = tab_ft)

obnd = report_add_doc_content(obnd,
           type     = "flextable_object",
           content  = list(ft=tab_fto,
                           caption  = "Flextable object created by the user."))

obnd = report_add_doc_content(obnd,
           type     = "table",
           content  = tab_cont)

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
test_that("Word: Creating Tables",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_doc_content(obnd,
           type     = "ggplot",
           content  = list(image   = p,
                           caption = "This is an example of an image from a ggplot object."))

obnd = report_add_doc_content(obnd,
           type     = "imagefile",
           content  = list(image   = imgfile,
                           caption = "This is an example of an image from a file."))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
test_that("Word: Markdown in Captions",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))


obnd = report_add_doc_content(obnd,
           type     = "ggplot",
           content  = list(image           = p,
                           caption_format  = "md",
                           caption         = "<shade:#33ff33>Markdown in a </shade><shade:#33ff33>*caption*.</shade>"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
test_that("Word: Markdown in flextables",{
# Loading the default Word template
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                      mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))


data = data.frame(property = c("mean",   "variance"),
                  length     = c(200,      0.13),
                  width      = c(12,       0.05),
                  area       = c(240,      0.11),
                  volume     = c(1200,     0.32))

header = list(property = c("",             ""),
              length   = c("Length",       "cm"),
              width    = c("Wdith",        "cm"),
              area     = c("Area",         "cm2"),
              volume   = c("Volume",       "cm3"))

ft = flextable::flextable(data)                     |> 
     flextable::delete_part(part = "header")        |> 
     flextable::add_header(values =as.list(header)) |> 
     flextable::theme_zebra()


dft      = fetch_md_def(obnd, style="Table_Labels")$md_def
dft_body = fetch_md_def(obnd, style="Table")$md_def


ft = ft |>  
  flextable::compose(j     = "area",
                     part  = "header",
                     value = c(md_to_oo("Area", dft)$oo, md_to_oo("cm^2^", dft)$oo))   |> 
  flextable::compose(j     = "volume",
                     part  = "header",
                     value = c(md_to_oo("Volume", dft)$oo, md_to_oo("cm^3^", dft)$oo)) |> 
  flextable::compose(j     = "property",
                     i     = match("mean", data$property),
                     part  = "body",
                     value = c(md_to_oo("**<ff:symbol>m</ff>**", dft_body)$oo))    |> 
  flextable::compose(j     = "property",
                     i     = match("variance", data$property),
                     part  = "body",
                     value = c(md_to_oo("**<ff:symbol>s</ff>**^**2**^", dft_body)$oo))

obnd = report_add_doc_content(obnd,
           type     = "flextable_object",
           content  = list(ft=ft,
                           caption  = "After applying formatting"))

sr = save_report(obnd, docxsavefile)
isgood = sr[["isgood"]] & obnd[["isgood"]]
expect_true(isgood) })
#--------------
