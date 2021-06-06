test_that("PowerPoint Layout",{
obnd = view_layout(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
             output_file   = file.path(tempdir(), "layout.pptx"))
expect_true(obnd[["isgood"]])})

test_that("Word Layout",{
obnd = view_layout(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
            output_file   = file.path(tempdir(), "layout.docx"))
expect_true(obnd[["isgood"]])})

test_that("PowerPoint Preview",{
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                     mapping = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
obnd = preview_template(obnd)

expect_true(obnd[["isgood"]])})


# JMH add Word Preview

