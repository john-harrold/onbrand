test_that("Load Default PowerPoint Template",{
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
                    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
expect_true(obnd[["isgood"]])})

test_that("Load Default Word Template",{
obnd = read_template(template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
                 mapping     = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
expect_true(obnd[["isgood"]])})
