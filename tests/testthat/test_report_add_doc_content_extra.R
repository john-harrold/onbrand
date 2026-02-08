test_that("Word: Adding page break", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "break",
    content = list())

  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding table of contents", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "toc",
    content = list(level = 3))

  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - continuous", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "text",
    content = list(text = "Before section"))

  obnd = report_add_doc_content(obnd,
    type    = "section",
    content = list(section_type = "continuous"))

  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - landscape", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "section",
    content = list(section_type = "landscape",
                   h = 8.5, w = 11))

  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - portrait", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "section",
    content = list(section_type = "portrait",
                   h = 11, w = 8.5))

  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - columns", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  obnd = report_add_doc_content(obnd,
    type    = "section",
    content = list(section_type = "columns",
                   widths = c(3, 3)))

  expect_true(obnd[["isgood"]])
})

test_that("Word: Error - PowerPoint object passed to doc_content", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "text",
      content = list(text = "test"),
      verbose = FALSE))
})

test_that("Word: Error - bad onbrand object", {
  obnd = list(isgood = FALSE)

  expect_error(
    report_add_doc_content(obnd,
      type    = "text",
      content = list(text = "test"),
      verbose = FALSE))
})

test_that("Word: Error - invalid section type", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "section",
      content = list(section_type = "invalid_type"),
      verbose = FALSE))
})

test_that("Word: Error - invalid style name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "text",
      content = list(text = "test", style = "Nonexistent_Style"),
      verbose = FALSE))
})

test_that("Word: Error - non-existent image file", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "imagefile",
      content = list(image = "/nonexistent/path/image.png"),
      verbose = FALSE))
})

test_that("Word: Error - non-ggplot object as ggplot", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "ggplot",
      content = list(image = "not_a_ggplot"),
      verbose = FALSE))
})

test_that("Word: Error - non-data.frame as table", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_doc_content(obnd,
      type    = "table",
      content = list(table = "not_a_dataframe"),
      verbose = FALSE))
})

test_that("Word: Adding text with fpar format", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  ftext_obj = officer::fpar(
    officer::ftext("Formatted ", prop = officer::fp_text(color = "red")),
    officer::ftext("text", prop = NULL))

  obnd = report_add_doc_content(obnd,
    type    = "text",
    content = list(text   = ftext_obj,
                   format = "fpar",
                   style  = "Normal"))

  expect_true(obnd[["isgood"]])
})
