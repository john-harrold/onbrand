test_that("save_report: bad onbrand object", {
  obnd = list(isgood = FALSE)

  expect_error(save_report(obnd, tempfile(fileext = ".pptx"), verbose = FALSE))
})

test_that("save_report: NULL output_file", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(save_report(obnd, output_file = NULL, verbose = FALSE))
})

test_that("save_report: mismatched file extension - pptx object to docx file", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(save_report(obnd, tempfile(fileext = ".docx"), verbose = FALSE))
})

test_that("save_report: mismatched file extension - docx object to pptx file", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(save_report(obnd, tempfile(fileext = ".pptx"), verbose = FALSE))
})

test_that("save_report: invalid file extension", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(save_report(obnd, tempfile(fileext = ".pdf"), verbose = FALSE))
})

test_that("save_report: Word post_processing - successful", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  # Set post_processing to valid R code
  obnd[["meta"]][["rdocx"]][["post_processing"]] = "x = 1"

  sr = save_report(obnd, tempfile(fileext = ".docx"), verbose = FALSE)
  expect_true(sr$isgood)
})

test_that("save_report: Word post_processing - failure", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  # Set post_processing to invalid R code
  obnd[["meta"]][["rdocx"]][["post_processing"]] = "stop('test error')"

  expect_error(save_report(obnd, tempfile(fileext = ".docx"), verbose = FALSE))
})

test_that("save_report: PowerPoint post_processing - successful", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  # Set post_processing to valid R code
  obnd[["meta"]][["rpptx"]][["post_processing"]] = "x = 1"

  sr = save_report(obnd, tempfile(fileext = ".pptx"), verbose = FALSE)
  expect_true(sr$isgood)
})

test_that("save_report: PowerPoint post_processing - failure", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  # Set post_processing to invalid R code
  obnd[["meta"]][["rpptx"]][["post_processing"]] = "stop('test error')"

  expect_error(save_report(obnd, tempfile(fileext = ".pptx"), verbose = FALSE))
})

test_that("save_report: verbose error messages", {
  obnd = list(isgood = FALSE)

  expect_error(
    expect_message(
      save_report(obnd, tempfile(fileext = ".pptx"), verbose = TRUE),
      "Bad onbrand object"))
})
