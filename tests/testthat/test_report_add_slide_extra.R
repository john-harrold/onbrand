test_that("PowerPoint: Error - missing template", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_slide(obnd,
      template = NULL,
      elements = list(
        title = list(content = "Test", type = "text")),
      verbose = FALSE))
})

test_that("PowerPoint: Error - invalid template name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_slide(obnd,
      template = "nonexistent_template",
      elements = list(
        title = list(content = "Test", type = "text")),
      verbose = FALSE))
})

test_that("PowerPoint: Error - Word object passed to slide function", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_slide(obnd,
      template = "title_slide",
      elements = list(
        title = list(content = "Test", type = "text")),
      verbose = FALSE))
})

test_that("PowerPoint: Error - bad onbrand object", {
  obnd = list(isgood = FALSE)

  expect_error(
    report_add_slide(obnd,
      template = "title_slide",
      elements = list(
        title = list(content = "Test", type = "text")),
      verbose = FALSE))
})

test_that("PowerPoint: Error - neither elements nor user_location provided", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  expect_error(
    report_add_slide(obnd,
      template = "title_slide",
      verbose = FALSE))
})
