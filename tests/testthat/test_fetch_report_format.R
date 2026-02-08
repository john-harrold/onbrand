test_that("fetch_report_format: default format from PowerPoint", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_report_format(obnd, verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$format_detials))
})

test_that("fetch_report_format: default format from Word", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_report_format(obnd, verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$format_detials))
})

test_that("fetch_report_format: non-existent format name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_report_format(obnd, format_name = "nonexistent_format", verbose = FALSE)

  expect_true(any(grepl("nonexistent_format", res$msgs)))
})

test_that("fetch_report_format: bad onbrand object", {
  obnd = list(isgood = FALSE, rpttype = "Word",
              meta = list(rdocx = list(md_def = list())))

  res = fetch_report_format(obnd, verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("Bad onbrand object", res$msgs)))
})
