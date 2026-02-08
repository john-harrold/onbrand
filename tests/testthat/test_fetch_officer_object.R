test_that("fetch_officer_object: PowerPoint object", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_officer_object(obnd, verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$rpt))
  expect_true(inherits(res$rpt, "rpptx"))
})

test_that("fetch_officer_object: Word object", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_officer_object(obnd, verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$rpt))
  expect_true(inherits(res$rpt, "rdocx"))
})

test_that("fetch_officer_object: bad onbrand object", {
  obnd = list(isgood = FALSE)

  res = fetch_officer_object(obnd, verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$rpt)
  expect_true(any(grepl("Bad onbrand object", res$msgs)))
})
