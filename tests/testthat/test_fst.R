test_that("fst: valid style name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fst(obnd, osn = "Heading_1", verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$wsn))
  expect_false(is.null(res$dff))
})

test_that("fst: invalid style name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fst(obnd, osn = "nonexistent_style", verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$wsn)
  expect_true(any(grepl("nonexistent_style", res$msgs)))
})

test_that("fst: bad onbrand object", {
  obnd = list(isgood = FALSE)

  res = fst(obnd, osn = "Heading_1", verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$wsn)
  expect_true(any(grepl("Bad onbrand object", res$msgs)))
})

test_that("fst: no meta in onbrand object", {
  obnd = list(isgood = TRUE)

  res = fst(obnd, osn = "Heading_1", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("No mapping information", res$msgs)))
})

test_that("fst: no rdocx in meta", {
  obnd = list(isgood = TRUE, meta = list())

  res = fst(obnd, osn = "Heading_1", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("No Word mapping", res$msgs)))
})
