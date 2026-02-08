test_that("fetch_md_def: valid style from Word object", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_md_def(obnd, style = "default", verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$md_def))
})

test_that("fetch_md_def: valid style from PowerPoint object", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_md_def(obnd, style = "default", verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$md_def))
})

test_that("fetch_md_def: nonexistent style", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fetch_md_def(obnd, style = "nonexistent_style", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("nonexistent_style", res$msgs)))
})

test_that("fetch_md_def: bad onbrand object", {
  obnd = list(isgood = FALSE)

  res = fetch_md_def(obnd, style = "default", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("Bad onbrand object", res$msgs)))
})
