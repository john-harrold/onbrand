test_that("fph: valid template and placeholder", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fph(obnd, template = "content_text", pn = "title", verbose = FALSE)

  expect_true(res$isgood)
  expect_false(is.null(res$pl))
  expect_false(is.null(res$type))
})

test_that("fph: invalid template name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fph(obnd, template = "nonexistent_template", pn = "title", verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$pl)
  expect_true(any(grepl("nonexistent_template", res$msgs)))
})

test_that("fph: invalid placeholder name", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  res = fph(obnd, template = "content_text", pn = "nonexistent_ph", verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$pl)
  expect_true(any(grepl("nonexistent_ph", res$msgs)))
})

test_that("fph: bad onbrand object", {
  obnd = list(isgood = FALSE)

  res = fph(obnd, template = "content_text", pn = "title", verbose = FALSE)

  expect_false(res$isgood)
  expect_null(res$pl)
  expect_true(any(grepl("Bad onbrand object", res$msgs)))
})

test_that("fph: no meta in onbrand object", {
  obnd = list(isgood = TRUE)

  res = fph(obnd, template = "content_text", pn = "title", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("No mapping information", res$msgs)))
})

test_that("fph: no rpptx in meta", {
  obnd = list(isgood = TRUE, meta = list())

  res = fph(obnd, template = "content_text", pn = "title", verbose = FALSE)

  expect_false(res$isgood)
  expect_true(any(grepl("No PowerPoint mapping", res$msgs)))
})
