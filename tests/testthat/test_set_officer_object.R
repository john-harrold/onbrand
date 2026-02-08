test_that("set_officer_object: round-trip PowerPoint", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  rpt = fetch_officer_object(obnd, verbose = FALSE)$rpt
  obnd2 = set_officer_object(obnd, rpt, verbose = FALSE)

  expect_true(obnd2[["isgood"]])
})

test_that("set_officer_object: round-trip Word", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

  rpt = fetch_officer_object(obnd, verbose = FALSE)$rpt
  obnd2 = set_officer_object(obnd, rpt, verbose = FALSE)

  expect_true(obnd2[["isgood"]])
})

test_that("set_officer_object: bad onbrand object", {
  obnd = list(isgood = FALSE)

  obnd2 = set_officer_object(obnd, NULL, verbose = FALSE)

  expect_false(obnd2[["isgood"]])
})
