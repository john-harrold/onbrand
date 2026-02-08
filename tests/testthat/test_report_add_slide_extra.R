# Helper to load a PowerPoint onbrand object
load_pptx = function() {
  read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
}

# --- Error cases ---

test_that("PowerPoint: Error - missing template", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = NULL,
    elements = list(title = list(content = "Test", type = "text")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - invalid template name", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "nonexistent_template",
    elements = list(title = list(content = "Test", type = "text")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - Word object passed to slide function", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = "Test", type = "text")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - bad onbrand object", {
  obnd = list(isgood = FALSE)
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = "Test", type = "text")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - neither elements nor user_location provided", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "title_slide",
    verbose = FALSE))
})

test_that("PowerPoint: Error - wrong content type for placeholder", {
  obnd = load_pptx()
  # title placeholder is content_type = "text", so passing "list" should fail
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = c("1", "a", "2", "b"),
                                 type = "list")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - element not found in template mapping", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(nonexistent_element = list(content = "Test", type = "text")),
    verbose = FALSE))
})

test_that("PowerPoint: Error - user_location x start > x stop", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = "Test", type = "text")),
    user_location = list(
      bad_loc = list(content = "text", type = "text",
                     start = c(0.8, 0.1), stop = c(0.2, 0.5))),
    verbose = FALSE))
})

test_that("PowerPoint: Error - user_location y start > y stop", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = "Test", type = "text")),
    user_location = list(
      bad_loc = list(content = "text", type = "text",
                     start = c(0.1, 0.8), stop = c(0.5, 0.2))),
    verbose = FALSE))
})

test_that("PowerPoint: Error - user_location non-numeric start/stop", {
  obnd = load_pptx()
  expect_error(report_add_slide(obnd, template = "title_slide",
    elements = list(title = list(content = "Test", type = "text")),
    user_location = list(
      bad_loc = list(content = "text", type = "text",
                     start = c("a", "b"), stop = c("c", "d"))),
    verbose = FALSE))
})
