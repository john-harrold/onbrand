# Helper to load a Word onbrand object
load_docx = function() {
  read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
}

# --- Success cases ---

test_that("Word: Adding page break", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "break", content = list())
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding table of contents", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "toc", content = list(level = 3))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding TOC with style", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "toc",
    content = list(level = 2, style = "TOC"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - continuous", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "text", content = list(text = "Before"))
  obnd = report_add_doc_content(obnd, type = "section",
    content = list(section_type = "continuous"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - landscape", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "section",
    content = list(section_type = "landscape", h = 8.5, w = 11))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - portrait", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "section",
    content = list(section_type = "portrait", h = 11, w = 8.5))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding section - columns", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "section",
    content = list(section_type = "columns", widths = c(3, 3)))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding text with fpar format", {
  obnd = load_docx()
  ftext_obj = officer::fpar(
    officer::ftext("Formatted ", prop = officer::fp_text(color = "red")),
    officer::ftext("text", prop = NULL))
  obnd = report_add_doc_content(obnd, type = "text",
    content = list(text = ftext_obj, format = "fpar", style = "Normal"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding text with default style (no style specified)", {
  obnd = load_docx()
  obnd = report_add_doc_content(obnd, type = "text",
    content = list(text = "Default style text"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding ggplot with caption", {
  obnd = load_docx()
  p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "test")
  obnd = report_add_doc_content(obnd, type = "ggplot",
    content = list(image = p, caption = "Test caption", key = "fig_test_1",
                   width = 4, height = 3))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding ggplot with md caption", {
  obnd = load_docx()
  p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "test")
  obnd = report_add_doc_content(obnd, type = "ggplot",
    content = list(image = p, caption = "**Bold** caption",
                   caption_format = "md", key = "fig_test_2"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding ggplot with notes", {
  obnd = load_docx()
  p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "test")
  obnd = report_add_doc_content(obnd, type = "ggplot",
    content = list(image = p, caption = "Fig with notes",
                   notes = "Some notes here", key = "fig_test_3"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding ggplot with md notes", {
  obnd = load_docx()
  p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "test")
  obnd = report_add_doc_content(obnd, type = "ggplot",
    content = list(image = p, caption = "Fig with md notes",
                   notes = "**Bold** notes", notes_format = "md",
                   key = "fig_test_4"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding table with caption", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "table",
    content = list(table = tdf, caption = "Test table", key = "tab_test_1"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding table with default style", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "table",
    content = list(table = tdf, header = FALSE, first_row = FALSE))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding flextable_object without caption", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  ft = flextable::flextable(tdf)
  obnd = report_add_doc_content(obnd, type = "flextable_object",
    content = list(ft = ft))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding flextable with headers", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "flextable",
    content = list(table = tdf,
                   header_top = list(A = "Col A", B = "Col B"),
                   cwidth = 0.8, table_autofit = TRUE,
                   table_theme = "theme_zebra"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding flextable with md headers", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "flextable",
    content = list(table = tdf,
                   header_top = list(A = "**A**", B = "**B**"),
                   header_format = "md",
                   cwidth = 0.8, table_autofit = TRUE))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Adding flextable with caption and notes", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "flextable",
    content = list(table = tdf,
                   header_top = list(A = "Col A", B = "Col B"),
                   caption = "Table caption",
                   notes = "Table notes",
                   key = "tab_test_2",
                   cwidth = 0.8))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Reusing reference key", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  # First use of key
  obnd = report_add_doc_content(obnd, type = "flextable",
    content = list(table = tdf,
                   header_top = list(A = "Col A", B = "Col B"),
                   caption = "First table",
                   key = "tab_reuse",
                   cwidth = 0.8))
  # Second use of same key
  obnd = report_add_doc_content(obnd, type = "flextable",
    content = list(table = tdf,
                   header_top = list(A = "Col A", B = "Col B"),
                   caption = "Second table",
                   key = "tab_reuse",
                   cwidth = 0.8))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Figure with tab_start_at and fig_start_at", {
  obnd = load_docx()
  p = ggplot2::ggplot() + ggplot2::annotate("text", x=0, y=0, label = "test")
  obnd = report_add_doc_content(obnd, type = "ggplot",
    content = list(image = p, caption = "Restart numbering",
                   key = "fig_restart"),
    fig_start_at = 5)
  expect_true(obnd[["isgood"]])
})

test_that("Word: Table with tab_start_at", {
  obnd = load_docx()
  tdf = data.frame(A = 1:3, B = 4:6)
  obnd = report_add_doc_content(obnd, type = "table",
    content = list(table = tdf, caption = "Table restart",
                   key = "tab_restart"),
    tab_start_at = 10)
  expect_true(obnd[["isgood"]])
})

# --- Error cases ---

test_that("Word: Error - PowerPoint object passed to doc_content", {
  obnd = read_template(
    template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
    mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
  expect_error(report_add_doc_content(obnd, type = "text",
    content = list(text = "test"), verbose = FALSE))
})

test_that("Word: Error - bad onbrand object", {
  obnd = list(isgood = FALSE)
  expect_error(report_add_doc_content(obnd, type = "text",
    content = list(text = "test"), verbose = FALSE))
})

test_that("Word: Error - NULL content", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "text",
    content = NULL, verbose = FALSE))
})

test_that("Word: Error - invalid section type", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "section",
    content = list(section_type = "invalid_type"), verbose = FALSE))
})

test_that("Word: Error - section without section_type", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "section",
    content = list(), verbose = FALSE))
})

test_that("Word: Error - invalid style name", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "text",
    content = list(text = "test", style = "Nonexistent_Style"), verbose = FALSE))
})

test_that("Word: Error - wrong style type for text (using table style)", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "text",
    content = list(text = "test", style = "Table"), verbose = FALSE))
})

test_that("Word: Error - unsupported content type", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "unsupported_type",
    content = list(text = "test"), verbose = FALSE))
})

test_that("Word: Error - non-existent image file", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "imagefile",
    content = list(image = "/nonexistent/path/image.png"), verbose = FALSE))
})

test_that("Word: Error - non-ggplot object as ggplot", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "ggplot",
    content = list(image = "not_a_ggplot"), verbose = FALSE))
})

test_that("Word: Error - non-data.frame as table", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "table",
    content = list(table = "not_a_dataframe"), verbose = FALSE))
})

test_that("Word: Error - placeholder missing fields", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "ph",
    content = list(name = "TEST"), verbose = FALSE))
})

test_that("Word: Error - placeholder invalid location", {
  obnd = load_docx()
  # Invalid location should still produce a message but location check
  # doesn't set isgood=FALSE, it just adds a message
  obnd = report_add_doc_content(obnd, type = "ph",
    content = list(name = "TEST", value = "val", location = "invalid_loc"))
  expect_true(obnd[["isgood"]])
})

test_that("Word: Error - flextable_object missing ft", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "flextable_object",
    content = list(caption = "No ft"), verbose = FALSE))
})

test_that("Word: Error - TOC with invalid style", {
  obnd = load_docx()
  expect_error(report_add_doc_content(obnd, type = "toc",
    content = list(style = "Nonexistent_Style"), verbose = FALSE))
})
