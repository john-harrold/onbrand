#--- mk_lg_tbl tests ---

test_that("mk_lg_tbl: returns correct structure", {
  res = mk_lg_tbl()

  expect_true(is.list(res))
  expect_true(all(c("lg_tbl_body", "lg_tbl_row_common",
                     "lg_tbl_body_head", "lg_tbl_row_common_head") %in% names(res)))

  expect_equal(nrow(res$lg_tbl_body), 51)
  expect_equal(ncol(res$lg_tbl_body), 63)
  expect_equal(nrow(res$lg_tbl_row_common), 51)
  expect_equal(ncol(res$lg_tbl_row_common), 2)
  expect_equal(nrow(res$lg_tbl_body_head), 3)
  expect_equal(ncol(res$lg_tbl_body_head), 63)
  expect_equal(nrow(res$lg_tbl_row_common_head), 3)
  expect_equal(ncol(res$lg_tbl_row_common_head), 2)
})

test_that("mk_lg_tbl: contains BQL and NC markers", {
  res = mk_lg_tbl()

  # Row 1 should all be BQL
  expect_true(all(res$lg_tbl_body[1, ] == "BQL"))
  # Row 20 should all be BQL
  expect_true(all(res$lg_tbl_body[20, ] == "BQL"))
  # Row 5, col 8 should be NC
  expect_equal(res$lg_tbl_body[5, 8], "NC")
  # Row 25, col 2 should be NC
  expect_equal(res$lg_tbl_body[25, 2], "NC")
})

#--- span_table validation tests ---

test_that("span_table: missing table_body", {
  row_common = data.frame(ID = 1:5)

  res = span_table(table_body = NULL, row_common = row_common)

  expect_false(res$isgood)
  expect_true(any(grepl("table_body must be specified", res$msgs)))
})

test_that("span_table: missing row_common", {
  table_body = data.frame(A = 1:5, B = 6:10)

  res = span_table(table_body = table_body, row_common = NULL)

  expect_false(res$isgood)
  expect_true(any(grepl("row_common must be specified", res$msgs)))
})

test_that("span_table: mismatched rows between table_body and row_common", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:3)

  res = span_table(table_body = table_body, row_common = row_common)

  expect_false(res$isgood)
  expect_true(any(grepl("different number of rows", res$msgs)))
})

test_that("span_table: mismatched columns table_body_head vs table_body", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:5)
  table_body_head = data.frame(A = "A", B = "B", C = "C")  # 3 cols vs 2
  row_common_head = data.frame(ID = "ID")

  res = span_table(table_body = table_body, row_common = row_common,
                   table_body_head = table_body_head,
                   row_common_head = row_common_head)

  expect_false(res$isgood)
  expect_true(any(grepl("table_body and table_body_head have a different number of columns", res$msgs)))
})

test_that("span_table: mismatched columns row_common_head vs row_common", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:5)
  table_body_head = data.frame(A = "A", B = "B")
  row_common_head = data.frame(ID = "ID", Extra = "X")  # 2 cols vs 1

  res = span_table(table_body = table_body, row_common = row_common,
                   table_body_head = table_body_head,
                   row_common_head = row_common_head)

  expect_false(res$isgood)
  expect_true(any(grepl("row_common and row_common_head have a different number of columns", res$msgs)))
})

test_that("span_table: only row_common_head specified (not table_body_head)", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:5)
  row_common_head = data.frame(ID = "ID")

  res = span_table(table_body = table_body, row_common = row_common,
                   row_common_head = row_common_head)

  expect_false(res$isgood)
  expect_true(any(grepl("You must specify both or neither", res$msgs)))
})

test_that("span_table: only table_body_head specified (not row_common_head)", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:5)
  table_body_head = data.frame(A = "A", B = "B")

  res = span_table(table_body = table_body, row_common = row_common,
                   table_body_head = table_body_head)

  expect_false(res$isgood)
  expect_true(any(grepl("You must specify both or neither", res$msgs)))
})

test_that("span_table: mismatched rows in headers", {
  table_body = data.frame(A = 1:5, B = 6:10)
  row_common = data.frame(ID = 1:5)
  table_body_head = data.frame(A = c("A", "a"), B = c("B", "b"))  # 2 rows
  row_common_head = data.frame(ID = "ID")  # 1 row

  res = span_table(table_body = table_body, row_common = row_common,
                   table_body_head = table_body_head,
                   row_common_head = row_common_head)

  expect_false(res$isgood)
  expect_true(any(grepl("same number of rows", res$msgs)))
})

#--- span_table functional tests ---
# Note: row_common needs 2+ columns to avoid R's single-column drop behavior

test_that("span_table: small table fits in one table", {
  table_body = data.frame(A = 1:3, B = 4:6, C = 7:9)
  row_common = data.frame(ID = c("S1", "S2", "S3"), GRP = c("X", "X", "Y"))

  res = span_table(table_body = table_body, row_common = row_common,
                   max_row = 20, max_col = 10)

  expect_true(res$isgood)
  expect_equal(length(res$tables), 1)
  expect_false(is.null(res$one_table))
  expect_false(is.null(res$one_body))
})

test_that("span_table: table with headers", {
  table_body = data.frame(A = 1:3, B = 4:6, C = 7:9)
  row_common = data.frame(ID = c("S1", "S2", "S3"), GRP = c("X", "X", "Y"))
  table_body_head = data.frame(A = "Col A", B = "Col B", C = "Col C")
  row_common_head = data.frame(ID = "Subject", GRP = "Group")

  res = span_table(table_body = table_body, row_common = row_common,
                   table_body_head = table_body_head,
                   row_common_head = row_common_head,
                   max_row = 20, max_col = 10)

  expect_true(res$isgood)
  expect_equal(length(res$tables), 1)
  expect_false(is.null(res$one_header))
})

test_that("span_table: pagination by rows", {
  table_body = data.frame(A = 1:10, B = 11:20, C = 21:30)
  row_common = data.frame(ID = 1:10, GRP = rep(c("X", "Y"), 5))

  res = span_table(table_body = table_body, row_common = row_common,
                   max_row = 5, max_col = 10)

  expect_true(res$isgood)
  expect_equal(length(res$tables), 2)
})

test_that("span_table: pagination by columns", {
  table_body = data.frame(A = 1:3, B = 4:6, C = 7:9,
                          D = 10:12, E = 13:15, F = 16:18,
                          G = 19:21, H = 22:24, I = 25:27)
  row_common = data.frame(ID = c("S1", "S2", "S3"), GRP = c("X", "X", "Y"))

  res = span_table(table_body = table_body, row_common = row_common,
                   max_row = 20, max_col = 5)

  expect_true(res$isgood)
  expect_true(length(res$tables) >= 2)
})

test_that("span_table: notes_detect finds markers", {
  table_body = data.frame(A = c("1", "BQL", "3"),
                          B = c("NC", "5", "6"),
                          C = c("7", "8", "9"))
  row_common = data.frame(ID = c("S1", "S2", "S3"), GRP = c("X", "X", "Y"))

  res = span_table(table_body = table_body, row_common = row_common,
                   notes_detect = c("BQL", "NC"))

  expect_true(res$isgood)
  expect_true("BQL" %in% res$tables[["Table 1"]]$notes)
  expect_true("NC" %in% res$tables[["Table 1"]]$notes)
})

test_that("span_table: NULL max_row and max_col use full dimensions", {
  table_body = data.frame(A = 1:3, B = 4:6, C = 7:9)
  row_common = data.frame(ID = 1:3, GRP = c("X", "X", "Y"))

  res = span_table(table_body = table_body, row_common = row_common,
                   max_row = NULL, max_col = NULL,
                   max_height = NULL, max_width = NULL)

  expect_true(res$isgood)
  expect_equal(length(res$tables), 1)
})

test_that("span_table: large table with mk_lg_tbl", {
  lg = mk_lg_tbl()

  res = span_table(
    table_body      = lg$lg_tbl_body,
    row_common      = lg$lg_tbl_row_common,
    table_body_head = lg$lg_tbl_body_head,
    row_common_head = lg$lg_tbl_row_common_head,
    max_row = 20, max_col = 10,
    notes_detect = c("BQL", "NC"))

  expect_true(res$isgood)
  expect_true(length(res$tables) > 1)
})

test_that("span_table: border options", {
  table_body = data.frame(A = 1:3, B = 4:6, C = 7:9)
  row_common = data.frame(ID = 1:3, GRP = c("X", "X", "Y"))

  res = span_table(table_body = table_body, row_common = row_common,
                   set_header_inner_border_v = FALSE,
                   set_header_inner_border_h = FALSE,
                   set_header_outer_border   = FALSE,
                   set_body_inner_border_v   = FALSE,
                   set_body_inner_border_h   = TRUE,
                   set_body_outer_border     = FALSE)

  expect_true(res$isgood)
})
