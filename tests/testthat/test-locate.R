test_that("strip_output_block removes a full embedded block", {
  code <- paste(
    "head(mtcars)",
    "# >>> output",
    "# a b c",
    "# <<< output",
    sep = "\n"
  )
  expect_equal(.ioa_strip_output_block(code), "head(mtcars)")
})

test_that("strip_output_block removes an inline appended marker", {
  expect_equal(
    .ioa_strip_output_block("head(mtcars) # >>> output"),
    "head(mtcars)"
  )
})

test_that("strip_output_block keeps ordinary comments", {
  code <- "x <- 1 # a comment"
  expect_equal(.ioa_strip_output_block(code), "x <- 1 # a comment")
})

test_that("locate finds a block immediately below the code", {
  contents <- c("head(mtcars)", "# >>> output", "# a", "# <<< output")
  loc <- .ioa_locate_existing_block(contents, end_line = 1)
  expect_equal(loc, list(start = 2, end = 4))
})

test_that("locate skips blank lines between code and block", {
  contents <- c("x", "", "", "# >>> output", "# 1", "# <<< output")
  loc <- .ioa_locate_existing_block(contents, end_line = 1)
  expect_equal(loc, list(start = 4, end = 6))
})

test_that("locate returns NULL when there is no block", {
  contents <- c("x", "y", "z")
  expect_null(.ioa_locate_existing_block(contents, end_line = 1))
})

test_that("locate returns NULL for an unterminated block", {
  contents <- c("x", "# >>> output", "# 1")
  expect_null(.ioa_locate_existing_block(contents, end_line = 1))
})

test_that("locate does not match a block separated by a non-blank line", {
  contents <- c("x", "# a normal comment", "# >>> output", "# 1", "# <<< output")
  expect_null(.ioa_locate_existing_block(contents, end_line = 1))
})

test_that("first_block_line finds a block within a range", {
  contents <- c("a", "b", "# >>> output", "# 1", "# <<< output")
  expect_equal(.ioa_first_block_line(contents, 1, 5), 3L)
  expect_true(is.na(.ioa_first_block_line(contents, 1, 2)))
})
