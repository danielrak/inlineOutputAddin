test_that("bounds cover a single-line expression", {
  lines <- c("a <- 1", "b <- 2", "c <- 3")
  expect_equal(.ioa_find_expression_bounds(lines, 2), c(2, 2))
})

test_that("bounds span a multi-line call", {
  lines <- c(
    "data.frame(",
    "  a = 1:3,",
    "  b = 4:6",
    ")",
    "nrow(x)"
  )
  expect_equal(.ioa_find_expression_bounds(lines, 2), c(1, 4))
  expect_equal(.ioa_find_expression_bounds(lines, 5), c(5, 5))
})

test_that("bounds span a piped expression across lines", {
  lines <- c(
    "mtcars |>",
    "  head() |>",
    "  nrow()"
  )
  expect_equal(.ioa_find_expression_bounds(lines, 2), c(1, 3))
})

test_that("a syntax error elsewhere does not break bounds for valid code", {
  lines <- c(
    "x <- 1",
    "y <- function(z) {",   # unterminated function below the cursor
    "  z +"
  )
  # cursor on the first, valid statement
  expect_equal(.ioa_find_expression_bounds(lines, 1), c(1, 1))
})

test_that("a new multi-line expression at the bottom is detected", {
  lines <- c(
    "broken <- (",   # earlier broken expression
    "f(",
    "  1,",
    "  2",
    ")"
  )
  expect_equal(.ioa_find_expression_bounds(lines, 2), c(2, 5))
})
