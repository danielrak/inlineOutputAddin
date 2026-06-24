test_that("evaluating a visible expression captures its print output", {
  out <- .ioa_evaluate("1 + 1", envir = new.env())
  expect_equal(out, "[1] 2")
})

test_that("invisible assignments produce no output", {
  e <- new.env()
  out <- .ioa_evaluate("x <- 41", envir = e)
  expect_equal(out, character())
  expect_equal(get("x", envir = e), 41)
})

test_that("assignments persist in the supplied environment across calls", {
  e <- new.env()
  .ioa_evaluate("df <- data.frame(a = 1:2)", envir = e)
  out <- .ioa_evaluate("nrow(df)", envir = e)
  expect_equal(out, "[1] 2")
})

test_that("cat() output is captured", {
  out <- .ioa_evaluate("cat('hello\\n')", envir = new.env())
  expect_equal(out, "hello")
})

test_that("multiple top-level expressions accumulate output", {
  out <- .ioa_evaluate("1\n2", envir = new.env())
  expect_equal(out, c("[1] 1", "[1] 2"))
})

test_that("data frames print with the requested width", {
  out <- .ioa_evaluate("head(mtcars, 2)", width = 80, envir = new.env())
  expect_true(any(grepl("Mazda RX4", out)))
})

test_that("warnings are reported and do not stop evaluation", {
  out <- .ioa_evaluate("as.integer('x'); 99", envir = new.env())
  expect_true(any(grepl("^Warning: ", out)))
  expect_true(any(out == "[1] 99"))
})

test_that("messages are reported", {
  out <- .ioa_evaluate("message('hi there'); 1", envir = new.env())
  expect_true(any(out == "Message: hi there"))
})

test_that("errors are captured and stop further evaluation", {
  out <- .ioa_evaluate("stop('boom')\n123", envir = new.env())
  expect_true(any(grepl("^Error: boom", out)))
  expect_false(any(out == "[1] 123"))
})

test_that("parse errors are reported as output", {
  out <- .ioa_evaluate("1 +", envir = new.env())
  expect_true(any(grepl("^Error: ", out)))
})

test_that("native pipe is supported", {
  out <- .ioa_evaluate("c(1, 2, 3) |> sum()", envir = new.env())
  expect_equal(out, "[1] 6")
})

test_that("magrittr pipe works when magrittr is attached", {
  skip_if_not_installed("magrittr")
  # Evaluate in an environment whose search path includes magrittr, mirroring
  # a session where the user has run library(magrittr).
  e <- new.env(parent = asNamespace("magrittr"))
  out <- .ioa_evaluate("c(1, 2, 3) %>% sum()", envir = e)
  expect_equal(out, "[1] 6")
})

test_that("an unavailable operator errors just like the console", {
  # baseenv() as parent keeps any attached magrittr out of scope, so `%>%`
  # is unresolvable and must error rather than being silently injected.
  e <- new.env(parent = baseenv())
  out <- .ioa_evaluate("c(1, 2) %>% sum()", envir = e)
  expect_true(any(grepl("^Error:", out)))
})

test_that("evaluation injects nothing into the environment", {
  e <- new.env(parent = baseenv())
  .ioa_evaluate("x <- 1", envir = e)
  expect_equal(ls(e, all.names = TRUE), "x")
})
