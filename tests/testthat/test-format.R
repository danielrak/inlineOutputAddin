test_that("build_block wraps output in markers and comments each line", {
  expect_equal(
    .ioa_build_block(c("[1] 1", "[1] 2")),
    c("# >>> output", "# [1] 1", "# [1] 2", "# <<< output")
  )
})

test_that("build_block produces a single bare comment for empty output", {
  expect_equal(
    .ioa_build_block(character()),
    c("# >>> output", "#", "# <<< output")
  )
})

test_that("blank output lines become a bare comment without trailing space", {
  block <- .ioa_build_block(c("a", "", "b"))
  expect_equal(block, c("# >>> output", "# a", "#", "# b", "# <<< output"))
  expect_false(any(grepl("[ \t]+$", block)))
})

test_that("trailing whitespace in output is stripped", {
  block <- .ioa_build_block("value   ")
  expect_equal(block, c("# >>> output", "# value", "# <<< output"))
})

test_that("split_lines expands embedded newlines", {
  expect_equal(.ioa_split_lines("a\nb\nc"), c("a", "b", "c"))
  expect_equal(.ioa_split_lines(character()), character())
  expect_equal(.ioa_split_lines(c("a", "b\nc")), c("a", "b", "c"))
})

test_that("truncate keeps max_lines and adds a note", {
  out <- as.character(1:100)
  tr <- .ioa_truncate(out, 10)
  expect_equal(length(tr), 11)
  expect_match(tr[11], "truncated to 10 lines")
})

test_that("truncate leaves short output untouched", {
  expect_equal(.ioa_truncate(c("a", "b"), 60), c("a", "b"))
})
