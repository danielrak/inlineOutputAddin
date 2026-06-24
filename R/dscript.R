#' Update inline output blocks in an R script (RStudio addin)
#'
#' Evaluates selected code (or the current top-level expression) and inserts or
#' replaces a marked output block delimited by `# >>> output` and
#' `# <<< output`.
#'
#' Code is evaluated in the global environment, so objects created while
#' building or analysing data persist between runs (exactly as if the code had
#' been sent to the console). Printed values, `cat()` output, messages and
#' warnings are captured; errors are reported in the block instead of
#' interrupting the editor. Code that prints nothing (such as a plain
#' assignment like `x <- c(1, 2)`) simply runs without inserting any block,
#' and any stale block left over from a previous run is removed.
#'
#' @param max_lines Maximum number of output lines to keep.
#' @param width Print width used when formatting output.
#'
#' @return Invisibly returns `TRUE` if an update was performed, otherwise
#'   `FALSE`.
#'
#' @export
#' @importFrom utils capture.output getFromNamespace
dscript <- function(max_lines = 60, width = 80) {
  if (!requireNamespace("rstudioapi", quietly = TRUE) ||
      !rstudioapi::isAvailable()) {
    stop("RStudio required for this addin.", call. = FALSE)
  }

  ctx <- rstudioapi::getActiveDocumentContext()
  if (length(ctx$selection) == 0) {
    return(invisible(FALSE))
  }

  sel <- ctx$selection[[1]]
  contents <- ctx$contents

  start_line <- sel$range$start[1]
  end_line <- sel$range$end[1]
  cursor_line <- start_line

  code <- sel$text
  use_selection <- !.ioa_is_blank_text(code)

  if (use_selection) {
    code <- paste(code, collapse = "\n")

    # If the selection itself contains a previously generated output block,
    # treat the code as ending just before that block so it gets replaced
    # rather than duplicated.
    block_in_sel <- .ioa_first_block_line(contents, start_line, end_line)
    if (!is.na(block_in_sel)) {
      end_line <- max(start_line, block_in_sel - 1L)
    }
  } else {
    bounds <- .ioa_find_expression_bounds(contents, cursor_line)
    start_line <- bounds[1]
    end_line <- bounds[2]
    code <- paste(contents[start_line:end_line], collapse = "\n")

    if (.ioa_is_blank_text(code)) {
      return(invisible(FALSE))
    }
  }

  code <- .ioa_strip_output_block(code)

  out <- .ioa_evaluate(code, width = width)
  out <- .ioa_split_lines(out)
  out <- .ioa_truncate(out, max_lines)

  # Locate an existing block immediately below the code (skipping blank lines).
  loc <- .ioa_locate_existing_block(contents, end_line)

  if (length(out) == 0) {
    # The code printed nothing to the console (e.g. `x <- c(1, 2)`): behave
    # exactly like running the code. Never insert an output block; only drop a
    # stale one left over from a previous run.
    if (!is.null(loc)) {
      del_start <- loc$start - 1L
      rng <- rstudioapi::document_range(
        rstudioapi::document_position(
          del_start, nchar(contents[[del_start]]) + 1
        ),
        rstudioapi::document_position(
          loc$end, nchar(contents[[loc$end]]) + 1
        )
      )
      rstudioapi::modifyRange(rng, "")
      return(invisible(TRUE))
    }
    return(invisible(FALSE))
  }

  block_text <- paste(.ioa_build_block(out), collapse = "\n")

  if (!is.null(loc)) {
    rng <- rstudioapi::document_range(
      rstudioapi::document_position(loc$start, 1),
      rstudioapi::document_position(loc$end, nchar(contents[[loc$end]]) + 1)
    )
    rstudioapi::modifyRange(rng, block_text)
    return(invisible(TRUE))
  }

  pos <- rstudioapi::document_position(
    end_line,
    nchar(contents[[end_line]]) + 1
  )
  rstudioapi::insertText(pos, paste0("\n", block_text, "\n"))

  invisible(TRUE)
}


# ---------------------------------------------------------------------------
# Internal helpers (not exported). Kept as standalone functions so the core
# logic can be unit tested without a running RStudio session.
# ---------------------------------------------------------------------------

#' @noRd
.ioa_is_blank_text <- function(x) {
  length(x) == 0 || !nzchar(trimws(paste(x, collapse = "\n")))
}

#' Split any multi-line strings into one element per line.
#' @noRd
.ioa_split_lines <- function(x) {
  if (length(x) == 0) {
    return(character())
  }
  unlist(
    lapply(x, function(s) {
      if (length(s) == 0 || is.na(s) || !grepl("\n", s, fixed = TRUE)) {
        return(s)
      }
      strsplit(s, "\n", fixed = TRUE)[[1]]
    }),
    use.names = FALSE
  )
}

#' Remove a marked output block (and inline `# >>> output` markers) from code.
#' @noRd
.ioa_strip_output_block <- function(code) {
  lines <- strsplit(paste(code, collapse = "\n"), "\n", fixed = TRUE)[[1]]
  if (length(lines) == 0) {
    return("")
  }

  keep <- rep(TRUE, length(lines))
  in_block <- FALSE
  for (k in seq_along(lines)) {
    trimmed <- trimws(lines[k])
    if (!in_block && grepl("^#\\s*>>>\\s*output\\s*$", trimmed)) {
      in_block <- TRUE
      keep[k] <- FALSE
      next
    }
    if (in_block) {
      keep[k] <- FALSE
      if (grepl("^#\\s*<<<\\s*output\\s*$", trimmed)) {
        in_block <- FALSE
      }
    }
  }
  lines <- lines[keep]

  # Strip an inline appended marker, e.g. `head(mtcars) # >>> output`.
  lines <- sub("\\s*#\\s*>>>\\s*output.*$", "", lines)

  paste(lines, collapse = "\n")
}

#' Index (1-based) of the first `# >>> output` marker within `contents`
#' between `from` and `to` inclusive, or `NA` if none.
#' @noRd
.ioa_first_block_line <- function(contents, from, to) {
  from <- max(1L, from)
  to <- min(length(contents), to)
  if (to < from) {
    return(NA_integer_)
  }
  for (k in from:to) {
    if (grepl("^#\\s*>>>\\s*output\\s*$", trimws(contents[[k]]))) {
      return(as.integer(k))
    }
  }
  NA_integer_
}

#' Find the bounds of the top-level expression covering `line_idx`.
#'
#' Uses source references from a full-document parse when possible, and falls
#' back to progressively smaller parses so that a syntax error elsewhere in the
#' script (e.g. an unfinished expression) does not prevent evaluation.
#' @noRd
.ioa_find_expression_bounds <- function(lines, line_idx, max_window = 200L) {
  n <- length(lines)
  if (n == 0) {
    return(c(line_idx, line_idx))
  }
  line_idx <- max(1L, min(n, line_idx))

  srcref_bounds <- function(exprs) {
    srcrefs <- attr(exprs, "srcref")
    if (is.null(srcrefs) || length(srcrefs) == 0) {
      return(NULL)
    }
    candidates <- lapply(srcrefs, function(ref) c(ref[[1]], ref[[3]]))
    hits <- vapply(
      candidates,
      function(x) isTRUE(x[1] <= line_idx && line_idx <= x[2]),
      logical(1)
    )
    if (!any(hits)) {
      return(NULL)
    }
    spans <- vapply(candidates[hits], function(x) x[2] - x[1], numeric(1))
    best <- candidates[hits][[which.min(spans)]]
    c(best[1], best[2])
  }

  # 1. Whole-document parse (the common, fully valid case).
  exprs <- tryCatch(
    parse(text = paste(lines, collapse = "\n"), keep.source = TRUE),
    error = function(e) NULL
  )
  if (!is.null(exprs)) {
    b <- srcref_bounds(exprs)
    if (!is.null(b)) {
      return(b)
    }
  }

  # 2. Valid prefix: drop trailing lines until the document parses, then look
  #    for the expression containing the cursor.
  if (n > line_idx) {
    for (end in n:line_idx) {
      exprs <- tryCatch(
        parse(text = paste(lines[seq_len(end)], collapse = "\n"),
              keep.source = TRUE),
        error = function(e) NULL
      )
      if (!is.null(exprs)) {
        b <- srcref_bounds(exprs)
        if (!is.null(b)) {
          return(b)
        }
        break
      }
    }
  }

  # 3. Growing window starting at the cursor (a new expression being typed).
  hi <- min(n, line_idx + max_window)
  for (end in line_idx:hi) {
    parsed <- tryCatch(
      parse(text = paste(lines[line_idx:end], collapse = "\n")),
      error = function(e) NULL
    )
    if (!is.null(parsed) && length(parsed) >= 1) {
      return(c(line_idx, end))
    }
  }

  c(line_idx, line_idx)
}

#' Make magrittr pipe operators available if they are not already resolvable.
#'
#' Operators are only injected when they cannot already be resolved from
#' `envir` (e.g. magrittr is installed but not attached). The names of any
#' operators actually injected are returned so the caller can remove them
#' again, leaving the environment exactly as it found it.
#' @noRd
.ioa_make_pipes_available <- function(envir) {
  injected <- character()
  if (!requireNamespace("magrittr", quietly = TRUE)) {
    return(injected)
  }
  ops <- c("%>%", "%T>%", "%$%", "%<>%")
  for (op in ops) {
    if (exists(op, envir = envir, inherits = TRUE)) {
      next
    }
    if (exists(op, envir = asNamespace("magrittr"), inherits = FALSE)) {
      assign(op, getFromNamespace(op, "magrittr"), envir = envir)
      injected <- c(injected, op)
    }
  }
  injected
}

#' Print a value, rendering plots to the active device but emitting a textual
#' placeholder (graphics produce no capturable text).
#' @noRd
.ioa_print_value <- function(value) {
  if (inherits(value, c("ggplot", "gg", "patchwork"))) {
    print(value)
    cat("<plot rendered>")
  } else {
    print(value)
  }
}

#' Evaluate a single expression, capturing output, messages and warnings.
#' @noRd
.ioa_eval_one <- function(expr, envir) {
  warnings <- character()
  messages <- character()
  err <- NULL

  text <- tryCatch(
    withCallingHandlers(
      capture.output({
        res <- withVisible(eval(expr, envir = envir))
        if (isTRUE(res$visible)) {
          .ioa_print_value(res$value)
        }
      }),
      warning = function(w) {
        warnings <<- c(warnings, conditionMessage(w))
        invokeRestart("muffleWarning")
      },
      message = function(m) {
        messages <<- c(messages, sub("\n$", "", conditionMessage(m)))
        invokeRestart("muffleMessage")
      }
    ),
    error = function(e) {
      err <<- conditionMessage(e)
      character()
    }
  )

  lines <- text
  if (length(messages) > 0) {
    lines <- c(lines, paste0("Message: ", .ioa_split_lines(messages)))
  }
  if (length(warnings) > 0) {
    lines <- c(lines, paste0("Warning: ", .ioa_split_lines(warnings)))
  }
  if (!is.null(err)) {
    lines <- c(lines, paste0("Error: ", .ioa_split_lines(err)))
  }

  list(lines = lines, error = !is.null(err))
}

#' Parse and evaluate `code`, returning the captured output as a character
#' vector (one element per line).
#' @noRd
.ioa_evaluate <- function(code, width = 80, envir = globalenv()) {
  exprs <- tryCatch(
    parse(text = code, keep.source = TRUE),
    error = function(e) structure(
      paste0("Error: ", conditionMessage(e)),
      class = "ioa_parse_error"
    )
  )

  old_opt <- options(width = max(20L, as.integer(width)))
  on.exit(options(old_opt), add = TRUE)

  if (inherits(exprs, "ioa_parse_error")) {
    return(.ioa_split_lines(unclass(exprs)))
  }
  if (length(exprs) == 0) {
    return(character())
  }

  # Pipe operators are made available only for the duration of this
  # evaluation and removed again afterwards (even on error), so running code
  # never leaves magrittr operators behind in the user's environment.
  injected_ops <- .ioa_make_pipes_available(envir)
  if (length(injected_ops) > 0) {
    on.exit(
      suppressWarnings(rm(list = injected_ops, envir = envir)),
      add = TRUE
    )
  }

  out <- character()
  for (expr in exprs) {
    res <- .ioa_eval_one(expr, envir)
    out <- c(out, res$lines)
    if (res$error) {
      break
    }
  }
  out
}

#' Truncate output to `max_lines`, appending a note when lines are dropped.
#' @noRd
.ioa_truncate <- function(out, max_lines) {
  max_lines <- max(1L, as.integer(max_lines))
  if (length(out) > max_lines) {
    out <- c(
      out[seq_len(max_lines)],
      sprintf("... (truncated to %d lines)", max_lines)
    )
  }
  out
}

#' Build the commented output block from a character vector of output lines.
#' @noRd
.ioa_build_block <- function(out) {
  out <- .ioa_split_lines(out)
  out <- sub("[ \t]+$", "", out)
  body <- if (length(out) > 0) {
    ifelse(nzchar(out), paste0("# ", out), "#")
  } else {
    "#"
  }
  c("# >>> output", body, "# <<< output")
}

#' Locate an existing output block immediately below `end_line` (skipping blank
#' lines and an inline marker). Returns a list with `start`/`end` line indices,
#' or `NULL` if no complete block is found.
#' @noRd
.ioa_locate_existing_block <- function(contents, end_line) {
  n <- length(contents)
  if (end_line < 1 || end_line > n) {
    return(NULL)
  }

  # Skip blank lines between the code and a candidate output block.
  i <- end_line + 1L
  while (i <= n && trimws(contents[[i]]) == "") {
    i <- i + 1L
  }

  has_start <- i <= n &&
    grepl("^#\\s*>>>\\s*output\\s*$", trimws(contents[[i]]))
  if (!has_start) {
    return(NULL)
  }

  j <- i + 1L
  while (j <= n &&
         !grepl("^#\\s*<<<\\s*output\\s*$", trimws(contents[[j]]))) {
    j <- j + 1L
  }
  if (j > n) {
    return(NULL)
  }

  list(start = i, end = j)
}
