#' Update inline output blocks in an R script (RStudio addin)
#'
#' Evaluates selected code (or the current top-level expression) and inserts or
#' replaces a marked output block delimited by `# >>> output` and `# <<< output`.
#'
#' @param max_lines Maximum number of output lines to keep.
#' @param width Print width used when formatting output.
#'
#' @return Invisibly returns `TRUE` if an update was performed, otherwise `FALSE`.
#'
#' @export
#' @importFrom utils capture.output getFromNamespace
dscript <- function(max_lines = 60, width = 80) {
  if (!rstudioapi::isAvailable()) {
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

  is_blank_text <- function(x) {
    length(x) == 0 || !nzchar(trimws(paste(x, collapse = "\n")))
  }

  # Split a string into individual physical lines so that the leading `# `
  # comment prefix applies to every line. Condition messages (and some
  # errors/warnings) carry embedded or trailing newlines which, left intact,
  # would leak uncommented text into the `.R` script.
  split_lines <- function(x) {
    x <- gsub("\r", "", x, fixed = TRUE)
    x <- sub("\n+$", "", x)
    if (!nzchar(x)) {
      return("")
    }
    strsplit(x, "\n", fixed = TRUE)[[1]]
  }

  find_expression_bounds <- function(lines, line_idx) {
    full_text <- paste(lines, collapse = "\n")

    exprs <- tryCatch(
      parse(text = full_text, keep.source = TRUE),
      error = function(e) NULL
    )

    if (is.null(exprs) || length(exprs) == 0) {
      return(c(line_idx, line_idx))
    }

    srcrefs <- attr(exprs, "srcref")
    if (is.null(srcrefs) || length(srcrefs) == 0) {
      return(c(line_idx, line_idx))
    }

    candidates <- lapply(srcrefs, function(ref) c(ref[[1]], ref[[3]]))

    hits <- vapply(
      candidates,
      function(x) isTRUE(x[1] <= line_idx && line_idx <= x[2]),
      logical(1)
    )

    if (!any(hits)) {
      return(c(line_idx, line_idx))
    }

    spans <- vapply(candidates[hits], function(x) x[2] - x[1], numeric(1))
    best <- candidates[hits][[which.min(spans)]]

    c(best[1], best[2])
  }

  # Evaluate in the global environment so the addin behaves exactly like
  # running the code (assignments and other side effects persist). magrittr
  # pipe operators are temporarily made available when the package is
  # installed but not attached, then removed again on exit.
  inject_pipe_operators <- function() {
    injected <- character()

    if (requireNamespace("magrittr", quietly = TRUE)) {
      ns <- asNamespace("magrittr")
      ops <- c("%>%", "%T>%", "%$%", "%<>%")

      for (op in ops) {
        if (exists(op, envir = ns, inherits = FALSE) &&
            !exists(op, envir = .GlobalEnv, inherits = TRUE)) {
          assign(op, getFromNamespace(op, "magrittr"), envir = .GlobalEnv)
          injected <- c(injected, op)
        }
      }
    }

    injected
  }

  code <- sel$text
  use_selection <- !is_blank_text(code)

  if (use_selection) {
    code <- paste(code, collapse = "\n")
  } else {
    bounds <- find_expression_bounds(contents, cursor_line)
    start_line <- bounds[1]
    end_line <- bounds[2]
    code <- paste(contents[start_line:end_line], collapse = "\n")

    if (is_blank_text(code)) {
      return(invisible(FALSE))
    }
  }

  code <- sub("\\s*#\\s*>>>\\s*output.*$", "", code)

  exprs <- tryCatch(
    parse(text = code, keep.source = TRUE),
    error = function(e) structure(
      paste0("Error: ", conditionMessage(e)),
      class = "dscript_parse_error"
    )
  )

  old_opt <- options(width = width)
  on.exit(options(old_opt), add = TRUE)

  out <- character()

  if (inherits(exprs, "dscript_parse_error")) {
    out <- unclass(exprs)
  } else {
    injected_ops <- inject_pipe_operators()
    if (length(injected_ops)) {
      on.exit(
        suppressWarnings(rm(list = injected_ops, envir = .GlobalEnv)),
        add = TRUE
      )
    }

    for (expr in exprs) {
      warnings <- character()
      messages <- character()
      value <- NULL
      visible <- FALSE

      expr_out <- tryCatch(
        withCallingHandlers(
          capture.output({
            res <- withVisible(eval(expr, envir = .GlobalEnv))
            value <- res$value
            visible <- isTRUE(res$visible)

            if (visible) {
              print(value)
            }
          }),
          warning = function(w) {
            warnings <<- c(warnings, conditionMessage(w))
            invokeRestart("muffleWarning")
          },
          message = function(m) {
            messages <<- c(messages, conditionMessage(m))
            invokeRestart("muffleMessage")
          }
        ),
        error = function(e) structure(
          paste0("Error: ", conditionMessage(e)),
          class = "dscript_eval_error"
        )
      )

      if (inherits(expr_out, "dscript_eval_error")) {
        out <- c(out, unclass(expr_out))
        break
      }

      if (length(messages) > 0) {
        expr_out <- c(expr_out, paste0("Message: ", messages))
      }

      if (length(warnings) > 0) {
        expr_out <- c(expr_out, paste0("Warning: ", warnings))
      }

      out <- c(out, expr_out)
    }

  }

  # Normalise every entry to a single physical line so the `# ` comment
  # prefix below applies to all of it (messages, warnings and multi-line
  # errors may otherwise contain embedded newlines).
  out <- unlist(lapply(out, split_lines), use.names = FALSE)
  if (is.null(out)) {
    out <- character()
  }

  if (length(out) > max_lines) {
    out <- c(
      out[seq_len(max_lines)],
      sprintf("... (truncated to %d lines)", max_lines)
    )
  }

  has_output <- length(out) > 0

  # Locate an existing output block immediately following the code.
  same_line_has_marker <- grepl("#\\s*>>>\\s*output", contents[[end_line]])

  if (same_line_has_marker) {
    i <- end_line + 1
  } else {
    i <- end_line + 1
    while (i <= length(contents) && trimws(contents[[i]]) == "") {
      i <- i + 1
    }
  }

  has_start <- i <= length(contents) &&
    grepl("^#\\s*>>>\\s*output\\s*$", trimws(contents[[i]]))

  block_end <- NA_integer_
  if (has_start) {
    j <- i + 1
    while (j <= length(contents) &&
           !grepl("^#\\s*<<<\\s*output\\s*$", trimws(contents[[j]]))) {
      j <- j + 1
    }

    if (j <= length(contents)) {
      block_end <- j
    }
  }

  has_block <- !is.na(block_end)

  if (!has_output) {
    # The code printed nothing to the console (e.g. `x <- c(1, 2)`): behave
    # exactly like running the code. Never insert an output block; only drop
    # a stale one left over from a previous run.
    if (has_block) {
      del_start_row <- i - 1
      rng <- rstudioapi::document_range(
        rstudioapi::document_position(
          del_start_row,
          nchar(contents[[del_start_row]]) + 1
        ),
        rstudioapi::document_position(
          block_end,
          nchar(contents[[block_end]]) + 1
        )
      )
      rstudioapi::modifyRange(rng, "")
      return(invisible(TRUE))
    }

    return(invisible(FALSE))
  }

  block <- c(
    "# >>> output",
    paste0("# ", out),
    "# <<< output"
  )

  if (has_block) {
    rng <- rstudioapi::document_range(
      rstudioapi::document_position(i, 1),
      rstudioapi::document_position(block_end, nchar(contents[[block_end]]) + 1)
    )
    rstudioapi::modifyRange(rng, paste(block, collapse = "\n"))
    return(invisible(TRUE))
  }

  pos <- rstudioapi::document_position(
    end_line,
    nchar(contents[[end_line]]) + 1
  )

  rstudioapi::insertText(
    pos,
    paste0("\n", paste(block, collapse = "\n"), "\n")
  )

  invisible(TRUE)
}
