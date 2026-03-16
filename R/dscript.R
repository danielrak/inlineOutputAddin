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

  build_eval_env <- function() {
    eval_env <- new.env(parent = .GlobalEnv)

    if (requireNamespace("magrittr", quietly = TRUE)) {
      assign("%>%", getFromNamespace("%>%", "magrittr"), envir = eval_env)

      if (exists("%T>%", envir = asNamespace("magrittr"), inherits = FALSE)) {
        assign("%T>%", getFromNamespace("%T>%", "magrittr"), envir = eval_env)
      }

      if (exists("%$%", envir = asNamespace("magrittr"), inherits = FALSE)) {
        assign("%$%", getFromNamespace("%$%", "magrittr"), envir = eval_env)
      }

      if (exists("%<>%", envir = asNamespace("magrittr"), inherits = FALSE)) {
        assign("%<>%", getFromNamespace("%<>%", "magrittr"), envir = eval_env)
      }
    }

    eval_env
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
    eval_env <- build_eval_env()

    for (expr in exprs) {
      warnings <- character()
      messages <- character()
      value <- NULL
      visible <- FALSE

      expr_out <- tryCatch(
        withCallingHandlers(
          capture.output({
            res <- withVisible(eval(expr, envir = eval_env))
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

    if (length(out) == 0) {
      out <- ""
    }
  }

  if (length(out) > max_lines) {
    out <- c(
      out[seq_len(max_lines)],
      sprintf("... (truncated to %d lines)", max_lines)
    )
  }

  block <- c(
    "# >>> output",
    if (length(out)) paste0("# ", out) else "#",
    "# <<< output"
  )

  insert_at <- end_line + 1

  same_line_has_marker <- grepl("#\\s*>>>\\s*output", contents[[end_line]])

  if (same_line_has_marker) {
    i <- end_line + 1
  } else {
    i <- insert_at
    while (i <= length(contents) && trimws(contents[[i]]) == "") {
      i <- i + 1
    }
  }

  has_start <- i <= length(contents) &&
    grepl("^#\\s*>>>\\s*output\\s*$", trimws(contents[[i]]))

  if (has_start) {
    j <- i + 1
    while (j <= length(contents) &&
           !grepl("^#\\s*<<<\\s*output\\s*$", trimws(contents[[j]]))) {
      j <- j + 1
    }

    if (j <= length(contents)) {
      rng <- rstudioapi::document_range(
        rstudioapi::document_position(i, 1),
        rstudioapi::document_position(j, nchar(contents[[j]]) + 1)
      )
      rstudioapi::modifyRange(rng, paste(block, collapse = "\n"))
      return(invisible(TRUE))
    }
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
