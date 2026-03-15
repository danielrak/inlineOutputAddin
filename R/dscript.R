#' Update inline output blocks in an R script (RStudio addin)
#'
#' Evaluates selected code (or the current line) and inserts or replaces a
#' marked output block delimited by `# >>> output` and `# <<< output`.
#'
#' @param max_lines Maximum number of output lines to keep.
#' @param width Print width used when formatting output.
#' @return Invisibly returns TRUE if an update was performed, otherwise FALSE.
#' @export
#' @importFrom utils capture.output
dscript <- function(max_lines = 60, width = 80) {
  if (!rstudioapi::isAvailable()) {
    stop("RStudio required for this addin.", call. = FALSE)
  }

  ctx <- rstudioapi::getActiveDocumentContext()
  sel <- ctx$selection[[1]]
  contents <- ctx$contents

  cursor_line <- sel$range$start[1]
  start_line <- sel$range$start[1]
  end_line <- sel$range$end[1]

  # Helper: find full top-level expression around the cursor if nothing is selected
  find_expression_bounds <- function(lines, line_idx) {
    n <- length(lines)

    is_expr_complete <- function(txt) {
      if (!nzchar(trimws(txt))) return(FALSE)
      res <- tryCatch(parse(text = txt), error = function(e) NULL)
      !is.null(res)
    }

    # Expand upward
    start <- line_idx
    while (start > 1) {
      candidate <- paste(lines[start - 1:line_idx], collapse = "\n")
      parsed <- tryCatch(parse(text = candidate), error = function(e) NULL)
      if (is.null(parsed)) {
        start <- start - 1
      } else {
        break
      }
    }

    # Expand downward until parse succeeds
    end <- line_idx
    repeat {
      candidate <- paste(lines[start:end], collapse = "\n")
      if (is_expr_complete(candidate)) break
      if (end >= n) break
      end <- end + 1
    }

    # If still not valid, try a more exhaustive local search
    candidate <- paste(lines[start:end], collapse = "\n")
    if (!is_expr_complete(candidate)) {
      best <- NULL
      for (s in seq_len(line_idx)) {
        for (e in seq(line_idx, n)) {
          txt <- paste(lines[s:e], collapse = "\n")
          if (is_expr_complete(txt)) {
            best <- c(s, e)
          }
        }
      }
      if (!is.null(best)) {
        start <- best[1]
        end <- best[2]
      }
    }

    c(start, end)
  }

  code <- sel$text

  # If nothing selected, detect the whole expression around cursor
  if (!nzchar(code)) {
    bounds <- find_expression_bounds(contents, cursor_line)
    start_line <- bounds[1]
    end_line <- bounds[2]
    code <- paste(contents[start_line:end_line], collapse = "\n")

    if (!nzchar(trimws(code))) return(invisible(FALSE))
  }

  # Remove inline marker if present on last line of code
  code <- sub("\\s*#\\s*>>>\\s*output.*$", "", code)

  old_opt <- options(width = width)
  on.exit(options(old_opt), add = TRUE)

  out <- tryCatch(
    {
      val <- eval(parse(text = code), envir = .GlobalEnv)
      capture.output(print(val))
    },
    error = function(e) paste0("Error: ", conditionMessage(e))
  )

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

  # ---- FIND EXISTING BLOCK ----

  # Case 1: inline marker on the last line of the code block
  same_line_has_marker <- grepl("#\\s*>>>\\s*output", contents[[end_line]])

  if (same_line_has_marker) {
    i <- end_line + 1
  } else {
    # Case 2: marker below code block
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

  # ---- NO EXISTING BLOCK -> INSERT ----

  # Insert at end of last code line, forcing a newline before the block
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
