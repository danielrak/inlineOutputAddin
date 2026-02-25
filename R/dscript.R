#' Update (or insert) an inline output block under selected code / current line.
#'
#' @importFrom utils capture.output
#'
#' @param max_lines Maximum lines live-printed within your current script
#' @param width Print width within your current script
#'
#' Markers used:
#'   # >>> output
#'   # <<< output
#'
#' Behavior:
#' - If an output block exists immediately below the code, it is replaced.
#' - Otherwise, a new block is inserted directly below.
#'
#' Notes:
#' - Evaluates in .GlobalEnv (like typical interactive use).
#' - Captures printed output; errors are captured as "Error: ...".
dscript <- function(max_lines = 60, width = 80) {
  if (!rstudioapi::isAvailable()) stop("RStudio required for this addin.")

  ctx <- rstudioapi::getActiveDocumentContext()
  sel <- ctx$selection[[1]]
  contents <- ctx$contents

  start_line <- sel$range$start[1]
  end_line   <- sel$range$end[1]

  code <- sel$text

  # If nothing selected, use current line
  if (!nzchar(code)) {
    code <- contents[[start_line]]
    if (!nzchar(trimws(code))) return(invisible(FALSE))
  }

  # Remove inline marker if user wrote:
  # head(cars) # >>> output
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
    out <- c(out[seq_len(max_lines)],
             sprintf("... (truncated to %d lines)", max_lines))
  }

  block <- c(
    "# >>> output",
    if (length(out)) paste0("# ", out) else "#",
    "# <<< output"
  )

  insert_at <- end_line + 1

  # ---- FIND EXISTING BLOCK ----

  # Case 1: marker is on same line
  same_line_has_marker <- grepl("#\\s*>>>\\s*output", contents[[start_line]])

  if (same_line_has_marker) {
    i <- start_line + 1
  } else {
    # Case 2: marker is below (skip blank lines)
    i <- insert_at
    while (i <= length(contents) && trimws(contents[[i]]) == "") i <- i + 1
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

  # ---- NO EXISTING BLOCK → INSERT ----

  pos <- rstudioapi::document_position(insert_at, 1)
  rstudioapi::insertText(pos, paste0(paste(block, collapse = "\n"), "\n"))
  invisible(TRUE)
}
