# Utilitaires ----------------------------------------------------------------
col2num <- function(col_str) {
  letters <- strsplit(col_str, "")[[1]]
  sum((match(letters, LETTERS)) * 26^((length(letters) - 1):0))
}

vlookup_r <- function(lookup, df, col_index) {
  idx <- match(lookup, df[[1]])
  ifelse(is.na(idx), NA, df[[col_index]][idx])
}

extract_deps <- function(form, sheet_name) {
  pattern <- "(?:'([^']+)'!)?([A-Z]+[0-9]+)"
  m <- gregexpr(pattern, form, perl = TRUE)
  refs <- regmatches(form, m)[[1]]
  deps <- vapply(refs, function(x) {
    parts <- regmatches(x, regexec(pattern, x, perl = TRUE))[[1]]
    sh <- if (nzchar(parts[2])) parts[2] else sheet_name
    paste0(sh, "!", parts[3])
  }, character(1))
  unique(deps)
}

topo_order <- function(ids, deps_list) {
  remaining <- setNames(deps_list, ids)
  result <- character(0)
  while (length(remaining) > 0) {
    ready <- names(remaining)[vapply(remaining, length, integer(1)) == 0]
    if (length(ready) == 0) stop("Boucle détectée dans les dépendances !")
    result <- c(result, ready)
    remaining[ready] <- NULL
    for (i in seq_along(remaining)) {
      remaining[[i]] <- setdiff(remaining[[i]], ready)
    }
  }
  result
}

evaluate_cells <- function(sheets, form_cells) {
  ids <- paste0(form_cells$sheet, "!", form_cells$address)
  order_ids <- topo_order(ids, form_cells$deps)
  pos <- setNames(
    Map(function(sh, addr, row, col) {
      list(sheet = sh, row = row, col = col)
    }, form_cells$sheet, form_cells$address, form_cells$row, form_cells$col),
    ids
  )
  for (id in order_ids) {
    info <- pos[[id]]
    code <- form_cells$R_code[ids == id]
    val <- eval(parse(text = code))
    sheets[[info$sheet]][info$row, info$col] <- val
  }
  sheets
}