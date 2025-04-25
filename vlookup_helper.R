
vlookup_r <- function(lookup, df, col_index) {
  idx <- match(lookup, df[[1]])
  ifelse(is.na(idx), NA, df[[col_index]][idx])
}
