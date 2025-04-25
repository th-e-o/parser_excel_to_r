# parse_excel_formulas_with_assignments.R

# Dépendances
# install.packages(c("tidyxl","openxlsx","dplyr"))
library(tidyxl)
library(openxlsx)
library(dplyr)

# Utilitaires ----------------------------------------------------------------
# Convertit une colonne Excel ("AB") en numéro
col2num <- function(col_str) {
  letters <- strsplit(col_str, "")[[1]]
  sum((match(letters, LETTERS)) * 26^((length(letters) - 1):0))
}

# VLOOKUP exact comme Excel
vlookup_r <- function(lookup, df, col_index) {
  idx <- match(lookup, df[[1]])# parse_excel_formulas_with_assignments.R

# Dépendances
# install.packages(c("tidyxl","openxlsx","dplyr"))
library(tidyxl)
library(openxlsx)
library(dplyr)

# Utilitaires ----------------------------------------------------------------
# Convertit une colonne Excel ("AB") en numéro
col2num <- function(col_str) {
  letters <- strsplit(col_str, "")[[1]]
  sum((match(letters, LETTERS)) * 26^((length(letters) - 1):0))
}

# VLOOKUP exact comme Excel
vlookup_r <- function(lookup, df, col_index) {
  idx <- match(lookup, df[[1]])
  ifelse(is.na(idx), NA, df[[col_index]][idx])
}

# Fonction principale --------------------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  # 1) Charger toutes les feuilles
  wb_sheets <- getSheetNames(path)
  sheets <- setNames(
    lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)),
    wb_sheets
  )
  
  # 2) Extraire toutes les cellules (avec formules)
  cells_all <- xlsx_cells(path)
  form_cells <- cells_all %>% filter(!is.na(formula))
  
  # 3) Helpers ---------------------------------------------------------------
  # Convertit une référence Excel (A1 ou A1:B2) en valeurs debug
  convert_ref <- function(ref) {
    if (grepl(":", ref)) {
      parts <- strsplit(ref, ":")[[1]]
      start <- strcapture("^([A-Z]+)([0-9]+)$", parts[1], proto = list(C = "", R = 0))
      end   <- strcapture("^([A-Z]+)([0-9]+)$", parts[2], proto = list(C = "", R = 0))
      sprintf("values[%d:%d, %d:%d]",
              as.integer(start$R), as.integer(end$R),
              col2num(start$C),   col2num(end$C))
    } else {
      xy <- strcapture("^([A-Z]+)([0-9]+)$", ref, proto = list(C = "", R = 0))
      sprintf("values[%d, %d]",
              as.integer(xy$R), col2num(xy$C))
    }
  }
  
  # Convertit une formule Excel -> expression R
  convert_formula <- function(form) {
    f <- sub("^=", "", form)
    
    # IFERROR -> tryCatch
    f <- gsub(
      'IFERROR\\(([^,]+),([^)]*)\\)',
      'tryCatch({\1}, error = function(e) {\2})',
      f, perl = TRUE
    )
    # Fonctions Excel -> équivalents R
    func_map <- c(
      SUM = "sum", AVERAGE = "mean", IF = "ifelse",
      MIN = "min", MAX = "max", COUNT = "length",
      COUNTA = "length", ROUND = "round", CONCATENATE = "paste0"
    )
    for (xl in names(func_map)) {
      f <- gsub(paste0("\\b", xl, "\\b"), func_map[xl], f, ignore.case = TRUE)
    }
    # TEXT -> as.character
    f <- gsub(
      'TEXT\\(([^,]+),"[^"]*"\\)',
      'as.character(\1)',
      f, perl = TRUE
    )
    # VLOOKUP(..., FALSE)
    # SUMIF(range, critère, sum_range)
    f <- gsub(
      "SUMIF\\(\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+),[[:space:]]*\"?&?\"?([^,]+),[[:space:]]*\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+)\\)",
      "sum(ifelse(\1 == \2, \3, 0))",
      f,
      perl = TRUE
    )
    
    # Opérateur & -> paste0 & -> paste0 & -> paste0
    f <- gsub('([^[:space:]]+)\\&([^[:space:]]+)', 'paste0(\1, \2)', f, perl = TRUE)
    
    # Références A1/B2 -> convert_ref
    refs <- unique(unlist(
      regmatches(f, gregexpr("([A-Z]+[0-9]+)(?::[A-Z]+[0-9]+)?", f, perl = TRUE))
    ))
    for (r in refs) {
      f <- gsub(r, convert_ref(r), f, fixed = TRUE)
    }
    
    # Egalités simples -> ==
    f <- gsub("(?<!=)=(?!=)", "==", f, perl = TRUE)
    
    f
  }
  
  # 4) Appliquer conversion et extraire indices
  form_cells <- form_cells %>%
    mutate(
      R_code = vapply(formula, convert_formula, ""),
      row    = as.integer(sub("^[A-Z]+", "", address)),
      col    = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  
  # 5) Exporter formules brutes
  openxlsx::write.xlsx(
    form_cells %>% select(sheet, address, formula),
    file     = paste0(tools::file_path_sans_ext(basename(path)), "_raw_formulas.xlsx"),
    rowNames = FALSE
  )
  
  # 6) Générer script assignant
  if (emit_script) {
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    lines <- c(
      "# Script généré par parse_excel_formulas()",
      "library(tidyxl); library(openxlsx); library(dplyr)",
      deparse(col2num),
      deparse(vlookup_r),
      "# Charger classeur",
      sprintf("path <- '%s'", path),
      "wb_sheets <- getSheetNames(path)",
      "sheets <- setNames(lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)), wb_sheets)",
      ""
    )
    for (i in seq_len(nrow(form_cells))) {
      r <- form_cells[i, ]
      lines <- c(lines,
                 sprintf("# %s!%s -> %s", r$sheet, r$address, r$formula),
                 sprintf("sheets[['%s']][%d, %d] <- %s", r$sheet, r$row, r$col, r$R_code),
                 ""
      )
    }
    writeLines(lines, script_file)
    message("Script écrit dans : ", script_file)
  }
  
  invisible(form_cells)
}

# Exemple d'utilisation ------------------------------------------------------
parse_excel_formulas("mon_fichier.xlsm", emit_script = TRUE)
  ifelse(is.na(idx), NA, df[[col_index]][idx])
}

# Fonction principale --------------------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  # 1) Charger toutes les feuilles
  wb_sheets <- getSheetNames(path)
  sheets <- setNames(
    lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)),
    wb_sheets
  )
  
  # 2) Extraire toutes les cellules (avec formules)
  cells_all <- xlsx_cells(path)
  form_cells <- cells_all %>% filter(!is.na(formula))
  
  # 3) Helpers ---------------------------------------------------------------
  # Convertit une référence Excel (A1 ou A1:B2) en valeurs debug
  convert_ref <- function(ref) {
    if (grepl(":", ref)) {
      parts <- strsplit(ref, ":")[[1]]
      start <- strcapture("^([A-Z]+)([0-9]+)$", parts[1], proto = list(C = "", R = 0))
      end   <- strcapture("^([A-Z]+)([0-9]+)$", parts[2], proto = list(C = "", R = 0))
      sprintf("values[%d:%d, %d:%d]",
              as.integer(start$R), as.integer(end$R),
              col2num(start$C),   col2num(end$C))
    } else {
      xy <- strcapture("^([A-Z]+)([0-9]+)$", ref, proto = list(C = "", R = 0))
      sprintf("values[%d, %d]",
              as.integer(xy$R), col2num(xy$C))
    }
  }
  
  # Convertit une formule Excel -> expression R
  convert_formula <- function(form) {
    f <- sub("^=", "", form)
    
    # IFERROR -> tryCatch
    f <- gsub(
      'IFERROR\\(([^,]+),([^)]*)\\)',
      'tryCatch({\1}, error = function(e) {\2})',
      f, perl = TRUE
    )
    # Fonctions Excel -> équivalents R
    func_map <- c(
      SUM = "sum", AVERAGE = "mean", IF = "ifelse",
      MIN = "min", MAX = "max", COUNT = "length",
      COUNTA = "length", ROUND = "round", CONCATENATE = "paste0"
    )
    for (xl in names(func_map)) {
      f <- gsub(paste0("\\b", xl, "\\b"), func_map[xl], f, ignore.case = TRUE)
    }
    # TEXT -> as.character
    f <- gsub(
      'TEXT\\(([^,]+),"[^"]*"\\)',
      'as.character(\1)',
      f, perl = TRUE
    )
    # VLOOKUP(..., FALSE)
    # SUMIF(range, critère, sum_range)
    f <- gsub(
      "SUMIF\\(\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+),[[:space:]]*\"?&?\"?([^,]+),[[:space:]]*\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+)\\)",
      "sum(ifelse(\1 == \2, \3, 0))",
      f,
      perl = TRUE
    )
    
    # Opérateur & -> paste0 & -> paste0 & -> paste0
    f <- gsub('([^[:space:]]+)\\&([^[:space:]]+)', 'paste0(\1, \2)', f, perl = TRUE)
    
    # Références A1/B2 -> convert_ref
    refs <- unique(unlist(
      regmatches(f, gregexpr("([A-Z]+[0-9]+)(?::[A-Z]+[0-9]+)?", f, perl = TRUE))
    ))
    for (r in refs) {
      f <- gsub(r, convert_ref(r), f, fixed = TRUE)
    }
    
    # Egalités simples -> ==
    f <- gsub("(?<!=)=(?!=)", "==", f, perl = TRUE)
    
    f
  }
  
  # 4) Appliquer conversion et extraire indices
  form_cells <- form_cells %>%
    mutate(
      R_code = vapply(formula, convert_formula, ""),
      row    = as.integer(sub("^[A-Z]+", "", address)),
      col    = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  
  # 5) Exporter formules brutes
  openxlsx::write.xlsx(
    form_cells %>% select(sheet, address, formula),
    file     = paste0(tools::file_path_sans_ext(basename(path)), "_raw_formulas.xlsx"),
    rowNames = FALSE
  )
  
  # 6) Générer script assignant
  if (emit_script) {
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    lines <- c(
      "# Script généré par parse_excel_formulas()",
      "library(tidyxl); library(openxlsx); library(dplyr)",
      deparse(col2num),
      deparse(vlookup_r),
      "# Charger classeur",
      sprintf("path <- '%s'", path),
      "wb_sheets <- getSheetNames(path)",
      "sheets <- setNames(lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)), wb_sheets)",
      ""
    )
    for (i in seq_len(nrow(form_cells))) {
      r <- form_cells[i, ]
      lines <- c(lines,
                 sprintf("# %s!%s -> %s", r$sheet, r$address, r$formula),
                 sprintf("sheets[['%s']][%d, %d] <- %s", r$sheet, r$row, r$col, r$R_code),
                 ""
      )
    }
    writeLines(lines, script_file)
    message("Script écrit dans : ", script_file)
  }
  
  invisible(form_cells)
}

# Exemple d'utilisation ------------------------------------------------------
parse_excel_formulas("mon_fichier.xlsm", emit_script = TRUE)