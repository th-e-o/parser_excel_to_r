# parse_excel_formulas_with_assignments.R

# Dépendances
# install.packages(c("tidyxl", "openxlsx", "dplyr"))
library(tidyxl)
library(openxlsx)
library(dplyr)

# Source des helpers
source("~/work/excel_formula_to_R.R")
source("~/work/utils.R")


# Fonction principale --------------------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  wb_sheets <- getSheetNames(path)
  sheets <- setNames(lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)), wb_sheets)
  
  cells_all <- xlsx_cells(path)
  form_cells <- cells_all %>%
    filter(!is.na(formula)) %>%
    mutate(
      deps = purrr::pmap(list(formula, sheet), extract_deps)
    ) %>%
    filter(!grepl("\\bLEFT\\b", formula, ignore.case = TRUE)) %>%
    mutate(
      R_code = vapply(formula, convert_formula, ""),
      row    = as.integer(sub("^[A-Z]+", "", address)),
      col    = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  
  if (!emit_script) {
    sheets <- evaluate_cells(sheets, form_cells)
  }
  
  openxlsx::write.xlsx(
    form_cells %>% select(sheet, address, formula),
    file = paste0(tools::file_path_sans_ext(basename(path)), "_raw_formulas.xlsx"),
    rowNames = FALSE
  )
  
  if (emit_script) {
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    lines <- c(
      "# Script généré par parse_excel_formulas()",
      "library(tidyxl); library(openxlsx); library(dplyr)",
      deparse(col2num),
      deparse(vlookup_r),
      "# Charger le classeur",
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

# Exemple
parse_excel_formulas("mon_fichier.xlsm", emit_script = TRUE)
