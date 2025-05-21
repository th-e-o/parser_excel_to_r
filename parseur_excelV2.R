# parse_excel_formulas_with_assignments.R

library(tidyxl)
library(openxlsx)
library(dplyr)

source("~/work/excel_formula_to_R.R")
source("~/work/utils.R")



# Fonction principale --------------------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  message("[parse_excel_formulas] Chargement du fichier Excel...")
  wb_sheets <- getSheetNames(path)
  sheets <- setNames(
    lapply(wb_sheets, function(sh) {
      message(sprintf(" - Lecture de la feuille '%s'", sh))
      read.xlsx(path, sheet = sh, colNames = FALSE)
    }), wb_sheets
  )
  message("[parse_excel_formulas] Extraction des variables globales")
  globals <- get_excel_globals(path)
  nom_cellules <- build_named_cell_map(globals)
  print(nom_cellules)

  

  message("[parse_excel_formulas] Extraction des cellules contenant des formules...")
  cells_all <- xlsx_cells(path)
  
  form_cells <- cells_all %>% filter(!is.na(formula))
  message(sprintf("[parse_excel_formulas] %d formules extraites.", nrow(form_cells)))
  message("[parse_excel_formulas] Vérification des variables globales utilisées dans les formules...")
  check_globals_usage(formules = form_cells$formula, noms_cellules = nom_cellules)
  
  message("[parse_excel_formulas] Extraction des dépendances...")
  form_cells <- form_cells %>%
    mutate(
      deps = purrr::pmap(list(formula, sheet), extract_deps)
    )
  
  message("[parse_excel_formulas] Filtrage des formules contenant 'LEFT'...")
  initial_count <- nrow(form_cells)
  form_cells <- form_cells %>%
    filter(!grepl("\\bLEFT\\b", formula, ignore.case = TRUE))
  message(sprintf("[parse_excel_formulas] %d formules restantes après filtrage (supprimées: %d).", 
                  nrow(form_cells), initial_count - nrow(form_cells)))
  
  message("[parse_excel_formulas] Conversion des formules Excel en code R...")
  form_cells <- form_cells %>%
    mutate(
      R_code = vapply(formula, function(f) convert_formula(f, nom_cellules), character(1)),
      row    = as.integer(sub("^[A-Z]+", "", address)),
      col    = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  
  message("[parse_excel_formulas] Conversion terminée.")
  
  if (!emit_script) {
    message("[parse_excel_formulas] Evaluation des cellules dans le bon ordre...")
    sheets <- evaluate_cells(sheets, form_cells)
  }
  
  message("[parse_excel_formulas] Export des formules brutes...")
  openxlsx::write.xlsx(
    form_cells %>% select(sheet, address, formula),
    file = paste0(tools::file_path_sans_ext(basename(path)), "_raw_formulas.xlsx"),
    rowNames = FALSE
  )
  
  if (emit_script) {
    message("[parse_excel_formulas] Génération du script R...")
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
                 # on définit values avant chaque ligne
                 sprintf("values <- sheets[['%s']]", r$sheet),
                 sprintf("sheets[['%s']][%d, %d] <- %s", r$sheet, r$row, r$col, r$R_code),
                 ""
      )
    }
    
    
    writeLines(lines, script_file)
    message("Script écrit dans : ", script_file)
  }
  
  message("[parse_excel_formulas] Terminé.")
  invisible(form_cells)
}


# Exemple d'appel
parse_excel_formulas("mon_fichier.xlsx", emit_script = TRUE)
