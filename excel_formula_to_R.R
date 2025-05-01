# excel_formula_to_r.R

# --- Gestion correcte du & ---
handle_ampersand <- function(f) {
  while (grepl("&", f)) {
    f <- sub(
      '(".*?"|[^&"]+)\\s*&\\s*(".*?"|[^&"]+)',
      'paste0(\\1, \\2)',
      f,
      perl = TRUE
    )
  }
  f
}

# Convertit une référence Excel A1 ou A1:B2 vers syntaxe R
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

# Convertit une formule Excel en code R
convert_formula <- function(form) {
  f <- sub("^=", "", form)
  f <- gsub("\\$", "", f)
  
  repeat {
    f_old <- f  # Sauvegarde avant transformation
    
    # IFERROR -> tryCatch
    f <- gsub(
      'IFERROR\\(([^()]+(?:\\([^()]*\\))?[^()]*)\\s*,\\s*([^()]*)\\)',
      'tryCatch({\\1}, error = function(e) {\\2})',
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
    
    # SUMIF -> sum(ifelse(...))
    f <- gsub(
      "SUMIF\\(\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+),\\s*\"?&?\"?([^,]+),\\s*\\$?([A-Z]+\\$?[0-9]+:[A-Z]+\\$?[0-9]+)\\)",
      "sum(ifelse(\\1 == \\2, \\3, 0))",
      f,
      perl = TRUE
    )
    
    # VLOOKUP exact
    f <- gsub(
      "VLOOKUP\\(([^,]+),\\s*'([^']+)'!\\$?([A-Z]+)\\$?([0-9]+):\\$?([A-Z]+)\\$?([0-9]+),\\s*([0-9]+),\\s*FALSE\\)",
      "vlookup_r(\\1, sheets[['\\2']][\\4:\\6, col2num('\\3'):col2num('\\5')], \\7)",
      f,
      perl = TRUE
    )
    
    # Gestion du &
    f <- handle_ampersand(f)
    
    # TEXT -> as.character
    f <- gsub(
      'TEXT\\(([^,]+?),\\s*"[^"]*"\\)',
      'as.character(\\1)',
      f,
      perl = TRUE
    )
    
    # Si rien n'a changé, on sort
    if (identical(f, f_old)) break
  }
  
  # Remplacer toutes les références (A1, A1:B2)
  refs <- unique(unlist(
    regmatches(f, gregexpr("([A-Z]+[0-9]+)(?::[A-Z]+[0-9]+)?", f, perl = TRUE))
  ))
  for (r in refs) {
    f <- gsub(r, convert_ref(r), f, fixed = TRUE)
  }
  
  # Remplacer = par ==
  f <- gsub("(?<!=)=(?!=)", "==", f, perl = TRUE)
  
  f
}

