# excel2R_refactored.R
# -----------------------------------
# Convertisseur de formules Excel en code R
# Architecture : nettoyage → split générique (multi-char ops) → parsing récursif → codegen via table de fonctions

# 1) UTILITAIRES DE SPLITTING -------------------------------

# Découpe f en deux parties au premier opérateur de ops au niveau racine,
# en gérant d'abord les opérateurs à deux caractères puis à un seul.
split_at_top <- function(f, ops) {
  chars <- strsplit(f, "", fixed = TRUE)[[1]]
  depth <- 0L
  i <- 1L
  while (i <= length(chars)) {
    ch <- chars[i]
    if (ch == "(") {
      depth <- depth + 1L
    } else if (ch == ")") {
      depth <- depth - 1L
    } else if (depth == 0L) {
      # tenter les opérateurs à 2 caractères
      two <- if (i < length(chars)) paste0(chars[i], chars[i+1]) else ""
      if (two %in% ops) {
        left  <- trimws(substr(f, 1, i-1))
        right <- trimws(substr(f, i+2, nchar(f)))
        return(list(left = left, op = two, right = right))
      }
      # puis opérateurs à 1 caractère
      if (ch %in% ops) {
        left  <- trimws(substr(f, 1, i-1))
        right <- trimws(substr(f, i+1, nchar(f)))
        return(list(left = left, op = ch, right = right))
      }
    }
    i <- i + 1L
  }
  NULL
}

# Scinde une chaîne s sur les virgules de profondeur 0
split_top <- function(s) {
  if (length(s) != 1 || is.na(s) || !nzchar(s)) return(character())
  chars <- strsplit(s, "")[[1]]
  depth <- 0L; args <- character(); curr <- ""
  for (ch in chars) {
    if (ch == "(")      depth <- depth + 1L
    else if (ch == ")") depth <- depth - 1L
    else if (ch == "," && depth == 0L) {
      args <- c(args, curr); curr <- ""; next
    }
    curr <- paste0(curr, ch)
  }
  c(args, curr)
}

# Nettoie les parenthèses externes appariées
clean_parentheses <- function(f) {
  # First check if the string starts and ends with parentheses
  if (!grepl("^\\(.*\\)$", f)) return(f)
  
  # Now check for balanced parentheses
  chars <- strsplit(f, "", fixed = TRUE)[[1]]
  depth <- cumsum(chars == "(") - cumsum(chars == ")")
  
  # Only strip if the last parenthesis closes the first one
  if (depth[length(depth)] == 0 && all(depth[-length(depth)] > 0)) {
    return(substr(f, 2, nchar(f) - 1))
  }
  
  return(f)
}

# 2) UTILITAIRES DE RÉFÉRENCES ----------------------------------

# Conversion de colonne Excel ("A", "AB") en numéro
col2num <- function(col) {
  letters <- strsplit(col, "")[[1]]
  sum(sapply(seq_along(letters), function(i) {
    (match(letters[i], LETTERS)) * 26^(length(letters) - i)
  }))
}

# Convertit une référence Excel en indices R
convert_ref <- function(ref) {
  # Extraire le nom de feuille si présent
  if (grepl("!", ref)) {
    parts <- strsplit(ref, "!", fixed = TRUE)[[1]]
    sheet <- parts[1]
    ref <- parts[2]
  } else {
    sheet <- NULL
  }
  
  # Conversion des coordonnées
  if (grepl(":", ref, fixed = TRUE)) {
    prts <- strsplit(ref, ":", fixed = TRUE)[[1]]
    s <- strcapture("^([A-Z]+)([0-9]+)$", prts[1], proto = list(C = "", R = 0))
    e <- strcapture("^([A-Z]+)([0-9]+)$", prts[2], proto = list(C = "", R = 0))
    ref_str <- sprintf("%d:%d, %d",
                       as.integer(s$R), as.integer(e$R),
                       col2num(s$C))
  } else {
    xy <- strcapture("^([A-Z]+)([0-9]+)$", ref, proto = list(C = "", R = 0))
    ref_str <- sprintf("%d, %d", 
                       as.integer(xy$R), col2num(xy$C))
  }
  ref_str <- gsub(",[0-9]+:[0-9]+", "", ref_str)
  if (!is.null(sheet)) {
    sprintf("sheets[['%s']][%s]", sheet, ref_str)
  } else {
    sprintf("values[%s]", ref_str)
  }
}
offset_r <- function(addr, row_off, col_off) {
  m     <- regexec("^([A-Z]+)([0-9]+)$", addr)
  parts <- regmatches(addr, m)[[1]]
  r0    <- as.integer(parts[3]) + as.integer(row_off)
  c0    <- col2num(parts[2]) + as.integer(col_off)
  sprintf("values[%d, %d]", r0, c0)
}


convert_criteria <- function(crit, noms_cellules = NULL) {
  crit <- str_trim(crit)
  
  # 1) Si l’utilisateur a écrit =… en début, c’est juste une égalité simple
  if (str_starts(crit, "=")) {
    return(convert_criteria(sub("^=", "", crit), noms_cellules))
  }
  
  # 2) Cas concaténation Excel "=\"&C7&H8" ou ="&C7&H8
  if (str_starts(crit, '"="&') || str_starts(crit, '="&')) {
    # On enlève précisément le préfixe '"="&' ou '="&'
    reste <- if (str_starts(crit, '"="&')) {
      sub('^"="&', "", crit)
    } else {
      sub('^="&', "", crit)
    }
    
    # On découpe sur & pour isoler chaque morceau, puis on convertit
    parts <- str_split(reste, fixed("&"))[[1]] %>% str_trim()
    parts_r <- lapply(parts, function(x) {
      if (str_detect(x, '^".*"$')) {
        sub('^"(.*)"$', '\\1', x)
      } else {
        convert_formula(x, noms_cellules)
      }
    })
    
    # On reconstruit l’expression R de concaténation
    val_expr <- paste0("paste0(", paste(parts_r, collapse = ","), ")")
    
    # On renvoie la seule liste correspondant à un compare pour SUMIF
    return(list(type = "compare", op = "==", value = val_expr))
  }
  
  # 3) Comparateurs classiques >, <, !=, <=, >=
  if (str_detect(crit, '^[<>]=?|^!=')) {
    op  <- str_extract(crit, '^[<>]=?|^!=')
    val <- sub('^[<>]=?|^!=', '', crit)
    return(list(type = "compare", op = op, value = val))
  }
  
  # 4) Nombre pur
  if (str_detect(crit, '^[0-9.]+$')) {
    return(list(type = "number", value = crit))
  }
  
  # 5) Référence isolée (A1, B2, etc.)
  if (str_detect(crit, '^[A-Za-z]+[0-9]+$')) {
    return(list(type = "ref", value = convert_formula(crit, noms_cellules)))
  }
  
  # 6) Wildcards Excel (si besoin)…
  # if (str_detect(crit, "[*?]")) { … }
  
  # 7) Tout le reste, c’est du texte simple
  txt <- str_replace_all(crit, '^"|"$', '')
  return(list(type = "text", value = paste0('"', txt, '"')))
}


# 3) TABLE DE MAPPING DES FONCTIONS ----------------------------------
fun_map <- list(
  SUM = function(args, noms_cellules = NULL) paste0("sum(", paste(args, collapse = ","), ", na.rm = TRUE)"),
  IF = function(args, noms_cellules = NULL) paste0("ifelse(", paste(args, collapse = ","), ")"),
  SUMIF = function(args, noms_cellules = NULL, raw_args) {
    if (length(args) < 2) stop("SUMIF requires at least 2 arguments")
    range     <- args[[1]]
    crit_raw  <- raw_args[[2]]          # <-- on prend le 2ᵉ argument EXCEL brut
    sum_range <- if (length(args) >= 3) args[[3]] else range
    
    crit <- convert_criteria(crit_raw, noms_cellules)
    expr_test <- switch(
      crit$type,
      compare  = sprintf("(%s %s %s)", range, crit$op, crit$value),
      wildcard = sprintf("grepl(%s, %s)", crit$pattern, range),
      sprintf("(%s == %s)", range, crit$value)
    )
    paste0("sum((", expr_test, ")*", sum_range, ", na.rm=TRUE)")
  },
  IFERROR = function(args, noms_cellules = NULL) sprintf("tryCatch({%s}, error=function(e){%s})", args[1], args[2]),
  TEXT = function(args, noms_cellules = NULL) paste0("as.character(", args[1], ")"),
  ROUND = function(args, noms_cellules = NULL) {
    d <- if (length(args) >= 2) args[2] else "0"
    paste0("round(", args[1], ",", d, ")")
  },
  INT = function(args, noms_cellules = NULL) paste0("floor(", args[1], ")"),
  DATE = function(args, noms_cellules = NULL) paste0("as.Date(ISOdate(", paste(args, collapse = ","), "))"),
  OFFSET = function(args, noms_cellules = NULL) paste0("offset_r('", args[1], "',", args[2], ",", args[3], ")")
)

vlookup_r <- function(x, table, col) {
  table[x, col]
}

convert_formula <- function(form, noms_cellules = list()) {
  print(form)
  # Initialize logging
  cat("\n=== NEW CONVERSION ===\n")
  cat("Original formula:", deparse(form), "\n")
  
  # Input validation
  if (!is.character(form)) {
    cat("! Input is not character, type:", typeof(form), "\n")
    return(NA_character_)
  }
  if (is.na(form) || !nzchar(form)) {
    cat("! Empty or NA input\n")
    return(NA_character_)
  }
  
  # Cleanup global
  f_original <- form
  f <- gsub("[\r\n]+", "", form)
  f <- sub("^=", "", f)
  f <- gsub("\\$", "", f)
  f <- gsub("<>", "!=", f, fixed = TRUE)
  f <- trimws(f)
  f <- clean_parentheses(f)
  
  cat("After cleaning:", f, "\n")
  
  # 2) Handle comparisons
  eq <- split_at_top(f, c("!=", "==", "<=", ">=", "<", ">", "="))
  if (!is.null(eq)) {
    cat("Found comparison operator:", eq$op, "\n")
    cat("Left part:", eq$left, "\n")
    cat("Right part:", eq$right, "\n")
    
    left_r <- convert_formula(eq$left, noms_cellules)
    right_r <- convert_formula(eq$right, noms_cellules)
    
    # Gestion spéciale pour comparaison avec chaîne vide
    if (right_r == '""') {
      left_r <- paste0("!is.na(", left_r, ") & ", left_r, " != \"\"")
    }
    
    op_r <- switch(eq$op,
                   "=" = "==", "==" = "==", "!=" = "!=",
                   "<" = "<", ">" = ">", "<=" = "<=", ">=" = ">=")
    
    result <- paste0(left_r, " ", op_r, " ", right_r) 
    cat("Comparison result:", result, "\n")
    return(result)
  }
  
  # 3) Handle binary operations
  bin <- split_at_top(f, c("+", "-", "*", "/", "&"))
  if (!is.null(bin)) {
    cat("Found binary operator:", bin$op, "\n")
    cat("Left part:", bin$left, "\n")
    cat("Right part:", bin$right, "\n")
    
    l <- convert_formula(bin$left, noms_cellules)
    r <- convert_formula(bin$right, noms_cellules)
    
    if (bin$op == "&") {
      result <- paste0("paste0(", l, ",", r, ")")
    } else {
      result <- paste(l, bin$op, r)
    }
    
    cat("Binary operation result:", result, "\n")
    return(result)
  }
  
  # 4) Handle literals, numbers, references
  if (!grepl("^[A-Z]+\\(", f)) {
    cat("Processing as literal/reference\n")
    
    # Named ranges
    if (f %in% names(noms_cellules)) {
      cat("Found named range:", f, "\n")
      cat("Reference:", noms_cellules[[f]]$ref, "\n")
      cat("Sheet:", noms_cellules[[f]]$sheet, "\n")
      
      cell_part <- sub("^values", "", noms_cellules[[f]]$ref)
      sheet <- noms_cellules[[f]]$sheet
      result <- sprintf("sheets[['%s']]%s", sheet, cell_part)
      
      cat("Named range result:", result, "\n")
      return(result)
    }
    
    # Direct references
    if (grepl('^[^!]+![A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', f) ||
        grepl('^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', f)) {
      cat("Found direct reference:", f, "\n")
      result <- convert_ref(f)
      cat("Direct reference result:", result, "\n")
      return(result)
    }
    
    # Literals
    if (f %in% c("TRUE", "FALSE")) {
      cat("Found boolean:", f, "\n")
      return(f)
    }
    if (grepl('^".*"$', f) || grepl('^[0-9\\.]+$', f)) {
      cat("Found literal:", f, "\n")
      return(f)
    }
    
    cat("Returning as-is:", f, "\n")
    return(f)
  }
  
  # Ajoutez ceci avant la section des fonctions dans convert_formula()
  if (grepl("^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$", f)) {
    return(convert_ref(f))
  }
  
  # 5) Function calls
  cat("Processing function call\n")
  m <- regexec("^([A-Z]+)\\((.*)\\)$", f)
  parts <- regmatches(f, m)[[1]]
  
  if (length(parts) < 3) {
    cat("! Invalid function format:", f, "\n")
    return(NA_character_)
  }
  
  fname <- parts[2]
  inside <- parts[3]
  cat("Function name:", fname, "\n")
  cat("Arguments string:", inside, "\n")
  
  raw_args <- split_top(inside)
  raw_args <- vapply(raw_args, function(a) sub("^\\((.*)\\)$", "\\1", trimws(a)), "")
  cat("Split arguments:", paste(raw_args, collapse = "|"), "\n")
  
  conv_args <- vapply(raw_args, function(x) {
    cat("Converting argument:", x, "\n")
    convert_formula(x, noms_cellules)
  }, character(1))
  
  # 6) Generate code
  # 6) Generate code
  up <- toupper(fname)
  
  if (up == "SUMIF") {
    # on passe raw_args en 3e paramètre
    out <- fun_map[[up]](conv_args, noms_cellules, raw_args)
    
  } else if (up == "VLOOKUP") {
    cat("Handling VLOOKUP specially\n")
    out <- sprintf("vlookup_r(%s, %s, %s)",
                   conv_args[1],
                   convert_ref(raw_args[2]),
                   conv_args[3])
    
  } else if (up %in% names(fun_map)) {
    # toutes les autres fonctions de la map
    if ("noms_cellules" %in% names(formals(fun_map[[up]]))) {
      out <- fun_map[[up]](conv_args, noms_cellules)
    } else {
      out <- fun_map[[up]](conv_args)
    }
    
  } else {
    # appel générique
    cat("Generic function handling\n")
    out <- paste0(fname, "(", paste(conv_args, collapse = ","), ")")
  }
  
  
  # 7) Final cleanup
  #result <- gsub("(?<!=)=(?!=)", "==", out, perl = TRUE)
  cat("Final result:", out, "\n")
  return(out)
}