# excel2R_refactored.R
# -----------------------------------
# Convertisseur de formules Excel en code R
# Architecture : nettoyage → split générique (multi-char ops) → parsing récursif → codegen via table de fonctions

# 1) UTILITAIRES DE SPLITTING -------------------------------

# Découpe f en deux parties au premier opérateur de ops au niveau racine,
# en gérant d'abord les opérateurs à deux caractères puis à un seul.
# ────────────────────────────────────────────────────────────────────────────────
# split_at_top : split au premier opérateur (hors parenthèses et hors quotes)
split_at_top <- function(f, ops) {
  chars <- strsplit(f, "", fixed = TRUE)[[1]]
  depth     <- 0L
  in_quote  <- FALSE
  quote_chr <- ""
  i         <- 1L
  
  while (i <= length(chars)) {
    ch <- chars[i]
    
    # → gestion des quotes : on ignore tout jusqu'à la quote fermante
    if (!in_quote && (ch == "'" || ch == '"')) {
      in_quote  <- TRUE
      quote_chr <- ch
      i <- i + 1L
      next
    }
    if (in_quote && ch == quote_chr) {
      in_quote  <- FALSE
      quote_chr <- ""
      i <- i + 1L
      next
    }
    if (in_quote) {
      i <- i + 1L
      next
    }
    
    # hors quotes, on gère la profondeur des parenthèses
    if (ch == "(")      depth <- depth + 1L
    else if (ch == ")") depth <- depth - 1L
    else if (depth == 0L) {
      # test des ops à 2 caractères
      two <- if (i < length(chars)) paste0(chars[i], chars[i+1]) else ""
      if (two %in% ops) {
        left  <- trimws(substr(f, 1, i-1))
        right <- trimws(substr(f, i+2, nchar(f)))
        return(list(left = left, op = two, right = right))
      }
      # test des ops à 1 caractère
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
# ────────────────────────────────────────────────────────────────────────────────
# split_top : découpe sur les virgules de profondeur 0 (hors parenthèses & hors quotes)
split_top <- function(s) {
  if (length(s) != 1 || is.na(s) || !nzchar(s)) return(character())
  chars <- strsplit(s, "", fixed = TRUE)[[1]]
  depth    <- 0L
  in_quote <- FALSE
  quote_chr<- ""
  args     <- character()
  curr     <- ""
  
  for (ch in chars) {
    # gestion des quotes
    if (!in_quote && (ch == "'" || ch == '"')) {
      in_quote  <- TRUE
      quote_chr <- ch
      curr      <- paste0(curr, ch)
      next
    }
    if (in_quote && ch == quote_chr) {
      in_quote  <- FALSE
      quote_chr <- ""
      curr      <- paste0(curr, ch)
      next
    }
    if (in_quote) {
      curr <- paste0(curr, ch)
      next
    }
    
    # hors quotes, gestion des parenthèses
    if (ch == "(")      depth <- depth + 1L
    else if (ch == ")") depth <- depth - 1L
    else if (ch == "," && depth == 0L) {
      args <- c(args, curr)
      curr <- ""
      next
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
  # on capture soit 'nom de feuille' (avec '' pour les apostrophes),
  # soit un nom sans apostrophes (lettres, chiffres, espace, - ou _),
  # suivi de ! et d'une coordonnée (A1 ou A1:B2)
  m <- regexec(
    "^(?:'((?:[^']|'')+)'|([A-Za-z0-9 _-]+))!([A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?)$",
    ref, perl = TRUE
  )
  parts <- regmatches(ref, m)[[1]]
  
  if (length(parts) > 0) {
    # parts[2] = texte entre '…' (avec '' pour chaque '), 
    # parts[3] = texte non-apostrophé
    sheet_raw <- if (nzchar(parts[2])) parts[2] else parts[3]
    # on remplace les doubles apostrophes par une seule
    sheet <- gsub("''", "'", sheet_raw)
    coord <- parts[4]
  } else {
    # pas de "!", c'est juste une coordonnée
    sheet <- NULL
    coord <- ref
  }
  
  # maintenant on transforme coord (ex. "A1" ou "A1:B2") en indices R
  if (grepl(":", coord, fixed = TRUE)) {
    bounds <- strsplit(coord, ":", fixed = TRUE)[[1]]
    start  <- strcapture("^([A-Z]+)([0-9]+)$", bounds[1],
                         proto = list(C="", R = 0))
    end    <- strcapture("^([A-Z]+)([0-9]+)$", bounds[2],
                         proto = list(C="", R = 0))
    rows <- sprintf("%d:%d", as.integer(start$R), as.integer(end$R))
    cols <- sprintf("%d:%d",
                    col2num(start$C),
                    col2num(end$C))
  } else {
    xy   <- strcapture("^([A-Z]+)([0-9]+)$", coord,
                       proto = list(C="", R = 0))
    rows <- sprintf("%d",   as.integer(xy$R))
    cols <- sprintf("%d",   col2num(xy$C))
  }
  
  idx <- sprintf("%s, %s", rows, cols)
  if (!is.null(sheet)) {
    sprintf("sheets[[\"%s\"]][%s]", sheet, idx)
  } else {
    sprintf("values[%s]", idx)
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
  
  # 6bis) Si le critère est un appel de fonction (ex. MID(...)), on le compare à ==
  if (grepl("^[A-Za-z]+\\(.*\\)$", crit)) {
    return(list(type = "compare",
                op   = "==",
                value = convert_formula(crit, noms_cellules)))
  }
  
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
  CONCATENATE = function(args, noms_cellules = NULL) {
    # Equivalent Excel CONCATENATE -> R paste0
    paste0("paste0(", paste(args, collapse = ","), ")")
  }, 
  OFFSET = function(args, noms_cellules = NULL, raw_args) {
    # raw_args[[1]] = référence Excel "'IV - Flux effectifs'!M29"
    ref_raw <- raw_args[[1]]
    ref_r   <- convert_ref(ref_raw)  # e.g. "sheets[['IV - Flux effectifs']][29, 13]"
    # on extrait feuille et indices
    m     <- regexec("^(.*)\\[([0-9]+),\\s*([0-9]+)\\]$", ref_r)
    parts <- regmatches(ref_r, m)[[1]]
    sheet <- parts[2]
    row0  <- as.integer(parts[3])
    col0  <- as.integer(parts[4])
    # args[2], args[3] sont déjà l’expression R du row-off et col-off
    row_off <- args[2]
    col_off <- args[3]
    sprintf("%s[%d + (%s), %d + (%s)]",
            sheet, row0, row_off, col0, col_off)
  }, 
  MID = function(args, noms_cellules = NULL) {
    # Excel MID(text, start, n) → R substr(text, start, start + n - 1)
    paste0(
      "substr(",
      args[1], ", ",
      args[2], ", ",
      args[2], " + ", args[3], " - 1)"
    )
  },
  ISBLANK = function(args, noms_cellules = NULL) {
    # Excel ISBLANK(x) → R : is.na(x) | x == ""
    paste0("(", args[[1]], " %in% c(NA, \"\"))")
  },
  INDIRECT = function(args, noms_cellules = NULL) {
    # Excel INDIRECT("A1") → R : évaluer dynamiquement l’expression
    # ici on parse le texte et on évalue
    paste0("eval(parse(text=", args[[1]], "))")
  },
  AND    = function(args, noms_cellules = NULL) paste0("(", paste(args, collapse = " & "), ")"), 
  OR = function(args, noms_cellules = NULL) paste0("(", paste(args, collapse = " | "), ")")
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
  
  # --- nettoyage habituel de form → f
  f <- gsub("[\r\n]+", "", form)
  f <- sub("^=",      "", f)
  f <- gsub("\\$",    "", f)
  f <- gsub("<>",    "!=", f, fixed = TRUE)
  f <- trimws(f)
  #f <- clean_parentheses(f)
  
  # si f est entouré par une paire de (… ) parfaitement appariée,
  # on enlève ces parenthèses le temps d’analyser l’intérieur,
  # puis on les remet autour du code R généré
  if (grepl("^\\(.*\\)$", f)) {
    chars <- strsplit(f, "")[[1]]
    depth <- cumsum(chars == "(") - cumsum(chars == ")")
    if (depth[length(depth)] == 0L && all(depth[-length(depth)] > 0L)) {
      inner <- substr(f, 2, nchar(f) - 1)
      conv  <- convert_formula(inner, noms_cellules)
      return(paste0("(", conv, ")"))
    }
  }
  
  if (grepl("^[-]?[0-9]+(\\.[0-9]+)?%$", f)) {
    num <- sub("%$", "", f)
    return(paste0("(", num, "/100)"))
  }
  
  cat("After cleaning :", f, "\n")
  
  # détection de FEUILLE!CELL ou 'FEUILLE'!CELL, avec ou sans plage
  if (grepl("^(?:'[^']+'|[A-Za-z0-9 _]+)![A-Za-z]+[0-9]+$", f, perl = TRUE)) {
    return(convert_ref(f))
  }
  
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
      # On renvoie directement le test sans jamais recoller != ""
      result <- paste0("!is.na(", left_r, ") & ", left_r, " != \"\"")
      cat("Empty-string comparison result:", result, "\n")
      return(result)
    }
    
    # Sinon, on fait le test normal
    op_r <- switch(eq$op,
                   "="  = "==", "==" = "==", "!=" = "!=",
                   "<"  = "<",  ">"  = ">",  "<=" = "<=",  ">=" = ">=")
    
    result <- paste0(left_r, " ", op_r, " ", right_r) 
    
    cat("Comparison result:", result, "\n")
    return(result)
  }
  
  # 3) Handle binary operations avec priorité (& > +- > */)
  bin <- NULL
  for (ops in list(c("&"), c("+", "-"), c("*", "/"))) {
    bin <- split_at_top(f, ops)
    if (!is.null(bin)) break
  }
  if (!is.null(bin)) {
    cat("Found binary operator:", bin$op, "\n")
    cat("Left part:", bin$left, "\n")
    cat("Right part:", bin$right, "\n")
    
    l <- convert_formula(bin$left,  noms_cellules)
    r <- convert_formula(bin$right, noms_cellules)
    
    if (bin$op == "&") {
      # concaténation texte
      result <- paste0("paste0(", l, ",", r, ")")
    } else if (bin$op == "-" && trimws(bin$left) == "") {
      # unaire
      result <- paste0("-", r)
    } else {
      # binaire (+, -, *, /)
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
  
  cat("Split arguments:", paste(raw_args, collapse = "|"), "\n")
  
  conv_args <- vapply(raw_args, function(x) {
    cat("Converting argument:", x, "\n")
    convert_formula(x, noms_cellules)
  }, character(1))
  
  # 6) Generate code
  up <- toupper(fname)
  
  if (up == "SUMIF") {
    # on passe raw_args en 3e paramètre
    out <- fun_map[[up]](conv_args, noms_cellules, raw_args)
    
  } else if (up == "OFFSET") {
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