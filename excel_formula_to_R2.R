# excel_formula_to_R.R

# -----------------------------------
# split_operator_top : si f contient un + - * / au niveau racine,
# renvoie list(left,op,right), sinon NULL
# avant : c("+", "-", "*", "/")
# après :  c("+", "-", "*", "/", "&")
split_operator_top <- function(f) {
  chars <- strsplit(f, "", fixed = TRUE)[[1]]
  depth <- 0L
  for (i in seq_along(chars)) {
    ch <- chars[i]
    if      (ch == "(") depth <- depth + 1L
    else if (ch == ")") depth <- depth - 1L
    else if (depth == 0L && ch %in% c("+", "-", "*", "/", "&")) {
      left  <- substr(f, 1, i-1)
      right <- substr(f, i+1, nchar(f))
      return(list(left = trimws(left),
                  op   = ch,
                  right= trimws(right)))
    }
  }
  NULL
}

# ———————— split_eq_top —————————————————————————
# si f contient un = au niveau racine (et pas ==), renvoie list(left, right), sinon NULL
split_eq_top <- function(f) {
  chars <- strsplit(f, "", fixed = TRUE)[[1]]
  depth <- 0L
  for (i in seq_along(chars)) {
    ch <- chars[i]
    if      (ch == "(") depth <- depth + 1L
    else if (ch == ")") depth <- depth - 1L
    else if (depth == 0L && ch == "=") {
      # s'assurer que ce n'est pas un "==" 
      prev <- if (i > 1) chars[i-1] else ""
      nextc<- if (i < length(chars)) chars[i+1] else ""
      if (prev != "=" && nextc != "=") {
        left  <- paste(chars[1:(i-1)], collapse = "")
        right <- paste(chars[(i+1):length(chars)], collapse = "")
        return(list(left = trimws(left), right = trimws(right)))
      }
    }
  }
  NULL
}


# -----------------------------------

# split_top et replace_sum comme avant
# ———————— split_top —————————————————————————
# scinde une chaîne sur les virgules de profondeur 0, ignore les NA
split_top <- function(s) {
  # si entrée vide ou NA, on renvoie vecteur vide
  if (length(s) != 1 || is.na(s) || nzchar(s, keepNA = FALSE) == FALSE) {
    return(character())
  }
  chars <- strsplit(s, "")[[1]]
  depth <- 0L
  args  <- character()
  curr  <- ""
  for (ch in chars) {
    # on saute les caractères NA
    if (is.na(ch)) next
    if (ch == "(") {
      depth <- depth + 1L
    } else if (ch == ")") {
      depth <- depth - 1L
    } else if (ch == "," && depth == 0L) {
      # virgule au niveau racine : on clôt un argument
      args <- c(args, curr)
      curr <- ""
      next
    }
    # sinon on accumulate
    curr <- paste0(curr, ch)
  }
  # on rajoute le dernier
  args <- c(args, curr)
  args
}

replace_sum <- function(f) {
  # on découpe chaque SUM(...) racine et on remplace
  pat <- "SUM\\("
  while (grepl(pat, f, perl=TRUE)) {
    m <- regexpr(pat, f, perl=TRUE)
    start <- m[1]
    ml    <- attr(m,"match.length")
    depth <- 1L; i <- start + ml; n <- nchar(f)
    while (depth>0 && i<=n) {
      ch <- substr(f,i,i)
      if (ch=="(") depth<-depth+1L
      if (ch==")") depth<-depth-1L
      i <- i+1L
    }
    endpos  <- i-1L
    content <- substr(f, start+ml, endpos-1L)
    args    <- split_top(content)
    f <- paste0(
      substr(f,1,start-1),
      "sum(", paste(args, collapse=","), ")",
      substr(f,endpos+1L,n)
    )
  }
  f
}

# helper pour OFFSET sur une seule cellule
offset_r <- function(addr, row_off, col_off) {
  # addr : chaîne Excel A1
  m     <- regexec("^([A-Z]+)([0-9]+)$", addr, perl=TRUE)
  parts <- regmatches(addr, m)[[1]]
  # calcul de la nouvelle ligne/col
  r0    <- as.integer(parts[3]) + as.integer(row_off)
  c0    <- col2num(parts[2])   + as.integer(col_off)
  sprintf("values[%d, %d]", r0, c0)
}


convert_ref <- function(ref) {
  if (grepl(":", ref, fixed=TRUE)) {
    prts  <- strsplit(ref,":",fixed=TRUE)[[1]]
    s <- strcapture("^([A-Z]+)([0-9]+)$", prts[1], proto=list(C="",R=0))
    e <- strcapture("^([A-Z]+)([0-9]+)$", prts[2], proto=list(C="",R=0))
    sprintf("values[%d:%d, %d:%d]",
            as.integer(s$R), as.integer(e$R),
            col2num(s$C),   col2num(e$C))
  } else {
    xy <- strcapture("^([A-Z]+)([0-9]+)$", ref, proto=list(C="",R=0))
    sprintf("values[%d, %d]",
            as.integer(xy$R), col2num(xy$C))
  }
}
# excel_formula_to_R.R

# … vos autres helpers (col2num, convert_ref, split_top, replace_sum, etc.) …

convert_formula <- function(form) {
  #message("[convert_formula] formule initiale :", form)
  
  # 1) nettoyage global : virer sauts de ligne, '$' et le '=' de début
  f <- gsub("[\r\n]+", "", form)     
  f <- sub("^=",       "", f)        
  f <- gsub("\\$",     "", f)
  f <- gsub("<>",     "!=", f, fixed = TRUE) 
  f <- trimws(f)
  #message("[convert_formula] après cleanup       :", f)
  # 1bis) SI c'est une égalité racine, on split et on traite récursivement
  eq <- split_eq_top(f)
  if (!is.null(eq)) {
    left_r  <- convert_formula(eq$left)
    right_r <- convert_formula(eq$right)
    return(paste0("(", left_r, ") == (", right_r, ")"))
  }
  # 1bis) SI la chaîne entière est entourée de parenthèses appariées,
  #      on les enlève et on rappelle convert_formula() pour retraiter
  
  if (grepl("^\\(.*\\)$", f)) {
    # on enlève la première et la dernière parenthèse
    inner <- substr(f, 2, nchar(f) - 1)
    # seulement si le premier “(” correspond bien au dernier “)”
    # on pourrait ajouter un test plus strict sur le “depth” ici
    return(convert_formula(inner))
  }
  # 2a) Cas opération binaire simple (A1/B2, C3+D4, AB7&H11, etc.)
  op_split <- split_operator_top(f)
  if (!is.null(op_split)) {
    left_r  <- convert_formula(op_split$left)
    right_r <- convert_formula(op_split$right)
    if (op_split$op == "&") {
      # concaténation Excel -> paste0(...) en R
      return(paste0("paste0(", left_r, ",", right_r, ")"))
    } else {
      # opérateur arithmétique normal
      return(paste0(left_r, " ", op_split$op, " ", right_r))
    }
  }
  
  # 2b) sinon on continue…
  
  # 2) remplacer tout '"=" & X' par 'X' (cas SUMIF critère dynamique)
  f <- gsub("\"=\"\\s*&\\s*", "", f, perl = TRUE)
  #message("[convert_formula] après suppression \"=\" &   :", f)
  
  # 3) si ce n'est pas un appel FOO(...), retourner tel quel ou convert_ref()
  if (!grepl("^[A-Z]+\\(", f, perl = TRUE)) {
    if (grepl("^\".*\"$", f) || grepl("^[0-9\\.]+$", f)) {
      return(f)
    }
    if (grepl("^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", f)) {
      return(convert_ref(f))
    }
    return(f)
  }
  
  # 4) extraire NOM(...) et son contenu
  m    <- regexec("^([A-Z]+)\\((.*)\\)$", f, perl = TRUE)
  prts <- regmatches(f, m)[[1]]
  fname  <- prts[2]
  inside <- prts[3]
  
  # 5) découper au niveau top-level
  raw_args <- split_top(inside)
  #    => retirer les parenthèses externes éventuelles
  raw_args <- vapply(raw_args, function(a) {
    a <- trimws(a)
    sub("^\\((.*)\\)$", "\\1", a)
  }, FUN.VALUE = "")
  #message("[convert_formula] args bruts     :", paste(raw_args, collapse = " | "))
  
  # 6) récursion sur chaque argument
  conv_args <- vapply(raw_args, convert_formula, FUN.VALUE = "")
  #message("[convert_formula] args convertis :", paste(conv_args, collapse = " | "))
  
  # 7) reconstituer selon la fonction
  fname <- toupper(fname)
  out <- switch(fname,
                "SUM"      = paste0("sum(",  paste(conv_args, collapse = ","), ")"),
                "SUMIF"    = paste0("sum(ifelse(", conv_args[1], "==", conv_args[2],
                                    ",", conv_args[3], ",0))"),
                "IF"       = paste0("ifelse(",  conv_args[1], ",", conv_args[2], ",", conv_args[3], ")"),
                "IFERROR"  = paste0("tryCatch({", conv_args[1], "}, error=function(e){", conv_args[2], "})"),
                "VLOOKUP"  = {
                  # args_raw = lookup, "'Feuille'!A1:B2", col, FALSE
                  ms <- regexec("^'([^']+)'!([A-Z]+[0-9]+:[A-Z]+[0-9]+)$", raw_args[2], perl=TRUE)
                  gr <- regmatches(raw_args[2], ms)[[1]]
                  sheet <- gr[2]; rng <- gr[3]
                  parts <- strsplit(rng, ":", fixed=TRUE)[[1]]
                  st <- strcapture("^([A-Z]+)([0-9]+)$", parts[1], proto=list(C="",R=0))
                  ed <- strcapture("^([A-Z]+)([0-9]+)$", parts[2], proto=list(C="",R=0))
                  prng <- sprintf("sheets[['%s']][%d:%d, %d:%d]",
                                  sheet,
                                  as.integer(st$R), as.integer(ed$R),
                                  col2num(st$C), col2num(ed$C))
                  paste0("vlookup_r(", conv_args[1], ", ", prng, ", ", conv_args[3], ")")
                },
                "TEXT"     = paste0("as.character(", conv_args[1], ")"),
                "ROUND"    = {
                  digits <- if (length(conv_args)>=2) conv_args[2] else "0"
                  paste0("round(", conv_args[1], ", ", digits, ")")
                },
                "INT"      = paste0("floor(", conv_args[1], ")"),
                "DATE"     = {
                  # on utilise ISOdate puis as.Date
                  paste0("as.Date(ISOdate(", 
                         conv_args[1], ",", conv_args[2], ",", conv_args[3], 
                         "))")
                },
                "OFFSET"   = {
                  # args bruts : référence, décalage lignes, décalage colonnes, ... on ne gère que les 3 premiers
                  paste0("offset_r('", raw_args[1], "', ", conv_args[2], ", ", conv_args[3], ")")
                },
                # fallback pour les autres
                paste0(fname, "(", paste(conv_args, collapse=","), ")")
  )
  #message("[convert_formula] avant replace_sum   :", out)
  
  # 8) replier encore d’éventuels SUM(…) imbriqués
  out <- replace_sum(out)
  # 9) transformer '=' isolés en '=='
  out <- gsub("(?<!=)=(?!=)", "==", out, perl=TRUE)
  
  message("[convert_formula] formule R finale :", out)
  out
}

