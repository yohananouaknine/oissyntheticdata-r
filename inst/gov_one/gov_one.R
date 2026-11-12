# gov_one.R -- one file, base R only (no package to install), fully offline.
#
# The R twin of gov_one.py: same flow, same rules, same JSON files.
#
# THE FLOW
#   1. ON THE PREMISES   Rscript gov_one.R profile
#         Reads every Excel / CSV file in the folder, finds how the files relate to each
#         other, and writes a folder  to_send/  holding
#             profile_<table>.json   aggregates only (counts, shares, quantiles)
#             schema.json            which column of which file points to which
#             profile_summary.md     a plain list of exactly what leaves -- for the
#                                    person who approves the e-mail
#         No row, no identifier and no value seen fewer than MIN_CELL_COUNT times is written.
#   2. AT HOME           Rscript gov_one.R synth
#         Rebuilds files with the same names, sheets, columns, types, distributions and
#         working joins from the JSON files.  Every value is invented.
#   3. AT HOME           write your own analysis script (analysis_template.R is a start):
#             g <- new.env(); sys.source("gov_one.R", envir = g)
#             tabs <- g$open_tables(".")    # real files if present, else the synthetic ones
#         and use base R (table, chisq.test, t.test, lm ...) plus the helpers in this file.
#   4. BACK ON THE PREMISES   copy gov_one.R + your script, run the script on the real files.
#
#   Rscript gov_one.R            does step 1 when it sees data files, step 2 when it sees only JSON files
#   Rscript gov_one.R check      lists the files, sheets and headers it would read (no data shown)
#   Rscript gov_one.R settings   writes gov_one_settings.json (hide a column, declare a link, ...)
#
# The JSON written here is read by gov_one.py and the other way round, so the profile can be made
# with R on the premises and the synthetic files with Python at home (or the reverse).
# Only packages that come with every R installation are used (base, stats, utils).
# This file is plain ASCII: Hebrew words are written as \u escapes so it reads the same everywhere.

VERSION <- "gov_one.R 3.0.0 offline (2026-11-12)"

# ==============================================================================
#  Settings.  Edit here, or -- on the premises -- in gov_one_settings.json.
# ==============================================================================
MIN_CELL_COUNT <- 5            # a value / cell seen fewer times than this never leaves (k)
MIN_TABLE_ROWS <- 20           # a table smaller than this is described by its shape only
MAX_CATEGORIES <- 500          # a text column with more distinct values is treated as free text / id
MIN_GROUP_KEYS <- 20           # a key column needs more distinct values than this
QUANTILE_GRID <- c(0, 5, 10, 25, 50, 75, 90, 95, 100)
USE_ROBUST_BOUNDS <- TRUE
P_LOW <- 1.0
P_HIGH <- 99.0
OUTPUT_PREFIXES <- c("synthetic_", "profile_", "~$")      # never read our own outputs or Excel lock files
SEND_DIR <- "to_send"
SYNTH_DIR <- "synthetic"
SETTINGS_FILE <- "gov_one_settings.json"
DEFAULT_SEED <- 12345
SERIAL_MIN <- 20000
SERIAL_MAX <- 60000            # Excel day numbers of about 1954-09 .. 2064-03
EXCEL_EPOCH_DAYS <- -25569     # 1899-12-30 counted in days from 1970-01-01
DATE_SAMPLE <- 600
ID_DISTINCT_RATIO <- 0.5       # an integer column with at least this distinct-ratio looks like a key

# joint (two-way) tables kept so that the synthetic columns move together like the real ones
JOINT_MAX_PARENT_LEVELS <- 60
JOINT_MAX_CHILD_LEVELS <- 300
JOINT_MAX_TABLES <- 40
JOINT_MAX_SECOND <- 20         # how many columns may follow a second column as well
JOINT_SAMPLE_ROWS <- 20000
JOINT_MIN_V <- 0.05
JOINT_MIN_V_SECOND <- 0.03     # the same, for what a second column adds beyond the first
CROSSTAB_NUMERIC_BINS <- 12

# A column whose HEADER says it holds a person's name, an id number, a phone, an address or free
# text is described by its shape only (length, letters / digits) -- its values never leave.
# (The lists below are the lists of gov_one.py, generated from it.)

MISSING_TOKENS <- c("", "#div/0!", "#n/a", "#null!", "#ref!", "#value!", "-", ".", "n/a", "na", "nan",
                    "none", "null")
ID_NAME_HINTS <- c("id", "key", "code", "\u05de\u05d6\u05d4\u05d4", "\u05de\u05d9\u05e1\u05e4\u05d5\u05e8",
                   "\u05ea\u05d6")
SENSITIVE_PHRASES <- c("\u05e9\u05dd \u05e4\u05e8\u05d8\u05d9",
                       "\u05e9\u05dd \u05de\u05e9\u05e4\u05d7\u05d4", "\u05e9\u05dd \u05de\u05dc\u05d0",
                       "\u05e9\u05dd \u05d4\u05d0\u05d1", "\u05e9\u05dd \u05d4\u05d0\u05dd",
                       "\u05e9\u05dd \u05d4\u05dc\u05e7\u05d5\u05d7",
                       "\u05e9\u05dd \u05dc\u05e7\u05d5\u05d7",
                       "\u05e9\u05dd \u05d4\u05de\u05d8\u05d5\u05e4\u05dc",
                       "\u05e9\u05dd \u05de\u05d8\u05d5\u05e4\u05dc",
                       "\u05e9\u05dd \u05d4\u05e4\u05d5\u05e0\u05d4",
                       "\u05e9\u05dd \u05e4\u05d5\u05e0\u05d4",
                       "\u05e9\u05dd \u05d4\u05e2\u05d5\u05d1\u05d3",
                       "\u05e9\u05dd \u05e2\u05d5\u05d1\u05d3",
                       "\u05e9\u05dd \u05d4\u05de\u05d8\u05e4\u05dc",
                       "\u05e9\u05dd \u05de\u05d8\u05e4\u05dc",
                       "\u05e9\u05dd \u05d4\u05e8\u05d5\u05e4\u05d0",
                       "\u05e9\u05dd \u05e8\u05d5\u05e4\u05d0",
                       "\u05e9\u05dd \u05d4\u05ea\u05dc\u05de\u05d9\u05d3",
                       "\u05e9\u05dd \u05ea\u05dc\u05de\u05d9\u05d3",
                       "\u05e9\u05dd \u05d4\u05d0\u05e1\u05d9\u05e8",
                       "\u05e9\u05dd \u05d0\u05e1\u05d9\u05e8",
                       "\u05e9\u05dd \u05d4\u05de\u05e9\u05e4\u05d7\u05d4",
                       "\u05e9\u05dd \u05d0\u05d9\u05e9 \u05e7\u05e9\u05e8",
                       "\u05ea\u05e2\u05d5\u05d3\u05ea \u05d6\u05d4\u05d5\u05ea",
                       "\u05de\u05e1\u05e4\u05e8 \u05d6\u05d4\u05d5\u05ea",
                       "\u05de\u05e1 \u05d6\u05d4\u05d5\u05ea", "\u05de\u05e1' \u05d6\u05d4\u05d5\u05ea",
                       "\u05ea.\u05d6", "\u05ea\"\u05d6", "\u05ea''\u05d6", "\u05d3\u05e8\u05db\u05d5\u05df",
                       "\u05d8\u05dc\u05e4\u05d5\u05df", "\u05e0\u05d9\u05d9\u05d3",
                       "\u05e4\u05dc\u05d0\u05e4\u05d5\u05df", "\u05db\u05ea\u05d5\u05d1\u05ea",
                       "\u05e8\u05d7\u05d5\u05d1", "\u05d3\u05d5\u05d0\"\u05dc", "\u05d3\u05d5\u05d0\u05dc",
                       "\u05de\u05d9\u05d9\u05dc", "\u05d4\u05e2\u05e8\u05d5\u05ea",
                       "\u05d4\u05e2\u05e8\u05d4", "\u05de\u05dc\u05dc \u05d7\u05d5\u05e4\u05e9\u05d9",
                       "\u05ea\u05d9\u05d0\u05d5\u05e8 \u05d4\u05de\u05e7\u05e8\u05d4",
                       "\u05e4\u05d9\u05e8\u05d5\u05d8", "\u05d0\u05d9\u05e9 \u05e7\u05e9\u05e8",
                       "\u05d0\u05e0\u05e9\u05d9 \u05e7\u05e9\u05e8",
                       "\u05d0\u05d9\u05e9 \u05d4\u05e7\u05e9\u05e8",
                       "\u05e7\u05e8\u05d5\u05d1 \u05de\u05e9\u05e4\u05d7\u05d4",
                       "\u05d0\u05e4\u05d5\u05d8\u05e8\u05d5\u05e4\u05d5\u05e1", "contact person",
                       "next of kin", "guardian", "first name", "last name", "full name", "surname",
                       "family name", "given name", "patient name", "client name", "contact name",
                       "passport", "phone", "mobile", "address", "street", "email", "e-mail", "remarks",
                       "comments", "free text", "national id", "id number", "ssn")
SENSITIVE_WORDS <- c("\u05e9\u05dd", "name", "notes", "note", "comment", "\u05ea\u05d9\u05d0\u05d5\u05e8",
                     "description")
SENSITIVE_ROLE_WORDS <- c("\u05e2\u05d5\u05d1\u05d3", "\u05e2\u05d5\u05d1\u05d3\u05ea",
                          "\u05e2\u05d5\"\u05e1", "\u05e2\u05d5\u05e1", "\u05e8\u05d5\u05e4\u05d0",
                          "\u05e8\u05d5\u05e4\u05d0\u05d4", "\u05de\u05d8\u05e4\u05dc",
                          "\u05de\u05d8\u05e4\u05dc\u05ea", "\u05d0\u05d7\u05e8\u05d0\u05d9",
                          "\u05d0\u05d7\u05e8\u05d0\u05d9\u05ea", "\u05de\u05e0\u05d4\u05dc",
                          "\u05de\u05e0\u05d4\u05dc\u05ea", "\u05e4\u05e7\u05d9\u05d3",
                          "\u05e4\u05e7\u05d9\u05d3\u05d4", "\u05de\u05d3\u05e8\u05d9\u05da",
                          "\u05de\u05d3\u05e8\u05d9\u05db\u05d4", "\u05de\u05d5\u05e8\u05d4",
                          "\u05de\u05dc\u05d5\u05d5\u05d4", "\u05e8\u05db\u05d6", "\u05e8\u05db\u05d6\u05ea",
                          "\u05de\u05e4\u05e7\u05d7", "\u05de\u05e4\u05e7\u05d7\u05ea",
                          "\u05d7\u05d5\u05e7\u05e8", "\u05d7\u05d5\u05e7\u05e8\u05ea",
                          "\u05e7\u05e6\u05d9\u05df", "\u05e7\u05e6\u05d9\u05e0\u05d4", "\u05d0\u05d7",
                          "\u05d0\u05d7\u05d5\u05ea", "worker", "doctor", "physician", "therapist",
                          "manager", "officer", "employee", "teacher", "nurse", "caseworker")
DATE_NAME_TOKENS <- c("date", "\u05ea\u05d0\u05e8\u05d9\u05da", "yyyy", "dt")
DATE_TOKENS <- c("date", "time", "\u05d9\u05d5\u05dd", "\u05ea\u05d0\u05e8\u05d9\u05da",
                 "\u05e9\u05e0\u05d4", "\u05e9\u05e0\u05ea", "\u05d7\u05d5\u05d3\u05e9", "year", "month")
MEASURE_NAME_TOKENS <- c("count", "age", "amount", "total", "sum", "price", "salary", "income", "weight",
                         "height", "score", "percent", "days", "hours", "month", "year", "number of", "qty",
                         "quantity", "\u05db\u05de\u05d5\u05ea", "\u05d2\u05d9\u05dc",
                         "\u05e1\u05db\u05d5\u05dd", "\u05e1\u05d4\"\u05db", "\u05e1\u05da",
                         "\u05de\u05d7\u05d9\u05e8", "\u05e9\u05db\u05e8", "\u05d4\u05db\u05e0\u05e1\u05d4",
                         "\u05de\u05e9\u05e7\u05dc", "\u05d2\u05d5\u05d1\u05d4", "\u05e6\u05d9\u05d5\u05df",
                         "\u05d0\u05d7\u05d5\u05d6", "\u05d9\u05de\u05d9\u05dd", "\u05e9\u05e2\u05d5\u05ea",
                         "\u05d7\u05d5\u05d3\u05e9", "\u05e9\u05e0\u05d4", "\u05e9\u05e0\u05ea",
                         "\u05de\u05e1\u05e4\u05e8 \u05d9\u05dc\u05d3\u05d9\u05dd",
                         "\u05de\u05e1\u05e4\u05e8 \u05e0\u05e4\u05e9\u05d5\u05ea",
                         "\u05de\u05e1\u05e4\u05e8 \u05d9\u05de\u05d9")
GENERIC_NAME_WORDS <- c("code", "id", "key", "no", "num", "number", "of", "the", "\u05de\u05d6\u05d4\u05d4",
                        "\u05de\u05e1", "\u05de\u05e1\u05e4\u05e8", "\u05e7\u05d5\u05d3", "\u05e9\u05dc")
KEY_HEADER_WORDS <- c("code", "id", "key", "no", "num", "number", "\u05de\u05d6\u05d4\u05d4", "\u05de\u05e1",
                      "\u05de\u05e1'", "\u05de\u05e1\u05e4\u05d5\u05e8", "\u05de\u05e1\u05e4\u05e8",
                      "\u05e1\u05d9\u05d3\u05d5\u05e8\u05d9", "\u05e7\u05d5\u05d3")
BUILTIN_DATE <- c(14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58)
BUILTIN_TIME <- c(18, 19, 20, 21, 45, 46, 47)
RARE_LABEL <- "(rare value)"
# settings template examples
EX_HIDE <- c("\u05e9\u05dd \u05e2\u05d5\u05d1\u05d3 \u05e1\u05d5\u05e6\u05d9\u05d0\u05dc\u05d9",
             "cases::worker")
EX_SHOW <- c("\u05ea\u05d9\u05d0\u05d5\u05e8")


# ==============================================================================
#  Value helpers.  A table column is a character vector; "" is an empty cell.
# ==============================================================================
`%||%` <- function(a, b) if (is.null(a)) b else a

# every character Python's str.strip() / str.split() treats as white space (written as escapes: this
# also marks the pattern as UTF-8, so it is matched character by character, not byte by byte)
.WS <- "[\\s\u001c-\u001f\u0085\u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000]"

pystrip <- function(x) {
  x <- enc2utf8(as.character(x))
  gsub(paste0("^", .WS, "+|", .WS, "+$"), "", x, perl = TRUE)
}

is_missing <- function(v) tolower(pystrip(v)) %in% MISSING_TOKENS

# text -> number, NA when it is not a number ("1,234" counts as 1234)
to_float <- function(v) {
  s <- pystrip(gsub(",", "", as.character(v), fixed = TRUE))
  out <- rep(NA_real_, length(s))
  ok <- grepl("^[+-]?(\\d+\\.?\\d*|\\.\\d+)([eE][+-]?\\d+)?$", s, perl = TRUE)
  out[ok] <- suppressWarnings(as.numeric(s[ok]))
  inf <- !ok & grepl("^[+-]?(inf|infinity)$", tolower(s))
  out[inf] <- ifelse(startsWith(s[inf], "-"), -Inf, Inf)
  out
}

looks_int <- function(v) {
  f <- to_float(v)
  !is.na(f) & is.finite(f) & f == floor(f)
}

# whole number as text without exponent ("100000", never "1e+05")
int_str <- function(x) sprintf("%.0f", x)

c_sort <- function(x) sort(enc2utf8(as.character(x)), method = "radix")   # code-point order, as Python sorts text
c_gt <- function(a, b) a != b && order(enc2utf8(c(a, b)), method = "radix")[1] == 2L

# counts in order of first appearance (a Python dict filled row by row)
count_first <- function(x) {
  u <- unique(x)
  structure(tabulate(match(x, u), length(u)), names = u)
}

stride_idx <- function(n, k = DATE_SAMPLE) {
  if (n <= k) return(seq_len(n))
  step <- n / k
  floor((0:(k - 1)) * step) + 1
}
stride_sample <- function(xs, k = DATE_SAMPLE) xs[stride_idx(length(xs), k)]

percentile <- function(sorted_xs, p) {
  n <- length(sorted_xs)
  if (n == 0) return(0)
  k <- (n - 1) * p / 100
  f <- floor(k)
  cc <- k - f
  if (f + 2 <= n) return(sorted_xs[f + 1] + cc * (sorted_xs[f + 2] - sorted_xs[f + 1]))
  sorted_xs[f + 1]
}

# round to 4 significant figures (a quantile is a real value; its last digits do not leave)
# round() as Python does it: the nearest decimal to the exact binary value, exact ties to even
# (R's own round() differs in the last digit for a few values such as 0.25015)
pyround <- function(x, d = 0) if (d > 0) as.numeric(sprintf(paste0("%.", d, "f"), x)) else round(x, d)

sig4 <- function(x) {
  if (is.na(x) || x == 0) return(x)
  pyround(x, 3 - floor(log10(abs(x))))
}

signature <- function(v) {
  v <- enc2utf8(as.character(v))
  gsub("\\p{Nd}", "D", gsub("\\p{L}", "A", v, perl = TRUE), perl = TRUE)
}

# ---- dates: a date-time is a number of seconds from 1970-01-01 00:00 (no time zone) ----
days_from_civil <- function(y, m, d) {
  y <- y - (m <= 2)
  era <- floor(y / 400)
  yoe <- y - era * 400
  doy <- (153 * (m + ifelse(m > 2, -3, 9)) + 2) %/% 5 + d - 1
  doe <- yoe * 365 + yoe %/% 4 - yoe %/% 100 + doy
  era * 146097 + doe - 719468
}

civil_from_days <- function(z) {
  z <- z + 719468
  era <- floor(z / 146097)
  doe <- z - era * 146097
  yoe <- (doe - doe %/% 1460 + doe %/% 36524 - doe %/% 146096) %/% 365
  y <- yoe + era * 400
  doy <- doe - (365 * yoe + yoe %/% 4 - yoe %/% 100)
  mp <- (5 * doy + 2) %/% 153
  d <- doy - (153 * mp + 2) %/% 5 + 1
  m <- mp + ifelse(mp < 10, 3, -9)
  list(y = y + (m <= 2), m = m, d = d)
}

days_in_month <- function(y, m) {
  dm <- c(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)[m]
  dm + (m == 2 & ((y %% 4 == 0 & y %% 100 != 0) | y %% 400 == 0))
}

make_dt <- function(y, m, d, H = 0, M = 0, S = 0) days_from_civil(y, m, d) * 86400 + H * 3600 + M * 60 + S
dt_days <- function(dt) floor(dt / 86400)
dt_parts <- function(dt) {
  dd <- dt_days(dt)
  p <- civil_from_days(dd)
  s <- dt - dd * 86400
  list(y = p$y, m = p$m, d = p$d, H = s %/% 3600, M = (s %/% 60) %% 60, S = s %% 60, sec = s)
}
dt_weekday <- function(dt) (dt_days(dt) + 3) %% 7              # Monday = 0, as Python

dt_format <- function(dt, fmt = "%Y-%m-%d") {
  if (!length(dt)) return(character(0))
  ok <- !is.na(dt)
  p <- dt_parts(ifelse(ok, dt, 0))
  part <- list(Y = sprintf("%04d", as.integer(p$y)), m = sprintf("%02d", as.integer(p$m)), d = sprintf("%02d", as.integer(p$d)),
               H = sprintf("%02d", as.integer(p$H)), M = sprintf("%02d", as.integer(p$M)), S = sprintf("%02d", as.integer(p$S)))
  out <- rep("", length(dt))
  ch <- strsplit(fmt, "", fixed = TRUE)[[1]]
  i <- 1
  while (i <= length(ch)) {
    if (ch[i] == "%" && i < length(ch) && !is.null(part[[ch[i + 1]]])) {
      out <- paste0(out, part[[ch[i + 1]]])
      i <- i + 2
    } else {
      out <- paste0(out, ch[i])
      i <- i + 1
    }
  }
  out[!ok] <- ""
  out
}

.DATE_FORMATS <- list(
  list("^(\\d{4})-(\\d{1,2})-(\\d{1,2})$", c(1, 2, 3), "%Y-%m-%d"),
  list("^(\\d{1,2})/(\\d{1,2})/(\\d{4})$", c(3, 2, 1), "%d/%m/%Y"),
  list("^(\\d{1,2})\\.(\\d{1,2})\\.(\\d{4})$", c(3, 2, 1), "%d.%m.%Y"),
  list("^(\\d{1,2})/(\\d{1,2})/(\\d{4})$", c(3, 1, 2), "%m/%d/%Y"),
  list("^(\\d{4})/(\\d{1,2})/(\\d{1,2})$", c(1, 2, 3), "%Y/%m/%d"),
  list("^(\\d{1,2})-(\\d{1,2})-(\\d{4})$", c(3, 2, 1), "%d-%m-%Y"),
  list("^(\\d{4})-(\\d{1,2})-(\\d{1,2})\\s+(\\d{1,2}):(\\d{1,2}):(\\d{1,2})$", c(1, 2, 3, 4, 5, 6), "%Y-%m-%d %H:%M:%S"),
  list("^(\\d{1,2})/(\\d{1,2})/(\\d{4})\\s+(\\d{1,2}):(\\d{1,2}):(\\d{1,2})$", c(3, 2, 1, 4, 5, 6), "%d/%m/%Y %H:%M:%S"),
  list("^(\\d{1,2})/(\\d{1,2})/(\\d{4})\\s+(\\d{1,2}):(\\d{1,2})$", c(3, 2, 1, 4, 5), "%d/%m/%Y %H:%M"),
  list("^(\\d{4})-(\\d{2})-(\\d{2})[ T](\\d{2}):(\\d{2}):(\\d{2})(?:\\.\\d+)?$", c(1, 2, 3, 4, 5, 6), "%Y-%m-%d %H:%M:%S"))

.parse_unique <- function(sv) {
  n <- length(sv)
  dt <- rep(NA_real_, n)
  fm <- rep(NA_character_, n)
  for (f in .DATE_FORMATS) {
    todo <- which(is.na(dt))
    if (!length(todo)) break
    hit <- todo[grepl(f[[1]], sv[todo], perl = TRUE)]
    if (!length(hit)) next
    m <- regmatches(sv[hit], regexec(f[[1]], sv[hit], perl = TRUE))
    g <- suppressWarnings(matrix(as.numeric(unlist(lapply(m, function(z) z[-1]))), nrow = length(hit), byrow = TRUE))
    ord <- f[[2]]
    y <- g[, ord[1]]; mo <- g[, ord[2]]; d <- g[, ord[3]]
    H <- if (length(ord) >= 4) g[, ord[4]] else 0
    M <- if (length(ord) >= 5) g[, ord[5]] else 0
    S <- if (length(ord) >= 6) g[, ord[6]] else 0
    ok <- y >= 1 & mo >= 1 & mo <= 12 & d >= 1 & H < 24 & M < 60 & S < 62
    ok[ok] <- d[ok] <= days_in_month(y[ok], mo[ok])
    ok[is.na(ok)] <- FALSE
    if (any(ok)) {
      dt[hit[ok]] <- make_dt(y[ok], mo[ok], d[ok], rep_len(H, length(ok))[ok], rep_len(M, length(ok))[ok], rep_len(S, length(ok))[ok])
      fm[hit[ok]] <- f[[3]]
    }
  }
  list(dt = dt, fmt = fm)
}

# text -> list(dt = seconds or NA, fmt = the format it was written in)
parse_date <- function(v) {
  sv <- pystrip(v)
  u <- unique(sv)
  r <- .parse_unique(u)
  i <- match(sv, u)
  list(dt = r$dt[i], fmt = r$fmt[i])
}
parse_date_dt <- function(v) parse_date(v)$dt

excel_serial_to_dt <- function(v) {
  f <- to_float(v)
  f[!is.finite(f)] <- NA
  (round(f) + EXCEL_EPOCH_DAYS) * 86400
}
dt_to_excel_serial <- function(dt) dt_days(dt) - EXCEL_EPOCH_DAYS

to_dt <- function(v, serial = FALSE) if (isTRUE(serial)) excel_serial_to_dt(v) else parse_date_dt(v)

name_is_dateish <- function(name) {
  nm <- tolower(pystrip(name))
  any(vapply(DATE_NAME_TOKENS, function(tok) grepl(tok, nm, fixed = TRUE), logical(1)))
}

looks_like_serial <- function(v) {
  f <- to_float(v)
  !is.na(f) & is.finite(f) & f == floor(f) & f >= SERIAL_MIN & f <= SERIAL_MAX
}

# TRUE only when the column name signals a date and the values are whole numbers in the Excel
# day-number range.  Year (2024) and month (1-12) columns fall below the range.
is_serial_date_col <- function(present, name = "") {
  if (!name_is_dateish(name) || !length(present)) return(FALSE)
  sum(looks_like_serial(present)) / length(present) > 0.9
}

# ==============================================================================
#  JSON (written and read here, so no package is needed)
#    object  = a list or vector WITH names        array  = an unnamed list
#    scalar  = a vector of length 1 without names  null   = NULL / NA
# ==============================================================================
obj <- function(...) {
  x <- list(...)
  if (!length(x)) return(structure(list(), names = character(0)))
  x
}
arr <- function(x) unname(as.list(x))

.json_str <- function(s) {
  s <- enc2utf8(as.character(s))
  s <- gsub("\\", "\\\\", s, fixed = TRUE)
  s <- gsub("\"", "\\\"", s, fixed = TRUE)
  s <- gsub("\n", "\\n", s, fixed = TRUE)
  s <- gsub("\r", "\\r", s, fixed = TRUE)
  s <- gsub("\t", "\\t", s, fixed = TRUE)
  if (any(grepl("[\\x01-\\x1f]", s, perl = TRUE))) {
    for (code in c(1:8, 11, 12, 14:31)) s <- gsub(intToUtf8(code), sprintf("\\u%04x", code), s, fixed = TRUE)
  }
  paste0("\"", s, "\"")
}

.json_num <- function(x) {
  if (is.na(x) || !is.finite(x)) return("null")
  if (x == floor(x) && abs(x) < 1e15) return(sprintf("%.0f", x))
  s <- sprintf("%.15g", x)
  if (as.numeric(s) != x) s <- sprintf("%.17g", x)
  s
}

.json_scalar <- function(x) {
  if (is.null(x) || length(x) == 0) return("null")
  if (is.logical(x)) return(if (is.na(x)) "null" else if (x) "true" else "false")
  if (is.numeric(x)) return(.json_num(x))
  if (is.na(x)) return("null")
  .json_str(x)
}

to_json <- function(x, indent = 1, level = 0) {
  pad <- function(l) if (indent > 0) paste0("\n", strrep(" ", indent * l)) else ""
  if (is.null(x)) return("null")
  nm <- names(x)
  if (!is.null(nm)) {                                           # object
    if (!length(x)) return("{}")
    items <- vapply(seq_along(x), function(i) {
      v <- if (is.list(x)) x[[i]] else unname(x[i])
      paste0(.json_str(nm[i]), ": ", to_json(v, indent, level + 1))
    }, character(1))
    return(paste0("{", pad(level + 1), paste(items, collapse = paste0(",", pad(level + 1))), pad(level), "}"))
  }
  if (is.list(x) || length(x) != 1) {                           # array
    if (!length(x)) return("[]")
    items <- if (is.list(x)) vapply(x, function(v) to_json(v, indent, level + 1), character(1))
             else vapply(seq_along(x), function(i) .json_scalar(x[i]), character(1))
    return(paste0("[", pad(level + 1), paste(items, collapse = paste0(",", pad(level + 1))), pad(level), "]"))
  }
  .json_scalar(x)
}

write_json <- function(x, path, indent = 1) {
  con <- file(path, open = "wb")
  on.exit(close(con))
  writeBin(charToRaw(enc2utf8(to_json(x, indent))), con)
}

.json_unescape <- function(s) {
  if (!grepl("\\", s, fixed = TRUE)) return(s)
  out <- character(0)
  ch <- strsplit(s, "", fixed = TRUE)[[1]]
  i <- 1
  n <- length(ch)
  while (i <= n) {
    if (ch[i] != "\\" || i == n) { out <- c(out, ch[i]); i <- i + 1; next }
    e <- ch[i + 1]
    if (e == "u" && i + 5 <= n) {
      code <- strtoi(paste(ch[(i + 2):(i + 5)], collapse = ""), 16L)
      i <- i + 6
      if (!is.na(code) && code >= 0xD800 && code <= 0xDBFF && i + 5 <= n && ch[i] == "\\" && ch[i + 1] == "u") {
        low <- strtoi(paste(ch[(i + 2):(i + 5)], collapse = ""), 16L)
        code <- 0x10000 + (code - 0xD800) * 0x400 + (low - 0xDC00)
        i <- i + 6
      }
      out <- c(out, intToUtf8(code))
      next
    }
    out <- c(out, switch(e, n = "\n", r = "\r", t = "\t", b = "\b", f = "\f", e))
    i <- i + 2
  }
  paste(out, collapse = "")
}

# JSON text -> R: objects are named lists, arrays are unnamed lists, null is NULL
parse_json <- function(txt) {
  txt <- enc2utf8(paste(txt, collapse = "\n"))
  txt <- sub("^\ufeff", "", txt)
  m <- gregexpr("\"(?:[^\"\\\\]|\\\\.)*\"|-?\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?|true|false|null|NaN|-?Infinity|[\\[\\]{}:,]", txt, perl = TRUE)
  tok <- regmatches(txt, m)[[1]]
  pos <- 0L
  nxt <- function() { pos <<- pos + 1L; tok[pos] }
  peek <- function() tok[pos + 1L]
  value <- function() {
    t <- nxt()
    if (is.na(t)) stop("unexpected end of JSON")
    if (t == "{") {
      keys <- character(0); vals <- list()
      if (peek() == "}") { nxt(); return(structure(list(), names = character(0))) }
      repeat {
        k <- nxt()
        key <- .json_unescape(substr(k, 2, nchar(k) - 1))
        if (nxt() != ":") stop("JSON: ':' expected")
        v <- value()
        keys <- c(keys, key)
        vals[length(vals) + 1] <- list(v)
        s <- nxt()
        if (s == "}") break
        if (s != ",") stop("JSON: ',' expected")
      }
      names(vals) <- keys
      return(vals)
    }
    if (t == "[") {
      vals <- list()
      if (peek() == "]") { nxt(); return(vals) }
      repeat {
        v <- value()
        vals[length(vals) + 1] <- list(v)
        s <- nxt()
        if (s == "]") break
        if (s != ",") stop("JSON: ',' expected")
      }
      return(vals)
    }
    if (startsWith(t, "\"")) return(.json_unescape(substr(t, 2, nchar(t) - 1)))
    if (t == "true") return(TRUE)
    if (t == "false") return(FALSE)
    if (t == "null") return(NULL)
    if (t == "NaN") return(NaN)
    if (t == "Infinity") return(Inf)
    if (t == "-Infinity") return(-Inf)
    as.numeric(t)
  }
  value()
}

read_json <- function(path) {
  sz <- file.info(path)$size
  raw <- readBin(path, "raw", sz)
  if (length(raw) >= 3 && raw[1] == as.raw(0xef) && raw[2] == as.raw(0xbb) && raw[3] == as.raw(0xbf)) raw <- raw[-(1:3)]
  txt <- rawToChar(raw)
  Encoding(txt) <- "UTF-8"
  parse_json(txt)
}

jnum <- function(x) vapply(x, function(v) if (is.null(v)) NA_real_ else as.numeric(v), numeric(1))   # JSON array -> numbers
jchr <- function(x) vapply(x, function(v) if (is.null(v)) NA_character_ else as.character(v), character(1))


# ==============================================================================
#  Reading Excel (.xlsx / .xlsm) and CSV with base R
# ==============================================================================
zip_names <- function(path) utils::unzip(path, list = TRUE)$Name

zip_text <- function(path, part) {
  tmp <- tempfile("gov_one_zip_")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE, force = TRUE))
  utils::unzip(path, files = part, exdir = tmp)
  f <- file.path(tmp, part)
  if (!file.exists(f)) return("")
  x <- rawToChar(readBin(f, "raw", file.info(f)$size))
  Encoding(x) <- "UTF-8"
  x
}

# every match of a pattern in one long text, with its first capture group
re_all <- function(x, pattern) {
  m <- gregexpr(pattern, x, perl = TRUE)[[1]]
  if (m[1] < 0) return(list(start = integer(0), end = integer(0), group = character(0)))
  st <- attr(m, "capture.start")[, 1]
  ln <- attr(m, "capture.length")[, 1]
  g <- substring(x, st, st + ln - 1)
  g[st < 1] <- ""
  list(start = as.integer(m), end = as.integer(m) + attr(m, "match.length") - 1L, group = g)
}

# first capture group of a pattern in each element (NA when it does not match)
re_first <- function(x, pattern) {
  out <- rep(NA_character_, length(x))
  if (!length(x)) return(out)
  m <- regexpr(pattern, x, perl = TRUE)
  ok <- which(m > 0)
  if (length(ok)) {
    st <- attr(m, "capture.start")[ok, 1]
    ln <- attr(m, "capture.length")[ok, 1]
    out[ok] <- substring(x[ok], st, st + ln - 1)
  }
  out
}

xml_unesc <- function(x) {
  hit <- grepl("&", x, fixed = TRUE)
  if (!any(hit)) return(x)
  y <- gsub("&#13;", "", x[hit], fixed = TRUE)
  num <- unique(unlist(regmatches(y, gregexpr("&#[xX]?[0-9A-Fa-f]+;", y, perl = TRUE))))
  for (e in num) {
    body <- substr(e, 3, nchar(e) - 1)
    code <- if (substr(body, 1, 1) %in% c("x", "X")) strtoi(substr(body, 2, nchar(body)), 16L) else suppressWarnings(as.integer(body))
    if (!is.na(code) && code > 0) y <- gsub(e, intToUtf8(code), y, fixed = TRUE)
  }
  y <- gsub("&lt;", "<", y, fixed = TRUE)
  y <- gsub("&gt;", ">", y, fixed = TRUE)
  y <- gsub("&quot;", "\"", y, fixed = TRUE)
  y <- gsub("&apos;", "'", y, fixed = TRUE)
  y <- gsub("&amp;", "&", y, fixed = TRUE)
  x[hit] <- y
  x
}

# the text of every <t>...</t> run inside each element, joined
t_concat <- function(inner) {
  out <- re_first(inner, "(?s)<t\\b[^>]*>(.*?)</t>")
  out[is.na(out)] <- ""
  many <- lengths(gregexpr("<t\\b", inner, perl = TRUE)) > 1 & grepl("</t>.*</t>", inner, perl = TRUE)   # rich text: several runs
  if (any(many)) out[many] <- gsub("(?s)<t\\b[^>]*>(.*?)</t>|.", "\\1", inner[many], perl = TRUE)
  out
}

# the inside of every <tag ...>...</tag> block of one long text.  The text is cut at the closing
# tags first, so the work grows with its length and not with its square.
xml_blocks <- function(x, tag) {
  close_tag <- paste0("</", tag, ">")
  parts <- strsplit(x, close_tag, fixed = TRUE)[[1]]
  if (!endsWith(x, close_tag)) parts <- parts[-length(parts)]     # what follows the last block
  if (!length(parts)) return(character(0))
  hit <- regexpr(paste0("(?s)^.*<", tag, "\\b[^>]*>"), parts, perl = TRUE)
  out <- substring(parts, attr(hit, "match.length") + 1)
  out[hit < 0] <- ""
  out[nzchar(out)]
}

col_idx <- function(letters_) {
  u <- unique(letters_)
  v <- vapply(u, function(s) {
    ci <- 0
    for (code in utf8ToInt(s)) ci <- ci * 26 + (code - 64)
    ci - 1
  }, numeric(1))
  unname(v[match(letters_, u)])
}

# [(sheet name, zip part)] in workbook order; falls back to sorted parts
xlsx_sheets <- function(path) {
  names <- zip_names(path)
  out <- list()
  if ("xl/workbook.xml" %in% names && "xl/_rels/workbook.xml.rels" %in% names) {
    wb <- zip_text(path, "xl/workbook.xml")
    rl <- zip_text(path, "xl/_rels/workbook.xml.rels")
    rtags <- regmatches(rl, gregexpr("(?s)<Relationship\\b[^>]*>", rl, perl = TRUE))[[1]]
    rid <- re_first(rtags, "\\bId=\"([^\"]*)\"")
    rtarget <- re_first(rtags, "\\bTarget=\"([^\"]*)\"")
    keep <- !is.na(rid) & !is.na(rtarget)
    rid <- rid[keep]; rtarget <- rtarget[keep]
    stags <- regmatches(wb, gregexpr("(?s)<sheet\\b[^>]*>", wb, perl = TRUE))[[1]]
    for (tag in stags) {
      sname <- re_first(tag, "\\bname=\"([^\"]*)\"")
      r <- re_first(tag, "\\br:id=\"([^\"]*)\"")
      if (is.na(sname) || is.na(r)) next
      j <- which(rid == r)
      target <- if (length(j)) rtarget[j[length(j)]] else ""
      part <- paste0("xl/", sub("xl/", "", sub("^/+", "", target), fixed = TRUE))
      if (!(part %in% names)) {
        cand <- names[endsWith(names, basename(target))]
        part <- if (length(cand) && nzchar(target)) cand[1] else ""
      }
      if (nzchar(part)) out[[length(out) + 1]] <- list(name = xml_unesc(sname), part = part)
    }
  }
  if (!length(out)) {
    parts <- c_sort(names[startsWith(names, "xl/worksheets/sheet") & endsWith(names, ".xml")])
    out <- lapply(parts, function(p) list(name = basename(p), part = p))
  }
  out
}

# {cell style index: 'date' | 'time'} for the styles Excel shows as dates or times
date_styles <- function(path, names) {
  if (!("xl/styles.xml" %in% names)) return(character(0))
  xml <- zip_text(path, "xl/styles.xml")
  custom <- character(0)
  for (tag in regmatches(xml, gregexpr("<numFmt\\b[^>]*>", xml, perl = TRUE))[[1]]) {
    i <- re_first(tag, "\\bnumFmtId=\"(\\d+)\"")
    cc <- re_first(tag, "\\bformatCode=\"([^\"]*)\"")
    if (is.na(i) || is.na(cc)) next
    code <- tolower(gsub("\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.|_.|\\*.", "", xml_unesc(cc), perl = TRUE))
    if (grepl("general", code, fixed = TRUE) || !grepl("[dmyhs]", code)) next
    custom[i] <- if (grepl("[dy]", code)) "date" else "time"
  }
  block <- re_first(xml, "(?s)<cellXfs\\b[^>]*>(.*?)</cellXfs>")
  out <- character(0)
  if (is.na(block)) return(out)
  xfs <- regmatches(block, gregexpr("<xf\\b[^>]*>", block, perl = TRUE))[[1]]
  fid <- re_first(xfs, "\\bnumFmtId=\"(\\d+)\"")
  fid[is.na(fid)] <- "0"
  for (k in seq_along(xfs)) {
    f <- as.integer(fid[k])
    kind <- if (f %in% BUILTIN_DATE) "date" else if (f %in% BUILTIN_TIME) "time" else if (!is.na(custom[fid[k]])) custom[[fid[k]]] else NA
    if (!is.na(kind)) out[as.character(k - 1)] <- kind
  }
  out
}

# an Excel date / time number as text: 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM:SS' or 'HH:MM:SS' (NA = not a date)
serial_text <- function(val, kind) {
  f <- suppressWarnings(as.numeric(val))
  out <- rep(NA_character_, length(f))
  ok <- !is.na(f) & f >= 0 & f < 2958466
  tm <- ok & (kind == "time" | f < 1)
  if (any(tm)) {
    s <- round((f[tm] - floor(f[tm])) * 86400)
    out[tm] <- sprintf("%02d:%02d:%02d", as.integer((s %/% 3600) %% 24), as.integer((s %/% 60) %% 60), as.integer(s %% 60))
  }
  dd <- ok & !tm
  if (any(dd)) {
    dt <- EXCEL_EPOCH_DAYS * 86400 + round(f[dd] * 86400)
    whole <- dt %% 86400 == 0
    out[dd] <- ifelse(whole, dt_format(dt, "%Y-%m-%d"), dt_format(dt, "%Y-%m-%d %H:%M:%S"))
  }
  out
}

# one worksheet -> list(m = character matrix of the cells, datecells = number of date cells per column index)
# A cell Excel shows as a date comes back as 'YYYY-MM-DD' whatever the column is called.
read_sheet <- function(path, part, max_rows = NULL) {
  names <- zip_names(path)
  shared <- character(0)
  if ("xl/sharedStrings.xml" %in% names) shared <- xml_unesc(t_concat(xml_blocks(zip_text(path, "xl/sharedStrings.xml"), "si")))
  dstyles <- date_styles(path, names)
  empty <- list(m = matrix("", 0, 0), datecells = integer(0))
  bodies <- xml_blocks(zip_text(path, part), "row")
  cells <- regmatches(bodies, gregexpr("(?s)<c\\b[^>]*?(?:/>|>.*?</c>)", bodies, perl = TRUE))
  cells <- cells[lengths(cells) > 0]                              # a row without cells is not a row
  if (!is.null(max_rows)) cells <- utils::head(cells, max_rows)
  if (!length(cells)) return(empty)
  rid <- rep(seq_along(cells), lengths(cells))
  cells <- unlist(cells)
  attrs <- re_first(cells, "^<c\\b([^>]*?)(?:/>|>)")
  attrs[is.na(attrs)] <- ""
  inner <- re_first(cells, "(?s)^<c\\b[^>]*?>(.*)</c>$")
  inner[is.na(inner)] <- ""
  ref <- re_first(attrs, "\\br=\"([A-Z]+)\\d*\"")
  ci <- col_idx(ifelse(is.na(ref), "A", ref))
  if (anyNA(ref)) {                                             # no cell reference: the cells follow each other
    seq_in_row <- ave(seq_along(rid), rid, FUN = seq_along) - 1
    ci[is.na(ref)] <- seq_in_row[is.na(ref)]
  }
  tt <- re_first(attrs, "\\bt=\"([^\"]*)\"")
  tt[is.na(tt)] <- ""
  val <- rep("", length(rid))
  inline <- tt == "inlineStr" & nzchar(inner)
  if (any(inline)) val[inline] <- xml_unesc(t_concat(inner[inline]))
  v <- re_first(inner, "(?s)<v>(.*?)</v>")
  rest <- !inline & nzchar(inner) & !is.na(v)
  sh <- rest & tt == "s"
  if (any(sh)) {
    k <- suppressWarnings(as.integer(v[sh])) + 1L
    got <- ifelse(!is.na(k) & k >= 1 & k <= length(shared), shared[pmin(pmax(k, 1L), max(1L, length(shared)))], "")
    got[is.na(got)] <- ""
    val[sh] <- got
  }
  other <- rest & tt != "s"
  datecells <- integer(0)
  if (any(other)) {
    val[other] <- xml_unesc(v[other])
    if (length(dstyles)) {
      sidx <- re_first(attrs, "\\bs=\"(\\d+)\"")
      cand <- which(other & tt %in% c("", "n") & !is.na(sidx))
      if (length(cand)) {
        kind <- unname(dstyles[sidx[cand]])
        has <- !is.na(kind)
        if (any(has)) {
          txt <- serial_text(v[cand[has]], kind[has])
          good <- !is.na(txt)
          idx <- cand[has][good]
          val[idx] <- txt[good]
          tb <- table(ci[idx])
          datecells <- structure(as.integer(tb), names = names(tb))
        }
      }
    }
  }
  rr <- match(rid, unique(rid))
  m <- matrix("", nrow = max(rr), ncol = max(ci) + 1)
  m[cbind(rr, ci + 1)] <- val
  list(m = m, datecells = datecells)
}

# a CSV file -> character matrix.  Encoding (UTF-8, UTF-16, Windows Hebrew) and delimiter are detected.
read_csv_rows <- function(path) {
  raw <- readBin(path, "raw", file.info(path)$size)
  if (!length(raw)) return(matrix("", 0, 0))
  if (length(raw) >= 2 && ((raw[1] == as.raw(0xff) && raw[2] == as.raw(0xfe)) || (raw[1] == as.raw(0xfe) && raw[2] == as.raw(0xff)))) {
    text <- iconv(list(raw), from = "UTF-16", to = "UTF-8")
  } else {
    if (length(raw) >= 3 && raw[1] == as.raw(0xef) && raw[2] == as.raw(0xbb) && raw[3] == as.raw(0xbf)) raw <- raw[-(1:3)]
    text <- rawToChar(raw[raw != as.raw(0)])
    Encoding(text) <- "UTF-8"
    if (!validUTF8(text)) text <- iconv(list(raw), from = "CP1255", to = "UTF-8", sub = "\ufffd")
  }
  if (is.na(text)) return(matrix("", 0, 0))
  text <- sub("^\ufeff", "", text)
  head <- substr(text, 1, 20000)
  delims <- c(",", ";", "\t", "|")
  cnt <- vapply(delims, function(d) lengths(regmatches(head, gregexpr(d, head, fixed = TRUE))), numeric(1))
  delim <- delims[which.max(cnt)]
  # the line ends Python's splitlines() knows; fixed patterns only, so a large file stays fast
  b <- charToRaw(text)                                              # single bytes, never part of a UTF-8 letter
  cr <- which(b == as.raw(0x0d))
  crlf <- cr[cr < length(b)]
  crlf <- crlf[b[crlf + 1L] == as.raw(0x0a)]
  if (length(crlf)) b <- b[-crlf]
  b[b %in% as.raw(c(0x0d, 0x0b, 0x0c, 0x1c, 0x1d, 0x1e))] <- as.raw(0x0a)
  text <- rawToChar(b)
  Encoding(text) <- "UTF-8"
  for (sep in c("\u0085", "\u2028", "\u2029")) if (grepl(sep, text, fixed = TRUE)) text <- gsub(sep, "\n", text, fixed = TRUE)
  lines <- strsplit(text, "\n", fixed = TRUE)[[1]]
  if (!length(lines)) return(matrix("", 0, 0))
  rows <- csv_split(lines, delim)
  n <- lengths(rows)
  m <- matrix("", length(rows), max(c(n, 1)))
  m[cbind(rep.int(seq_along(rows), n), sequence(n))] <- unlist(rows, use.names = FALSE)
  m
}

# The fields of lines that hold a quote, by the rules of Python's csv module: a field is quoted only
# when it STARTS with a quote (so the quote in a header like  ID"NO  is an ordinary character), a doubled
# quote inside a quoted field is one quote, and text after the closing quote is kept as it is.
# Returns list(fields = one character vector per line, open = the line ends inside a quoted field).
.csv_quoted <- function(lines, delim) {
  no <- paste0("[^", delim, "]*")
  toks <- regmatches(paste0(lines, delim), gregexpr(paste0("(?:\"(?:[^\"]|\"\")*+\"", no, "|", no, ")(?:[", delim, "]|$)"),
                                                     paste0(lines, delim), perl = TRUE))
  id <- rep.int(seq_along(toks), lengths(toks))
  flat <- unlist(toks, use.names = FALSE)
  keep <- endsWith(flat, delim)                                  # every real field ends with a delimiter here
  flat <- flat[keep]; id <- id[keep]
  flat <- substr(flat, 1L, nchar(flat) - 1L)
  open <- logical(length(lines))
  q <- startsWith(flat, "\"")
  if (any(q)) {
    fq <- flat[q]
    pat <- "^\"((?:[^\"]|\"\")*+)\"(.*)$"
    ok <- grepl(pat, fq, perl = TRUE)
    fq[ok] <- paste0(gsub("\"\"", "\"", sub(pat, "\\1", fq[ok], perl = TRUE), fixed = TRUE), sub(pat, "\\2", fq[ok], perl = TRUE))
    flat[q] <- fq
    open[unique(id[q][!ok])] <- TRUE
  }
  list(fields = unname(split(flat, factor(id, levels = seq_along(lines)))), open = open)
}

# lines -> a list of rows (character vectors).  A quoted field that runs over several lines is one
# field; like Python's csv module on lines without their line ends, it is joined without a line break.
csv_split <- function(lines, delim) {
  nl <- length(lines)
  out <- vector("list", nl)
  hasq <- grepl("\"", lines, fixed = TRUE)
  plain <- which(!hasq)
  if (length(plain)) out[plain] <- strsplit(paste0(lines[plain], delim), delim, fixed = TRUE)
  qi <- which(hasq)
  drop <- logical(nl)
  if (length(qi)) {
    res <- .csv_quoted(lines[qi], delim)
    out[qi] <- res$fields
    last <- 0
    for (i in qi[res$open]) {
      if (i <= last) next                                        # already taken as the continuation of a field
      rec <- lines[i]
      j <- i
      repeat {
        j <- j + 1
        if (j > nl) {                                            # the file ends inside the quoted field
          one <- .csv_quoted(paste0(rec, "\""), delim)
          j <- nl
          break
        }
        rec <- paste0(rec, lines[j])
        one <- .csv_quoted(rec, delim)
        if (!one$open) break
      }
      out[[i]] <- one$fields[[1]]
      if (j > i) drop[(i + 1):j] <- TRUE
      last <- j
    }
  }
  out[!drop]
}

# Index (from 0) of the header row.  Ministry sheets often start with a title and a date; the header
# is the first row that is nearly as wide as the data below it and holds labels, not numbers.
find_header_row <- function(m) {
  look <- m[seq_len(min(60, nrow(m))), , drop = FALSE]
  if (nrow(look) < 3) return(0)
  st <- matrix(pystrip(look), nrow = nrow(look))
  widths <- sort(rowSums(st != ""))
  typical <- widths[length(widths) %/% 2 + 1]
  if (typical < 2) return(0)
  for (i in seq_len(min(30, nrow(look)))) {
    filled <- st[i, st[i, ] != ""]
    if (length(filled) < max(2, 0.6 * typical)) next
    labels <- sum(is.na(to_float(filled)) & is.na(parse_date_dt(filled)))
    if (labels >= 0.8 * length(filled) && length(unique(filled)) >= 0.8 * length(filled)) return(i - 1)
  }
  0
}

# duplicate or blank header names never collapse into one column
dedupe_header <- function(header) {
  out <- character(0); notes <- character(0); seen <- list()
  for (i in seq_along(header)) {
    h <- pystrip(header[i])
    if (!nzchar(h)) {
      h <- sprintf("col_%d", i)
      notes <- c(notes, sprintf("blank header at column %d -> %s", i, h))
    }
    if (!is.null(seen[[h]])) {
      seen[[h]] <- seen[[h]] + 1
      new <- sprintf("%s_%d", h, seen[[h]])
      notes <- c(notes, sprintf("duplicate header '%s' at column %d -> %s", h, i, new))
      h <- new
    } else seen[[h]] <- 1
    out <- c(out, h)
  }
  list(header = out, notes = notes)
}

file_stem <- function(fname) sub("(?<=.)\\.[^.]*$", "", fname, perl = TRUE)

# Every non-empty sheet of a workbook (or the one table of a CSV) as a table:
#   base, file, sheet, sheet_index, file_type, header, cols (named list of character vectors),
#   n (rows), notes, excel_dates (columns stored as real Excel dates), header_row.
# The first sheet is named after the file; further sheets are 'file#sheet'.
read_file_tables <- function(path) {
  fname <- basename(path)
  stem <- file_stem(fname)
  low <- tolower(path)
  if (endsWith(low, ".csv") || endsWith(low, ".tsv")) {
    sheets <- list(list(name = "", m = read_csv_rows(path), datecells = integer(0)))
    ftype <- "csv"
  } else {
    ftype <- "xlsx"
    sheets <- lapply(xlsx_sheets(path), function(s) {
      r <- read_sheet(path, s$part)
      list(name = s$name, m = r$m, datecells = r$datecells)
    })
  }
  out <- list()
  for (s in sheets) {
    m <- s$m
    if (!nrow(m) || !ncol(m)) next
    st <- matrix(pystrip(m), nrow = nrow(m))
    keep <- rowSums(st != "") > 0
    st <- st[keep, , drop = FALSE]
    if (!nrow(st)) next
    h <- find_header_row(st)
    head <- st[h + 1, ]
    while (length(head) && !nzchar(head[length(head)])) head <- head[-length(head)]
    if (!length(head)) next
    dd <- dedupe_header(head)
    header <- dd$header; notes <- dd$notes
    if (h) notes <- c(notes, sprintf("%d row(s) above the header were skipped (titles)", h))
    width <- length(header)
    body <- st[-seq_len(h + 1), , drop = FALSE]
    cols <- lapply(seq_len(width), function(j) if (j <= ncol(body)) unname(body[, j]) else rep("", nrow(body)))
    names(cols) <- header
    if (ncol(body) > width) {
      wide <- sum(rowSums(body[, (width + 1):ncol(body), drop = FALSE] != "") > 0)
      if (wide) notes <- c(notes, sprintf("%d row(s) have values to the right of the last header cell; those cells were ignored", wide))
    }
    excel_dates <- character(0)
    for (jn in names(s$datecells)) {
      j <- as.integer(jn)
      if (j < width) {
        filled <- sum(cols[[j + 1]] != "")
        if (filled && s$datecells[[jn]] >= 0.9 * filled) excel_dates <- c(excel_dates, header[j + 1])
      }
    }
    out[[length(out) + 1]] <- list(base = if (!length(out)) stem else sprintf("%s#%s", stem, s$name), file = fname, sheet = s$name,
                                   sheet_index = length(out), file_type = ftype, header = header, cols = cols, n = nrow(body),
                                   notes = notes, excel_dates = excel_dates, header_row = h)
  }
  out
}

# (data files to read, files that cannot be read and why).  Our own outputs are never inputs.
find_inputs <- function(folder) {
  good <- character(0); bad <- list()
  for (name in c_sort(list.files(folder, all.files = TRUE, no.. = TRUE))) {
    p <- file.path(folder, name)
    low <- tolower(name)
    if (dir.exists(p) || !file.exists(p) || any(startsWith(low, OUTPUT_PREFIXES))) next
    if (grepl("\\.(xlsx|xlsm|csv|tsv)$", low)) good <- c(good, p)
    else if (endsWith(low, ".xls")) bad[[length(bad) + 1]] <- c(name, "old .xls format: open it in Excel and save as .xlsx")
    else if (endsWith(low, ".xlsb")) bad[[length(bad) + 1]] <- c(name, "binary .xlsb format: open it in Excel and save as .xlsx")
  }
  list(good = good, bad = bad)
}

# ==============================================================================
#  Writing Excel (several sheets, real date cells) and CSV with base R
# ==============================================================================
.le <- function(x, bytes) as.raw(bitwAnd(floor(x / 256^(0:(bytes - 1))) %% 256, 255L))

# one zip entry: the bytes are deflated by R's own gzip writer, which also gives the checksum
.deflate <- function(data) {
  tmp <- tempfile("gov_one_gz_")
  on.exit(unlink(tmp))
  con <- gzfile(tmp, "wb")
  writeBin(data, con)
  close(con)
  gz <- readBin(tmp, "raw", file.info(tmp)$size)
  n <- length(gz)
  flg <- as.integer(gz[4])
  pos <- 11
  if (bitwAnd(flg, 4L)) pos <- pos + 2 + as.integer(gz[pos]) + 256 * as.integer(gz[pos + 1])
  if (bitwAnd(flg, 8L)) pos <- pos + which(gz[pos:n] == as.raw(0))[1]
  if (bitwAnd(flg, 16L)) pos <- pos + which(gz[pos:n] == as.raw(0))[1]
  if (bitwAnd(flg, 2L)) pos <- pos + 2
  list(data = gz[pos:(n - 8)], crc = gz[(n - 7):(n - 4)])
}

zip_write <- function(path, entries) {
  con <- file(path, "wb")
  on.exit(close(con))
  central <- raw(0); offset <- 0
  for (e in entries) {
    name <- charToRaw(enc2utf8(e$name))
    z <- .deflate(e$data)
    head <- c(.le(20, 2), .le(2048, 2), .le(8, 2), .le(0, 2), .le(33, 2), z$crc, .le(length(z$data), 4), .le(length(e$data), 4),
              .le(length(name), 2), .le(0, 2))
    writeBin(c(as.raw(c(0x50, 0x4b, 0x03, 0x04)), head, name, z$data), con)
    central <- c(central, as.raw(c(0x50, 0x4b, 0x01, 0x02)), .le(20, 2), head, .le(0, 2), .le(0, 2), .le(0, 2), .le(0, 4), .le(offset, 4), name)
    offset <- offset + 30 + length(name) + length(z$data)
  }
  writeBin(c(central, as.raw(c(0x50, 0x4b, 0x05, 0x06)), .le(0, 2), .le(0, 2), .le(length(entries), 2), .le(length(entries), 2),
             .le(length(central), 4), .le(offset, 4), .le(0, 2)), con)
}

.STYLES_XML <- paste0('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
  '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>',
  '</styleSheet>')

xl_esc <- function(t) gsub(">", "&gt;", gsub("<", "&lt;", gsub("&", "&amp;", as.character(t), fixed = TRUE), fixed = TRUE), fixed = TRUE)

xl_colref <- function(ci) {                                       # 0 -> A, 25 -> Z, 26 -> AA
  vapply(ci, function(cc) {
    s <- ""; cc <- cc + 1
    while (cc > 0) { r <- (cc - 1) %% 26; s <- paste0(LETTERS[r + 1], s); cc <- (cc - 1) %/% 26 }
    s
  }, character(1))
}

.sheet_names <- function(names) {
  used <- character(0); out <- character(0)
  for (name in names) {
    s <- substr(pystrip(gsub("[\\[\\]:*?/\\\\]", " ", if (nzchar(name)) name else "Sheet", perl = TRUE)), 1, 31)
    if (!nzchar(s)) s <- "Sheet"
    base <- s; k <- 2
    while (tolower(s) %in% used) { s <- sprintf("%s %d", substr(base, 1, 28), k); k <- k + 1 }
    used <- c(used, tolower(s)); out <- c(out, s)
  }
  out
}

# One sheet = list(name, header, cols): each column is a character vector (text), a numeric vector
# (numbers; NA = empty) or a numeric vector with attribute "dt" = TRUE (date-times in seconds; NA = empty).
.col_cells <- function(v, ref) {
  out <- rep("", length(v))
  if (isTRUE(attr(v, "dt"))) {
    ok <- !is.na(v)
    f <- (v[ok] - EXCEL_EPOCH_DAYS * 86400) / 86400
    whole <- v[ok] %% 86400 == 0
    out[ok] <- ifelse(whole, sprintf('<c r="%s" s="1"><v>%.0f</v></c>', ref[ok], f), sprintf('<c r="%s" s="2"><v>%.8f</v></c>', ref[ok], f))
  } else if (is.numeric(v)) {
    ok <- !is.na(v) & is.finite(v)
    out[ok] <- sprintf('<c r="%s"><v>%s</v></c>', ref[ok], vapply(v[ok], .json_num, character(1)))
  } else {
    num <- isTRUE(attr(v, "num"))                                # text that holds numbers is written as numbers
    v <- as.character(v)
    ok <- !is.na(v) & v != ""
    isnum <- ok & num & grepl("^[+-]?([0-9]+\\.?[0-9]*|\\.[0-9]+)$", v)
    if (any(isnum)) {
      out[isnum] <- sprintf('<c r="%s"><v>%s</v></c>', ref[isnum], vapply(as.numeric(v[isnum]), .json_num, character(1)))
      ok <- ok & !isnum
    }
    out[ok] <- sprintf('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>', ref[ok], xl_esc(v[ok]))
  }
  out
}

write_workbook <- function(path, sheets) {
  names <- .sheet_names(vapply(sheets, function(s) s$name, character(1)))
  n <- length(sheets)
  X <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  ctypes <- paste0(X, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    paste0(sprintf('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', seq_len(n)), collapse = ""),
    '</Types>')
  rels <- paste0(X, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>')
  wb <- paste0(X, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>',
    paste0(sprintf('<sheet name="%s" sheetId="%d" r:id="rId%d"/>', gsub("\"", "&quot;", xl_esc(names), fixed = TRUE), seq_len(n), seq_len(n)), collapse = ""),
    '</sheets></workbook>')
  wbrels <- paste0(X, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    paste0(sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>', seq_len(n), seq_len(n)), collapse = ""),
    sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', n + 1),
    '</Relationships>')
  u8 <- function(s) charToRaw(enc2utf8(s))
  entries <- list(list(name = "[Content_Types].xml", data = u8(ctypes)), list(name = "_rels/.rels", data = u8(rels)),
                  list(name = "xl/workbook.xml", data = u8(wb)), list(name = "xl/_rels/workbook.xml.rels", data = u8(wbrels)),
                  list(name = "xl/styles.xml", data = u8(.STYLES_XML)))
  for (i in seq_len(n)) {
    s <- sheets[[i]]
    nr <- if (length(s$cols)) length(s$cols[[1]]) else 0
    letters_ <- xl_colref(seq_along(s$header) - 1)
    head <- paste0('<row r="1">', paste0(sprintf('<c r="%s1" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>', letters_, xl_esc(s$header)), collapse = ""), '</row>')
    body <- ""
    if (nr) {
      rowno <- seq_len(nr) + 1
      cells <- rep("", nr)
      for (j in seq_along(s$header)) cells <- paste0(cells, .col_cells(s$cols[[s$header[j]]], paste0(letters_[j], rowno)))
      body <- paste0(sprintf('<row r="%d">', rowno), cells, "</row>", collapse = "")
    }
    xml <- paste0(X, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>', head, body, '</sheetData></worksheet>')
    entries[[length(entries) + 1]] <- list(name = sprintf("xl/worksheets/sheet%d.xml", i), data = u8(xml))
  }
  zip_write(path, entries)
}

write_csv <- function(path, header, cols) {
  q <- function(v) {
    if (isTRUE(attr(v, "dt"))) v <- ifelse(v %% 86400 == 0, dt_format(v, "%Y-%m-%d"), dt_format(v, "%Y-%m-%d %H:%M:%S"))
    else if (is.numeric(v)) v <- ifelse(is.na(v), "", vapply(v, .json_num, character(1)))
    else if (isTRUE(attr(v, "num"))) {                           # 007 -> 7, 5.10 -> 5.1, as a spreadsheet would show it
      isnum <- grepl("^[+-]?([0-9]+\\.?[0-9]*|\\.[0-9]+)$", v)
      v[isnum] <- vapply(as.numeric(v[isnum]), .json_num, character(1))
    }
    v <- as.character(v)
    v[is.na(v)] <- ""
    need <- grepl("[\",\r\n]", v)
    v[need] <- paste0("\"", gsub("\"", "\"\"", v[need], fixed = TRUE), "\"")
    v
  }
  lines <- paste(q(header), collapse = ",")
  if (length(cols) && length(cols[[1]])) lines <- c(lines, do.call(paste, c(lapply(header, function(h) q(cols[[h]])), sep = ",")))
  con <- file(path, "wb")
  on.exit(close(con))
  writeBin(c(as.raw(c(0xef, 0xbb, 0xbf)), charToRaw(enc2utf8(paste0(paste(lines, collapse = "\r\n"), "\r\n")))), con)
}


# ==============================================================================
#  Relationships between tables (found from the values, on the premises; only
#  column names and group-size distributions are written down)
# ==============================================================================
.split_ws <- function(s) {
  p <- strsplit(s, paste0(.WS, "+"), perl = TRUE)[[1]]
  p[nzchar(p)]
}
.has_tok <- function(nm, toks) any(vapply(toks, function(tk) grepl(tk, nm, fixed = TRUE), logical(1)))
.is_ascii <- function(s) !grepl("[^\\x01-\\x7f]", s, perl = TRUE)

# trim, and treat '_' and ' ' as the same
norm_name <- function(s) paste(.split_ws(gsub("_", " ", pystrip(s), fixed = TRUE)), collapse = " ")

nonmissing <- function(v) v[!is_missing(v)]

all_int_text <- function(v) length(v) > 0 && all(grepl("^[+-]?[0-9]+(_[0-9]+)*$", v, perl = TRUE))

# TRUE for date / time columns (by name or by readable dates), so they are never taken for keys
dateish <- function(name, present) {
  nm <- tolower(pystrip(name))
  if (.has_tok(nm, DATE_TOKENS)) return(TRUE)
  pv <- utils::head(present, 200)
  if (length(pv) && !all_int_text(pv)) {
    hits <- sum(!is.na(parse_date_dt(pv)))
    if (hits && hits / length(pv) > 0.9) return(TRUE)
  }
  FALSE
}

# A header that names a measurement (a count, an age, an amount) -- never a key.
# English words must match whole words ('age' does not match 'stage'); Hebrew matches as text.
measureish <- function(name) {
  nm <- tolower(pystrip(name))
  words <- strsplit(nm, "[^a-z0-9]+", perl = TRUE)[[1]]
  for (tok in MEASURE_NAME_TOKENS) {
    if (.is_ascii(tok)) {
      if ((grepl(" ", tok, fixed = TRUE) && grepl(tok, nm, fixed = TRUE)) || tok %in% words) return(TRUE)
    } else if (grepl(tok, nm, fixed = TRUE)) return(TRUE)
  }
  FALSE
}

is_sensitive_name <- function(name) {
  nm <- paste(.split_ws(gsub("_", " ", tolower(pystrip(name)), fixed = TRUE)), collapse = " ")
  if (nm %in% SENSITIVE_WORDS || .has_tok(nm, SENSITIVE_PHRASES)) return(TRUE)
  any(strsplit(nm, paste0("(?:", .WS, "|[-./,()])+"), perl = TRUE)[[1]] %in% SENSITIVE_ROLE_WORDS)
}

name_words <- function(name) {
  w <- strsplit(tolower(pystrip(name)), paste0("(?:", .WS, "|[_./'\"-])+"), perl = TRUE)[[1]]
  setdiff(unique(w[nchar(w) >= 2]), GENERIC_NAME_WORDS)
}

# does column a decide column b (the same a always comes with the same b)?
determines <- function(av, bv) {
  ok <- !is_missing(av)
  av <- av[ok]; bv <- bv[ok]
  all(bv == bv[match(av, av)])
}

# distribution of the number of rows per key value, as robust quantiles
fanout <- function(col) {
  sizes <- sort(as.numeric(count_first(nonmissing(col))))
  q <- vapply(QUANTILE_GRID, function(p) percentile(sizes, p), numeric(1))
  if (USE_ROBUST_BOUNDS && length(sizes)) {
    q[1] <- percentile(sizes, P_LOW)
    q[length(q)] <- percentile(sizes, P_HIGH)
  }
  list(grid = QUANTILE_GRID, q = pyround(q, 6), n = length(sizes))
}

own_key <- function(header, uniq) {
  uniques <- header[uniq[header]]
  if (!length(uniques)) return(NULL)
  for (cc in uniques) {
    cl <- tolower(cc)
    if (cl %in% c("id", "url") || grepl("(_id|_key|_code)$", cl)) return(cc)
  }
  uniques[1]
}

# parents before children; parent: named list, base -> parent base or NULL
topo <- function(bases, parent) {
  children <- stats::setNames(lapply(bases, function(b) character(0)), bases)
  indeg <- stats::setNames(rep(0, length(bases)), bases)
  for (b in bases) {
    p <- parent[[b]]
    if (!is.null(p) && nzchar(p)) {
      children[[p]] <- c(children[[p]], b)
      indeg[[b]] <- indeg[[b]] + 1
    }
  }
  q <- c_sort(bases[indeg == 0])
  ord <- character(0)
  while (length(q)) {
    b <- q[1]; q <- q[-1]
    ord <- c(ord, b)
    for (ch in c_sort(children[[b]])) {
      indeg[[ch]] <- indeg[[ch]] - 1
      if (indeg[[ch]] == 0) q <- c(q, ch)
    }
    q <- c_sort(q)
  }
  c(ord, bases[!(bases %in% ord)])
}

# what a column needs to be a key: enough distinct values, whole numbers or text, not a date, not a measure
col_facts <- function(name, col) {
  vals <- nonmissing(col)
  if (!length(vals)) return(NULL)
  s <- unique(vals)
  if (length(s) <= MIN_GROUP_KEYS) return(NULL)
  f <- to_float(stride_sample(vals, 400))
  numeric <- !anyNA(f)
  is_int <- numeric && all(is.finite(f) & f == floor(f))
  if ((numeric && !is_int) || dateish(name, vals) || measureish(name)) return(NULL)
  out <- list(set = s, n = length(vals), distinct = length(s), ratio = length(s) / length(vals), int = is_int)
  if (is_int) {
    nums <- to_float(s)
    if (anyNA(nums)) out$int <- FALSE else { out$lo <- min(nums); out$hi <- max(nums) }
  }
  out
}

.lex_gt <- function(a, b) {
  d <- which(a != b)
  length(d) > 0 && a[d[1]] > b[d[1]]
}

# tables: named list (base -> table).  A column of one table is a link to a column of another when
# at least 80% of its distinct values are found there and the other side is close to unique.  Two
# whole-number columns with different names must also cover a similar range, so an age column is
# never 'linked' to a running number.  forced: list of c(child, column, parent, column);
# forbidden: 'table.column'.  Returns the schema written to schema.json.
detect_links <- function(tables, forced = list(), forbidden = character(0)) {
  bases <- names(tables)
  header <- lapply(tables, function(t) t$header)
  facts <- list(); uniq <- list()
  for (b in bases) {
    fb <- list(); ub <- logical(0)
    for (cc in header[[b]]) {
      v <- nonmissing(tables[[b]]$cols[[cc]])
      ub[cc] <- length(v) > 0 && !anyDuplicated(v)
      f <- col_facts(cc, tables[[b]]$cols[[cc]])
      if (!is.null(f)) fb[[cc]] <- f
    }
    facts[[b]] <- fb
    uniq[[b]] <- ub
  }
  cand <- stats::setNames(lapply(bases, function(b) list()), bases)
  for (b in bases) for (cc in names(facts[[b]])) {
    cf <- facts[[b]][[cc]]
    if (paste0(b, ".", cc) %in% forbidden) next
    best <- NULL
    for (p in bases) {
      if (p == b) next
      for (pc in names(facts[[p]])) {
        pf <- facts[[p]][[pc]]
        if (pf$ratio < ID_DISTINCT_RATIO || pf$distinct < cf$distinct) next
        overlap <- sum(cf$set %in% pf$set) / cf$distinct
        if (overlap < 0.8) next
        same_name <- tolower(norm_name(cc)) == tolower(norm_name(pc))
        if (uniq[[b]][[cc]] && !same_name) next                 # two running numbers are not a link
        if (cf$int && pf$int && !same_name) {
          span_p <- pf$hi - pf$lo
          if (span_p == 0) span_p <- 1
          if ((cf$hi - cf$lo) / span_p < 0.5) next              # e.g. ages 18-90 inside ids 1-5000
        }
        score <- c(same_name, length(intersect(name_words(cc), name_words(pc))), overlap, uniq[[p]][[pc]], pf$distinct)
        if (is.null(best) || .lex_gt(score, best$score)) best <- list(score = score, p = p, pc = pc)
      }
    }
    if (!is.null(best)) {
      ps <- facts[[best$p]][[best$pc]]$set
      mirror <- uniq[[b]][[cc]] && length(ps) == length(cf$set) && all(cf$set %in% ps) && c_gt(b, best$p)
      if (!mirror) cand[[b]][[cc]] <- list(p = best$p, pc = best$pc, ov = best$score[3])
    }
  }
  for (a in bases) {                                            # A->B and B->A: keep the stronger direction
    snap_a <- cand[[a]]
    for (ca in names(snap_a)) {
      pa <- snap_a[[ca]]$p; oa <- snap_a[[ca]]$ov
      snap_p <- cand[[pa]]
      for (cb in names(snap_p)) {
        if (snap_p[[cb]]$p == a) {
          if (snap_p[[cb]]$ov <= oa) cand[[pa]][[cb]] <- NULL else cand[[a]][[ca]] <- NULL
        }
      }
    }
  }
  for (f in forced) {
    if (f[1] %in% bases && f[3] %in% bases && f[2] %in% header[[f[1]]] && f[4] %in% header[[f[3]]])
      cand[[f[1]]][[f[2]]] <- list(p = f[3], pc = f[4], ov = 1)
  }
  files <- list(); parent <- list()
  for (b in bases) {
    own <- own_key(header[[b]], uniq[[b]])
    links <- list()
    for (cc in names(cand[[b]])) {
      L <- cand[[b]][[cc]]
      fo <- fanout(tables[[b]]$cols[[cc]])
      links[[length(links) + 1]] <- list(parent = L$p, link = cc, parent_key = L$pc, share_found = pyround(L$ov, 3),
                                         fanout_grid = arr(fo$grid), fanout_quantiles = arr(fo$q))
    }
    if (!length(links)) {
      files[[b]] <- list(file = tables[[b]]$file, parent = NULL, link = NULL, parent_key = NULL, inherited = list(), key = own,
                         fanout_grid = NULL, fanout_quantiles = NULL, links = list())
      parent[b] <- list(NULL)
      next
    }
    nd <- vapply(links, function(L) length(unique(nonmissing(tables[[b]]$cols[[L$link]]))), numeric(1))
    pi <- which.max(nd)
    for (i in seq_along(links)) links[[i]]$primary <- (i == pi)
    pr <- links[[pi]]
    phead <- header[[pr$parent]]
    inherited <- character(0)
    for (cc in header[[b]]) {
      if (cc != pr$link && cc %in% phead && determines(tables[[b]]$cols[[pr$link]], tables[[b]]$cols[[cc]])) inherited <- c(inherited, cc)
    }
    files[[b]] <- list(file = tables[[b]]$file, parent = pr$parent, link = pr$link, parent_key = pr$parent_key,
                       inherited = arr(inherited), key = own, fanout_grid = pr$fanout_grid, fanout_quantiles = pr$fanout_quantiles,
                       links = links)
    parent[[b]] <- pr$parent
  }
  list(tool = VERSION, files = files, order = arr(topo(bases, parent)))
}

# every column that takes part in a detected relationship, per table (profiled as identifiers)
key_columns_per_base <- function(schema) {
  keys <- stats::setNames(lapply(names(schema$files), function(b) character(0)), names(schema$files))
  for (b in names(schema$files)) {
    for (L in schema$files[[b]]$links) {
      keys[[b]] <- union(keys[[b]], L$link)
      keys[[L$parent]] <- union(keys[[L$parent]], L$parent_key)
    }
  }
  keys
}


# ==============================================================================
#  Profiling one column (aggregates only)
# ==============================================================================
numeric_bins <- function(values, nbins = CROSSTAB_NUMERIC_BINS) {
  nums <- to_float(values)
  nums <- sort(nums[!is.na(nums)])
  if (!length(nums)) return(numeric(0))
  sort(unique(vapply(seq_len(nbins - 1), function(i) percentile(nums, 100 * i / nbins), numeric(1))))
}

bin_of <- function(v, edges) {
  f <- to_float(v)
  out <- rep("", length(f))
  ok <- !is.na(f)
  out[ok] <- as.character(findInterval(f[ok], edges, left.open = TRUE))
  out
}

.fobj <- function(v) if (length(v)) v else obj()                 # a {label: count} object, also when empty

.group_profile <- function(present) {
  sizes <- sort(as.numeric(count_first(present)))
  gq <- vapply(QUANTILE_GRID, function(p) percentile(sizes, p), numeric(1))
  if (USE_ROBUST_BOUNDS) {
    gq[1] <- percentile(sizes, P_LOW)
    gq[length(gq)] <- percentile(sizes, P_HIGH)
  }
  list(kind = "identifier_group", n_groups = length(sizes), group_size_grid = arr(QUANTILE_GRID), group_size_quantiles = arr(pyround(gq, 6)))
}

.rare_freqs <- function(present, k) {
  cnt <- count_first(present)
  cnt <- cnt[order(-cnt)]
  rare <- cnt < k
  names(cnt)[rare] <- sprintf("RARE_%03d", seq_len(sum(rare)))
  list(freqs = cnt, n_rare = sum(rare))
}

profile_column <- function(name, values, min_cell_count = MIN_CELL_COUNT, force_key = FALSE) {
  present <- values[!is_missing(values)]
  n <- length(values)
  nd <- length(unique(present))
  col <- list(name = name, n = n, missing_rate = if (n) pyround((n - length(present)) / n, 6) else 0, n_unique = nd)
  if (!length(present)) {
    col$kind <- "empty"
    return(col)
  }
  # Keys found by the relationship pass are profiled as identifiers whatever their values:
  # only counts and text lengths are stored -- no key value ever leaves the premises.
  if (force_key) {
    if (nd == length(present)) {
      L <- nchar(present, type = "chars")
      return(c(col, list(kind = "identifier_unique", len_min = min(L), len_max = max(L))))
    }
    return(c(col, .group_profile(present)))
  }
  f <- to_float(present)
  all_numeric <- !anyNA(f)
  all_int <- all_numeric && all(is.finite(f) & f == floor(f))
  smp <- stride_sample(present)
  is_date <- !all_numeric && sum(!is.na(parse_date_dt(smp))) / length(smp) > 0.9
  is_bool <- all(unique(tolower(present)) %in% c("0", "1", "true", "false", "yes", "no", "y", "n")) && nd <= 3
  nm <- tolower(pystrip(name))
  nm_words <- .split_ws(gsub("[/_-]", " ", nm))
  is_id_name <- any(vapply(ID_NAME_HINTS, function(h) nm == h || endsWith(nm, paste0("_", h)) || h %in% nm_words, logical(1)))
  if (measureish(name)) is_id_name <- FALSE

  if (endsWith(nm, "_month")) {                                  # month codes are labels, not magnitudes
    r <- .rare_freqs(present, min_cell_count)
    return(c(col, list(kind = "categorical", frequencies = .fobj(r$freqs), rare_levels_suppressed = r$n_rare, derived_month = TRUE)))
  }
  if (is_serial_date_col(present, name)) {                       # Excel day numbers under a date header
    dts <- excel_serial_to_dt(present)
    dts <- dts[!is.na(dts)]
    if (length(dts)) {
      yfreq <- count_first(sprintf("%d", as.integer(dt_parts(dts)$y)))
      return(c(col, list(kind = "datetime", format = "%Y-%m-%d", serial = TRUE, min = dt_format(min(dts)), max = dt_format(max(dts)),
                         year_freqs = .fobj(yfreq))))
    }
  }
  # whole-number identifiers are not measurements (an id-like NAME is required here)
  if (all_int && is_id_name && nd == length(present)) {
    L <- nchar(present, type = "chars")
    return(c(col, list(kind = "identifier_unique", len_min = min(L), len_max = max(L))))
  }
  if (all_int && is_id_name && nd > MIN_GROUP_KEYS) return(c(col, .group_profile(present)))
  if (all_int && is_id_name) {
    cnt <- count_first(present)
    return(c(col, list(kind = "identifier_group", n_groups = length(cnt), group_size_grid = arr(c(0, 100)),
                       group_size_quantiles = arr(c(min(cnt), max(cnt))))))
  }
  if (all_numeric && !is_bool) {
    nums <- sort(f)
    qv <- vapply(QUANTILE_GRID, function(p) percentile(nums, p), numeric(1))
    if (USE_ROBUST_BOUNDS) {
      qv[1] <- percentile(nums, P_LOW)
      qv[length(qv)] <- percentile(nums, P_HIGH)
    }
    m <- sum(nums) / length(nums)
    sdv <- if (length(nums) < 2) 0 else sqrt(sum((nums - m)^2) / length(nums))     # population SD
    col <- c(col, list(kind = if (all_int) "integer" else "float", is_integer = all_int, mean = pyround(m, 6), std = pyround(sdv, 6),
                       quantile_grid = arr(QUANTILE_GRID), quantile_values = arr(pyround(qv, 6)), has_negative = any(nums < 0),
                       has_zero = any(nums == 0), zero_rate = pyround(sum(nums == 0) / length(nums), 6)))
    if (!all_int) {                                              # how many decimals the values are written with
      dec <- rep(0, length(present))
      has <- grepl(".", present, fixed = TRUE) & !grepl("e", tolower(present), fixed = TRUE)
      if (any(has)) dec[has] <- pmin(6, nchar(sub("0+$", "", sub("^[^.]*\\.([^.]*).*$", "\\1", present[has]))))
      dec <- sort(dec)
      col$decimals <- dec[floor(0.95 * (length(dec) - 1)) + 1]
    }
  } else if (is_date) {
    parsed <- parse_date(present)
    fm <- count_first(parsed$fmt[!is.na(parsed$fmt)])
    fmt <- names(fm)[which.max(fm)]
    dts <- parsed$dt[!is.na(parsed$dt)]
    yfreq <- count_first(sprintf("%d", as.integer(dt_parts(dts)$y)))
    col <- c(col, list(kind = "datetime", format = fmt, min = dt_format(min(dts), fmt), max = dt_format(max(dts), fmt),
                       year_freqs = .fobj(yfreq)))
  } else if (is_bool) {
    col <- c(col, list(kind = "boolean", frequencies = .fobj(count_first(present))))
  } else if (nd <= MAX_CATEGORIES) {
    r <- .rare_freqs(present, min_cell_count)
    col <- c(col, list(kind = "categorical", frequencies = .fobj(r$freqs), rare_levels_suppressed = r$n_rare))
  } else {                                                       # many different values: an id or free text
    sigs <- count_first(signature(present))
    dom <- names(sigs)[which.max(sigs)]
    L <- nchar(present, type = "chars")
    col <- c(col, list(kind = "id_or_text", dominant_signature = dom, signature_coverage = pyround(max(sigs) / length(present), 4),
                       len_min = min(L), len_max = max(L), is_unique_key = nd == length(present)))
  }
  col
}

# The table's own unique column is a key (counts only) when its values are whole numbers under a
# header that says 'number / id / code', or a dense run of numbers, or codes that are neither
# numbers nor dates.  A unique measurement (an amount) is not a key.
is_running_key <- function(name, col) {
  vals <- nonmissing(col)
  if (length(vals) < MIN_TABLE_ROWS || anyDuplicated(vals)) return(FALSE)
  smp <- stride_sample(vals, 400)
  if (all(looks_int(smp) & !grepl(".", smp, fixed = TRUE))) {
    words <- strsplit(tolower(pystrip(name)), paste0("(?:", .WS, "|[_./()-])+"), perl = TRUE)[[1]]
    nums <- to_float(vals)
    return(any(words %in% KEY_HEADER_WORDS) || (max(nums, na.rm = TRUE) - min(nums, na.rm = TRUE) + 1) <= 3 * length(vals))
  }
  if (any(!is.na(to_float(smp)))) return(FALSE)
  s60 <- utils::head(smp, 60)
  sum(!is.na(parse_date_dt(s60))) < 0.5 * length(s60)
}

# Shape only: how many, how many missing, and the commonest pattern of letters / digits.
# Used for hidden columns and for tables too small to summarise.  No value is kept.
shape_profile <- function(name, values, why) {
  present <- values[!is_missing(values)]
  n <- length(values)
  col <- list(name = name, n = n, missing_rate = if (n) pyround((n - length(present)) / n, 6) else 0, n_unique = 0, kind = "empty",
              shape_only = why)
  if (!length(present)) return(col)
  sigs <- count_first(signature(present))
  L <- nchar(present, type = "chars")
  col$kind <- "id_or_text"
  c(col, list(dominant_signature = names(sigs)[which.max(sigs)], len_min = min(L), len_max = max(L), is_unique_key = FALSE))
}

# whole numbers, nearly all different, all of one length of 7+ digits: an id / phone number, not a measure
looks_like_id_number <- function(present) {
  if (length(present) < MIN_TABLE_ROWS) return(FALSE)
  smp <- stride_sample(present, 600)
  if (!all(looks_int(smp) & !grepl(".", smp, fixed = TRUE))) return(FALSE)
  lens <- count_first(nchar(sub("^-+", "", smp), type = "chars"))
  L <- as.numeric(names(lens)[which.max(lens)])
  L >= 7 && max(lens) >= 0.9 * length(smp) && length(unique(present)) >= 0.9 * length(present)
}

# Labels seen fewer than k times lose their name and their own count: they become one pooled
# count, kept only if it reaches k.  NULL when no label is left.
pool_rare <- function(col, k) {
  fr <- unlist(col$frequencies)
  keep <- fr[fr >= k]
  if (!length(keep)) return(NULL)
  pooled <- sum(fr[fr < k])
  col$rare_levels_suppressed <- length(fr) - length(keep)
  if (pooled >= k) {
    if (RARE_LABEL %in% names(keep)) keep[[RARE_LABEL]] <- keep[[RARE_LABEL]] + pooled
    else keep <- c(keep, stats::setNames(pooled, RARE_LABEL))
  }
  col$frequencies <- keep
  col
}

# profile_column plus the protections a general tool needs: hidden columns and id-number
# look-alikes never have their values summarised; no count below k is written (years, months,
# yes / no); the ends of a numeric range are the 5th / 95th percentile when there are fewer than
# 500 values; quantiles keep 4 significant figures; dates keep their month-by-month shape.
profile_any <- function(name, values, k = MIN_CELL_COUNT, force_key = FALSE, hidden = FALSE, excel_date = FALSE) {
  present <- values[!is_missing(values)]
  if (hidden && !force_key) return(shape_profile(name, values, "hidden: the header names personal or free-text content"))
  auto_id <- !force_key && !excel_date && looks_like_id_number(present)
  col <- profile_column(name, values, min_cell_count = k, force_key = force_key || auto_id)
  if (auto_id) col$auto_identifier <- TRUE
  col$all_numeric <- length(present) > 0 && !anyNA(to_float(stride_sample(present, 400)))
  kind <- col$kind
  if (kind == "boolean") {
    fr <- unlist(col$frequencies)
    keep <- fr[fr >= k]
    if (!length(keep)) return(shape_profile(name, values, sprintf("every value is rarer than %d", k)))
    col$rare_levels_suppressed <- length(fr) - length(keep)
    col$frequencies <- keep
  } else if (kind == "categorical") {
    col <- pool_rare(col, k)
    if (is.null(col)) return(shape_profile(name, values, sprintf("every value is rarer than %d", k)))
  } else if (kind %in% c("integer", "float")) {
    qv <- unlist(col$quantile_values)
    nq <- length(qv)
    if (length(present) < 500) {
      qv[1] <- qv[2]
      qv[nq] <- qv[nq - 1]
    }
    if (col$n_unique > 50) {
      qv <- vapply(qv, sig4, numeric(1))
      col$mean <- sig4(col$mean)
      col$std <- sig4(col$std)
    }
    col$quantile_values <- arr(qv)
    if (kind == "integer" && col$n_unique <= 30) {
      cnt <- count_first(int_str(trunc(to_float(present)) + 0))
      keep <- cnt[cnt >= k]
      if (length(keep)) col$value_freqs <- keep
    }
  } else if (kind == "datetime") {
    serial <- isTRUE(col$serial)
    dts <- to_dt(present, serial)
    dts <- sort(dts[!is.na(dts)])
    if (!length(dts)) return(shape_profile(name, values, "dates could not be read"))
    fmt <- col$format %||% "%Y-%m-%d"
    col$excel_date <- isTRUE(excel_date)
    col$has_time <- any(utils::head(dts, 2000) %% 86400 != 0)
    yf <- unlist(col$year_freqs)
    col$year_freqs <- .fobj(yf[yf >= k])
    nd <- length(dts)
    if (nd >= 100) {
      lo <- dts[floor(0.01 * (nd - 1)) + 1]
      hi <- dts[floor(0.99 * (nd - 1)) + 1]
    } else {
      lo <- dts[1]
      hi <- dts[nd]
    }
    pl <- dt_parts(lo); ph <- dt_parts(hi)
    lo <- make_dt(pl$y, pl$m, 1)                                 # month resolution: no exact day leaves
    hi <- make_dt(ph$y, ph$m, days_in_month(ph$y, ph$m))
    col$min <- dt_format(lo, fmt)
    col$max <- dt_format(hi, fmt)
    col$first_month <- dt_format(lo, "%Y-%m")
    col$last_month <- dt_format(hi, "%Y-%m")
    ym <- table(dt_format(dts, "%Y-%m"))
    ym <- stats::setNames(as.numeric(ym), names(ym))
    ym <- ym[order(names(ym), method = "radix")]
    keep <- ym[ym >= k]
    other <- sum(ym[ym < k])
    if (other >= k) keep <- c(keep, "_other" = other)            # a remainder below k is itself a small cell
    col$ym_freqs <- .fobj(keep)
    wd <- tabulate(dt_weekday(dts) + 1, 7)                       # rows per day of the week (Monday first)
    if (min(wd) >= k) col$weekday_freqs <- arr(wd)
  }
  col
}

.col_names <- function(columns) vapply(columns, function(co) co$name, "")

# A column that is empty exactly when another column has certain values (no closing date while the
# status is 'open'): sets empty_when = {column, values}.  Only labels the profile already shows are used.
find_empty_rules <- function(cols, n, columns) {
  is_cat <- vapply(columns, function(co) co$kind %in% c("categorical", "boolean") && is.null(co[["helper"]]) && is.null(co[["shape_only"]]) &&
                     length(co$frequencies) >= 2 && length(co$frequencies) <= JOINT_MAX_PARENT_LEVELS, logical(1))
  idx <- stride_idx(n, 20000)
  for (i in seq_along(columns)) {
    co <- columns[[i]]
    mr <- co$missing_rate %||% 0
    if (!is.null(co[["helper"]]) || mr < 0.02 || mr > 0.98) next
    cmiss <- is_missing(cols[[co$name]][idx])
    for (j in which(is_cat)) {
      if (j == i) next
      x <- columns[[j]]
      xv <- cols[[x$name]][idx]
      ok <- !is_missing(xv)
      tot <- count_first(xv[ok])
      levels <- names(x$frequencies)
      levels <- levels[levels %in% names(tot) & levels != RARE_LABEL]
      if (!length(levels)) next
      mis <- count_first(xv[ok & cmiss])
      mcount <- mis[match(levels, names(mis))]
      mcount[is.na(mcount)] <- 0
      share <- as.numeric(mcount) / as.numeric(tot[match(levels, names(tot))])
      empty <- levels[share >= 0.98]
      if (length(empty) && length(empty) < length(levels) && all(share <= 0.02 | share >= 0.98)) {
        columns[[i]]$empty_when <- list(column = x$name, values = arr(empty))
        break
      }
    }
  }
  columns
}

# A whole-number column that is simply the years between two dates of the same row (an age at the
# event): list(column = c(earlier date, later date)).  It is then recomputed, not drawn.
find_derived_years <- function(header, cols, n, columns) {
  nm <- .col_names(columns)
  kind <- vapply(columns, function(co) co$kind, "")
  dcols <- header[header %in% nm[kind == "datetime"]]
  serial <- function(a) isTRUE(columns[[match(a, nm)]]$serial)
  idx <- stride_idx(n, 3000)
  out <- list()
  for (name in header) {
    i <- match(name, nm)
    if (is.na(i)) next
    co <- columns[[i]]
    qv <- unlist(co$quantile_values)
    if (co$kind != "integer" || !is.null(co[["helper"]]) || !length(qv) || qv[1] < 0 || qv[length(qv)] > 130) next
    v <- to_float(cols[[name]][idx])
    for (a in dcols) for (b in dcols) {
      if (a == b || !is.null(out[[name]])) next
      va <- cols[[a]][idx]; vb <- cols[[b]][idx]
      ok <- !is.na(v) & !is_missing(va) & !is_missing(vb)
      da <- to_dt(va[ok], serial(a)); db <- to_dt(vb[ok], serial(b))
      ok2 <- !is.na(da) & !is.na(db)
      nn <- sum(ok2)
      hit <- sum(abs(floor((db[ok2] - da[ok2]) / 86400) / 365.25 - 0.5 - v[ok][ok2]) <= 1.0)
      if (nn >= MIN_TABLE_ROWS && hit >= 0.97 * nn) out[[name]] <- c(a, b)
    }
  }
  out
}

# How finely a date column can be cut into periods and stay under JOINT_MAX_PARENT_LEVELS:
# quarters, else years, else 5-year bands, else decades.
period_label <- function(dt, scheme) {
  p <- dt_parts(dt)
  y <- as.integer(p$y)
  if (scheme == "q") return(sprintf("%04d-Q%d", y, as.integer((p$m - 1) %/% 3 + 1)))
  if (scheme == "y") return(sprintf("Y%04d", y))
  if (scheme == "y5") return(sprintf("Y%04d-%02d", y - y %% 5L, (y - y %% 5L + 4L) %% 100L))
  sprintf("Y%04ds", y - y %% 10L)
}

period_scheme <- function(dts) {
  if (!length(dts)) return("q")
  for (scheme in c("q", "y", "y5", "y10")) if (length(unique(period_label(dts, scheme))) <= JOINT_MAX_PARENT_LEVELS) return(scheme)
  "y10"
}

# the (year, month) pairs a period label covers, as a two-column matrix; 0 rows when it is not a period
period_months <- function(label) {
  m <- regmatches(label, regexec("^(\\d{4})-Q([1-4])$", label, perl = TRUE))[[1]]
  if (length(m)) {
    y <- as.numeric(m[2]); q <- as.numeric(m[3])
    return(cbind(y, c(3 * q - 2, 3 * q - 1, 3 * q)))
  }
  m <- regmatches(label, regexec("^Y(\\d{4})(?:(-\\d{2})|(s))?$", label, perl = TRUE))[[1]]
  if (length(m)) {
    y <- as.numeric(m[2])
    years <- y:(y + (if (nzchar(m[3])) 5 else if (nzchar(m[4])) 10 else 1) - 1)
    return(cbind(rep(years, each = 12), rep(1:12, length(years))))
  }
  matrix(numeric(0), 0, 2)
}

# list of c(earlier date column, later date column) for a later date that normally follows an
# earlier one (opened -> closed).  The days between them are then profiled like any other number.
gap_pairs <- function(header, cols, n, columns, k = MIN_CELL_COUNT) {
  nm <- .col_names(columns)
  kind <- vapply(columns, function(co) co$kind, "")
  dcols <- header[header %in% nm[kind == "datetime"]]
  serial <- function(a) isTRUE(columns[[match(a, nm)]]$serial)
  idx <- stride_idx(n, 20000)
  out <- list()
  for (j in seq_along(dcols)[-1]) {
    b <- dcols[j]; best <- NULL
    for (a in dcols[seq_len(j - 1)]) {
      va <- cols[[a]][idx]; vb <- cols[[b]][idx]
      ok <- !is_missing(va) & !is_missing(vb)
      da <- to_dt(va[ok], serial(a)); db <- to_dt(vb[ok], serial(b))
      ok2 <- !is.na(da) & !is.na(db)
      da <- da[ok2]; db <- db[ok2]
      nn <- length(da)
      if (nn < max(MIN_TABLE_ROWS, 4 * k) || sum(db >= da) < 0.9 * nn) next
      gaps <- sort(floor((db - da) / 86400)); own <- sort(dt_days(db))
      i3 <- (3 * nn) %/% 4 + 1; i1 <- nn %/% 4 + 1
      if (gaps[i3] - gaps[i1] > 0.9 * (own[i3] - own[i1])) next     # too loose to pin the later date (birth -> event)
      if (is.null(best) || nn > best$n) best <- list(a = a, n = nn)
    }
    if (!is.null(best)) out[[length(out) + 1]] <- c(best$a, b)
  }
  out
}

# c(mutual information, bias-corrected Cramer's V) of two vectors of labels; "" is skipped
assoc <- function(xs, ys) {
  ok <- xs != "" & ys != ""
  n <- sum(ok)
  if (n < 50) return(c(0, 0))
  fx <- match(xs[ok], unique(xs[ok])); fy <- match(ys[ok], unique(ys[ok]))
  r <- max(fx); cc <- max(fy)
  if (r < 2 || cc < 2) return(c(0, 0))
  cx <- tabulate(fx, r); cy <- tabulate(fy, cc)
  cxy <- tabulate((fx - 1) * cc + fy, r * cc)
  w <- which(cxy > 0)
  v <- cxy[w]; ix <- (w - 1) %/% cc + 1; iy <- (w - 1) %% cc + 1
  mi <- sum(v / n * log(v * n / (cx[ix] * cy[iy])))
  mi <- max(0, mi - (r - 1) * (cc - 1) / (2 * n))               # minus what chance alone gives to long lists of labels
  phi2 <- sum(v * v / (cx[ix] * cy[iy])) - 1
  if (stats::pchisq(n * phi2, (r - 1) * (cc - 1), lower.tail = FALSE) >= 0.01) return(c(0, 0))   # could be chance: not kept
  phi2c <- max(0, phi2 - (r - 1) * (cc - 1) / (n - 1))
  rc <- r - (r - 1)^2 / (n - 1); c2 <- cc - (cc - 1)^2 / (n - 1)
  den <- min(rc - 1, c2 - 1)
  c(mi, if (den > 0) sqrt(phi2c / den) else 0)
}

# How strongly x and y still move together once z is accounted for: the x-by-y counts are compared
# with what they would be if y followed z alone.  A corrected Cramer's V, 0 when it could be chance.
resid_assoc <- function(zs, xs, ys) {
  ok <- zs != "" & xs != "" & ys != ""
  n <- sum(ok)
  if (n < 50) return(0)
  fz <- match(zs[ok], unique(zs[ok])); fx <- match(xs[ok], unique(xs[ok])); fy <- match(ys[ok], unique(ys[ok]))
  nz <- max(fz); r <- max(fx); cc <- max(fy)
  df <- (r - 1) * (cc - 1)
  if (df <= 0) return(0)
  nxz <- matrix(tabulate((fz - 1) * r + fx, r * nz), r, nz)
  nyz <- matrix(tabulate((fz - 1) * cc + fy, cc * nz), cc, nz)
  ex <- nxz %*% (t(nyz) / tabulate(fz, nz))
  nxy <- matrix(tabulate((fy - 1) * r + fx, r * cc), r, cc)
  pos <- ex > 0
  x2 <- sum((nxy[pos] - ex[pos])^2 / ex[pos])
  if (stats::pchisq(x2, df, lower.tail = FALSE) >= 0.01) return(0)
  sqrt(max(0, x2 / n - df / (n - 1)) / min(r - 1, cc - 1))
}

.desc_rank <- function(x) -match(x, c_sort(unique(x)))

# Which pairs of columns to keep a joint table for.  The categorical columns (including the period
# of each date and the columns brought over from the parent table) are linked into the tree that
# keeps the most shared information (Chow-Liu); every other column (a number, a duration, a long
# list of labels) hangs from the categorical column it depends on most.  A column of the parent
# table is only ever the 'given' side, and its strength is measured on ONE row per parent.
# A column may then get a SECOND 'given' column: the one that tells most about it beyond the
# first, among the columns drawn before it.  Returns a list of list(a = given, b = column, second).
choose_joint_pairs <- function(cols, n, columns, key_cols, cluster = NULL) {
  idx <- stride_idx(n, JOINT_SAMPLE_ROWS)
  pick <- NULL
  if (!is.null(cluster) && length(cluster) && nzchar(cluster)) {
    kv <- cols[[cluster]][idx]
    one <- which(kv != "" & !duplicated(kv))
    if (length(one) >= 50) pick <- one
  }
  lab <- list(); parents <- character(0); leaves <- character(0); fixed <- character(0)
  for (co in columns) {
    name <- co$name; kind <- co$kind
    if (name %in% key_cols || !is.null(co[["shape_only"]]) || isTRUE(co[["auto_identifier"]])) next
    helper <- co[["helper"]] %||% ""
    if (kind %in% c("categorical", "boolean") || (kind == "integer" && length(co[["value_freqs"]]))) {
      v <- cols[[name]][idx]
      v[is_missing(v)] <- ""
      lab[[name]] <- v
      levels <- length(setdiff(unique(v), ""))
      if (levels >= 2 && levels <= JOINT_MAX_PARENT_LEVELS) {
        parents <- c(parents, name)
        if (helper == "parent") fixed <- c(fixed, name)
      } else if (levels <= JOINT_MAX_CHILD_LEVELS && helper != "parent") leaves <- c(leaves, name)
    } else if (kind %in% c("integer", "float")) {
      vals <- cols[[name]][idx]
      edges <- numeric_bins(vals)
      if (length(edges) && helper != "parent") {
        v <- bin_of(vals, edges)
        v[is_missing(vals)] <- ""
        lab[[name]] <- v
        leaves <- c(leaves, name)
      }
    }
  }
  parents <- utils::head(parents, 40)
  nlev <- vapply(parents, function(p) length(setdiff(unique(lab[[p]]), "")), numeric(1))
  names(nlev) <- parents
  if (!is.null(pick)) lab <- lapply(lab, function(v) v[pick])     # one row per parent for every comparison
  memo <- new.env()
  assoc_ <- function(a, b) {
    key <- paste(a, b, sep = "\r")
    r <- memo[[key]]
    if (is.null(r)) {
      r <- assoc(lab[[a]], lab[[b]])
      memo[[key]] <- r
      memo[[paste(b, a, sep = "\r")]] <- r
    }
    r
  }
  em <- numeric(0); ea <- character(0); eb <- character(0)
  np <- length(parents)
  if (np >= 2) for (i in 1:(np - 1)) for (j in (i + 1):np) {
    a <- parents[i]; b <- parents[j]
    if (a %in% fixed && b %in% fixed) next
    r <- assoc_(a, b)
    if (r[2] >= JOINT_MIN_V && r[1] > 0) { em <- c(em, r[1]); ea <- c(ea, a); eb <- c(eb, b) }
  }
  comp <- stats::setNames(parents, parents)
  fx <- parents[parents %in% fixed]
  for (p in fx[-1]) comp[[p]] <- fx[1]          # the parent's columns already hang together
  root <- function(x) {
    while (comp[[x]] != x) x <- comp[[x]]
    x
  }
  adj <- stats::setNames(lapply(parents, function(p) character(0)), parents)
  for (e in order(-em, .desc_rank(ea), .desc_rank(eb))) {        # Kruskal: the heaviest tree
    ra <- root(ea[e]); rb <- root(eb[e])
    if (ra != rb) {
      comp[[ra]] <- rb
      adj[[ea[e]]] <- c(adj[[ea[e]]], eb[e])
      adj[[eb[e]]] <- c(adj[[eb[e]]], ea[e])
    }
  }
  pa <- character(0); pb <- character(0); seen <- fixed
  free <- parents[!(parents %in% fixed)]
  for (start in c(fx, free[order(nlev[free])])) {
    if (start %in% seen && !(start %in% fixed)) next
    seen <- c(seen, start)
    queue <- start
    while (length(queue)) {
      a <- queue[1]; queue <- queue[-1]
      for (b in adj[[a]]) if (!(b %in% seen)) {
        seen <- c(seen, b)
        pa <- c(pa, a); pb <- c(pb, b)
        queue <- c(queue, b)
      }
    }
  }
  for (leaf in leaves) {
    best <- NULL
    for (p in parents) {
      r <- assoc_(p, leaf)
      if (r[2] >= JOINT_MIN_V && r[1] > 0 && (is.null(best) || r[2] > best$v)) best <- list(v = r[2], p = p)
    }
    if (!is.null(best)) { pa <- c(pa, best$p); pb <- c(pb, leaf) }
  }
  pa <- utils::head(pa, JOINT_MAX_TABLES); pb <- utils::head(pb, JOINT_MAX_TABLES)
  # The second 'given' column.  A candidate must already be drawn when the column is drawn, must be
  # related to the column on its own with at least 5 rows per cell on average, and must still be
  # related to it once the first is accounted for.
  n_smp <- if (length(lab)) length(lab[[1]]) else 0
  se <- numeric(0); sp <- character(0); sb <- character(0)
  for (i in seq_along(pb)) {
    a <- pa[i]; b <- pb[i]
    lev_b <- length(setdiff(unique(lab[[b]]), ""))
    cm <- numeric(0); cp <- character(0)
    for (p in parents) {
      pp <- match(p, pb)
      if (p == a || p == b || (!is.na(pp) && pp >= i) || 5 * nlev[[p]] * lev_b > n_smp) next
      r <- assoc_(p, b)
      if (r[2] >= JOINT_MIN_V && r[1] > 0) { cm <- c(cm, r[1]); cp <- c(cp, p) }
    }
    best <- NULL
    for (j in utils::head(order(-cm, .desc_rank(cp)), 8)) {
      extra <- resid_assoc(lab[[a]], lab[[cp[j]]], lab[[b]])
      if (extra >= JOINT_MIN_V_SECOND && (is.null(best) || extra > best$e)) best <- list(e = extra, p = cp[j])
    }
    if (!is.null(best)) { se <- c(se, best$e); sp <- c(sp, best$p); sb <- c(sb, b) }
  }
  out <- lapply(seq_along(pb), function(i) list(a = pa[i], b = pb[i], second = FALSE))
  for (j in utils::head(order(-se, .desc_rank(sp), .desc_rank(sb)), JOINT_MAX_SECOND))
    out[[length(out) + 1]] <- list(a = sp[j], b = sb[j], second = TRUE)
  out
}

# Two-way counts for the chosen pairs.  A cell below k is dropped and only its volume is kept
# ('_suppressed'); a parent value seen fewer than k times is dropped whole.
build_joint_tables <- function(cols, n, columns, key_cols, k = MIN_CELL_COUNT, cluster = NULL) {
  nm <- .col_names(columns)
  lab_of <- function(co, v) {
    if (length(co[["value_freqs"]])) {
      f <- to_float(v)
      ok <- !is.na(f)
      v[ok] <- int_str(trunc(f[ok]) + 0)
    }
    v
  }
  out <- list(); first_kept <- character(0)
  for (pr in choose_joint_pairs(cols, n, columns, key_cols, cluster)) {
    a <- pr$a; b <- pr$b
    if (pr$second && !(b %in% first_kept)) next
    ca <- columns[[match(a, nm)]]; cb <- columns[[match(b, nm)]]
    edges <- NULL
    if (cb$kind %in% c("integer", "float") && !length(cb[["value_freqs"]])) {
      edges <- numeric_bins(cols[[b]])
      if (!length(edges)) next
    }
    av <- cols[[a]]; bv <- cols[[b]]
    ok <- !is_missing(av) & !is_missing(bv)
    if (!any(ok)) next
    av <- lab_of(ca, av[ok])
    bv <- if (!is.null(edges)) bin_of(bv[ok], edges) else lab_of(cb, bv[ok])
    ua <- unique(av); ub <- unique(bv)
    na_ <- length(ua); nb <- length(ub)
    key <- (match(bv, ub) - 1) * na_ + match(av, ua)
    M <- matrix(tabulate(key, na_ * nb), na_, nb)
    first <- matrix(match(seq_len(na_ * nb), key), na_, nb)       # the row where each cell is first met
    row_tot <- rowSums(M); col_tot <- colSums(M)
    rank_a <- match(ua, c_sort(ua))
    V <- M
    V[V < k] <- 0
    alive <- row_tot >= k
    V[!alive, ] <- 0
    changed <- TRUE
    while (changed) {                    # a hidden remainder of 1..k-1 in a row or a column could be found by
      changed <- FALSE                   # subtraction from the totals, so one more cell of that line is hidden
      for (i in which(alive)) {
        vis <- which(V[i, ] > 0)
        if (!length(vis)) next
        rem <- row_tot[i] - sum(V[i, vis])
        if (rem > 0 && rem < k) {
          V[i, vis[order(V[i, vis], first[i, vis])[1]]] <- 0
          changed <- TRUE
        }
      }
      for (j in seq_len(nb)) {
        vis <- which(V[, j] > 0)
        if (!length(vis)) next
        rem <- col_tot[j] - sum(V[vis, j])
        if (rem > 0 && rem < k) {
          V[vis[order(V[vis, j], rank_a[vis])[1]], j] <- 0
          changed <- TRUE
        }
      }
    }
    cells <- list()
    for (i in which(alive)) {
      vis <- which(V[i, ] > 0)
      if (!length(vis)) next
      vis <- vis[order(first[i, vis])]
      keep <- stats::setNames(V[i, vis], ub[vis])
      s <- row_tot[i] - sum(keep)
      if (s) keep <- c(keep, "_suppressed" = s)
      cells[[ua[i]]] <- keep
    }
    if (length(cells)) {
      rec <- list(parent = a, child = b, cells = cells)
      if (!is.null(edges)) rec$child_bins <- arr(pyround(edges, 6))
      if (pr$second) rec$second <- TRUE else first_kept <- c(first_kept, b)
      out[[length(out) + 1]] <- rec
    }
  }
  out
}


# ==============================================================================
#  Step 1 -- profile a folder (on the premises)
# ==============================================================================
say <- function(...) cat(sprintf(...), "\n", sep = "")

fail <- function(...) stop(structure(class = c("gov_one_exit", "error", "condition"), list(message = sprintf(...), call = NULL)))

settings_template <- function() list(
  "_how" = paste0("Optional. Put this file next to the data files and run 'Rscript gov_one.R profile' again. ",
                  "A column is named as it appears in the header, or as 'table::column' to mean one table only."),
  min_cell = MIN_CELL_COUNT, hide = list(), show = list(), public_tables = list(), links = list(), no_links = list(),
  skip_files = list(),
  "_examples" = list(hide = arr(EX_HIDE), show = arr(EX_SHOW), public_tables = arr("districts"),
                     links = arr("payments::client = clients::client_id"), no_links = arr("cases::room"),
                     skip_files = arr("old_copy.xlsx"),
                     "_meaning" = paste0("hide: values never leave (shape only). show: cancel an automatic hide. public_tables: a ",
                                         "reference table with no personal data, copied as it is. links: a relationship the tool ",
                                         "missed. no_links: a relationship it found by mistake. skip_files: files not to read.")))

load_settings <- function(folder) {
  st <- list(min_cell = MIN_CELL_COUNT, hide = character(0), show = character(0), public_tables = character(0),
             links = character(0), no_links = character(0), skip_files = character(0))
  p <- file.path(folder, SETTINGS_FILE)
  if (file.exists(p)) {
    user <- tryCatch(read_json(p), error = function(e) fail("ERROR: %s could not be read (%s). Fix or delete it.", SETTINGS_FILE, conditionMessage(e)))
    for (a in names(st)) {
      if (!is.null(user[[a]])) st[[a]] <- if (a == "min_cell") as.numeric(user[[a]]) else as.character(unlist(user[[a]]))
    }
    say("  settings : %s", SETTINGS_FILE)
  }
  st
}

.spec <- function(specs, base, col) any(pystrip(specs) %in% c(col, paste0(base, "::", col)))

parse_links <- function(specs) {
  out <- list()
  for (s in specs) {
    m <- regmatches(s, regexec("^\\s*(.+?)::(.+?)\\s*=\\s*(.+?)::(.+?)\\s*$", s, perl = TRUE))[[1]]
    if (length(m)) out[[length(out) + 1]] <- pystrip(m[-1])
    else say("  [!] link not understood (use \"child::column = parent::column\"): %s", s)
  }
  out
}

# every table of every readable file in the folder -> list(tables, bad = unreadable files)
read_folder <- function(folder, skip = character(0)) {
  fi <- find_inputs(folder)
  bad <- fi$bad
  tables <- list(); seen <- character(0)
  for (p in fi$good) {
    name <- basename(p)
    if (name %in% skip) next
    say("  reading  : %s", name)
    tabs <- tryCatch(read_file_tables(p), error = function(e) e)
    if (inherits(tabs, "error")) {
      bad[[length(bad) + 1]] <- c(name, sprintf("could not be read (password-protected or damaged? %s)", conditionMessage(tabs)))
      next
    }
    if (!length(tabs)) bad[[length(bad) + 1]] <- c(name, "no table found (empty)")
    for (tb in tabs) {
      if (tb$base %in% seen) tb$base <- sprintf("%s (%s)", tb$base, tb$file_type)       # a.xlsx and a.csv in one folder
      seen <- c(seen, tb$base)
      tables[[length(tables) + 1]] <- tb
    }
  }
  list(tables = tables, bad = bad)
}

# The profile of one table.  entry: its line in the schema; built: list(base = list(tb, rep)) of the
# tables already profiled (parents come first), so a column of the parent table that this table
# depends on can be kept as a joint count.  Returns list(rep, tb): tb now also holds the helper columns.
build_report <- function(tb, key_cols, k, st, entry = NULL, built = NULL, kids = NULL) {
  header <- tb$header; base <- tb$base; n <- tb$n
  rp <- list(tool = VERSION, source_file = tb$file, file_type = tb$file_type, sheet = tb$sheet, sheet_index = tb$sheet_index %||% 0,
             base = base, generated_at = format(Sys.time(), "%Y-%m-%dT%H:%M:%S"), n_rows = n, n_cols = length(header),
             header = arr(header), header_row = tb$header_row + 1, load_notes = arr(tb$notes),
             config = list(min_cell = k, min_table_rows = MIN_TABLE_ROWS), crosstabs = list(), date_gaps = list())
  if (base %in% st$public_tables) {
    rp$public_table <- TRUE
    m <- do.call(cbind, unname(tb$cols[header]))
    rp$verbatim_rows <- lapply(seq_len(n), function(i) arr(m[i, ]))
    rp$columns <- lapply(header, function(h) list(name = h, kind = "verbatim", n = n))
    return(list(rep = rp, tb = tb))
  }
  small <- n < MIN_TABLE_ROWS
  rp$structure_only <- small
  columns <- list()
  for (name in header) {
    vals <- tb$cols[[name]]
    if (small) co <- shape_profile(name, vals, sprintf("the table has fewer than %d rows", MIN_TABLE_ROWS))
    else {
      hidden <- (is_sensitive_name(name) || .spec(st$hide, base, name)) && !.spec(st$show, base, name)
      co <- profile_any(name, vals, k = k, force_key = name %in% key_cols, hidden = hidden, excel_date = name %in% tb$excel_dates)
    }
    columns[[length(columns) + 1]] <- co
  }
  # Helper columns: never written to a synthetic file, they only carry structure.
  #   <date>__period   the quarter / year of a date, so other columns can depend on time
  #   <date>__gap      whole days from an earlier date of the same row (a duration, a waiting time)
  #   <parent>::<col>  a column of the parent table, so this table can depend on it
  helpers <- list()
  add <- function(h) helpers[[length(helpers) + 1]] <<- h
  if (!small) {
    nm <- .col_names(columns)
    serial <- function(a) isTRUE(columns[[match(a, nm)]]$serial)
    dy <- find_derived_years(header, tb$cols, n, columns)
    for (name in names(dy)) columns[[match(name, nm)]]$derived_years <- arr(dy[[name]])
    columns <- find_empty_rules(tb$cols, n, columns)
    for (kid in kids) {                                           # has::<child>: does this row have rows in the child table?
      pv <- tb$cols[[kid$pkey]]
      have <- ifelse(pv != "" & pv %in% kid$keys, "yes", "no")
      n_yes <- sum(have == "yes")
      if (n_yes < k || n - n_yes < max(k, 0.02 * n)) next          # (nearly) every row has some: nothing to keep
      hn <- paste0("has::", kid$base)
      tb$cols[[hn]] <- have
      h <- profile_any(hn, have, k = k)
      if (h$kind %in% c("categorical", "boolean")) {
        h$helper <- "has"; h$child <- kid$base; h$missing_rate <- 0
        add(h)
      }
    }
    for (co in columns) {
      if (co$kind != "datetime") next
      name <- co$name
      v <- tb$cols[[name]]
      dts <- to_dt(v, isTRUE(co$serial))
      dts[is_missing(v)] <- NA
      ok <- !is.na(dts)
      scheme <- period_scheme(dts[ok])
      hn <- paste0(name, "__period")
      pv <- rep("", n)
      if (any(ok)) pv[ok] <- period_label(dts[ok], scheme)
      tb$cols[[hn]] <- pv
      h <- profile_column(hn, pv, min_cell_count = k)
      if (h$kind == "categorical") {
        h <- pool_rare(h, k)
        if (!is.null(h)) {
          fr <- h$frequencies
          h$frequencies <- .fobj(fr[names(fr) != RARE_LABEL])       # a period is drawn from the counted periods only
          h$helper <- "period"; h$of <- name; h$scheme <- scheme; h$missing_rate <- 0
          add(h)
        }
      }
    }
    for (pr in gap_pairs(header, tb$cols, n, columns, k)) {
      a <- pr[1]; b <- pr[2]
      hn <- paste0(b, "__gap")
      va <- tb$cols[[a]]; vb <- tb$cols[[b]]
      da <- to_dt(va, serial(a)); db <- to_dt(vb, serial(b))
      ok <- !is_missing(va) & !is_missing(vb) & !is.na(da) & !is.na(db)
      ok[ok] <- db[ok] >= da[ok]
      gap <- rep("", n)
      gap[ok] <- int_str(floor((db[ok] - da[ok]) / 86400))
      tb$cols[[hn]] <- gap
      h <- profile_any(hn, gap, k = k)
      if (h$kind == "integer") {
        h$helper <- "gap"; h$of <- b; h$from <- a; h$missing_rate <- 0
        add(h)
      }
    }
    par <- entry$parent
    if (!is.null(par) && !is.null(built[[par]])) {
      ptab <- built[[par]]$tb; prep <- built[[par]]$rep
      pidx <- match(tb$cols[[entry$link]], ptab$cols[[entry$parent_key]])
      pcol <- function(name) {
        v <- ptab$cols[[name]][pidx]
        if (is.null(v)) v <- rep("", n)
        v[is.na(pidx)] <- ""
        v
      }
      taken <- 0
      for (pc in prep$columns) {
        if (taken >= 16) break
        if (!is.null(pc[["shape_only"]]) || isTRUE(pc[["auto_identifier"]]) || (pc[["helper"]] %||% "") %in% c("gap", "pgap", "parent", "has")) next
        if (!(pc$kind %in% c("categorical", "boolean") || length(pc[["value_freqs"]]))) next
        shown <- names(if (length(pc[["value_freqs"]])) pc[["value_freqs"]] else pc[["frequencies"]])   # only labels the parent's own profile shows
        if (pc$name %in% header || length(shown) < 2 || length(shown) > JOINT_MAX_PARENT_LEVELS) next
        hn <- sprintf("%s::%s", par, pc$name)
        v <- pcol(pc$name)
        if (length(pc[["value_freqs"]])) {
          f <- to_float(v)
          num <- v != "" & !is.na(f)
          v[num] <- int_str(trunc(f[num]) + 0)
        }
        out <- v != "" & !(v %in% shown)                          # rare in the parent table: no name here either
        v[out] <- if (RARE_LABEL %in% shown) RARE_LABEL else ""
        tb$cols[[hn]] <- v
        h <- profile_any(hn, v, k = k)
        if (h$kind %in% c("categorical", "boolean") || length(h[["value_freqs"]])) {
          h$helper <- "parent"; h$parent_col <- pc$name
          add(h)
          taken <- taken + 1
        }
      }
      banded <- 0
      for (pc in prep$columns) {                                  # a number of the parent row (age, income): five bands
        if (banded >= 4 || !(pc$kind %in% c("integer", "float")) || !is.null(pc[["helper"]]) || length(pc[["value_freqs"]]) ||
            !is.null(pc[["shape_only"]]) || pc$name %in% header) next
        edges <- sort(unique(vapply(numeric_bins(ptab$cols[[pc$name]], nbins = 5), sig4, numeric(1))))
        if (length(edges) < 2) next
        hn <- sprintf("%s::%s", par, pc$name)
        f <- to_float(pcol(pc$name))
        v <- rep("", n)
        v[!is.na(f)] <- paste0("B", findInterval(f[!is.na(f)], edges, left.open = TRUE))
        tb$cols[[hn]] <- v
        h <- profile_any(hn, v, k = k)
        if (h$kind == "categorical") {
          h$helper <- "parent"; h$parent_col <- pc$name; h$parent_bins <- arr(edges)
          add(h)
          banded <- banded + 1
        }
      }
      # <date>__pgap: whole days from a date of the parent row to this date (a payment after the
      # case was opened, a contact before the death), so the two tables stay in step in time
      pdates <- Filter(function(pc) pc$kind == "datetime" && is.null(pc[["helper"]]), prep$columns)
      for (co in columns) {
        if (co$kind != "datetime" || !length(pdates)) next
        name <- co$name; best <- NULL
        vc <- tb$cols[[name]]
        dc <- to_dt(vc, isTRUE(co$serial))
        dc[is_missing(vc)] <- NA
        for (pc in pdates) {
          vp <- pcol(pc$name)
          dp <- to_dt(vp, isTRUE(pc$serial))
          dp[is_missing(vp)] <- NA
          ok <- !is.na(dc) & !is.na(dp)
          nn <- sum(ok)
          d <- dt_days(dc[ok]) - dt_days(dp[ok])
          if (nn < max(MIN_TABLE_ROWS, 4 * k) || max(sum(d >= 0), sum(d <= 0)) < 0.9 * nn) next      # no steady order between the two dates
          s <- sort(d); own <- sort(dt_days(dc[ok]))
          i3 <- (3 * nn) %/% 4 + 1; i1 <- nn %/% 4 + 1
          spread <- s[i3] - s[i1]
          if (spread > 0.9 * (own[i3] - own[i1])) next              # too loose to pin this date (e.g. a date of birth)
          if (is.null(best) || spread < best$spread) {
            gaps <- rep("", n)
            gaps[ok] <- int_str(d)
            best <- list(spread = spread, pcol = pc$name, gaps = gaps)
          }
        }
        if (!is.null(best)) {
          hn <- paste0(name, "__pgap")
          tb$cols[[hn]] <- best$gaps
          h <- profile_any(hn, best$gaps, k = k)
          if (h$kind == "integer") {
            h$helper <- "pgap"; h$of <- name; h$parent_col <- best$pcol; h$missing_rate <- 0
            add(h)
          }
        }
      }
    }
  }
  rp$columns <- c(columns, helpers)
  if (!small) rp$crosstabs <- build_joint_tables(tb$cols, n, rp$columns, key_cols, k, entry$link)
  list(rep = rp, tb = tb)
}

# ---- the plain-language list of what leaves (profile_summary.md) ----
fmt <- function(x, digits = 2) {
  if (is.null(x) || (is.numeric(x) && (is.na(x) || is.nan(x)))) return("\u2014")
  if (is.character(x)) return(x)
  if (x == floor(x) && abs(x) < 1e15) return(formatC(x, format = "f", digits = 0, big.mark = ","))
  if (abs(x) >= 0.01) return(formatC(x, format = "f", digits = digits, big.mark = ","))
  sub("e([+-])0*(\\d)", "e\\1\\2", sprintf("%.2g", x))
}

fmt_p <- function(p) {
  if (is.null(p) || is.na(p)) return("\u2014")
  if (p < 0.001) "< .001" else sub("0.", ".", sprintf("= %.3f", p), fixed = TRUE)
}

.num_text <- function(x) if (x == floor(x) && abs(x) < 1000) sprintf("%.1f", x) else if (x == floor(x)) sprintf("%.0f", x) else sprintf("%.15g", x)

.nice <- function(name, columns, nm) {
  i <- match(name, nm)
  if (is.na(i)) return(name)
  co <- columns[[i]]
  helper <- co[["helper"]] %||% ""
  if (helper == "period") return(sprintf("%s (period)", co$of))
  if (helper == "gap") return(sprintf("days from %s to %s", co$from, co$of))
  if (helper == "parent") {
    pcol <- co$parent_col
    if (endsWith(pcol, "__period")) return(sprintf("%s (period, from the linked table)", substr(pcol, 1, nchar(pcol) - 8)))
    return(sprintf("%s (from the linked table%s)", pcol, if (length(co[["parent_bins"]])) ", in 5 bands" else ""))
  }
  if (helper == "pgap") return(sprintf("days from %s (linked table) to %s", co$parent_col, co$of))
  if (helper == "has") return(sprintf("has rows in %s", co$child))
  name
}

# a guess, for the reviewer only: at least 8 labels, nearly all of them two or three short words of letters
.look_like_people <- function(labels) {
  labs <- labels[labels != RARE_LABEL]
  if (length(labs) < 8) return(FALSE)
  hit <- sum(vapply(labs, function(a) {
    w <- .split_ws(a)
    length(w) >= 2 && length(w) <= 3 && all(grepl("^\\p{L}+$", gsub("['-]", "", w), perl = TRUE) & nchar(w) <= 12)
  }, logical(1)))
  hit >= 0.85 * length(labs)
}

.what_leaves <- function(co) {
  kind <- co$kind
  if (kind == "verbatim") return(c("EVERYTHING", "declared a public table in the settings"))
  if (!is.null(co[["helper"]])) return(NULL)
  if (!is.null(co[["shape_only"]])) return(c("shape only", co[["shape_only"]]))
  if (kind %in% c("categorical", "boolean")) {
    labs <- names(co$frequencies)
    labs <- labs[!startsWith(labs, "RARE_")]
    show <- paste0(paste(utils::head(labs, 8), collapse = ", "), if (length(labs) > 8) sprintf(" ... +%d more", length(labs) - 8) else "")
    flag <- if (length(labs) > 50) "  **<- CHECK: many text values**" else ""
    if (.look_like_people(labs)) flag <- "  **<- CHECK: THESE LOOK LIKE NAMES OF PEOPLE. If they are, add this header to \"hide\".**"
    return(c("labels + counts", sprintf("%d labels (%d rare ones hidden): %s%s", length(labs), as.integer(co[["rare_levels_suppressed"]] %||% 0), show, flag)))
  }
  if (kind %in% c("integer", "float")) {
    how <- if (length(co[["value_freqs"]])) "values + counts" else "9 rounded quantiles"
    qv <- unlist(co$quantile_values)
    lo <- sig4(qv[1]); hi <- sig4(qv[length(qv)])
    if (isTRUE(co$is_integer) && 1800 <= lo && lo <= hi && hi <= 2100) return(c(how, sprintf("about %.0f to %.0f", lo, hi)))   # years
    return(c(how, sprintf("about %s to %s, mean %s", fmt(lo), fmt(hi), fmt(sig4(co$mean)))))
  }
  if (kind == "datetime") return(c("counts per month", sprintf("%s to %s (month and year only, never the day)", co$first_month %||% "None", co$last_month %||% "None")))
  if (kind %in% c("identifier_unique", "identifier_group"))
    return(c("nothing but counts", sprintf("key / id column%s", if (isTRUE(co[["auto_identifier"]])) " (looked like an id number)" else "")))
  if (kind == "id_or_text")
    return(c("shape only", sprintf("free text or codes: pattern %s, length %d-%d", substr(co$dominant_signature, 1, 24), as.integer(co$len_min), as.integer(co$len_max))))
  c("nothing", "empty column")
}

link_lines <- function(schema) {
  out <- character(0)
  for (b in unlist(schema$order)) {
    e <- schema$files[[b]]
    for (L in e$links) {
      q <- unlist(L$fanout_quantiles)
      mid <- if (length(q)) sprintf(" -- typically %s row(s) per %s", .num_text(sig4(q[length(q) %/% 2 + 1])), L$parent_key) else ""
      out <- c(out, sprintf("%s . %s  ->  %s . %s  (%.0f%% of its values found there%s)%s", b, L$link, L$parent, L$parent_key,
                            100 * (L$share_found %||% 1), if (isFALSE(L$primary)) "; second link" else "", mid))
    }
    if (length(e$inherited)) out <- c(out, sprintf("%s repeats these columns of %s: %s", b, e$parent, paste(unlist(e$inherited), collapse = ", ")))
  }
  out
}

write_text <- function(path, lines, eol = "\n") {
  con <- file(path, "wb")
  on.exit(close(con))
  writeBin(charToRaw(enc2utf8(paste(lines, collapse = eol))), con)
}

write_summary <- function(out_dir, reports, schema, bad, k, warnings) {
  md <- c("# What leaves the premises", "",
          sprintf("- Made %s by %s", format(Sys.time(), "%Y-%m-%dT%H:%M"), VERSION),
          sprintf("- %d table(s) from %d file(s)", length(reports), length(unique(vapply(reports, function(r) r$source_file, "")))),
          sprintf("- **Aggregates only.** No row is copied. No value or combination seen fewer than **%d** times is written.", k),
          "- Ids, names, phone numbers, addresses and free text are described by their shape only (length, letters / digits).",
          paste0("- Dates keep the month and the year, never the day. Numeric ranges are rounded and cut at the 5th / 95th ",
                 "percentile when there are fewer than 500 values."),
          "", "## Check before sending", "",
          paste0("1. In the tables below, every column marked **labels + counts** sends its values as written. ",
                 "They must be categories (district, status, type) -- never names of people."),
          sprintf("2. To hide a column, add its header to \"hide\" in %s and run the profile again.", SETTINGS_FILE),
          sprintf("3. Send the whole `%s` folder. Nothing else is needed.", SEND_DIR), "")
  if (length(warnings) || length(bad))
    md <- c(md, "## Warnings", "", sprintf("- [!] %s", warnings), vapply(bad, function(b) sprintf("- [!] not read: %s -- %s", b[1], b[2]), ""), "")
  rel <- link_lines(schema)
  md <- c(md, "## How the files relate", "", if (length(rel)) sprintf("- %s", rel) else "- no relationship was found between the tables", "")
  for (rp in reports) {
    title <- paste0(rp$source_file, if (nzchar(rp$sheet)) sprintf(" / sheet %s", rp$sheet) else "")
    md <- c(md, sprintf("## %s", rp$base), "",
            sprintf("%s -- %d rows, %d columns%s", title, as.integer(rp$n_rows), as.integer(rp$n_cols),
                    if (isTRUE(rp$structure_only)) " -- **too small: shape only**" else ""), "")
    for (note in unlist(rp$load_notes)) md <- c(md, sprintf("- note: %s", note))
    md <- c(md, "", "| Column | What leaves | Detail | Missing |", "|---|---|---|---|")
    nm <- .col_names(rp$columns)
    for (co in rp$columns) {
      w <- .what_leaves(co)
      if (is.null(w)) next
      md <- c(md, sprintf("| %s | %s | %s | %s |", gsub("|", "/", co$name, fixed = TRUE), w[1], gsub("|", "/", w[2], fixed = TRUE),
                          if (!is.null(co$missing_rate)) sprintf("%.0f%%", 100 * co$missing_rate) else ""))
    }
    gaps <- character(0)
    for (co in rp$columns) if (identical(co[["helper"]], "gap")) gaps <- c(gaps, sprintf("%s -> %s", co$from, co$of))
    for (co in rp$columns) if (identical(co[["helper"]], "pgap")) gaps <- c(gaps, sprintf("%s (linked table) -> %s", co$parent_col, co$of))
    if (length(gaps)) md <- c(md, "", sprintf("Durations kept (quantiles of whole days): %s", paste(gaps, collapse = "; ")))
    for (co in rp$columns) {
      if (identical(co[["helper"]], "has")) {
        yes <- co$frequencies[["yes"]] %||% 0
        if (!("yes" %in% names(co$frequencies))) yes <- 0
        md <- c(md, "", sprintf("Kept: how many rows have rows in %s (%.0f%% do)", co$child, 100 * yes / max(1, co$n)))
      }
      if (!is.null(co[["empty_when"]]))
        md <- c(md, "", sprintf("%s is empty when %s is: %s", co$name, co$empty_when$column, paste(unlist(co$empty_when$values), collapse = ", ")))
      if (!is.null(co[["derived_years"]])) {
        dy <- unlist(co$derived_years)
        md <- c(md, "", sprintf("%s is the number of years from %s to %s: it is recomputed in the synthetic file", co$name, dy[1], dy[2]))
      }
    }
    if (length(rp$crosstabs))
      md <- c(md, "", sprintf("Joint counts kept, so columns move together (cells below %d dropped): %s", k,
                              paste(vapply(rp$crosstabs, function(x) sprintf("%s x %s%s", .nice(x$parent, rp$columns, nm), .nice(x$child, rp$columns, nm),
                                                                             if (isTRUE(x[["second"]])) " (second link)" else ""), ""), collapse = "; ")))
    md <- c(md, "")
  }
  write_text(file.path(out_dir, "profile_summary.md"), md)
}

# Step 1.  Returns list(out_dir, reports).
run_profile <- function(data_dir, out_dir = NULL, k = NULL) {
  st <- load_settings(data_dir)
  k <- as.integer(k %||% st$min_cell %||% MIN_CELL_COUNT)
  if (is.na(k) || k < 3) fail("ERROR: the minimum cell count cannot be below 3.")
  out_dir <- out_dir %||% file.path(data_dir, SEND_DIR)
  rf <- read_folder(data_dir, skip = st$skip_files)
  tables <- rf$tables; bad <- rf$bad
  if (!length(tables)) {
    for (b in bad) say("  [!] %s: %s", b[1], b[2])
    fail("ERROR: no readable Excel (.xlsx / .xlsm) or CSV file in %s", data_dir)
  }
  names(tables) <- vapply(tables, function(tb) tb$base, "")
  say("  links    : comparing the columns of %d table(s)", length(tables))
  schema <- detect_links(tables, forced = parse_links(st$links), forbidden = gsub("::", ".", st$no_links, fixed = TRUE))
  keycols <- key_columns_per_base(schema)
  for (tb in tables) {                                            # a table's own running number is a key too
    own <- schema$files[[tb$base]]$key
    if (!is.null(own) && is_running_key(own, tb$cols[[own]])) keycols[[tb$base]] <- union(keycols[[tb$base]], own)
  }
  dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
  old <- list.files(out_dir, pattern = "^profile_.*\\.json$", full.names = TRUE)
  if (length(old)) file.remove(old)
  reports <- list(); warnings <- character(0); fps <- character(0); built <- list()
  for (base in unlist(schema$order)) {                            # parents before their children
    if (is.null(tables[[base]])) next
    tb <- tables[[base]]
    say("  profiling: %s (%d rows, %d columns)", base, tb$n, length(tb$header))
    kids <- list()
    for (cb in names(schema$files)) {
      e <- schema$files[[cb]]
      if (identical(e$parent, base) && !is.null(tables[[cb]]))
        kids[[length(kids) + 1]] <- list(base = cb, pkey = e$parent_key, keys = setdiff(unique(tables[[cb]]$cols[[e$link]]), ""))
    }
    br <- build_report(tb, keycols[[base]] %||% character(0), k, st, schema$files[[base]], built, kids)
    rp <- br$rep
    built[[base]] <- br
    fp <- to_json(lapply(rp$columns, function(co) co[intersect(c("name", "n", "kind", "n_unique", "frequencies", "quantile_values"), names(co))]), indent = 0)
    if (rp$n_rows >= MIN_TABLE_ROWS && fp %in% names(fps))
      warnings <- c(warnings, sprintf("%s has exactly the same profile as %s -- one may be a copy of the other", rp$base, fps[[fp]]))
    if (!(fp %in% names(fps))) fps[fp] <- rp$base
    if (isTRUE(rp$public_table)) warnings <- c(warnings, sprintf("%s is copied in full (declared public in the settings): %d rows", rp$base, rp$n_rows))
    write_json(rp, file.path(out_dir, sprintf("profile_%s.json", rp$base)))
    reports[[base]] <- rp
  }
  reports <- unname(reports[names(tables)[names(tables) %in% names(reports)]])      # back to file / sheet order
  schema$min_cell <- k
  write_json(schema, file.path(out_dir, "schema.json"))
  write_summary(out_dir, reports, schema, bad, k, warnings)
  cat("\n")
  for (r in link_lines(schema)) say("  link     : %s", r)
  for (w in warnings) say("  [!] %s", w)
  for (b in bad) say("  [!] not read: %s -- %s", b[1], b[2])
  say("\nDONE. %d table(s) profiled. Folder to send: %s", length(reports), out_dir)
  say("Read profile_summary.md first: it lists exactly what leaves. Aggregates only; no cell below %d.", k)
  invisible(list(out_dir = out_dir, reports = reports))
}


# ==============================================================================
#  Step 2 -- synthetic files from the JSON files (at home)
#  While a table is built every column is text ("" = empty), except dates, which
#  are numbers of seconds (NA = empty).  Numbers become real number cells on writing.
# ==============================================================================
ALPHA <- LETTERS
STUDY_DATE_RANGE <- c("2015-01-01", format(Sys.Date(), "%Y-%m-%d"))      # for date-shaped text in shape-only columns

num_text <- function(x) {
  x <- x + 0
  out <- ifelse(x == floor(x) & abs(x) < 1e15, sprintf("%.0f", x), sprintf("%.15g", x))
  out[is.na(x)] <- ""
  out
}

.STRICT_NUM <- "^[+-]?([0-9]+\\.?[0-9]*|\\.[0-9]+)$"

inv_cdf <- function(grid01, qv, u) {
  ng <- length(grid01)
  j <- pmin(findInterval(u, grid01[-1], left.open = TRUE), ng - 1) + 1
  j2 <- pmin(j + 1, ng)
  lo_g <- grid01[j]; hi_g <- grid01[j2]
  frac <- ifelse(hi_g == lo_g, 0, (u - lo_g) / (hi_g - lo_g))
  qv[j] + frac * (qv[pmin(j + 1, length(qv))] - qv[j])
}

sample_numeric <- function(co, n) {
  grid <- unlist(co$quantile_grid) / 100
  qv <- unlist(co$quantile_values)
  out <- inv_cdf(grid, qv, stats::runif(n))
  if (isTRUE(co$has_zero) && n >= 3) out[3] <- 0
  zr <- co$zero_rate %||% 0
  if (zr > 0 && isTRUE(co$is_integer) && !isTRUE(co$has_negative)) {
    # the point mass at 0 is reproduced exactly; the rest is drawn from the part of the
    # quantile function above the zero mass
    z <- stats::runif(n) < zr
    out[z] <- 0
    re <- !z & out <= 0
    out[re] <- inv_cdf(grid, qv, zr + (1 - zr) * stats::runif(sum(re)))
  }
  if (isTRUE(co$is_integer)) out <- round(out)
  out
}

.draw <- function(labels, weights, n) {
  if (!n || !length(labels)) return(rep("", n))
  labels[sample.int(length(labels), n, TRUE, weights)]
}

sample_categorical <- function(co, n) {
  fr <- unlist(co$frequencies)
  .draw(names(fr), fr, n)
}

# a text column known only by its pattern, when the pattern is a date / time stamp
sample_timestamp_like <- function(co, n) {
  sig <- co$dominant_signature %||% ""
  lo <- parse_date_dt(STUDY_DATE_RANGE[1]); hi <- parse_date_dt(STUDY_DATE_RANGE[2])
  d <- lo + floor(stats::runif(n) * (hi - lo + 1))
  s <- dt_format(d, "%Y-%m-%d")
  if (grepl(" ", sig, fixed = TRUE)) {
    s <- paste0(s, dt_format(d, " %H:%M:%S"))
    frac <- if (grepl(".", sig, fixed = TRUE)) nchar(strsplit(sig, ".", fixed = TRUE)[[1]][2]) else 0
    if (frac && n) s <- paste0(s, ".", apply(matrix(sample(0:9, n * frac, TRUE), n, frac), 1, paste, collapse = ""))
  }
  s
}

sample_id <- function(co, n) {
  sig <- co$dominant_signature
  if (!is.null(sig) && grepl("^DDDD-DD-DD( DD:DD:DD(\\.D+)?)?$", sig)) return(sample_timestamp_like(co, n))
  if (is.null(sig) || !nzchar(sig)) sig <- strrep("D", max(co$len_min %||% 6, 1))
  ch <- strsplit(sig, "", fixed = TRUE)[[1]]
  make <- function(m) {
    out <- rep("", m)
    for (cc in ch) out <- paste0(out, if (cc == "D") sample(0:9, m, TRUE) else if (cc == "A") sample(ALPHA, m, TRUE) else cc)
    out
  }
  out <- make(n)
  if (isTRUE(co$is_unique_key)) for (try in 1:50) {
    dup <- duplicated(out)
    if (!any(dup)) break
    out[dup] <- make(sum(dup))
  }
  out
}

# running numbers in random order, with as many digits as the real key has (the real values never left)
sample_identifier_unique <- function(co, n) {
  if (!n) return(character(0))
  width <- co$len_min %||% 1
  if (!width) width <- 1
  start <- if (isTRUE(co$all_numeric) && identical(as.numeric(width), as.numeric(co$len_max)) && width > 1 && width <= 15 &&
               n <= 9 * 10^(width - 1)) 10^(width - 1) else 1
  sprintf("%.0f", (start + seq_len(n) - 1)[sample.int(n)])
}

.fix_total <- function(scaled, n, floor_at) {
  ng <- length(scaled)
  diff <- n - sum(scaled)
  if (diff > 0) scaled <- scaled + tabulate(sample.int(ng, diff, TRUE), ng)
  while (diff < 0) {
    elig <- which(scaled > floor_at)
    if (!length(elig)) break
    take <- min(-diff, length(elig))
    pick <- elig[sample.int(length(elig), take)]
    scaled[pick] <- scaled[pick] - 1
    diff <- diff + take
  }
  scaled
}

sample_identifier_group <- function(co, n) {
  if (!n) return(character(0))
  ng <- min(n, max(1, floor(co$n_groups)))
  grid <- unlist(co$group_size_grid) / 100
  qv <- unlist(co$group_size_quantiles)
  sizes <- pmax(1, round(inv_cdf(grid, qv, stats::runif(ng))))
  scaled <- .fix_total(pmax(1, round(sizes * n / sum(sizes))), n, 1)
  out <- rep(seq_len(ng), scaled)
  sprintf("%.0f", out[sample.int(length(out))][seq_len(n)])
}

# Assign n child rows to parent rows; the number of children per parent follows the real
# group-size distribution, and a parent may receive none.
attach_indices <- function(n, n_parents, grid, quantiles) {
  if (n_parents <= 0) return(rep(1L, n))
  sizes <- if (length(quantiles)) pmax(0, round(inv_cdf(grid / 100, quantiles, stats::runif(n_parents)))) else rep(1, n_parents)
  if (sum(sizes) <= 0) sizes <- rep(1, n_parents)
  scaled <- .fix_total(round(sizes * n / sum(sizes)), n, 0)
  idx <- rep(seq_len(n_parents), scaled)
  if (length(idx) < n) idx <- c(idx, sample.int(n_parents, n - length(idx), TRUE))
  if (n) idx[sample.int(length(idx))][seq_len(n)] else integer(0)
}

load_profiles <- function(folder) {
  out <- list()
  for (f in c_sort(list.files(folder, pattern = "^profile_.*\\.json$"))) {
    r <- read_json(file.path(folder, f))
    if (is.list(r) && !is.null(r$columns) && !is.null(r$base)) out[[length(out) + 1]] <- r
  }
  out
}

.rand_time <- function(d) d + sample(7 * 3600 + 0:(12 * 3600), length(d), TRUE)

parse_fmt <- function(s, fmt) {
  if (is.null(s) || is.null(fmt)) return(NA_real_)
  tok <- regmatches(fmt, gregexpr("%[YmdHMS]", fmt))[[1]]
  num <- suppressWarnings(as.numeric(regmatches(s, gregexpr("[0-9]+", s))[[1]]))
  if (!length(tok) || length(num) < length(tok)) return(NA_real_)
  g <- function(k, d) {
    i <- match(k, tok)
    if (is.na(i)) d else num[i]
  }
  make_dt(g("%Y", 1970), g("%m", 1), g("%d", 1), g("%H", 0), g("%M", 0), g("%S", 0))
}

# Dates for one column.  The period (quarter / year) was already drawn, possibly depending on
# another column; the month inside it follows the month-by-month counts; the day follows the
# spread over the week.  "" in mask = empty.
synth_dates <- function(co, period_vals, mask, n) {
  fmt <- co$format %||% "%Y-%m-%d"
  lo <- parse_fmt(co$min, fmt); hi <- parse_fmt(co$max, fmt)
  if (is.na(lo) || is.na(hi)) {
    lo <- make_dt(2015, 1, 1)
    hi <- parse_date_dt(format(Sys.Date(), "%Y-%m-%d"))
  }
  span <- max(floor(hi - lo), 1)
  ymf <- unlist(co$ym_freqs); wdf <- unlist(co$weekday_freqs)
  out <- rep(NA_real_, n)
  act <- which(mask != "")
  if (!length(act)) return(out)
  labs <- if (!is.null(period_vals)) period_vals[act] else rep("", length(act))
  for (lab in unique(labs)) {
    rows <- act[labs == lab]
    months <- period_months(lab)
    if (nrow(months)) {
      w <- if (length(ymf)) ymf[sprintf("%04d-%02d", as.integer(months[, 1]), as.integer(months[, 2]))] else rep(0, nrow(months))
      w[is.na(w)] <- 0
      if (!any(w > 0)) w <- rep(1, nrow(months))
      mi <- sample.int(nrow(months), length(rows), TRUE, w)
      for (m in unique(mi)) {
        rr <- rows[mi == m]
        y <- months[m, 1]; mo <- months[m, 2]
        nd <- days_in_month(y, mo)
        day <- if (length(wdf) == 7) sample.int(nd, length(rr), TRUE, wdf[dt_weekday(make_dt(y, mo, 1:nd)) + 1])
               else sample.int(nd, length(rr), TRUE)
        out[rr] <- make_dt(y, mo, day)
      }
    } else out[rows] <- dt_days(lo + floor(stats::runif(length(rows)) * (span + 1))) * 86400
  }
  if (isTRUE(co$has_time)) out[act] <- .rand_time(out[act])
  out
}

# ---- columns that move together: redraw each column given the column(s) it follows ----
# list(given label = list(labs, wts, psup = share hidden)) for one joint table
.choosers <- function(xt) {
  lapply(xt$cells, function(d) {
    d <- unlist(d)
    keep <- names(d) != "_suppressed"
    sup <- sum(d[!keep])
    tot <- sum(d)
    list(labs = names(d)[keep], wts = as.numeric(d[keep]), psup = if (tot) sup / tot else 1)
  })
}

# One joint table as list(given label = c(child label = share)).  The hidden part of a line is
# spread over the child labels that line does not show, following the child's own distribution.
.shares <- function(xt, ccol) {
  bins <- xt[["child_bins"]]
  base <- if (!is.null(bins)) stats::setNames(rep(1, length(bins) + 1), as.character(0:length(bins)))
          else unlist(if (length(ccol[["value_freqs"]])) ccol[["value_freqs"]] else ccol[["frequencies"]])
  out <- list()
  for (av in names(xt$cells)) {
    d <- unlist(xt$cells[[av]])
    line <- d[names(d) != "_suppressed"]
    sup <- sum(d[names(d) == "_suppressed"])
    if (sup && length(base)) {
      rest <- base[!(names(base) %in% names(line))]
      if (!length(rest)) rest <- base
      tot <- sum(rest)
      if (!tot) tot <- 1
      add <- sup * rest / tot
      inl <- names(add) %in% names(line)
      if (any(inl)) line[names(add)[inl]] <- line[names(add)[inl]] + add[inl]
      line <- c(line, add[!inl])
    }
    s <- sum(line)
    if (s > 0) out[[av]] <- line / s
  }
  out
}

# A column that follows two others.  For every (first, second) pair of values found in the rows:
# the weights of the child's labels that agree with BOTH joint tables and add nothing more
# (iterative proportional fitting).  Returns list(keys, labs, W) or NULL.
.fit_two <- function(t1, t2, A, B) {
  ok <- A %in% names(t1) & B %in% names(t2)
  if (!any(ok)) return(NULL)
  key <- paste(A[ok], B[ok], sep = "\r")
  uk <- unique(key)
  cnt <- tabulate(match(key, uk), length(uk))
  ka <- A[ok][match(uk, key)]; kb <- B[ok][match(uk, key)]
  labs <- c_sort(unique(c(unlist(lapply(t1, names)), unlist(lapply(t2, names)))))
  L <- length(labs)
  if (!L || length(uk) * L > 400000) return(NULL)
  ua <- unique(ka); ub <- unique(kb)
  ia <- match(ka, ua); ib <- match(kb, ub)
  ma <- as.numeric(rowsum(cnt, ia)); mb <- as.numeric(rowsum(cnt, ib))
  mat <- function(tt, u, m) matrix(unlist(lapply(seq_along(u), function(i) {
    v <- tt[[u[i]]]
    x <- v[match(labs, names(v))]
    x[is.na(x)] <- 0
    as.numeric(x) * m[i]
  })), nrow = length(u), byrow = TRUE)
  R1 <- mat(t1, ua, ma); R2 <- mat(t2, ub, mb)
  W <- cnt * (R1[ia, , drop = FALSE] > 0 & R2[ib, , drop = FALSE] > 0)
  for (it in 1:25) {
    S <- rowsum(W, ia); f <- R1 / S; f[S <= 0] <- 1               # agree with each joint table in turn
    W <- W * f[ia, , drop = FALSE]
    S <- rowsum(W, ib); f <- R2 / S; f[S <= 0] <- 1
    W <- W * f[ib, , drop = FALSE]
    s <- rowSums(W)                                               # and with how often the pair itself occurs
    W <- W * ifelse(s > 0, cnt / s, 1)
  }
  keep <- rowSums(W) > 0
  if (!any(keep)) return(NULL)
  list(keys = uk[keep], labs = labs, W = W[keep, , drop = FALSE])
}

# a number inside quantile bin j (0-based) of a column
.bin_value <- function(co, bins, j) {
  qv <- unlist(co$quantile_values)
  nb <- length(bins)
  lo <- ifelse(j == 0, qv[1], bins[pmax(j, 1)])
  hi <- ifelse(j >= nb, qv[length(qv)], bins[pmin(j + 1, nb)])
  v <- stats::runif(length(j), pmin(lo, hi), pmax(lo, hi))
  if (isTRUE(co$is_integer)) round(v) else v
}

.marginal <- function(co, m) {
  kind <- co$kind
  if (kind %in% c("categorical", "boolean")) return(sample_categorical(co, m))
  if (kind %in% c("integer", "float")) return(num_text(sample_numeric(co, m)))
  if (kind == "id_or_text") return(sample_id(co, m))
  rep("", m)
}

apply_crosstabs <- function(cols, rp, n) {
  columns <- rp$columns
  nm <- .col_names(columns)
  xts <- rp$crosstabs
  second <- list()
  for (xt in xts) if (isTRUE(xt[["second"]]) && is.null(second[[xt$child]]) && !is.null(cols[[xt$parent]])) second[[xt$child]] <- xt
  done <- character(0)                                            # one first 'given' column per child
  for (xt in xts) {
    if (isTRUE(xt[["second"]])) next
    a <- xt$parent; b <- xt$child
    if (is.null(cols[[a]]) || is.null(cols[[b]]) || a == b || b %in% done) next
    done <- c(done, b)
    ci <- match(b, nm)
    out <- cols[[b]]; pa <- cols[[a]]
    if (is.na(ci) || !is.character(out) || !is.character(pa)) next
    ccol <- columns[[ci]]
    has_bins <- !is.null(xt[["child_bins"]])
    bins <- as.numeric(unlist(xt[["child_bins"]]))
    lab <- rep(NA_character_, n)                                   # the label (or bin) drawn for each row
    marg <- rep(FALSE, n)                                          # rows that take a draw from the column's own distribution
    todo <- out != ""                                              # an empty cell stays empty
    use <- function(chs, given, rows) {
      g <- match(given[rows], names(chs))
      for (grp in split(rows, g)) {
        ch <- chs[[g[match(grp[1], rows)]]]
        hide <- if (length(ch$labs)) stats::runif(length(grp)) < ch$psup else rep(TRUE, length(grp))
        marg[grp[hide]] <<- TRUE
        if (any(!hide)) lab[grp[!hide]] <<- .draw(ch$labs, ch$wts, sum(!hide))
      }
    }
    x2 <- second[[b]]
    if (!is.null(x2) && (x2$parent %in% c(a, b) || !identical(as.numeric(unlist(x2[["child_bins"]])), bins) ||
                         !is.character(cols[[x2$parent]]))) x2 <- NULL
    ch1 <- .choosers(xt)
    if (!is.null(x2)) {
      p2 <- cols[[x2$parent]]
      both <- .fit_two(.shares(xt, ccol), .shares(x2, ccol), pa, p2)
      if (!is.null(both)) {                                        # follows both columns
        g <- match(paste(pa, p2, sep = "\r"), both$keys)
        hit <- which(todo & !is.na(g))
        for (grp in split(hit, g[hit])) lab[grp] <- .draw(both$labs, both$W[g[grp[1]], ], length(grp))
        todo[hit] <- FALSE
      }
    }
    r1 <- which(todo & pa %in% names(ch1))
    if (length(r1)) use(ch1, pa, r1)
    if (!is.null(x2)) {                                            # the first value is unknown here: follow the second
      ch2 <- .choosers(x2)
      r2 <- which(todo & !(pa %in% names(ch1)) & p2 %in% names(ch2))
      if (length(r2)) use(ch2, p2, r2)
    }
    got <- which(!is.na(lab))
    if (length(got)) out[got] <- if (has_bins) num_text(.bin_value(ccol, bins, as.numeric(lab[got]))) else lab[got]
    mg <- which(marg)
    if (length(mg)) out[mg] <- .marginal(ccol, length(mg))
    cols[[b]] <- out
  }
  cols
}

# One synthetic table: list(header, cols, n, dates = date columns, num = columns written as numbers)
synth_table <- function(rp, entry, synth, scale) {
  columns <- rp$columns
  nm <- .col_names(columns)
  header <- if (length(rp$header)) as.character(unlist(rp$header)) else nm[vapply(columns, function(co) is.null(co[["helper"]]), TRUE)]
  if (!is.null(rp[["verbatim_rows"]])) {
    rows <- rp$verbatim_rows
    cols <- lapply(seq_along(header), function(i) vapply(rows, function(r) if (i <= length(r) && !is.null(r[[i]])) as.character(r[[i]]) else "", ""))
    names(cols) <- header
    return(list(header = header, n = length(rows), cols = cols, dates = list(), num = header))
  }
  n <- round(rp$n_rows * scale)
  n <- if (rp$n_rows) max(1, n) else 0
  parent <- entry$parent; link <- entry$link
  parent_key <- entry$parent_key %||% link
  inherited <- as.character(unlist(entry$inherited))
  attach <- NULL
  if (!is.null(parent) && !is.null(synth[[parent]]) && synth[[parent]]$n) {
    has <- synth[[parent]]$cols[[paste0("has::", rp$base)]]       # only some parent rows have rows here
    elig <- if (!is.null(has)) which(has == "yes") else integer(0)
    fg <- as.numeric(unlist(entry$fanout_grid)); fq <- as.numeric(unlist(entry$fanout_quantiles))
    attach <- if (length(elig)) elig[attach_indices(n, length(elig), fg, fq)] else attach_indices(n, synth[[parent]]$n, fg, fq)
  }
  cols <- list(); copied <- character(0)
  for (co in columns) {
    name <- co$name; kind <- co$kind
    if (!is.null(attach) && (identical(name, link) || name %in% inherited || identical(co[["helper"]], "parent"))) {
      src <- synth[[parent]]$cols[[if (identical(name, link)) parent_key else (co[["parent_col"]] %||% name)]]
      if (!is.null(src)) {                                         # the value of the parent row this row hangs from
        v <- src[attach]
        if (length(co[["parent_bins"]])) {                         # a number of the parent row, as its band
          f <- to_float(v)
          v <- ifelse(is.na(f), "", paste0("B", findInterval(f, as.numeric(unlist(co$parent_bins)), left.open = TRUE)))
        }
        cols[[name]] <- v
        if (kind == "datetime") copied <- c(copied, name)
        next
      }
    }
    vals <- if (kind == "datetime") rep("x", n)                    # filled in below, after the months are drawn
      else if (kind == "integer" && length(co[["value_freqs"]])) { vf <- unlist(co$value_freqs); .draw(names(vf), vf, n) }
      else if (kind == "integer") num_text(sample_numeric(co, n))
      else if (kind == "float") num_text(round(sample_numeric(co, n), 4))
      else if (kind %in% c("categorical", "boolean")) sample_categorical(co, n)
      else if (kind == "id_or_text") sample_id(co, n)
      else if (kind == "identifier_unique") sample_identifier_unique(co, n)
      else if (kind == "identifier_group") sample_identifier_group(co, n)
      else rep("", n)
    ruled <- !is.null(co[["empty_when"]]) && co$empty_when$column %in% nm     # emptied by rule below, not at random
    mr <- if (ruled) 0 else (co$missing_rate %||% 0)
    if (mr > 0 && n) vals[stats::runif(n) < mr] <- ""
    cols[[name]] <- vals
  }
  for (L in entry$links) {                                        # every other link: draw from the parent's keys
    if (isTRUE(L$primary) && !is.null(attach)) next
    if (!is.null(synth[[L$parent]]) && L$link %in% nm) {
      pool <- synth[[L$parent]]$cols[[L$parent_key]]
      pool <- pool[!is.na(pool) & pool != ""]
      if (length(pool)) {
        v <- pool[sample.int(length(pool), n, TRUE)]
        v[stats::runif(n) < (columns[[match(L$link, nm)]]$missing_rate %||% 0)] <- ""
        cols[[L$link]] <- v
      }
    }
  }
  cols <- apply_crosstabs(cols, rp, n)
  for (co in columns) {                                           # empty exactly when another column says so
    ew <- co[["empty_when"]]
    if (!is.null(ew) && !is.null(cols[[ew$column]]) && !is.null(cols[[co$name]]) && is.character(cols[[co$name]]))
      cols[[co$name]][as.character(cols[[ew$column]]) %in% as.character(unlist(ew$values))] <- ""
  }
  for (co in columns) if (co$kind == "datetime" && !(co$name %in% copied))
    cols[[co$name]] <- synth_dates(co, cols[[paste0(co$name, "__period")]], cols[[co$name]], n)
  has_time <- function(b) isTRUE(columns[[match(b, nm)]]$has_time)
  for (co in columns) {                                           # a date = the parent row's date + a real-looking number of days
    if (!identical(co[["helper"]], "pgap") || is.null(attach)) next
    b <- co$of
    src <- synth[[parent]]$cols[[co$parent_col]]
    if (is.null(src) || !is.numeric(src) || !is.numeric(cols[[b]]) || b %in% copied) next
    p <- src[attach]; G <- cols[[co$name]]
    ok <- !is.na(cols[[b]]) & G != "" & !is.na(p)
    d <- (dt_days(p[ok]) + trunc(as.numeric(G[ok]))) * 86400
    cols[[b]][ok] <- if (has_time(b)) .rand_time(d) else d
  }
  for (co in columns) {                                           # a later date = the earlier one + a real-looking wait
    if (!identical(co[["helper"]], "gap")) next
    a <- co$from; b <- co$of
    if (!is.numeric(cols[[a]]) || !is.numeric(cols[[b]]) || b %in% copied) next
    G <- cols[[co$name]]
    ok <- !is.na(cols[[a]]) & !is.na(cols[[b]]) & G != ""
    d <- cols[[a]][ok] + pmax(0, trunc(as.numeric(G[ok]))) * 86400
    cols[[b]][ok] <- if (has_time(b)) .rand_time(dt_days(d) * 86400) else d
  }
  for (co in columns) {                                           # an age = the years between two dates of the row
    dy <- as.character(unlist(co[["derived_years"]]))
    if (length(dy) == 2 && is.numeric(cols[[dy[1]]]) && is.numeric(cols[[dy[2]]]) && is.character(cols[[co$name]])) {
      A <- cols[[dy[1]]]; B <- cols[[dy[2]]]
      ok <- cols[[co$name]] != "" & !is.na(A) & !is.na(B)
      cols[[co$name]][ok] <- num_text(pmax(0, floor(floor((B[ok] - A[ok]) / 86400) / 365.25)))
    }
  }
  num <- character(0)
  for (co in columns) {
    name <- co$name
    if (identical(co[["helper"]], "period") && !is.null(cols[[co$of]])) {      # keep the period in step with the final date
      d <- cols[[co$of]]
      v <- rep("", n)
      if (is.numeric(d) && any(!is.na(d))) v[!is.na(d)] <- period_label(d[!is.na(d)], co$scheme)
      cols[[name]] <- v
    } else if (co$kind == "float" && is.character(cols[[name]])) {
      v <- cols[[name]]
      f <- suppressWarnings(as.numeric(v))
      ok <- v != "" & !is.na(f)
      v[ok] <- num_text(round(f[ok], co$decimals %||% 2))
      cols[[name]] <- v
      num <- c(num, name)
    } else if (co$kind == "integer" || (isTRUE(co$all_numeric) && co$kind != "datetime")) num <- c(num, name)
  }
  dates <- columns[vapply(columns, function(co) co$kind == "datetime", TRUE)]
  names(dates) <- vapply(dates, function(co) co$name, "")
  list(header = header, cols = cols, n = n, dates = dates, num = num)
}

# a date column as it is written: an Excel day number, a real Excel date cell, or text in its own format
.format_dates <- function(co, v, as_excel) {
  if (!is.numeric(v)) return(v)
  if (isTRUE(co$serial)) return(structure(num_text(dt_to_excel_serial(v)), num = TRUE))
  if (as_excel && isTRUE(co$excel_date)) return(structure(v, dt = TRUE))
  dt_format(v, co$format %||% "%Y-%m-%d")
}

# Step 2.  Returns list(out_dir, outputs).
run_synth <- function(profile_dir, out_dir, scale = 1, seed = DEFAULT_SEED) {
  reports <- load_profiles(profile_dir)
  if (!length(reports)) fail("ERROR: no profile_*.json in %s", profile_dir)
  schema <- list(files = list(), order = list())
  sp <- file.path(profile_dir, "schema.json")
  if (file.exists(sp)) schema <- read_json(sp)
  set.seed(seed)
  names(reports) <- vapply(reports, function(r) r$base, "")
  ord <- as.character(unlist(schema$order))
  ord <- c(ord[ord %in% names(reports)], names(reports)[!(names(reports) %in% ord)])
  synth <- list()
  for (base in ord) {
    say("  building : %s", base)
    synth[[base]] <- synth_table(reports[[base]], schema$files[[base]], synth, scale)
  }
  dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
  sf <- vapply(reports, function(r) r$source_file, "")
  si <- vapply(reports, function(r) as.numeric(r$sheet_index %||% 0), 0)
  reports <- reports[order(match(sf, c_sort(unique(sf))), si)]              # sheets in their real order
  sf <- vapply(reports, function(r) r$source_file, "")
  outputs <- list()
  for (fname in unique(sf)) {
    reps <- reports[sf == fname]
    stem <- file_stem(fname)
    is_csv <- identical(reps[[1]]$file_type, "csv")
    sheets <- list()
    for (r in reps) {
      tb <- synth[[r$base]]
      cols <- list()
      for (h in tb$header) {
        v <- tb$cols[[h]]
        if (is.null(v)) v <- rep("", tb$n)
        if (!is.null(tb$dates[[h]])) v <- .format_dates(tb$dates[[h]], v, as_excel = !is_csv)
        else if (h %in% tb$num && is.character(v)) attr(v, "num") <- TRUE
        cols[[h]] <- v
      }
      sheets[[length(sheets) + 1]] <- list(name = if (length(r$sheet) && nzchar(r$sheet)) r$sheet else "Sheet1", header = tb$header, cols = cols)
      outputs[[length(outputs) + 1]] <- list(file = fname, table = r$base, rows = tb$n, columns = length(tb$header))
    }
    if (is_csv) {
      out <- file.path(out_dir, sprintf("synthetic_%s.csv", stem))
      write_csv(out, sheets[[1]]$header, sheets[[1]]$cols)
    } else {
      out <- file.path(out_dir, sprintf("synthetic_%s.xlsx", stem))
      write_workbook(out, sheets)
    }
    say("  [OK] %s", basename(out))
  }
  write_json(list(synthetic = TRUE, tool = VERSION, made = format(Sys.time(), "%Y-%m-%dT%H:%M:%S"), seed = seed, scale = scale,
                  note = paste0("Every value is invented from aggregate profiles. Results on ",
                                "these files test the software; they are not findings."),
                  files = outputs), file.path(out_dir, "synthetic_manifest.json"))
  say("\nDONE. Synthetic files in: %s", out_dir)
  say("Every value is invented. Results on these files test your script; they are not findings.")
  invisible(list(out_dir = out_dir, outputs = outputs))
}


# ==============================================================================
#  Step 3 -- the toolbox for your own analysis script
#      g <- new.env(); sys.source("gov_one.R", envir = g)
#      tabs <- g$open_tables(".")
#  Every table is a data frame of TEXT columns ("" = empty cell), named as in the file.
#  All of base R is yours to use on it (table, tapply, aggregate, merge, t.test, lm, glm ...);
#  the functions below are the same helpers gov_one.py offers, with the same names and results.
# ==============================================================================
USING_SYNTHETIC <- FALSE

.as_frame <- function(tb, synthetic) {
  df <- structure(unname(tb$cols[tb$header]), names = tb$header, class = "data.frame", row.names = .set_row_names(tb$n))
  attr(df, "file") <- tb$file
  attr(df, "sheet") <- tb$sheet
  attr(df, "synthetic") <- synthetic
  df
}

# list(table name = data frame) for every Excel / CSV file in the folder.
# Real files win.  When the folder holds no real data file, the synthetic files (in the folder or
# in its 'synthetic' sub-folder) are used under the SAME table names, and USING_SYNTHETIC is set,
# so one script runs unchanged at home and on the premises.
open_tables <- function(folder = ".") {
  paths <- find_inputs(folder)$good
  synthetic <- FALSE
  if (!length(paths)) {
    for (d in c(folder, file.path(folder, SYNTH_DIR))) {
      f <- c_sort(list.files(d, pattern = "^synthetic_.*\\.(xlsx|csv)$", ignore.case = TRUE))
      if (length(f)) {
        paths <- file.path(d, f)
        synthetic <- TRUE
        break
      }
    }
  }
  if (!length(paths)) fail("ERROR: no data file (and no synthetic_ file) in %s", normalizePath(folder, winslash = "/", mustWork = FALSE))
  USING_SYNTHETIC <<- synthetic
  out <- list()
  for (p in paths) for (tb in read_file_tables(p)) {
    if (synthetic) tb$base <- sub("^synthetic_", "", tb$base)
    out[[tb$base]] <- .as_frame(tb, synthetic)
  }
  say("%s data: %d table(s) from %s", if (synthetic) "SYNTHETIC" else "REAL", length(out), normalizePath(folder, winslash = "/", mustWork = FALSE))
  out
}

# text -> number (NA when empty or not a number)
to_num <- function(v) {
  if (is.numeric(v)) return(v)
  f <- to_float(v)
  f[is_missing(v)] <- NA
  f
}

# text -> Date (NA when it is not a date).  Reads 2024-03-15, 15/03/2024, 15.03.2024 and date-times.
to_date <- function(v) {
  if (inherits(v, "Date")) return(v)
  as.Date(dt_days(parse_date_dt(v)), origin = "1970-01-01")
}

# 'YYYY-MM' of a date cell, or ""
month_key <- function(v) {
  d <- if (inherits(v, "Date")) as.numeric(v) * 86400 else parse_date_dt(v)
  out <- dt_format(d, "%Y-%m")
  out[is.na(d)] <- ""
  out
}

# rows per value (or per combination of values) of one or more columns, largest first
count_by <- function(df, ...) {
  cols <- c(...)
  key <- do.call(paste, c(unname(df[cols]), sep = "\r"))
  first <- !duplicated(key)
  out <- df[first, cols, drop = FALSE]
  out$n <- tabulate(match(key, key[first]))
  out <- out[order(-out$n), , drop = FALSE]
  rownames(out) <- NULL
  out
}

# Each row plus the columns of its parent row (matched on the key).  Rows without a parent get
# empty cells.  A parent column that already exists in the table gets the prefix.
join_parent <- function(df, parent, on, parent_on = on, prefix = "") {
  i <- match(df[[on]], parent[[parent_on]])
  i[df[[on]] == ""] <- NA
  for (a in setdiff(names(parent), parent_on)) {
    v <- parent[[a]][i]
    name <- if (a %in% names(df) || nzchar(prefix)) paste0(prefix, a) else a
    if (name %in% names(df)) v[is.na(i)] <- df[[name]][is.na(i)] else v[is.na(i)] <- ""
    df[[name]] <- v
  }
  df
}

# Rows (or the sum of a numeric column) per month: list(months = months in order, series = list(group = one
# number per month)).  Months with nothing are 0; the months run from the first to the last one seen.
monthly_counts <- function(df, date_col, by = NULL, value = NULL) {
  m <- month_key(df[[date_col]])
  ok <- m != ""
  if (!any(ok)) return(list(months = character(0), series = list()))
  g <- if (is.null(by)) rep("all", nrow(df)) else df[[by]]
  x <- if (is.null(value)) rep(1, nrow(df)) else { v <- to_num(df[[value]]); v[is.na(v)] <- 0; v }
  first <- as.numeric(strsplit(min(m[ok]), "-")[[1]]); last <- as.numeric(strsplit(max(m[ok]), "-")[[1]])
  k <- (first[1] * 12 + first[2] - 1):(last[1] * 12 + last[2] - 1)
  axis <- sprintf("%04d-%02d", as.integer(k %/% 12), as.integer(k %% 12 + 1))
  series <- list()
  for (grp in c_sort(unique(g[ok]))) {
    s <- tapply(x[ok & g == grp], factor(m[ok & g == grp], levels = axis), sum)
    s[is.na(s)] <- 0
    series[[grp]] <- as.numeric(s)
  }
  list(months = axis, series = series)
}

# A count for a report that may leave: '<k' instead of a number below k.  One count gives the
# number or the text; several counts give a list of them, ready to be cells of page$table().
safe <- function(n, k = MIN_CELL_COUNT) {
  if (length(n) != 1) return(lapply(unname(n), safe, k = k))
  if (!is.na(n) && n > 0 && n < k) sprintf("<%d", k) else n
}

# ---- statistics (the same formulas as gov_one.py, so both give the same numbers) ----
quantile_p <- function(xs, p) percentile(sort(xs), p)             # the p-th percentile (0-100)

describe <- function(xs) {
  xs <- sort(xs[!is.na(xs)])
  n <- length(xs)
  if (!n) return(list(n = 0))
  m <- sum(xs) / n
  list(n = n, mean = m, sd = if (n > 1) sqrt(sum((xs - m)^2) / (n - 1)) else 0, min = xs[1], p25 = percentile(xs, 25),
       median = percentile(xs, 50), p75 = percentile(xs, 75), max = xs[n])
}

p_normal <- function(z) 2 * stats::pnorm(-abs(z))                 # two-sided p values
p_t <- function(t, df) 2 * stats::pt(-abs(t), df)
p_chi2 <- function(x, df) stats::pchisq(x, df, lower.tail = FALSE)
p_f <- function(f, d1, d2) stats::pf(f, d1, d2, lower.tail = FALSE)
t_crit <- function(df, alpha = 0.05) stats::qt(1 - alpha / 2, df)

# counts of one column by another, as a matrix; empty cells of either column are skipped
crosstab <- function(df, row_col, col_col) {
  a <- df[[row_col]]; b <- df[[col_col]]
  ok <- !is_missing(a) & !is_missing(b)
  rl <- c_sort(unique(a[ok])); cl <- c_sort(unique(b[ok]))
  matrix(tabulate((match(b[ok], cl) - 1) * length(rl) + match(a[ok], rl), length(rl) * length(cl)), length(rl), length(cl),
         dimnames = list(rl, cl))
}

# Chi-square test of independence on a matrix of counts: chi2, df, p, n, cramers_v and the
# smallest expected count (below 5 = read with care)
chi_square <- function(tab) {
  tab <- as.matrix(tab)
  tab <- tab[rowSums(tab) > 0, colSums(tab) > 0, drop = FALSE]
  n <- sum(tab)
  if (n == 0 || nrow(tab) < 2 || ncol(tab) < 2) return(list(chi2 = NaN, df = 0, p = NaN, n = n, cramers_v = NaN, min_expected = NaN))
  e <- outer(rowSums(tab), colSums(tab)) / n
  chi2 <- sum((tab - e)^2 / e)
  df <- (nrow(tab) - 1) * (ncol(tab) - 1)
  list(chi2 = chi2, df = df, p = p_chi2(chi2, df), n = n, cramers_v = sqrt(chi2 / (n * (min(dim(tab)) - 1))), min_expected = min(e))
}

# Welch's t test for two groups of numbers: means, difference (a - b) with a 95% interval, t, df, two-sided p
t_test <- function(a, b) {
  da <- describe(a); db <- describe(b)
  if (da$n < 2 || db$n < 2) return(list(n_a = da$n, n_b = db$n, p = NaN))
  va <- da$sd^2 / da$n; vb <- db$sd^2 / db$n
  se <- sqrt(va + vb)
  diff <- da$mean - db$mean
  if (se == 0) return(list(n_a = da$n, n_b = db$n, mean_a = da$mean, mean_b = db$mean, diff = diff, p = NaN))
  df <- (va + vb)^2 / (va^2 / (da$n - 1) + vb^2 / (db$n - 1))
  t <- diff / se
  h <- t_crit(df) * se
  list(n_a = da$n, n_b = db$n, mean_a = da$mean, mean_b = db$mean, diff = diff, ci95 = c(diff - h, diff + h), t = t, df = df, p = p_t(t, df))
}

# Mann-Whitney U test (rank test for two groups; normal approximation with tie correction)
mann_whitney <- function(a, b) {
  a <- a[!is.na(a)]; b <- b[!is.na(b)]
  na <- length(a); nb <- length(b)
  if (na < 1 || nb < 1) return(list(p = NaN))
  r <- rank(c(a, b))
  u <- sum(r[seq_len(na)]) - na * (na + 1) / 2
  n <- na + nb
  ties <- tabulate(match(c(a, b), unique(c(a, b))))
  v <- if (n > 1) na * nb / 12 * ((n + 1) - sum(ties^3 - ties) / (n * (n - 1))) else 0
  if (v <= 0) return(list(u = u, p = NaN))
  z <- (u - na * nb / 2) / sqrt(v)
  list(u = u, z = z, p = p_normal(z), n_a = na, n_b = nb, prob_a_above_b = u / (na * nb))
}

pearson <- function(x, y) {
  n <- length(x)
  if (n < 3 || n != length(y)) return(list(r = NaN, n = n, p = NaN))
  mx <- sum(x) / n; my <- sum(y) / n
  sxx <- sum((x - mx)^2); syy <- sum((y - my)^2)
  if (sxx == 0 || syy == 0) return(list(r = NaN, n = n, p = NaN))
  r <- max(-1, min(1, sum((x - mx) * (y - my)) / sqrt(sxx * syy)))
  list(r = r, n = n, p = p_t(r * sqrt((n - 2) / max(1e-300, 1 - r * r)), n - 2))
}

spearman <- function(x, y) {
  out <- pearson(rank(x), rank(y))
  list(rho = out$r, n = out$n, p = out$p)
}

# two proportions x1/n1 and x2/n2: the difference with a 95% interval, z and the two-sided p
proportion_test <- function(x1, n1, x2, n2) {
  if (n1 <= 0 || n2 <= 0) return(list(p = NaN))
  p1 <- x1 / n1; p2 <- x2 / n2
  pool <- (x1 + x2) / (n1 + n2)
  se0 <- sqrt(pool * (1 - pool) * (1 / n1 + 1 / n2))
  se <- sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2)
  z <- if (se0 > 0) (p1 - p2) / se0 else NaN
  list(p1 = p1, p2 = p2, diff = p1 - p2, ci95 = c(p1 - p2 - 1.96 * se, p1 - p2 + 1.96 * se), z = z, p = if (se0 > 0) p_normal(z) else NaN)
}

# Ordinary least squares.  y: numbers; X: a matrix or data frame of predictors (one row per y).
# Returns names, coef, se, t, p, ci95 (matrix), r2, n, df, fitted, resid.
# hac_lags = NULL: classical standard errors and t tests (the same as lm).  hac_lags = k: Newey-West
# standard errors for a time series (about 6 for monthly data); tests then use the normal distribution.
ols <- function(y, X, names = NULL, intercept = TRUE, hac_lags = NULL) {
  y <- as.numeric(y)
  X <- as.matrix(X)
  storage.mode(X) <- "double"
  n <- length(y)
  names <- c(if (intercept) "intercept", names %||% colnames(X) %||% sprintf("x%d", seq_len(ncol(X))))
  if (intercept) X <- cbind(1, X)
  p <- ncol(X)
  if (n <= p || p == 0) return(list(error = sprintf("too few rows for %d coefficients", p), n = n))
  inv <- tryCatch(solve(crossprod(X)), error = function(e) NULL)
  if (is.null(inv)) return(list(error = "the predictors are linearly dependent", n = n))
  coef <- as.numeric(inv %*% crossprod(X, y))
  fitted <- as.numeric(X %*% coef)
  resid <- y - fitted
  df <- n - p
  sst <- sum((y - sum(y) / n)^2); sse <- sum(resid^2)
  if (is.null(hac_lags)) cov <- sse / df * inv
  else {
    L <- as.integer(hac_lags)
    Xe <- X * resid
    S <- crossprod(Xe)
    for (lag in seq_len(min(L, n - 1))) {
      G <- crossprod(Xe[(lag + 1):n, , drop = FALSE], Xe[1:(n - lag), , drop = FALSE])
      S <- S + (1 - lag / (L + 1)) * (G + t(G))
    }
    cov <- n / df * inv %*% S %*% inv
  }
  se <- sqrt(pmax(0, diag(cov)))
  tv <- ifelse(se > 0, coef / se, NaN)
  crit <- if (is.null(hac_lags)) t_crit(df) else 1.96
  list(names = names, coef = coef, se = se, t = tv, p = if (is.null(hac_lags)) p_t(tv, df) else p_normal(tv),
       ci95 = cbind(low = coef - crit * se, high = coef + crit * se), r2 = if (sst > 0) 1 - sse / sst else NaN, n = n, df = df,
       fitted = fitted, resid = resid)
}

# Interrupted time series on monthly (or other equally spaced) values.  break_index: position
# (counted from 1) of the first 'after' value.  Fits level + trend before, and the change in level
# and in trend after.  seasonal = TRUE adds a yearly sine / cosine (needs 24+ values).
# Returns the ols result plus level_change and slope_change as c(estimate, low, high, p).
its <- function(values, break_index, hac_lags = 6, seasonal = FALSE) {
  i <- seq_along(values) - 1
  t <- i - (break_index - 1)
  post <- as.numeric(t >= 0)
  X <- cbind(trend = t, level_change = post, slope_change = t * post)
  if (seasonal) X <- cbind(X, sin12 = sin(2 * pi * i / 12), cos12 = cos(2 * pi * i / 12))
  out <- ols(values, X, hac_lags = hac_lags)
  if (is.null(out$error)) for (a in c("level_change", "slope_change")) {
    j <- match(a, out$names)
    out[[a]] <- c(estimate = out$coef[[j]], low = out$ci95[[j, 1]], high = out$ci95[[j, 2]], p = out$p[[j]])
  }
  out
}

# ---- an offline HTML report ---------------------------------------------------
# Builds one self-contained HTML page (no internet, no external file).
#     page <- g$Report("Title"); page$h("Section"); page$p("text"); page$table(header, rows)
#     page$line(months, list(name = numbers)); page$bars(c(label = number)); page$save("results/report.html")
# The page says in a banner when it was built from synthetic files.
Report <- function(title, synthetic = NULL) {
  synthetic <- if (is.null(synthetic)) USING_SYNTHETIC else synthetic
  parts <- character(0)
  colors <- c("#2f66a8", "#b83a48", "#12766a", "#96660a", "#6a55b0", "#1f6f9c", "#55742f", "#7a7a7a")
  e <- function(s) {
    s <- gsub("&", "&amp;", as.character(s), fixed = TRUE)
    s <- gsub("<", "&lt;", s, fixed = TRUE)
    s <- gsub(">", "&gt;", s, fixed = TRUE)
    s <- gsub("\"", "&quot;", s, fixed = TRUE)
    gsub("'", "&#x27;", s, fixed = TRUE)
  }
  add <- function(x) parts[length(parts) + 1] <<- x
  cell <- function(v) {
    if (is.list(v)) v <- v[[1]]
    if (is.null(v) || length(v) == 0) v <- NA_real_
    sprintf("<td%s dir=\"auto\">%s</td>", if (is.numeric(v)) " class=\"n\"" else "", e(fmt(v)))
  }
  self <- list()
  self$h <- function(text, level = 2) add(sprintf("<h%d>%s</h%d>", level, e(text), level))
  self$p <- function(text) add(sprintf("<p dir=\"auto\">%s</p>", e(text)))
  self$note <- function(text) add(sprintf("<p class=\"note\" dir=\"auto\">%s</p>", e(text)))
  # rows: a data frame, or a list of rows (each a list or a vector of cells)
  self$table <- function(header, rows, caption = NULL) {
    if (is.data.frame(rows)) rows <- lapply(seq_len(nrow(rows)), function(i) lapply(rows, function(col) col[[i]]))
    body <- vapply(rows, function(r) paste0("<tr>", paste(vapply(seq_along(r), function(j) cell(r[j]), ""), collapse = ""), "</tr>"), "")
    add(paste0("<div class=\"tw\"><table>", if (!is.null(caption)) sprintf("<caption dir=\"auto\">%s</caption>", e(caption)) else "",
               "<thead><tr>", paste(sprintf("<th dir=\"auto\">%s</th>", e(header)), collapse = ""), "</tr></thead><tbody>",
               paste(body, collapse = ""), "</tbody></table></div>"))
  }
  # A line chart.  labels: x axis (e.g. months); series: list(name = numbers, NA for gaps);
  # mark: a label at which to draw a vertical line (e.g. the month of a reform).
  self$line <- function(labels, series, title = NULL, mark = NULL) {
    W <- 860; H <- 300; L <- 56; R <- 16; T <- 18; B <- 42
    vals <- unlist(series)
    vals <- vals[!is.na(vals)]
    if (!length(vals) || !length(labels)) return(invisible(NULL))
    lo <- min(0, min(vals)); hi <- max(vals)
    if (hi <= lo) hi <- lo + 1
    n <- length(labels)
    X <- function(i) L + (W - L - R) * ((i - 1) / max(1, n - 1))
    Y <- function(v) T + (H - T - B) * (1 - (v - lo) / (hi - lo))
    svg <- sprintf("<svg viewBox=\"0 0 %d %d\" role=\"img\" aria-label=\"%s\">", W, H, e(title %||% "chart"))
    for (q in 0:4) {
      v <- lo + (hi - lo) * q / 4
      svg <- c(svg, sprintf("<line x1=\"%d\" x2=\"%d\" y1=\"%.1f\" y2=\"%.1f\" class=\"grid\"/><text x=\"%d\" y=\"%.1f\" class=\"ax\" text-anchor=\"end\">%s</text>",
                            L, W - R, Y(v), Y(v), L - 6, Y(v) + 4, e(fmt(v, 1))))
    }
    for (i in seq(1, n, by = max(1, n %/% 10)))
      svg <- c(svg, sprintf("<text x=\"%.1f\" y=\"%d\" class=\"ax\" text-anchor=\"middle\">%s</text>", X(i), H - B + 16, e(labels[i])))
    if (!is.null(mark) && mark %in% labels) {
      x <- X(match(mark, labels))
      svg <- c(svg, sprintf("<line x1=\"%.1f\" x2=\"%.1f\" y1=\"%d\" y2=\"%d\" class=\"mark\"/><text x=\"%.1f\" y=\"%d\" class=\"ax\">%s</text>",
                            x, x, T, H - B, x + 4, T + 10, e(mark)))
    }
    legend <- character(0)
    for (k in seq_along(series)) {
      color <- colors[(k - 1) %% length(colors) + 1]
      s <- utils::head(as.numeric(series[[k]]), n)
      ok <- !is.na(s)
      pen <- c(FALSE, ok[-length(ok)])[ok]                         # a gap lifts the pen
      d <- sprintf("%s%.1f %.1f", ifelse(pen, "L", "M"), X(which(ok)), Y(s[ok]))
      svg <- c(svg, sprintf("<path d=\"%s\" fill=\"none\" stroke=\"%s\" stroke-width=\"2\"/>", paste(d, collapse = " "), color))
      legend <- c(legend, sprintf("<span><i style=\"background:%s\"></i>%s</span>", color, e(names(series)[k])))
    }
    add(sprintf("<figure>%s%s</svg><div class=\"lg\">%s</div></figure>",
                if (!is.null(title)) sprintf("<figcaption dir=\"auto\">%s</figcaption>", e(title)) else "", paste(svg, collapse = ""),
                paste(legend, collapse = "")))
  }
  # horizontal bars.  pairs: c(label = number), drawn in the order given
  self$bars <- function(pairs, title = NULL) {
    pairs <- unlist(pairs)
    pairs <- pairs[!is.na(pairs)]
    if (!length(pairs)) return(invisible(NULL))
    top <- max(abs(pairs))
    if (top == 0) top <- 1
    rows <- sprintf("<div class=\"br\"><span dir=\"auto\">%s</span><b style=\"width:%.1f%%\"></b><em>%s</em></div>",
                    e(names(pairs)), 100 * abs(pairs) / top, e(vapply(pairs, fmt, "")))
    add(sprintf("<figure>%s<div class=\"bars\">%s</div></figure>",
                if (!is.null(title)) sprintf("<figcaption dir=\"auto\">%s</figcaption>", e(title)) else "", paste(rows, collapse = "")))
  }
  self$save <- function(path) {
    dir.create(dirname(path), showWarnings = FALSE, recursive = TRUE)
    banner <- if (synthetic) paste0("<div class=\"syn\"><b>Synthetic test data \u2014 not findings.</b> Every number on this page comes from ",
                                    "invented files. It shows that the script works; it says nothing about the real world.</div>") else ""
    css <- paste0("body{font:15px/1.5 \"Segoe UI\",Arial,sans-serif;color:#1d2733;background:#f5f7fa;margin:0}",
      "main{max-width:960px;margin:0 auto;padding:24px 16px 60px}h1{font-size:26px;margin:0 0 4px}",
      "h2{font-size:19px;margin:30px 0 8px;padding-top:14px;border-top:3px solid #2f66a8;color:#2f66a8}h3{font-size:16px;margin:20px 0 6px}",
      ".meta{color:#5b6877;font-size:13px}.syn{background:#fdf4dc;border:1px solid #d9b24c;border-left:6px solid #96660a;padding:10px 14px;margin:14px 0;border-radius:6px}",
      ".note{color:#5b6877;font-size:13px}.tw{overflow-x:auto}table{border-collapse:collapse;background:#fff;margin:8px 0;font-size:14px}",
      "caption{text-align:start;font-weight:600;padding:4px 0}th,td{border:1px solid #d5dce5;padding:5px 9px;text-align:start}",
      "th{background:#eaf1f9}td.n{text-align:end;font-variant-numeric:tabular-nums}figure{margin:12px 0;background:#fff;border:1px solid #d5dce5;border-radius:8px;padding:10px}",
      "figcaption{font-weight:600;margin-bottom:6px}svg{width:100%;height:auto}.grid{stroke:#e3e8ee}.ax{font-size:11px;fill:#5b6877}",
      ".mark{stroke:#b83a48;stroke-dasharray:5 4}.lg span{display:inline-flex;align-items:center;gap:6px;margin-inline-end:14px;font-size:13px}",
      ".lg i{width:12px;height:12px;border-radius:2px;display:inline-block}.br{display:grid;grid-template-columns:minmax(90px,30%) 1fr 70px;gap:8px;align-items:center;margin:3px 0;font-size:13px}",
      ".br b{background:#2f66a8;height:14px;border-radius:3px;display:block}.br em{font-style:normal;text-align:end;font-variant-numeric:tabular-nums}",
      "@media print{body{background:#fff}.tw{overflow:visible}figure{break-inside:avoid}}")
    page <- paste0("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
                   "<title>", e(title), "</title><style>", css, "</style></head><body><main><h1 dir=\"auto\">", e(title),
                   "</h1><div class=\"meta\">built ", format(Sys.time(), "%Y-%m-%d %H:%M"), " by ", e(VERSION), "</div>", banner,
                   paste(parts, collapse = ""), "</main></body></html>")
    write_text(path, page)
    say("[OK] report: %s%s", normalizePath(path, winslash = "/", mustWork = FALSE), if (synthetic) "  (SYNTHETIC data - not findings)" else "")
    invisible(path)
  }
  self
}


# ==============================================================================
#  Command line
# ==============================================================================
.BASE_PACKAGES <- c("base", "stats", "utils", "methods", "graphics", "grDevices", "datasets", "tools", "compiler", "parallel",
                    "grid", "splines", "stats4", "tcltk")

# what would be read: files, sheets, header row, column names.  No data is shown.
run_check <- function(data_dir) {
  ext <- setdiff(loadedNamespaces(), .BASE_PACKAGES)
  say("packages loaded that do not come with R itself: %s", if (length(ext)) paste(sort(ext), collapse = ", ") else "none")
  fi <- find_inputs(data_dir)
  for (p in fi$good) {
    name <- basename(p)
    parts <- tryCatch({
      if (grepl("\\.(csv|tsv)$", tolower(p))) list(list(name = "", m = utils::head(read_csv_rows(p), 80)))
      else lapply(xlsx_sheets(p), function(s) list(name = s$name, m = read_sheet(p, s$part, max_rows = 80)$m))
    }, error = function(e) e)
    if (inherits(parts, "error")) {
      say("  [!] %s: cannot be read (%s)", name, conditionMessage(parts))
      next
    }
    for (s in parts) {
      label <- paste0(name, if (nzchar(s$name)) paste0(" / ", s$name) else "")
      st <- matrix(pystrip(s$m), nrow = nrow(s$m))
      st <- st[rowSums(st != "") > 0, , drop = FALSE]
      if (!nrow(st)) {
        say("  %s: empty", label)
        next
      }
      h <- find_header_row(st)
      head <- st[h + 1, ]
      head <- head[head != ""]
      say("  %s: header on row %d, %d columns: %s%s", label, h + 1, length(head), paste(utils::head(head, 14), collapse = ", "),
          if (length(head) > 14) " ..." else "")
    }
  }
  for (b in fi$bad) say("  [!] %s: %s", b[1], b[2])
  if (!length(fi$good)) say("  no Excel / CSV data file in %s", data_dir)
  say("  profile JSON files found: %d", length(list.files(data_dir, pattern = "^profile_.*\\.json$")) +
        length(list.files(file.path(data_dir, SEND_DIR), pattern = "^profile_.*\\.json$")))
}

.profile_dir <- function(data_dir) {
  for (d in c(data_dir, file.path(data_dir, SEND_DIR))) if (length(list.files(d, pattern = "^profile_.*\\.json$"))) return(d)
  NULL
}

.USAGE <- paste(
  "usage: Rscript gov_one.R [auto|profile|synth|check|settings] [--data DIR] [--out DIR] [--k K] [--rows R] [--seed N]",
  "",
  "Profile Excel / CSV files into aggregate JSON, rebuild synthetic files from the JSON, and give",
  "your own analysis script a statistics toolbox.  Base R only; offline.",
  "",
  "  --data DIR   folder with the data files (profile) or the JSON files (synth); default: here",
  "  --out DIR    output folder (default: to_send for profile, synthetic for synth)",
  "  --k K        minimum cell count (default 5)",
  "  --rows R     synth: size of the synthetic files relative to the real ones (0.1 = a tenth)",
  "  --seed N     synth: random seed", sep = "\n")

main <- function(argv = commandArgs(trailingOnly = TRUE)) {
  a <- list(command = "auto", data = ".", out = NULL, k = NULL, rows = 1, seed = DEFAULT_SEED)
  i <- 1
  while (i <= length(argv)) {
    x <- argv[i]
    if (x %in% c("-h", "--help")) {
      cat(.USAGE, "\n")
      return(invisible(NULL))
    }
    if (grepl("^--(data|out|k|rows|seed)(=|$)", x)) {
      key <- sub("^--([a-z]+).*$", "\\1", x)
      if (grepl("=", x, fixed = TRUE)) val <- sub("^[^=]*=", "", x)
      else {
        i <- i + 1
        val <- if (i <= length(argv)) argv[i] else fail("ERROR: %s needs a value", x)
      }
      a[[key]] <- if (key %in% c("k", "rows", "seed")) suppressWarnings(as.numeric(val)) else val
      if (key %in% c("k", "rows", "seed") && is.na(a[[key]])) fail("ERROR: --%s needs a number", key)
    } else if (x %in% c("auto", "profile", "synth", "check", "settings")) a$command <- x
    else fail("ERROR: argument not understood: %s\n%s", x, .USAGE)
    i <- i + 1
  }
  if (!dir.exists(a$data)) fail("ERROR: folder not found: %s", a$data)
  data <- normalizePath(a$data, winslash = "/")
  say("%s | R %s", VERSION, as.character(getRversion()))
  if (!isTRUE(l10n_info()[["UTF-8"]]))
    say("  [!] this R session does not work in UTF-8, so Hebrew file names and headers may be misread: use R 4.2 or later")
  cmd <- a$command
  if (cmd == "auto") {
    cmd <- if (length(find_inputs(data)$good)) "profile" else if (!is.null(.profile_dir(data))) "synth" else NULL
    if (is.null(cmd)) fail(paste0("Nothing to do in %s: no Excel / CSV data file and no profile_*.json.\n",
                                  "On the premises: put gov_one.R next to the data files and run it.\n",
                                  "At home: put it next to the JSON files (or the to_send folder) and run it."), data)
    say("auto     : %s", if (cmd == "profile") "data files found -> profile" else "only JSON files found -> synthetic files")
  }
  t0 <- proc.time()[["elapsed"]]
  if (cmd == "check") run_check(data)
  else if (cmd == "settings") {
    p <- file.path(data, SETTINGS_FILE)
    if (file.exists(p)) fail("%s already exists; edit it.", p)
    write_json(settings_template(), p, indent = 2)
    say("[OK] %s written. Edit it, then run the profile again.", p)
  } else if (cmd == "profile") run_profile(data, a$out, a$k)
  else if (cmd == "synth") {
    pd <- .profile_dir(data)
    if (is.null(pd)) fail("ERROR: no profile_*.json in %s (or its %s folder)", data, SEND_DIR)
    if (length(find_inputs(data)$good)) say("  [!] real data files are in this folder too; the synthetic files go to a separate folder")
    run_synth(pd, a$out %||% file.path(data, SYNTH_DIR), scale = a$rows, seed = a$seed)
  }
  say("(%.1f s)", proc.time()[["elapsed"]] - t0)
  invisible(NULL)
}

# run only when this file is started with Rscript (not when another script loads it with sys.source)
if (sys.nframe() == 0L && !interactive()) {
  tryCatch(main(), gov_one_exit = function(e) {
    message(conditionMessage(e))
    quit(save = "no", status = 1)
  })
}
