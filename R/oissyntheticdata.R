# oissyntheticdata -- a thin R package around the single file gov_one.R.
#
# The tool is inst/gov_one/gov_one.R.  It is the file r/gov_one.R of the project, copied here
# unchanged by r/build_package.R: edit that file, never the copy.  The functions below only load
# the file and pass a folder to it, so the package and the file carried into the premises always
# behave the same.

.CARRY <- c("gov_one.R", "analysis_template.R")

gov_one_file <- function(name = c("gov_one.R", "analysis_template.R")) {
  name <- match.arg(name)
  p <- system.file("gov_one", name, package = "oissyntheticdata")
  if (!nzchar(p)) stop(name, " was not found in this installation of oissyntheticdata", call. = FALSE)
  p
}

# The tool, loaded exactly as the README of gov_one.R says:  g <- new.env(); sys.source("gov_one.R", envir = g)
gov_one <- function() {
  g <- new.env(parent = globalenv())
  sys.source(gov_one_file("gov_one.R"), envir = g)
  g
}

gov_one_version <- function() gov_one()$VERSION

.folder <- function(x, what) {
  if (!is.character(x) || length(x) != 1L || is.na(x) || !nzchar(x)) stop("'", what, "' must be the path of one folder", call. = FALSE)
  x
}

.number <- function(x, what) {
  if (!is.numeric(x) || length(x) != 1L || is.na(x)) stop("'", what, "' must be one number", call. = FALSE)
  format(x, scientific = FALSE, trim = TRUE, digits = 15)
}

# Runs one command of gov_one.R.  Its own error (class gov_one_exit) becomes an ordinary R error.
.run <- function(g, argv, quiet = FALSE) {
  if (isTRUE(quiet)) {
    con <- file(nullfile(), open = "w")
    sink(con)
    on.exit({
      sink()
      close(con)
    }, add = TRUE)
  }
  tryCatch(g$main(argv), gov_one_exit = function(e) stop(conditionMessage(e), call. = FALSE))
  invisible(NULL)
}

.path <- function(p) normalizePath(p, winslash = "/", mustWork = FALSE)

gov_one_profile <- function(data, out = NULL, k = NULL, quiet = FALSE) {
  data <- .folder(data, "data")
  argv <- c("profile", "--data", data)
  if (!is.null(out)) argv <- c(argv, "--out", .folder(out, "out"))
  if (!is.null(k)) argv <- c(argv, "--k", .number(k, "k"))
  g <- gov_one()
  .run(g, argv, quiet)
  invisible(.path(if (is.null(out)) file.path(data, g$SEND_DIR) else out))
}

gov_one_synth <- function(data, out = NULL, rows = 1, seed = 12345, quiet = FALSE) {
  data <- .folder(data, "data")
  argv <- c("synth", "--data", data, "--rows", .number(rows, "rows"), "--seed", .number(seed, "seed"))
  if (!is.null(out)) argv <- c(argv, "--out", .folder(out, "out"))
  g <- gov_one()
  .run(g, argv, quiet)
  invisible(.path(if (is.null(out)) file.path(data, g$SYNTH_DIR) else out))
}

gov_one_check <- function(data) {
  data <- .folder(data, "data")
  .run(gov_one(), c("check", "--data", data))
}

gov_one_settings <- function(data) {
  data <- .folder(data, "data")
  g <- gov_one()
  .run(g, c("settings", "--data", data))
  invisible(.path(file.path(data, g$SETTINGS_FILE)))
}

# Copies gov_one.R and analysis_template.R into `dir`: the two files to carry into the premises.
gov_one_export <- function(dir, overwrite = FALSE) {
  dir <- .folder(dir, "dir")
  if (!dir.exists(dir) && !dir.create(dir, recursive = TRUE)) stop("the folder could not be created: ", dir, call. = FALSE)
  written <- character(0)
  for (name in .CARRY) {
    src <- gov_one_file(name)
    dst <- file.path(dir, name)
    if (file.exists(dst)) {
      same <- file.size(src) == file.size(dst) &&
        identical(readBin(src, "raw", file.size(src)), readBin(dst, "raw", file.size(dst)))
      if (same) {
        message("  [=] ", .path(dst), " is already there")
        next
      }
      if (!isTRUE(overwrite)) {
        message("  [!] ", .path(dst), " exists with other content; not replaced (use overwrite = TRUE to replace it)")
        next
      }
    }
    if (!file.copy(src, dst, overwrite = TRUE)) stop("the file could not be written: ", dst, call. = FALSE)
    written <- c(written, .path(dst))
    message("  [OK] ", .path(dst))
  }
  message("Carry these two files into the premises. Nothing else is needed there: no installation, no internet.")
  invisible(written)
}
