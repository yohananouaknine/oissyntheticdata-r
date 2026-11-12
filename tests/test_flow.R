# The whole flow on an invented office: profile, synthetic files, analysis, export.
# Base R only; every file is written under tempdir().
library(oissyntheticdata)

g <- gov_one()
stopifnot(is.environment(g), is.function(g$main), is.function(g$open_tables))
stopifnot(identical(gov_one_version(), g$VERSION))
# the shipped file and the package carry the same version number
stopifnot(grepl(paste0(" ", as.character(utils::packageVersion("oissyntheticdata")), " "), g$VERSION, fixed = TRUE))
# loading the file must not have run its command line
stopifnot(!dir.exists("to_send"), !dir.exists("synthetic"))

# ---- an invented office: clients and their cases --------------------------------------------
office <- file.path(tempdir(), "oissd_office")
unlink(office, recursive = TRUE)
dir.create(office)
nc <- 150
clients <- data.frame(
  CLIENT_ID = 1000 + seq_len(nc),
  DISTRICT = rep_len(c("North", "Centre", "South", "Centre", "North", "Centre"), nc),
  BIRTH_YEAR = 1950 + rep_len(c(3, 17, 29, 41, 8, 35, 22, 46, 12), nc),
  stringsAsFactors = FALSE)
n <- 600
cases <- data.frame(
  CASE_ID = 50000 + seq_len(n),
  CLIENT_ID = 1000 + rep_len(c(seq_len(nc), seq(1, nc, by = 3)), n),
  CASE_TYPE = rep_len(c("A", "B", "A", "C", "A", "B", "A"), n),
  AMOUNT = rep_len(c(120, 250, 90, 400, 310, 75, 180, 205, 660, 145, 98), n),
  OPENED = format(as.Date("2023-01-01") + rep_len(seq(0, 700, by = 3), n)),
  stringsAsFactors = FALSE)
cases$CASE_TYPE[c(17, 311)] <- "RAREKIND"          # seen twice: its name must not leave
write.csv(clients, file.path(office, "clients.csv"), row.names = FALSE)
write.csv(cases, file.path(office, "cases.csv"), row.names = FALSE)

# ---- on the premises: the profile -----------------------------------------------------------
to_send <- gov_one_profile(office, quiet = TRUE)
stopifnot(dir.exists(to_send), basename(to_send) == "to_send")
sent <- list.files(to_send)
stopifnot("schema.json" %in% sent, "profile_summary.md" %in% sent,
          "profile_cases.json" %in% sent, "profile_clients.json" %in% sent)
everything <- paste(unlist(lapply(file.path(to_send, sent), readLines, warn = FALSE, encoding = "UTF-8")), collapse = "\n")
stopifnot(!grepl("RAREKIND", everything, fixed = TRUE))            # a rare label does not leave
stopifnot(!grepl("50017", everything, fixed = TRUE))               # nor does an identifier
stopifnot(grepl("Centre", everything, fixed = TRUE))               # a common label does

# a wrong argument is an ordinary R error, not the end of the session
err <- tryCatch(gov_one_profile(file.path(tempdir(), "oissd_no_such_folder"), quiet = TRUE), error = function(e) conditionMessage(e))
stopifnot(is.character(err), grepl("folder not found", err, fixed = TRUE))
err <- tryCatch(gov_one_profile(office, k = 2, quiet = TRUE), error = function(e) conditionMessage(e))
stopifnot(is.character(err), grepl("cannot be below 3", err, fixed = TRUE))

# ---- at home: synthetic files ---------------------------------------------------------------
home <- file.path(tempdir(), "oissd_home")
unlink(home, recursive = TRUE)
dir.create(home)
stopifnot(file.copy(to_send, home, recursive = TRUE))
syn <- gov_one_synth(home, quiet = TRUE)
stopifnot(dir.exists(syn), basename(syn) == "synthetic")
tabs <- g$open_tables(home)
stopifnot(setequal(names(tabs), c("cases", "clients")))
stopifnot(identical(names(tabs$cases), names(cases)), identical(names(tabs$clients), names(clients)))
stopifnot(nrow(tabs$cases) == n, nrow(tabs$clients) == nc)
stopifnot(!any(tabs$cases$CASE_TYPE == "RAREKIND"))
stopifnot(all(tabs$cases$CLIENT_ID %in% tabs$clients$CLIENT_ID))   # the join still works
amount <- g$to_num(tabs$cases$AMOUNT)
stopifnot(!anyNA(amount), min(amount) >= 0)

# the same seed gives the same files; another seed gives other files
again <- file.path(tempdir(), "oissd_again")
other <- file.path(tempdir(), "oissd_other")
gov_one_synth(home, out = again, quiet = TRUE)
gov_one_synth(home, out = other, seed = 7, quiet = TRUE)
same_file <- function(a, b) identical(readLines(a, warn = FALSE), readLines(b, warn = FALSE))
f <- list.files(syn, pattern = "cases\\.csv$")
stopifnot(length(f) == 1L, same_file(file.path(syn, f), file.path(again, f)), !same_file(file.path(syn, f), file.path(other, f)))
half <- file.path(tempdir(), "oissd_half")
gov_one_synth(home, out = half, rows = 0.5, quiet = TRUE)
stopifnot(abs(nrow(g$open_tables(half)$cases) - n / 2) <= 1)

# ---- the helpers agree with R's own tests ---------------------------------------------------
a <- amount[tabs$cases$CASE_TYPE == "A"]
b <- amount[tabs$cases$CASE_TYPE == "B"]
stopifnot(isTRUE(all.equal(g$t_test(a, b)$p, stats::t.test(a, b)$p.value, tolerance = 1e-8)))
ct <- g$crosstab(tabs$cases, "CASE_TYPE", "OPENED")
stopifnot(is.matrix(ct), sum(ct) <= n)

# ---- the two files to carry in --------------------------------------------------------------
carry <- file.path(tempdir(), "oissd_carry")
unlink(carry, recursive = TRUE)
written <- gov_one_export(carry)
stopifnot(length(written) == 2L, all(file.exists(file.path(carry, c("gov_one.R", "analysis_template.R")))))
stopifnot(length(gov_one_export(carry)) == 0L)                      # already there: nothing written
cat("# changed\n", file = file.path(carry, "gov_one.R"), append = TRUE)
stopifnot(length(gov_one_export(carry)) == 0L)                      # other content: left alone
stopifnot(length(gov_one_export(carry, overwrite = TRUE)) == 1L)
stopifnot(identical(readLines(file.path(carry, "gov_one.R")), readLines(gov_one_file("gov_one.R"))))

# the exported file works on its own, as it will on the premises
h <- new.env()
sys.source(file.path(carry, "gov_one.R"), envir = h)
stopifnot(identical(h$VERSION, g$VERSION))

unlink(c(office, home, again, other, half, carry), recursive = TRUE)
cat("oissyntheticdata: the flow test passed\n")
