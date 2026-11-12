# analysis_template.R -- the starting point for YOUR analysis script.
#
#     Rscript analysis_template.R [folder]          (or open it in R / RStudio and press Source)
#
# Without a folder it works on the folder this script is in.
# The same script runs unchanged in both places:
#   * at home, the folder holds only synthetic files  -> it uses them and stamps the report "synthetic";
#   * on the premises, the folder holds the real files -> it uses those.
# It needs only gov_one.R next to it.  No package to install, no internet.
#
# Parts 1-3 work on any folder and give a first look.  Part 4 is where you write your own statistics;
# the examples there show the helpers of gov_one.R (g$describe, g$crosstab, g$chi_square, g$t_test,
# g$mann_whitney, g$pearson, g$spearman, g$proportion_test, g$ols, g$its, g$monthly_counts,
# g$join_parent).  Everything in base R (table, tapply, aggregate, t.test, lm, glm ...) works as well.

.here <- (function() {                         # the folder of this script, however it was started
  a <- grep("^--file=", commandArgs(FALSE), value = TRUE)
  if (length(a)) return(dirname(normalizePath(gsub("~+~", " ", sub("^--file=", "", a[1]), fixed = TRUE))))
  f <- tryCatch(sys.frame(1)$ofile, error = function(e) NULL)      # source("analysis_template.R") in R or RStudio
  if (is.null(f)) "." else dirname(normalizePath(f))
})()
g <- new.env()
sys.source(file.path(.here, "gov_one.R"), envir = g)

args <- commandArgs(trailingOnly = TRUE)
FOLDER <- if (length(args)) args[1] else .here         # no folder given: the folder this script is in
tabs <- g$open_tables(FOLDER)                 # list(table name = data frame of text columns); real files win over synthetic ones
page <- g$Report("First look at the data")

# ---- 1. what is there ---------------------------------------------------------------------------
page$h("Tables")
page$table(c("Table", "File", "Rows", "Columns"),
           lapply(names(tabs), function(name) list(name, attr(tabs[[name]], "file"), nrow(tabs[[name]]), ncol(tabs[[name]]))))

# ---- 2. every column of every table -------------------------------------------------------------
# Values are listed only for columns with few distinct values, and never below 5 rows, so this
# page does not show a person even when it is built from real files.
for (name in names(tabs)) {
  df <- tabs[[name]]
  page$h(name)
  lines <- list()
  first_date <- NULL
  for (col in names(df)) {
    vals <- df[[col]][!g$is_missing(df[[col]])]
    nums <- g$to_num(vals); nums <- nums[!is.na(nums)]
    dates <- g$to_date(vals); dates <- dates[!is.na(dates)]
    miss <- if (nrow(df)) sprintf("%.0f%%", 100 * (nrow(df) - length(vals)) / nrow(df)) else ""
    distinct <- length(unique(vals))
    if (!length(vals)) {
      kind <- "empty"; detail <- ""
    } else if (length(dates) >= 0.9 * length(vals)) {
      kind <- "date"; detail <- sprintf("%s to %s", format(min(dates), "%Y-%m"), format(max(dates), "%Y-%m"))
      if (is.null(first_date)) first_date <- col
    } else if (length(nums) >= 0.9 * length(vals) && distinct > 12) {
      d <- g$describe(nums)
      kind <- "number"
      detail <- sprintf("mean %s, median %s, middle half %s to %s", g$fmt(d$mean), g$fmt(d$median), g$fmt(d$p25), g$fmt(d$p75))
    } else if (distinct <= 12) {
      top <- g$count_by(data.frame(v = vals, stringsAsFactors = FALSE), "v")
      top <- utils::head(top[top$n >= 5, ], 6)
      kind <- "category"; detail <- paste(sprintf("%s (%s)", top$v, vapply(top$n, g$fmt, "")), collapse = ", ")
    } else {
      kind <- "text / id"; detail <- sprintf("%s distinct values", g$fmt(distinct))
    }
    lines[[length(lines) + 1]] <- list(col, kind, miss, detail)
  }
  page$table(c("Column", "Kind", "Empty", "Summary"), lines)

  # ---- 3. rows per month for the first date column --------------------------------------------
  if (!is.null(first_date)) {
    mc <- g$monthly_counts(df, first_date)
    if (length(mc$months) >= 6)
      page$line(mc$months, list("rows per month" = mc$series$all), title = sprintf("%s: rows per month by %s", name, first_date))
  }
}

# ---- 4. YOUR ANALYSIS ---------------------------------------------------------------------------
# Replace the names in capitals with your own table and column names, and delete what you do not need.
# Every column is text: use g$to_num() for numbers and g$to_date() for dates.
#
#   cases <- tabs[["CASES"]]                                   # a data frame; cases[["COLUMN"]] is text
#
#   # join a table to its parent (the link is listed in to_send/profile_summary.md)
#   cases <- g$join_parent(cases, tabs[["CLIENTS"]], on = "CLIENT ID")   # every case row now also has the client's columns
#
#   # two categories: are they related?
#   counts <- g$crosstab(cases, "DISTRICT", "CASE TYPE")       # a matrix of counts
#   test <- g$chi_square(counts)                               # the same numbers as chisq.test(counts, correct = FALSE)
#   page$h("Case type by district")
#   page$table(c("district", colnames(counts)), lapply(rownames(counts), function(a) c(list(a), as.list(counts[a, ]))))
#   page$p(sprintf("chi-square = %s, df = %d, p %s, Cramer V = %s", g$fmt(test$chi2), test$df, g$fmt_p(test$p), g$fmt(test$cramers_v)))
#
#   # a number in two groups
#   amount <- g$to_num(cases[["AMOUNT"]])
#   a <- amount[cases[["GENDER"]] == "A" & !is.na(amount)]
#   b <- amount[cases[["GENDER"]] == "B" & !is.na(amount)]
#   test <- g$t_test(a, b)           # Welch, as t.test(a, b); or g$mann_whitney(a, b) when the numbers are very skewed
#   page$p(sprintf("difference %s (95%% interval %s to %s), p %s", g$fmt(test$diff), g$fmt(test$ci95[1]), g$fmt(test$ci95[2]), g$fmt_p(test$p)))
#
#   # before / after a date, month by month (interrupted time series)
#   mc <- g$monthly_counts(cases, "OPENED")
#   k <- match("2024-01", mc$months)                           # the first month "after"
#   fit <- g$its(mc$series$all, k, hac_lags = 6)
#   page$line(mc$months, list(cases = mc$series$all, model = fit$fitted), title = "Cases per month", mark = "2024-01")
#   lc <- fit$level_change
#   page$p(sprintf("change in level at 2024-01: %s (95%% interval %s to %s), p %s", g$fmt(lc[["estimate"]]), g$fmt(lc[["low"]]), g$fmt(lc[["high"]]), g$fmt_p(lc[["p"]])))
#
#   # a regression with several predictors
#   fit <- g$ols(y, cbind(x1 = x1, x2 = x2))                   # the same coefficients and tests as lm(y ~ x1 + x2)
#   page$table(c("term", "b", "se", "t", "p"), lapply(seq_along(fit$names), function(i) list(fit$names[i], fit$coef[i], fit$se[i], fit$t[i], g$fmt_p(fit$p[i]))))

page$save(file.path(FOLDER, "results", "report.html"))
