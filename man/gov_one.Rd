\name{gov_one}
\alias{gov_one}
\alias{gov_one_version}
\alias{gov_one_file}
\title{Load the Tool for an Analysis Script}
\description{
\code{gov_one} loads the file \file{gov_one.R} into a new environment and
returns it. The functions of the tool are then called as \code{g$name(...)},
exactly as in a script that loads the file itself with
\code{g <- new.env(); sys.source("gov_one.R", envir = g)}.
}
\usage{
gov_one()

gov_one_version()

gov_one_file(name = c("gov_one.R", "analysis_template.R"))
}
\arguments{
  \item{name}{the file wanted: the tool, or the starting point for an analysis
    script.}
}
\details{
The main functions of the environment:
\describe{
  \item{\code{open_tables(folder)}}{reads the tables of a folder into a named
    list of data frames. The real files are used when they are there,
    otherwise the synthetic ones. Every column is text and an empty cell is
    \code{""}.}
  \item{\code{to_num}, \code{to_date}, \code{is_missing}}{turn a text column
    into numbers or dates, and mark the empty cells.}
  \item{\code{describe}, \code{crosstab}, \code{count_by},
    \code{monthly_counts}, \code{join_parent}}{descriptive helpers.}
  \item{\code{chi_square}, \code{t_test}, \code{mann_whitney}, \code{pearson},
    \code{spearman}, \code{proportion_test}, \code{ols}, \code{its}}{tests and
    models. They give the same numbers as the helpers of the 'Python' tool.}
  \item{\code{Report}}{builds an offline 'HTML' report. The report carries a
    banner when it was built from synthetic files.}
}
All of base R works on the data frames as well.

A script that must also run on the premises, where the package is not
installed, should load the file itself; see \code{\link{gov_one_export}}.
}
\value{
\code{gov_one} returns an environment holding the functions and settings of
\file{gov_one.R}.

\code{gov_one_version} returns the version line of the shipped file, a
character string.

\code{gov_one_file} returns the full path of the shipped file.
}
\seealso{
\code{\link{gov_one_profile}}, \code{\link{gov_one_export}}
}
\examples{
g <- gov_one()
gov_one_version()

x <- c(12, 15, 9, 22, 18, 30, 11)
y <- c(20, 25, 19, 31, 27, 35, 24)
g$describe(x)$mean
g$t_test(x, y)$p

basename(gov_one_file("analysis_template.R"))
}
