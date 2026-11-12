\name{gov_one_profile}
\alias{gov_one_profile}
\alias{gov_one_synth}
\alias{gov_one_check}
\alias{gov_one_settings}
\title{Profile a Folder, or Build Synthetic Files from a Profile}
\description{
\code{gov_one_profile} reads the 'Excel' and CSV files of a folder and writes
aggregate-only 'JSON' files. \code{gov_one_synth} builds synthetic files from
those 'JSON' files. \code{gov_one_check} lists what would be read, and
\code{gov_one_settings} writes a settings file to edit.
}
\usage{
gov_one_profile(data, out = NULL, k = NULL, quiet = FALSE)

gov_one_synth(data, out = NULL, rows = 1, seed = 12345, quiet = FALSE)

gov_one_check(data)

gov_one_settings(data)
}
\arguments{
  \item{data}{the folder to work on. For \code{gov_one_profile},
    \code{gov_one_check} and \code{gov_one_settings}: the folder with the data
    files. For \code{gov_one_synth}: the folder with the \file{profile_*.json}
    files, or the folder that holds \file{to_send}.}
  \item{out}{the output folder. The default is \file{to_send} inside
    \code{data} for the profile, and \file{synthetic} inside \code{data} for
    the synthetic files.}
  \item{k}{the minimum cell count: a value or a cell seen fewer times than
    this never leaves. The default is 5, or the value in
    \file{gov_one_settings.json}. It cannot be below 3.}
  \item{rows}{the size of the synthetic files relative to the real ones;
    \code{0.1} gives a tenth of the rows.}
  \item{seed}{the random seed. The same profile and the same seed give the
    same synthetic files.}
  \item{quiet}{if \code{TRUE}, the progress lines are not printed.}
}
\details{
These functions run the commands of \file{gov_one.R}; the same commands are
available on the premises as \command{Rscript gov_one.R profile} and so on.

The profile holds counts, shares and quantiles. A column whose header says it
holds a name, an identity number, a phone number, an address or free text is
described by its shape only. Dates keep the month only. The file
\file{profile_summary.md} in the output folder lists everything that was
written, in plain words.

\code{gov_one_synth} calls \code{\link{set.seed}} with \code{seed}, so it
changes the state of the random number generator of the session.

Synthetic files made by R and by the 'Python' tool from the same profile and
seed follow the same distributions, but they are not the same rows.
}
\value{
\code{gov_one_profile} and \code{gov_one_synth} return the path of the output
folder, invisibly. \code{gov_one_settings} returns the path of the settings
file, invisibly. \code{gov_one_check} returns \code{NULL}, invisibly; it is
called for the list it prints.
}
\seealso{
\code{\link{gov_one}} for the analysis script, \code{\link{gov_one_export}} to
copy the tool out.
}
\examples{
## an invented office: one file of cases
office <- file.path(tempdir(), "office")
dir.create(office)
n <- 300
cases <- data.frame(
  CASE_ID = seq_len(n),
  DISTRICT = rep_len(c("North", "Centre", "South", "Centre"), n),
  CASE_TYPE = rep_len(c("A", "B", "A", "C", "A"), n),
  AMOUNT = rep_len(c(120, 250, 90, 400, 310, 75, 180), n),
  OPENED = format(as.Date("2024-01-01") + rep_len(0:364, n)))
write.csv(cases, file.path(office, "cases.csv"), row.names = FALSE)

## on the premises: aggregates only
gov_one_check(office)
to_send <- gov_one_profile(office)
list.files(to_send)

## at home: invented files with the same shape
home <- file.path(tempdir(), "home")
dir.create(home)
file.copy(to_send, home, recursive = TRUE)
gov_one_synth(home, quiet = TRUE)

g <- gov_one()
tabs <- g$open_tables(home)
head(tabs[[1]])
g$crosstab(tabs[[1]], "DISTRICT", "CASE_TYPE")

unlink(c(office, home), recursive = TRUE)
}
