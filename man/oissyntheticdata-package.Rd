\name{oissyntheticdata-package}
\alias{oissyntheticdata-package}
\alias{oissyntheticdata}
\docType{package}
\title{Aggregate Profiles and Synthetic Files for Confidential Data}
\description{
A way to analyse confidential administrative files without taking them out of
the place where they are kept.
}
\details{
The work has four steps.
\enumerate{
  \item \strong{On the premises.} \code{\link{gov_one_profile}} reads every
    'Excel' and CSV file of a folder and writes a folder \file{to_send} that
    holds aggregates only: counts, shares and quantiles. No row, no identifier
    and no value seen fewer than five times is written. The file
    \file{profile_summary.md} lists exactly what leaves, for the person who
    approves it.
  \item \strong{At home.} \code{\link{gov_one_synth}} rebuilds files with the
    same names, sheets, columns, types, distributions and working joins from
    those aggregates. Every value is invented.
  \item \strong{At home.} The analysis script is written on the synthetic
    files, with base R and the helpers of the tool (see \code{\link{gov_one}}).
  \item \strong{Back on the premises.} The same script is run on the real
    files.
}

The tool is the single file \file{gov_one.R}. It uses base R only and works
offline. The package ships that file unchanged, and
\code{\link{gov_one_export}} copies it out, so that it can be carried into a
place where no package can be installed. There the file is run on its own:
\command{Rscript gov_one.R}.

The 'JSON' files are the same as those of the 'Python' package
\pkg{oissyntheticdata}: a profile made with R can be turned into synthetic
files with 'Python', and the other way round.

Results computed on synthetic files test the script. They are not findings.
}
\author{
Yohanan Ouaknine
}
\seealso{
\code{\link{gov_one}}, \code{\link{gov_one_profile}}, \code{\link{gov_one_export}}
}
\keyword{package}
