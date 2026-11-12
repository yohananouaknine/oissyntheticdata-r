\name{gov_one_export}
\alias{gov_one_export}
\title{Copy the Tool Out, to Carry into the Premises}
\description{
Copies \file{gov_one.R} and \file{analysis_template.R} into a folder. These two
files are all that is needed on the premises: no installation and no internet.
}
\usage{
gov_one_export(dir, overwrite = FALSE)
}
\arguments{
  \item{dir}{the folder to copy into. It is created when it does not exist.}
  \item{overwrite}{if \code{TRUE}, a file that is already there with other
    content is replaced. By default it is left alone and a message says so.}
}
\details{
On the premises the tool is run on its own, next to the data files:
\command{Rscript gov_one.R}. An analysis script that starts from
\file{analysis_template.R} loads the \file{gov_one.R} that sits next to it, so
the same script runs at home and on the premises.
}
\value{
The paths of the files written, invisibly: a character vector, empty when both
files were already there.
}
\seealso{
\code{\link{gov_one_file}} for the path of the shipped files.
}
\examples{
carry <- file.path(tempdir(), "carry")
gov_one_export(carry)
list.files(carry)
unlink(carry, recursive = TRUE)
}
