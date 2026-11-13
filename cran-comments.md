## Submission

This is a new submission.

The version number is 3.0.0 because the package ships the file 'gov_one.R'
3.0.0, the R counterpart of the 'Python' package of the same name (3.0.0 on
PyPI). The two write the same 'JSON' files and keep one version number.

## Test environments

* local: Windows 11, R 4.6.2

## R CMD check results

0 errors | 0 warnings | 1 note

* checking CRAN incoming feasibility ... NOTE
  New submission

## Notes for the reviewer

* The tool is one file, 'inst/gov_one/gov_one.R', written in base R. Its users
  work in offices where no package can be installed, so the whole tool has to
  stay one file that can be carried in and run with 'Rscript gov_one.R'. For
  that reason the code is shipped as a file and loaded by `gov_one()` with
  `sys.source()` into a new environment, and is not split over 'R/'.
  `gov_one_export()` copies the file out.
* The functions write only into the folder the user gives. The examples and
  the test write only under `tempdir()` and remove what they wrote.
* `gov_one_synth()` calls `set.seed()` with its `seed` argument. Synthetic
  files that can be rebuilt exactly are the purpose of that argument; the help
  page says that the state of the random number generator changes.
* The progress lines of `gov_one_profile()` and `gov_one_synth()` are switched
  off with `quiet = TRUE`.
* `quit()` appears once in 'gov_one.R', on the path that is taken only when
  the file is run with 'Rscript'. The functions of the package never reach it:
  an error of the tool becomes an ordinary R error.
