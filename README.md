# oissyntheticdata for R

Analyse confidential administrative files without taking them out of the place where they are kept.

The tool profiles a folder of Excel and CSV files into aggregate-only JSON files, rebuilds synthetic files from those JSON files, and gives your analysis script a small statistics toolbox. You write the script at home on the synthetic files and run it unchanged on the real files, on the premises.

**Results computed on synthetic files test the script. They are not findings.**

This is the R twin of the Python package [`oissyntheticdata`](https://pypi.org/project/oissyntheticdata/) (3.0.0 and later). The JSON files are the same: a profile made with R can be turned into synthetic files with Python, and the other way round.

## Install

```r
install.packages("remotes")
remotes::install_github("yohananouaknine/oissyntheticdata-r")
```

The package needs R 4.2 or later and nothing else: it uses base R only and works offline.

## The four steps

1. **On the premises: the profile.** Reads every Excel and CSV file of the folder and writes a folder `to_send` with counts, shares and quantiles. No row, no identifier and no value seen fewer than five times is written. `profile_summary.md` lists exactly what leaves, for the person who approves it.

   ```r
   library(oissyntheticdata)
   gov_one_profile("path/to/the/data")
   ```

2. **At home: synthetic files.** Rebuilds files with the same names, sheets, columns, types, distributions and working joins. Every value is invented.

   ```r
   gov_one_synth("path/to/the/folder/holding/to_send")
   ```

3. **At home: your script.** Written on the synthetic files, with base R and the helpers of the tool.

   ```r
   g <- gov_one()
   tabs <- g$open_tables("path/to/the/folder")     # real files if present, else the synthetic ones
   cases <- tabs[["cases"]]                        # a data frame; every column is text
   amount <- g$to_num(cases[["AMOUNT"]])
   g$crosstab(cases, "DISTRICT", "CASE_TYPE")
   ```

4. **Back on the premises: the same script on the real files.**

## Where no package can be installed

On the premises there is often no internet and no permission to install anything. The whole tool is one file, `gov_one.R`, and the package can copy it out:

```r
gov_one_export("path/to/a/folder")     # writes gov_one.R and analysis_template.R
```

Carry these two files in. There the tool runs on its own, next to the data files:

```
Rscript gov_one.R
```

and a script that starts from `analysis_template.R` loads the `gov_one.R` next to it, so the same script runs in both places.

## Functions

| Function | What it does |
|---|---|
| `gov_one()` | Loads the tool: `g <- gov_one()`, then `g$open_tables()`, `g$describe()`, `g$crosstab()`, `g$chi_square()`, `g$t_test()`, `g$mann_whitney()`, `g$pearson()`, `g$spearman()`, `g$proportion_test()`, `g$ols()`, `g$its()`, `g$monthly_counts()`, `g$join_parent()`, `g$Report()` |
| `gov_one_profile(data, out, k, quiet)` | The profile of a folder |
| `gov_one_synth(data, out, rows, seed, quiet)` | Synthetic files from a profile |
| `gov_one_check(data)` | Lists the files, sheets and headers that would be read; shows no data |
| `gov_one_settings(data)` | Writes `gov_one_settings.json`, to hide a column or declare a link |
| `gov_one_export(dir, overwrite)` | Copies the two files to carry in |
| `gov_one_file(name)`, `gov_one_version()` | The path and the version line of the shipped file |

Help for each: `?gov_one`, `?gov_one_profile`, `?gov_one_export`.

## What was tested

R 4.6.2 on Windows 11 only. `R CMD check` passes with no error, warning or note. The profile written by R is identical in content to the one written by the Python tool on the same files. Other R versions and other systems were not tested.

## Licence

MIT. An OIS tool: <https://ois.co.il>.
