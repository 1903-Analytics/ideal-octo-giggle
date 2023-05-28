  <!-- badges: start -->
  [![Lifecycle: experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
  <!-- badges: end -->

# workbookR

This `library` uses `openxlsx` to generate simple and custom `workbooks` without `java` depedencies. It uses `writeData()` for a greater `OpenOffice` support.

The `library` where specifically developped for the health economic investment model used at `VIVE`. Its repository can be found [here](https://github.com/serkor1/bionic-beaver).

## install

```R
devtools::install_github('https://github.com/1903-Analytics/ideal-octo-giggle')
```

## basic usage

The `syntax` is simple. The `data` to be stored as a `workbook` is given as nested `lists`,

```R
DT <- list(
  sheet1 = list(
    data.table::data.table(
      value = runif(10, 0,1)
    )
  ),
  sheet2 = list(
    data.table::data.table(
      value = runif(10, 0,1)
    )
  )
)
```

Each parent `list` indicates which `sheet` to store the `data` in. To create and store the `workbook`, the `generate_workbook` function is called,

```R
generate_workbook(
  file = 'sampleWorkbook.xlsx',
  list = DT,
  theme = list(
    compact = TRUE,
    combine = FALSE,
    color = 'Reds'
  )
)
```