## Basic Functions

| Formula                     | Action                                                              |
| --------------------------- | ------------------------------------------------------------------- |
| `=MIN(range)`               | Find the smallest value in `range`                                  |
| `=MAX(range)`               | Find the largest value in `range`                                   |
| `=SUM(range)`               | Sum up all the numbers in `range`                                   |
| `=AVERAGE(range)`           | Calculate the average of all numbers in `range`                     |
| `=COUNT(range)`             | Count the number of cells in `range` that contain _numbers_         |
| `=COUNTA(range)`            | Count the number of cells in `range` that are not empty             |
| `=COUNTIF(range, criteria)` | Count the number of cells in `range` that meet the given `criteria` |

## Logical Functions

- `=AND(condition1, condition2, ...)`
- `=OR(condition1, condition2, ...)`
- `=IF(condition, value if true, value if false)`
- Up to 7 `IF` functions can be nested by replacing the `value if false` with another `IF` function

## Lookup Functions

These functions are useful for finding specific data in a spreadsheet.

| Formula                                                            | Action                                                                                                                 |
| ------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------- |
| `=VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)` | Find `lookup_value` in the first column of `table_array` and return the value in the same row based on `col_index_num` |
| `=HLOOKUP(lookup_value, table_array, row_index_num, range_lookup)` | Find `lookup_value` in the first row of `table_array` and return the value in the same column based on `row_index_num` |
| `=INDEX(array, row_num, col_num)`                                  | Return the value in the cell at the intersection of `row_num` and `col_num`                                            |
| `=MATCH(lookup_value, lookup_array, match_type)`                   | Return the relative position of `lookup_value` in `lookup_array` according to `match_type`                             |

- `range_lookup` can be `TRUE` for an approximate match or `FALSE` for an exact match. If set to `TRUE` (the default setting), data in the first column/row must be sorted in ascending order for `VLOOKUP` or `HLOOKUP` (respectively) to work properly.
- `match_type` can be `1`, `0` or `-1`. Refer to the [documentation](https://support.office.com/en-us/article/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) for the exact behaviour of each.
- `INDEX` and `MATCH` are commonly used together to perform lookups: `MATCH` finds the position of an item in a list and `INDEX` gets the value at that position.

## Statistical Functions

### Division of Data

| Formula                     | Action                                                  |
| --------------------------- | ------------------------------------------------------- |
| `=PERCENTILE.INC(array, k)` | Compute the *k*th percentile, inclusive                 |
| `=PERCENTILE.EXC(array, k)` | Compute the *k*th percentile, exclusive                 |
| `=QUARTILE.INC(array, q)`   | Compute the *q*th quartile, inclusive (_q_ from 0 to 4) |
| `=QUARTILE.EXC(array, q)`   | Compute the *q*th quartile, exclusive (_q_ from 1 to 3) |

### Measures of Location

| Formula                                      | Action                                                           |
| -------------------------------------------- | ---------------------------------------------------------------- |
| `=AVERAGE(data_range)`                       | Compute the mean of all values in `data_range`                   |
| `=MEDIAN(data_range)`                        | Compute the median of all values in `data_range`                 |
| `=MODE.SNGL`                                 | Compute the mode of all values in `data_range` (unimodal case)   |
| `=MODE.MULT`                                 | Compute the mode of all values in `data_range` (multimodal case) |
| `=AVERAGE(MIN(data_range), MAX(data_range))` | Compute the midrange – average of greatest and least values      |

### Measures of Dispersion

| Formula                              | Action                                                                          |
| ------------------------------------ | ------------------------------------------------------------------------------- |
| `=MAX(data_range) - MIN(data_range)` | Compute the range of all values in `data_range`                                 |
| `=QUARTILE.INC(data_range)`          | Compute the interquartile range of all values in `data_range`                   |
| `=VAR.P(data_range)`                 | Compute the variance of all values in `data_range` for a _population_           |
| `=VAR.S(data_range)`                 | Compute the variance of all values in `data_range` for a _sample_               |
| `=STDEV.P(data_range)`               | Compute the standard deviation of all values in `data_range` for a _population_ |
| `=STDEV.S(data_range)`               | Compute the standard deviation of all values in `data_range` for a _sample_     |

### Measures of Shape

| Formula       | Action                                        |
| ------------- | --------------------------------------------- |
| `=SKEW(data)` | Compute the coefficient of skewness of `data` |
| `=KURT(data)` | Compute the coefficient of kurtosis of `data` |

### Probability Distributions

| Formula                                  | Action                                                   |
| ---------------------------------------- | -------------------------------------------------------- |
| `=NORM.DIST(x, mean, stdev, cumulative)` | Compute probability according to the normal distribution |
| `=NORM.INV(probability, mean, stdev)`    | Given a (cumulative) probability _F(x)_, find _x_        |
| `=NORM.S.DIST(z, cumulative)`            | Return the standard normal distribution                  |
