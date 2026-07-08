# Quality Check Documentation

The script includes a validation step to check whether the Clarity CSV exports are processed correctly before the final Excel output is created.

## Purpose

The quality check helps identify potential issues in the input CSV files and confirms whether the data was successfully extracted, merged, and written into the final Excel file.

The final Excel file contains a separate sheet called **Quality_Check** where all validation results are stored.

## Quality Check Sheet Structure

The **Quality_Check** sheet includes the following columns:

| Column      | Description                                   |
| ----------- | --------------------------------------------- |
| check type  | Type of validation performed                  |
| status      | Result of the validation: PASS, WARN, or FAIL |
| message     | Explanation of the result                     |
| source file | Related CSV file, if applicable               |
| Month       | Month of the affected record                  |
| url         | URL of the affected record                    |

## Status Definitions

**PASS** means the validation was successful and no issue was found.

**WARN** means the file was processed, but something should be reviewed. For example, multiple CSV files were merged for the same month and URL, or a scroll depth value is missing.

**FAIL** means the file could not be processed correctly. These rows should be checked first.

## Checks Performed

The script validates whether each CSV file can be read and parsed correctly.

It checks if required metadata fields are available, such as:

* Date range
* Visited URL regex
* Page views
* Metric

It also checks whether the scroll depth table exists and whether the expected scroll depth values are available.

Expected scroll depth values are:

```text
5, 10, 15, ..., 100
```

The script also checks whether multiple CSV files belong to the same **Month + URL** combination. If this happens, the files are merged and the merge is logged in the **Quality_Check** sheet.

## Merge Validation

If one month is split into multiple Clarity exports, the script merges them automatically.

For the **Pageviews** sheet:

```text
total_pageview = sum of total_pageview values from all partial CSV files
```

For the **Scroll** sheet:

```text
number of visitors = sum of number of visitors values for the same Month + URL + scroll depth
```

The `% drop off` column is not included in the final output and is not used in the merge calculation.

## Recommended Review Process

After running the script, open the final Excel file and check the **Quality_Check** sheet.

Start by filtering the `status` column:

1. Check all **FAIL** rows first.
2. Review all **WARN** rows.
3. Confirm that the summary rows match the expected number of processed CSV files.

A successful run should have:

```text
CSV files found = processed files + skipped files
```

Also, the number of scroll rows should usually be close to:

```text
Pageview rows × 20
```

because each URL-month combination is expected to have 20 scroll depth values from 5 to 100.
