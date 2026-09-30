// Worksheet function metadata for formula autocomplete and argument tooltips.
// Pure data: this module deliberately does not import the formula engine so the UI can load it
// cheaply. Every function the engine registers (formulas.ts + formula-library.ts) has an entry;
// scripts/qa-formula-lib-catalog.ts checks that the two stay in sync.
//
// Entries are stored as compact tuples and parsed once at module load:
//   [NAME, category, description, args, returnsArray?]
// `args` is "name: description | [optional]: description | [repeating]...: description".
// A name in [brackets] is optional; a trailing "..." marks an argument that may repeat.

export interface FunctionArgumentInfo {
  name: string;
  description: string;
  optional?: boolean;
  repeating?: boolean;
}

export type FunctionCategory =
  | "Math"
  | "Statistical"
  | "Text"
  | "Logical"
  | "Lookup"
  | "Date"
  | "Financial"
  | "Information"
  | "Engineering"
  | "Database"
  | "Array"
  | "Web"
  | "Operator";

export interface FunctionInfo {
  name: string;
  category: FunctionCategory | string;
  description: string;
  args: FunctionArgumentInfo[];
  returnsArray?: boolean;
}

export const FUNCTION_CATEGORIES: readonly FunctionCategory[] = [
  "Math",
  "Statistical",
  "Text",
  "Logical",
  "Lookup",
  "Date",
  "Financial",
  "Information",
  "Engineering",
  "Database",
  "Array",
  "Web",
  "Operator",
];

const MATH: FunctionCategory = "Math";
const STAT: FunctionCategory = "Statistical";
const TEXT: FunctionCategory = "Text";
const LOGIC: FunctionCategory = "Logical";
const LOOK: FunctionCategory = "Lookup";
const DATE: FunctionCategory = "Date";
const FIN: FunctionCategory = "Financial";
const INFO: FunctionCategory = "Information";
const ENG: FunctionCategory = "Engineering";
const DB: FunctionCategory = "Database";
const ARR: FunctionCategory = "Array";
const WEB: FunctionCategory = "Web";
const OP: FunctionCategory = "Operator";

type Row = readonly [name: string, category: FunctionCategory, description: string, args: string, returnsArray?: 1];

// ---- Shared argument lists -------------------------------------------------------------------

const NUMBERS = "number1: The first number, cell reference, or range | [number2]...: Additional numbers, cell references, or ranges";
const VALUES = "value1: The first value, cell reference, or range | [value2]...: Additional values, cell references, or ranges";
const LOGICALS = "logical1: The first condition that evaluates to TRUE or FALSE | [logical2]...: Additional conditions to test";
const TEXTS = "text1: The first text item, cell reference, or range | [text2]...: Additional text items to join";
const NUM = (what: string) => `number: ${what}`;
const ANGLE = "number: The angle in radians";
const ONE_VALUE = "value: The value to test";
const TWO_ARRAYS = "array1: The first range or array of values | array2: The second range or array of values";
const KNOWN_YX = "known_y's: The dependent data points | known_x's: The independent data points";
const XY_ARRAYS = "array_x: The first array or range of values | array_y: The second array or range of values";
const CRITERIA_PAIRS =
  "criteria_range1: The first range to evaluate | criteria1: The condition applied to criteria_range1 | [criteria_range2]...: Additional ranges to evaluate | [criteria2]...: Conditions for the additional ranges";
const DATABASE =
  "database: The range that makes up the list or database, including its header row | field: The column to use, as a quoted column label or a column position number | criteria: The range that holds the conditions, including a header row";
const INUMBER = "inumber: A complex number as text in the form x+yi or x+yj";
const BASE_PLACES = "[places]: The number of characters to use, padding with leading zeros";
const SETTLE = "settlement: The security's settlement date";
const MATURE = "maturity: The security's maturity date";
const FREQUENCY = "frequency: The number of coupon payments per year (1 = annual, 2 = semiannual, 4 = quarterly)";
const BASIS = "[basis]: The day count basis (0 = US 30/360, 1 = actual/actual, 2 = actual/360, 3 = actual/365, 4 = European 30/360)";
const COUPON_ARGS = `${SETTLE} | ${MATURE} | ${FREQUENCY} | ${BASIS}`;
const RATE_ARG = "rate: The interest rate per period";
const TYPE_ARG = "[type]: When payments are due (0 = end of period, 1 = beginning)";
const LAMBDA_ARG = "lambda: A LAMBDA that is called for each item";
const DIST_CUMULATIVE = "cumulative: TRUE for the cumulative distribution function, FALSE for the probability density or mass function";
const PROBABILITY = "probability: A probability between 0 and 1";
const DF = "deg_freedom: The number of degrees of freedom";
const F_DF = "deg_freedom1: The numerator degrees of freedom | deg_freedom2: The denominator degrees of freedom";
const NORMAL_PARAMS = "mean: The arithmetic mean of the distribution | standard_dev: The standard deviation of the distribution";
const LOGNORMAL_PARAMS = "mean: The mean of ln(x) | standard_dev: The standard deviation of ln(x)";
const GAMMA_PARAMS = "alpha: The shape parameter of the distribution | beta: The scale parameter of the distribution";
const BETA_PARAMS =
  "alpha: A shape parameter of the distribution | beta: A shape parameter of the distribution";
const BETA_BOUNDS = "[A]: The lower bound of the interval of x (default 0) | [B]: The upper bound of the interval of x (default 1)";
const BINOM_ARGS =
  "number_s: The number of successes in trials | trials: The number of independent trials | probability_s: The probability of success on each trial";
const HYPGEOM_ARGS =
  "sample_s: The number of successes in the sample | number_sample: The size of the sample | population_s: The number of successes in the population | number_pop: The population size";
const NEGBINOM_ARGS =
  "number_f: The number of failures | number_s: The threshold number of successes | probability_s: The probability of a success";
const TTEST_ARGS =
  "array1: The first data set | array2: The second data set | tails: The number of distribution tails (1 or 2) | type: The kind of t-test (1 = paired, 2 = equal variance, 3 = unequal variance)";
const CONF_ARGS =
  "alpha: The significance level used to compute the confidence level | standard_dev: The population standard deviation | size: The sample size";
const ZTEST_ARGS =
  "array: The array or range of data to test x against | x: The value to test | [sigma]: The population standard deviation, if known";
const CHISQ_TEST_ARGS =
  "actual_range: The range of observed values | expected_range: The range of expected values";
const TREND_ARGS =
  "known_y's: The y-values you already know | [known_x's]: The x-values you already know | [new_x's]: The new x-values to return predictions for | [const]: FALSE to force the constant b to 0";
const LINEST_ARGS =
  "known_y's: The y-values you already know | [known_x's]: The x-values you already know | [const]: FALSE to force the constant b to 0 | [stats]: TRUE to return additional regression statistics";
const WORKDAY_WEEKEND =
  "[weekend]: Which days are weekends, as a weekend number (1-7, 11-17) or a seven-character string of 0s and 1s starting Monday";
const DEPRECIATION = "cost: The initial cost of the asset | salvage: The value at the end of the depreciation | life: The number of periods over which the asset is depreciated";
const REGEX_CASE = "[case_sensitivity]: 0 for case-sensitive matching (default), 1 for case-insensitive";
const OPERANDS = "value1: The first operand | value2: The second operand";
const TRIG_H = (name: string) => `number: The number whose hyperbolic ${name} you want`;

// ---- Catalog --------------------------------------------------------------------------------

const ROWS: readonly Row[] = [
  // Logical
  ["IF", LOGIC, "Checks whether a condition is met, and returns one value if TRUE and another if FALSE.", "logical_test: The condition to test | value_if_true: The value returned when logical_test is TRUE | [value_if_false]: The value returned when logical_test is FALSE"],
  ["IFERROR", LOGIC, "Returns value_if_error if an expression evaluates to an error; otherwise returns the value of the expression.", "value: The value or expression to check for an error | value_if_error: The value returned when value is an error"],
  ["IFNA", LOGIC, "Returns the value you specify if an expression resolves to #N/A; otherwise returns the result of the expression.", "value: The value or expression to check for #N/A | value_if_na: The value returned when value is #N/A"],
  ["IFS", LOGIC, "Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.", "logical_test1: The first condition to test | value_if_true1: The result when logical_test1 is TRUE | [logical_test2]...: Additional conditions to test | [value_if_true2]...: Results for the additional conditions"],
  ["SWITCH", LOGIC, "Evaluates an expression against a list of values and returns the result corresponding to the first matching value.", "expression: The value to compare against the list | value1: The first value to compare with expression | result1: The result returned when value1 matches | [default_or_value2]...: A default result, or another value to compare | [result2]...: Results for the additional values"],
  ["AND", LOGIC, "Returns TRUE if all of its arguments are TRUE.", LOGICALS],
  ["OR", LOGIC, "Returns TRUE if any argument is TRUE.", LOGICALS],
  ["XOR", LOGIC, "Returns a logical exclusive OR of all arguments.", LOGICALS],
  ["NOT", LOGIC, "Reverses the logic of its argument.", "logical: A value or expression that evaluates to TRUE or FALSE"],
  ["TRUE", LOGIC, "Returns the logical value TRUE.", ""],
  ["FALSE", LOGIC, "Returns the logical value FALSE.", ""],
  ["LET", LOGIC, "Assigns names to calculation results so intermediate values can be reused inside a formula.", "name1: The first name to assign | name_value1: The value assigned to name1 | calculation_or_name2: A calculation using the names, or another name to assign | [name_value2]...: Values for additional names | [calculation_or_name3]...: The final calculation, or additional names"],
  ["LAMBDA", LOGIC, "Creates a custom, reusable function from parameters and a calculation.", "[parameter1]...: A name for a value passed to the function | calculation: The formula to evaluate, using the parameters"],
  ["ISOMITTED", LOGIC, "Checks whether a LAMBDA argument is missing and returns TRUE or FALSE.", "argument: The LAMBDA parameter to test"],
  ["ISBETWEEN", LOGIC, "Checks whether a value lies between two other values, with optional inclusive bounds.", "value_to_compare: The value to test | lower_value: The lower bound | upper_value: The upper bound | [lower_value_is_inclusive]: Whether the range includes lower_value (default TRUE) | [upper_value_is_inclusive]: Whether the range includes upper_value (default TRUE)"],

  // Array helpers and dynamic arrays
  ["MAP", ARR, "Returns an array formed by mapping each value in the arrays to a new value by applying a LAMBDA.", `array1: The first array to map | [array2]...: Additional arrays to map | ${LAMBDA_ARG}`, 1],
  ["REDUCE", ARR, "Reduces an array to an accumulated value by applying a LAMBDA to each value.", `[initial_value]: The starting value for the accumulator | array: The array to reduce | lambda: A LAMBDA called with the accumulator and each value`],
  ["SCAN", ARR, "Scans an array by applying a LAMBDA to each value and returns an array of each intermediate value.", `[initial_value]: The starting value for the accumulator | array: The array to scan | lambda: A LAMBDA called with the accumulator and each value`, 1],
  ["BYROW", ARR, "Applies a LAMBDA to each row and returns an array of the results.", `array: The array to process by row | lambda: A LAMBDA that takes a row as its parameter`, 1],
  ["BYCOL", ARR, "Applies a LAMBDA to each column and returns an array of the results.", `array: The array to process by column | lambda: A LAMBDA that takes a column as its parameter`, 1],
  ["MAKEARRAY", ARR, "Returns a calculated array of a specified row and column size by applying a LAMBDA.", "rows: The number of rows in the array | cols: The number of columns in the array | lambda: A LAMBDA called with the row and column index of each item", 1],
  ["FILTER", ARR, "Filters a range of data based on criteria you define.", "array: The array or range to filter | include: An array of TRUE/FALSE values the same height or width as array | [if_empty]: The value to return if nothing is included", 1],
  ["SORT", ARR, "Sorts the contents of a range or array.", "array: The range or array to sort | [sort_index]: The row or column to sort by | [sort_order]: 1 for ascending (default), -1 for descending | [by_col]: FALSE to sort by row (default), TRUE to sort by column", 1],
  ["SORTBY", ARR, "Sorts the contents of a range or array based on the values in a corresponding range or array.", "array: The range or array to sort | by_array1: The range or array to sort by | [sort_order1]: 1 for ascending (default), -1 for descending | [by_array2]...: Additional ranges or arrays to sort by | [sort_order2]...: Sort orders for the additional arrays", 1],
  ["SORTN", ARR, "Returns the first n items of a data set after sorting.", "range: The data to sort | [n]: The number of items to return (default 1) | [display_ties_mode]: How to handle ties (0-3) | [sort_column1]...: The column index or range to sort by | [is_ascending1]...: TRUE to sort ascending", 1],
  ["UNIQUE", ARR, "Returns a list of unique values in a list or range.", "array: The range or array from which to return unique rows or columns | [by_col]: TRUE to compare columns, FALSE to compare rows (default) | [exactly_once]: TRUE to return only values that occur exactly once", 1],
  ["SEQUENCE", ARR, "Generates a list of sequential numbers in an array.", "rows: The number of rows to return | [columns]: The number of columns to return | [start]: The first number in the sequence | [step]: The amount to increment each value", 1],
  ["TRANSPOSE", ARR, "Returns the transpose of an array, turning rows into columns and columns into rows.", "array: The array or range to transpose", 1],
  ["CHOOSECOLS", ARR, "Returns the specified columns from an array.", "array: The array containing the columns | col_num1: The first column to return | [col_num2]...: Additional columns to return", 1],
  ["CHOOSEROWS", ARR, "Returns the specified rows from an array.", "array: The array containing the rows | row_num1: The first row to return | [row_num2]...: Additional rows to return", 1],
  ["TAKE", ARR, "Returns a specified number of contiguous rows or columns from the start or end of an array.", "array: The array to take rows or columns from | rows: The number of rows to take; negative values take from the end | [columns]: The number of columns to take; negative values take from the end", 1],
  ["DROP", ARR, "Excludes a specified number of rows or columns from the start or end of an array.", "array: The array to drop rows or columns from | rows: The number of rows to drop; negative values drop from the end | [columns]: The number of columns to drop; negative values drop from the end", 1],
  ["VSTACK", ARR, "Appends arrays vertically and in sequence to return a larger array.", "array1: The first array to stack | [array2]...: Additional arrays to append below", 1],
  ["HSTACK", ARR, "Appends arrays horizontally and in sequence to return a larger array.", "array1: The first array to stack | [array2]...: Additional arrays to append to the right", 1],
  ["TOCOL", ARR, "Returns the array in a single column.", "array: The array or reference to return as a column | [ignore]: Values to skip (0 = none, 1 = blanks, 2 = errors, 3 = blanks and errors) | [scan_by_column]: TRUE to scan the array by column", 1],
  ["TOROW", ARR, "Returns the array in a single row.", "array: The array or reference to return as a row | [ignore]: Values to skip (0 = none, 1 = blanks, 2 = errors, 3 = blanks and errors) | [scan_by_column]: TRUE to scan the array by column", 1],
  ["WRAPROWS", ARR, "Wraps a row or column of values by rows after a specified number of elements.", "vector: The row or column to wrap | wrap_count: The maximum number of values in each row | [pad_with]: The value used to pad the last row (default #N/A)", 1],
  ["WRAPCOLS", ARR, "Wraps a row or column of values by columns after a specified number of elements.", "vector: The row or column to wrap | wrap_count: The maximum number of values in each column | [pad_with]: The value used to pad the last column (default #N/A)", 1],
  ["EXPAND", ARR, "Expands or pads an array to specified row and column dimensions.", "array: The array to expand | rows: The number of rows in the expanded array | [columns]: The number of columns in the expanded array | [pad_with]: The value used to pad (default #N/A)", 1],
  ["TRIMRANGE", ARR, "Excludes empty rows and columns from the outer edges of a range or array.", "range: The range or array to trim | [trim_rows]: Which empty rows to trim (0 = none, 1 = leading, 2 = trailing, 3 = both) | [trim_cols]: Which empty columns to trim (0 = none, 1 = leading, 2 = trailing, 3 = both)", 1],
  ["GROUPBY", ARR, "Groups the values of a table by one or more row fields and aggregates them.", "row_fields: The values to group by | values: The values to aggregate | function: The aggregation function, such as SUM or a LAMBDA | [field_headers]: Whether the data has headers and whether to show them | [total_depth]: Whether to show totals (0 = none, 1 = grand totals) | [sort_order]: The column index to sort by; negative for descending | [filter_array]: A TRUE/FALSE array of rows to include | [field_relationship]: 0 for hierarchy, 1 for table", 1],
  ["PIVOTBY", ARR, "Groups values by row and column fields and aggregates them into a summary table.", "row_fields: The values to group rows by | col_fields: The values to group columns by | values: The values to aggregate | function: The aggregation function, such as SUM or a LAMBDA | [field_headers]: Whether the data has headers and whether to show them | [row_total_depth]: Whether to show row totals | [row_sort_order]: How to sort rows | [col_total_depth]: Whether to show column totals | [col_sort_order]: How to sort columns | [filter_array]: A TRUE/FALSE array of rows to include | [relative_to]: What PERCENTOF-style calculations are relative to", 1],
  ["FLATTEN", ARR, "Flattens one or more ranges into a single column.", "range1: The first range to flatten | [range2]...: Additional ranges to flatten", 1],
  ["SPARKLINE", INFO, "Creates a miniature chart inside a single cell: line (default), column, winloss or bar.", "data: The range or array of values to plot | [options]: A two-column range or array of settings, such as {\"charttype\",\"column\";\"color\",\"red\"}"],
  ["ARRAYFORMULA", ARR, "Enables the display of values returned from an array formula into multiple rows and/or columns.", "array_formula: A range, a formula using ranges, or a function that returns more than one value", 1],
  ["ARRAY_CONSTRAIN", ARR, "Constrains an array result to a specified size.", "input_range: The range or array to constrain | num_rows: The number of rows to keep | num_cols: The number of columns to keep", 1],

  // Lookup and reference
  ["VLOOKUP", LOOK, "Looks for a value in the leftmost column of a table, and returns a value in the same row from a column you specify.", "lookup_value: The value to find in the first column of table_array | table_array: The table of information to search | col_index_num: The column number in table_array to return a value from | [range_lookup]: FALSE for an exact match, TRUE or omitted for an approximate match"],
  ["HLOOKUP", LOOK, "Looks for a value in the top row of a table, and returns a value in the same column from a row you specify.", "lookup_value: The value to find in the first row of table_array | table_array: The table of information to search | row_index_num: The row number in table_array to return a value from | [range_lookup]: FALSE for an exact match, TRUE or omitted for an approximate match"],
  ["XLOOKUP", LOOK, "Searches a range or array for a match and returns the corresponding item from a second range or array.", "lookup_value: The value to search for | lookup_array: The array or range to search | return_array: The array or range to return | [if_not_found]: The value returned when no match is found | [match_mode]: 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard | [search_mode]: 1 = first to last (default), -1 = last to first, 2 = binary ascending, -2 = binary descending"],
  ["XMATCH", LOOK, "Returns the relative position of an item in an array or range.", "lookup_value: The value to search for | lookup_array: The array or range to search | [match_mode]: 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard, 3 = regular expression | [search_mode]: 1 = first to last (default), -1 = last to first, 2 = binary ascending, -2 = binary descending"],
  ["MATCH", LOOK, "Returns the relative position of an item in an array that matches a specified value in a specified order.", "lookup_value: The value to match in lookup_array | lookup_array: The range of cells being searched | [match_type]: 1 = largest value less than or equal (default), 0 = exact match, -1 = smallest value greater than or equal"],
  ["INDEX", LOOK, "Returns a value or reference of the cell at the intersection of a particular row and column in a given range.", "array: A range of cells or an array constant | row_num: The row in array from which to return a value | [column_num]: The column in array from which to return a value"],
  ["LOOKUP", LOOK, "Looks up a value either from a one-row or one-column range or from an array.", "lookup_value: The value to search for | lookup_vector: A one-row or one-column range sorted in ascending order, or an array | [result_vector]: A one-row or one-column range the same size as lookup_vector"],
  ["CHOOSE", LOOK, "Chooses a value or action to perform from a list of values, based on an index number.", "index_num: Which value argument is selected (1 to 254) | value1: The first value to choose from | [value2]...: Additional values to choose from"],
  ["OFFSET", LOOK, "Returns a reference to a range that is a given number of rows and columns from a given reference.", "reference: The reference from which to base the offset | rows: The number of rows up or down to move | cols: The number of columns left or right to move | [height]: The height, in rows, of the returned reference | [width]: The width, in columns, of the returned reference"],
  ["INDIRECT", LOOK, "Returns the reference specified by a text string.", "ref_text: A text reference to a cell, such as \"A1\" or \"Sheet2!B3\" | [a1]: TRUE for A1-style references (default), FALSE for R1C1-style"],
  ["ROW", LOOK, "Returns the row number of a reference.", "[reference]: The cell or range whose row number you want; defaults to the formula's cell"],
  ["COLUMN", LOOK, "Returns the column number of a reference.", "[reference]: The cell or range whose column number you want; defaults to the formula's cell"],
  ["ROWS", LOOK, "Returns the number of rows in a reference or array.", "array: An array, array formula, or range reference"],
  ["COLUMNS", LOOK, "Returns the number of columns in a reference or array.", "array: An array, array formula, or range reference"],
  ["ADDRESS", LOOK, "Creates a cell reference as text, given specified row and column numbers.", "row_num: The row number to use in the reference | column_num: The column number to use in the reference | [abs_num]: The reference type (1 = absolute, 2 = absolute row, 3 = absolute column, 4 = relative) | [a1]: TRUE for A1-style (default), FALSE for R1C1-style | [sheet_text]: The name of the sheet to include"],
  ["AREAS", LOOK, "Returns the number of areas in a reference.", "reference: A reference to a cell or range of cells"],
  ["FORMULATEXT", LOOK, "Returns the formula at the given reference as text.", "reference: A reference to a cell"],
  ["HYPERLINK", LOOK, "Creates a shortcut that opens a document or web page; the cell shows the friendly name.", "link_location: The path or URL of the document to open | [friendly_name]: The text or value displayed in the cell"],
  ["ANCHORARRAY", LOOK, "Returns the entire spilled range of a dynamic array formula (the A1# spill reference).", "reference: The anchor cell of a spilled array"],

  // Information
  ["ISBLANK", INFO, "Checks whether a reference is to an empty cell, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISERR", INFO, "Checks whether a value is an error other than #N/A, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISERROR", INFO, "Checks whether a value is an error, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISNA", INFO, "Checks whether a value is #N/A, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISNUMBER", INFO, "Checks whether a value is a number, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISTEXT", INFO, "Checks whether a value is text, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISNONTEXT", INFO, "Checks whether a value is not text, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISLOGICAL", INFO, "Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.", ONE_VALUE],
  ["ISFORMULA", INFO, "Checks whether a reference is to a cell that contains a formula, and returns TRUE or FALSE.", "reference: A reference to the cell to test"],
  ["ISREF", INFO, "Checks whether a value is a reference, and returns TRUE or FALSE.", ONE_VALUE],
  ["ISEVEN", INFO, "Returns TRUE if the number is even.", NUM("The value to test; non-integers are truncated")],
  ["ISODD", INFO, "Returns TRUE if the number is odd.", NUM("The value to test; non-integers are truncated")],
  ["ISEMAIL", INFO, "Checks whether a value is a valid email address.", "value: The value to verify as an email address"],
  ["ISURL", INFO, "Checks whether a value is a valid URL.", "value: The value to verify as a URL"],
  ["TYPE", INFO, "Returns an integer representing the data type of a value: number = 1, text = 2, logical = 4, error = 16, array = 64.", "value: Any value"],
  ["ERROR.TYPE", INFO, "Returns a number corresponding to an error value.", "error_val: The error value whose identifying number you want"],
  ["N", INFO, "Converts a value to a number: numbers stay numbers, dates become serial numbers, TRUE becomes 1, and anything else becomes 0.", "value: The value you want converted"],
  ["NA", INFO, "Returns the error value #N/A.", ""],
  ["INFO", INFO, "Returns information about the current operating environment.", "type_text: Text specifying the type of information to return, such as 'osversion', 'system', or 'release'"],
  ["CELL", INFO, "Returns information about the formatting, location, or contents of a cell.", "info_type: A text value that specifies the type of cell information, such as 'address', 'row', 'col', 'contents', or 'type' | [reference]: The cell you want information about"],
  ["SHEET", INFO, "Returns the sheet number of the referenced sheet.", "[value]: The name of a sheet or a reference; defaults to the sheet containing the formula"],
  ["SHEETS", INFO, "Returns the number of sheets in a reference.", "[reference]: A reference for which you want the number of sheets; defaults to the whole workbook"],
  ["SINGLE", INFO, "Returns a single value using implicit intersection (the @ operator).", "value: The value, range, or array to reduce to a single value"],

  // Math and trigonometry
  ["SUM", MATH, "Adds all the numbers in a range of cells.", NUMBERS],
  ["SUMIF", MATH, "Adds the cells specified by a given condition or criteria.", "range: The range of cells you want evaluated by criteria | criteria: The condition that defines which cells are added | [sum_range]: The cells to add, if different from range"],
  ["SUMIFS", MATH, "Adds the cells specified by a given set of conditions or criteria.", `sum_range: The cells to sum | ${CRITERIA_PAIRS}`],
  ["SUMPRODUCT", MATH, "Returns the sum of the products of corresponding ranges or arrays.", "array1: The first array whose components you want to multiply and then add | [array2]...: Additional arrays of the same size"],
  ["SUMSQ", MATH, "Returns the sum of the squares of the arguments.", NUMBERS],
  ["SUMX2MY2", MATH, "Returns the sum of the difference of squares of corresponding values in two arrays.", XY_ARRAYS],
  ["SUMX2PY2", MATH, "Returns the sum of the sum of squares of corresponding values in two arrays.", XY_ARRAYS],
  ["SUMXMY2", MATH, "Returns the sum of squares of differences of corresponding values in two arrays.", XY_ARRAYS],
  ["PRODUCT", MATH, "Multiplies all the numbers given as arguments.", NUMBERS],
  ["SUBTOTAL", MATH, "Returns a subtotal in a list or database, ignoring other subtotals and optionally hidden rows.", "function_num: A number 1-11 or 101-111 that specifies the aggregate function | ref1: The first range or reference to subtotal | [ref2]...: Additional ranges or references"],
  ["AGGREGATE", MATH, "Returns an aggregate in a list or database, with options to ignore hidden rows and error values.", "function_num: A number 1-19 that specifies the aggregate function | options: A number 0-7 that specifies which values to ignore | ref1: The first range, array, or numeric argument | [ref2]...: Additional ranges, or the k value for LARGE, SMALL, PERCENTILE, and QUARTILE"],
  ["ABS", MATH, "Returns the absolute value of a number, a number without its sign.", NUM("The real number whose absolute value you want")],
  ["SIGN", MATH, "Returns the sign of a number: 1 if positive, 0 if zero, or -1 if negative.", NUM("Any real number")],
  ["INT", MATH, "Rounds a number down to the nearest integer.", NUM("The real number to round down to an integer")],
  ["TRUNC", MATH, "Truncates a number to an integer by removing the decimal, or fractional, part.", "number: The number to truncate | [num_digits]: The precision of the truncation (default 0)"],
  ["MOD", MATH, "Returns the remainder after a number is divided by a divisor.", "number: The number for which you want the remainder | divisor: The number by which to divide number"],
  ["QUOTIENT", MATH, "Returns the integer portion of a division.", "numerator: The dividend | denominator: The divisor"],
  ["POWER", MATH, "Returns the result of a number raised to a power.", "number: The base number | power: The exponent to raise the base number to"],
  ["SQRT", MATH, "Returns the square root of a number.", NUM("The number whose square root you want")],
  ["SQRTPI", MATH, "Returns the square root of (number * pi).", NUM("The number by which pi is multiplied")],
  ["EXP", MATH, "Returns e raised to the power of a given number.", NUM("The exponent applied to the base e")],
  ["LN", MATH, "Returns the natural logarithm of a number.", NUM("The positive real number whose natural logarithm you want")],
  ["LOG", MATH, "Returns the logarithm of a number to the base you specify.", "number: The positive real number whose logarithm you want | [base]: The base of the logarithm (default 10)"],
  ["LOG10", MATH, "Returns the base-10 logarithm of a number.", NUM("The positive real number whose base-10 logarithm you want")],
  ["PI", MATH, "Returns the value of pi, 3.14159265358979, accurate to 15 digits.", ""],
  ["ROUND", MATH, "Rounds a number to a specified number of digits.", "number: The number to round | num_digits: The number of digits to round to; negative values round to the left of the decimal point"],
  ["ROUNDUP", MATH, "Rounds a number up, away from zero.", "number: The number to round up | num_digits: The number of digits to round to"],
  ["ROUNDDOWN", MATH, "Rounds a number down, toward zero.", "number: The number to round down | num_digits: The number of digits to round to"],
  ["MROUND", MATH, "Returns a number rounded to the desired multiple.", "number: The value to round | multiple: The multiple to which to round number"],
  ["CEILING", MATH, "Rounds a number up to the nearest multiple of significance.", "number: The value to round | significance: The multiple to which to round"],
  ["CEILING.MATH", MATH, "Rounds a number up to the nearest integer or to the nearest multiple of significance.", "number: The value to round | [significance]: The multiple to which to round (default 1) | [mode]: For negative numbers, a nonzero value rounds away from zero"],
  ["CEILING.PRECISE", MATH, "Rounds a number up to the nearest integer or multiple of significance, regardless of the sign of the number.", "number: The value to round | [significance]: The multiple to which to round (default 1)"],
  ["ISO.CEILING", MATH, "Rounds a number up to the nearest integer or multiple of significance, regardless of the sign of the number.", "number: The value to round | [significance]: The multiple to which to round (default 1)"],
  ["FLOOR", MATH, "Rounds a number down, toward zero, to the nearest multiple of significance.", "number: The value to round | significance: The multiple to which to round"],
  ["FLOOR.MATH", MATH, "Rounds a number down to the nearest integer or to the nearest multiple of significance.", "number: The value to round | [significance]: The multiple to which to round (default 1) | [mode]: For negative numbers, a nonzero value rounds toward zero"],
  ["FLOOR.PRECISE", MATH, "Rounds a number down to the nearest integer or multiple of significance, regardless of the sign of the number.", "number: The value to round | [significance]: The multiple to which to round (default 1)"],
  ["EVEN", MATH, "Rounds a number up to the nearest even integer, away from zero.", NUM("The value to round")],
  ["ODD", MATH, "Rounds a number up to the nearest odd integer, away from zero.", NUM("The value to round")],
  ["RAND", MATH, "Returns a random number greater than or equal to 0 and less than 1.", ""],
  ["RANDBETWEEN", MATH, "Returns a random integer between the numbers you specify.", "bottom: The smallest integer RANDBETWEEN will return | top: The largest integer RANDBETWEEN will return"],
  ["RANDARRAY", MATH, "Returns an array of random numbers.", "[rows]: The number of rows to return | [columns]: The number of columns to return | [min]: The minimum number to return | [max]: The maximum number to return | [whole_number]: TRUE for whole numbers, FALSE for decimals (default)", 1],
  ["GCD", MATH, "Returns the greatest common divisor.", NUMBERS],
  ["LCM", MATH, "Returns the least common multiple.", NUMBERS],
  ["FACT", MATH, "Returns the factorial of a number.", NUM("The nonnegative number whose factorial you want")],
  ["FACTDOUBLE", MATH, "Returns the double factorial of a number.", NUM("The nonnegative number whose double factorial you want")],
  ["COMBIN", MATH, "Returns the number of combinations for a given number of items.", "number: The number of items | number_chosen: The number of items in each combination"],
  ["COMBINA", MATH, "Returns the number of combinations with repetitions for a given number of items.", "number: The number of items | number_chosen: The number of items in each combination"],
  ["MULTINOMIAL", MATH, "Returns the multinomial of a set of numbers.", NUMBERS],
  ["SERIESSUM", MATH, "Returns the sum of a power series based on the formula.", "x: The input value to the power series | n: The initial power to raise x to | m: The step by which to increase n for each term | coefficients: A set of coefficients by which each successive power of x is multiplied"],
  ["BASE", MATH, "Converts a number into a text representation with the given radix (base).", "number: The nonnegative integer to convert | radix: The base to convert into, from 2 to 36 | [min_length]: The minimum length of the returned string, padded with leading zeros"],
  ["DECIMAL", MATH, "Converts a text representation of a number in a given base into a decimal number.", "text: The text representation of the number | radix: The base of the number, from 2 to 36"],
  ["ROMAN", MATH, "Converts an Arabic numeral to Roman, as text.", "number: The Arabic numeral to convert | [form]: The type of Roman numeral, from 0 (classic, default) to 4 (simplified)"],
  ["ARABIC", MATH, "Converts a Roman numeral to an Arabic numeral.", "text: The Roman numeral text to convert"],
  ["PERCENTOF", MATH, "Returns the percentage that a subset makes up of a given data set.", "data_subset: The values in the subset | data_all: The values that make up the whole set"],
  ["MMULT", MATH, "Returns the matrix product of two arrays.", "array1: The first array to multiply | array2: The second array to multiply; its row count must equal array1's column count", 1],
  ["MINVERSE", MATH, "Returns the inverse matrix for the matrix stored in an array.", "array: A numeric array with an equal number of rows and columns", 1],
  ["MDETERM", MATH, "Returns the matrix determinant of an array.", "array: A numeric array with an equal number of rows and columns"],
  ["MUNIT", MATH, "Returns the unit matrix for the specified dimension.", "dimension: The dimension of the unit matrix", 1],
  ["DEGREES", MATH, "Converts radians to degrees.", "angle: The angle in radians to convert"],
  ["RADIANS", MATH, "Converts degrees to radians.", "angle: The angle in degrees to convert"],
  ["SIN", MATH, "Returns the sine of an angle.", ANGLE],
  ["COS", MATH, "Returns the cosine of an angle.", ANGLE],
  ["TAN", MATH, "Returns the tangent of an angle.", ANGLE],
  ["SEC", MATH, "Returns the secant of an angle.", ANGLE],
  ["CSC", MATH, "Returns the cosecant of an angle.", ANGLE],
  ["COT", MATH, "Returns the cotangent of an angle.", ANGLE],
  ["ASIN", MATH, "Returns the arcsine of a number, in radians in the range -pi/2 to pi/2.", NUM("The sine of the angle, from -1 to 1")],
  ["ACOS", MATH, "Returns the arccosine of a number, in radians in the range 0 to pi.", NUM("The cosine of the angle, from -1 to 1")],
  ["ATAN", MATH, "Returns the arctangent of a number, in radians in the range -pi/2 to pi/2.", NUM("The tangent of the angle")],
  ["ATAN2", MATH, "Returns the arctangent of the specified x- and y-coordinates, in radians between -pi and pi, excluding -pi.", "x_num: The x-coordinate of the point | y_num: The y-coordinate of the point"],
  ["ACOT", MATH, "Returns the arccotangent of a number, in radians in the range 0 to pi.", NUM("The cotangent of the angle")],
  ["SINH", MATH, "Returns the hyperbolic sine of a number.", TRIG_H("sine")],
  ["COSH", MATH, "Returns the hyperbolic cosine of a number.", TRIG_H("cosine")],
  ["TANH", MATH, "Returns the hyperbolic tangent of a number.", TRIG_H("tangent")],
  ["SECH", MATH, "Returns the hyperbolic secant of an angle.", TRIG_H("secant")],
  ["CSCH", MATH, "Returns the hyperbolic cosecant of an angle.", TRIG_H("cosecant")],
  ["COTH", MATH, "Returns the hyperbolic cotangent of a number.", TRIG_H("cotangent")],
  ["ASINH", MATH, "Returns the inverse hyperbolic sine of a number.", NUM("Any real number")],
  ["ACOSH", MATH, "Returns the inverse hyperbolic cosine of a number.", NUM("Any real number equal to or greater than 1")],
  ["ATANH", MATH, "Returns the inverse hyperbolic tangent of a number.", NUM("Any real number between -1 and 1, exclusive")],
  ["ACOTH", MATH, "Returns the inverse hyperbolic cotangent of a number.", NUM("A number whose absolute value is greater than 1")],

  // Statistical
  ["AVERAGE", STAT, "Returns the average (arithmetic mean) of its arguments.", NUMBERS],
  ["AVERAGEA", STAT, "Returns the average of its arguments, evaluating text and FALSE as 0 and TRUE as 1.", VALUES],
  ["AVERAGEIF", STAT, "Finds the average (arithmetic mean) for the cells specified by a given condition or criteria.", "range: The cells to evaluate against criteria | criteria: The condition that defines which cells are averaged | [average_range]: The cells to average, if different from range"],
  ["AVERAGEIFS", STAT, "Finds the average (arithmetic mean) for the cells specified by a given set of conditions or criteria.", `average_range: The cells to average | ${CRITERIA_PAIRS}`],
  ["AVERAGE.WEIGHTED", STAT, "Returns the weighted average of a set of values, given the values and their corresponding weights.", "values: The values to average | weights: The weights for each value | [additional_values]...: Additional values to average | [additional_weights]...: Weights for the additional values"],
  ["COUNT", STAT, "Counts the number of cells in a range that contain numbers.", VALUES],
  ["COUNTA", STAT, "Counts the number of cells in a range that are not empty.", VALUES],
  ["COUNTBLANK", STAT, "Counts the number of empty cells in a specified range of cells.", "range: The range in which to count blank cells"],
  ["COUNTIF", STAT, "Counts the number of cells within a range that meet the given condition.", "range: The range of cells to count | criteria: The condition that defines which cells are counted"],
  ["COUNTIFS", STAT, "Counts the number of cells specified by a given set of conditions or criteria.", CRITERIA_PAIRS],
  ["COUNTUNIQUE", STAT, "Counts the number of unique values in a list of specified values and ranges.", VALUES],
  ["MAX", STAT, "Returns the largest value in a set of values, ignoring logical values and text.", NUMBERS],
  ["MIN", STAT, "Returns the smallest number in a set of values, ignoring logical values and text.", NUMBERS],
  ["MAXA", STAT, "Returns the largest value in a set of values, including logical values and text.", VALUES],
  ["MINA", STAT, "Returns the smallest value in a set of values, including logical values and text.", VALUES],
  ["MAXIFS", STAT, "Returns the maximum value among cells specified by a given set of conditions or criteria.", `max_range: The cells in which to find the maximum | ${CRITERIA_PAIRS}`],
  ["MINIFS", STAT, "Returns the minimum value among cells specified by a given set of conditions or criteria.", `min_range: The cells in which to find the minimum | ${CRITERIA_PAIRS}`],
  ["MEDIAN", STAT, "Returns the median, or the number in the middle of the set of given numbers.", NUMBERS],
  ["MODE", STAT, "Returns the most frequently occurring, or repetitive, value in an array or range of data.", NUMBERS],
  ["MODE.SNGL", STAT, "Returns the most frequently occurring, or repetitive, value in an array or range of data.", NUMBERS],
  ["MODE.MULT", STAT, "Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data.", NUMBERS, 1],
  ["LARGE", STAT, "Returns the k-th largest value in a data set.", "array: The array or range of data | k: The position (from the largest) in the data to return"],
  ["SMALL", STAT, "Returns the k-th smallest value in a data set.", "array: The array or range of data | k: The position (from the smallest) in the data to return"],
  ["RANK", STAT, "Returns the rank of a number in a list of numbers.", "number: The number whose rank you want | ref: The list of numbers to rank against | [order]: 0 or omitted for descending, nonzero for ascending"],
  ["RANK.EQ", STAT, "Returns the rank of a number in a list of numbers; tied values get the top rank of the group.", "number: The number whose rank you want | ref: The list of numbers to rank against | [order]: 0 or omitted for descending, nonzero for ascending"],
  ["RANK.AVG", STAT, "Returns the rank of a number in a list of numbers; tied values get the average rank.", "number: The number whose rank you want | ref: The list of numbers to rank against | [order]: 0 or omitted for descending, nonzero for ascending"],
  ["PERCENTILE", STAT, "Returns the k-th percentile of values in a range.", "array: The array or range of data | k: The percentile value, from 0 to 1 inclusive"],
  ["PERCENTILE.INC", STAT, "Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.", "array: The array or range of data | k: The percentile value, from 0 to 1 inclusive"],
  ["PERCENTILE.EXC", STAT, "Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.", "array: The array or range of data | k: The percentile value, between 0 and 1 exclusive"],
  ["QUARTILE", STAT, "Returns the quartile of a data set.", "array: The array or range of numeric values | quart: Which quartile to return (0 = minimum, 1, 2 = median, 3, 4 = maximum)"],
  ["QUARTILE.INC", STAT, "Returns the quartile of a data set, based on percentile values from 0..1, inclusive.", "array: The array or range of numeric values | quart: Which quartile to return (0 = minimum, 1, 2 = median, 3, 4 = maximum)"],
  ["QUARTILE.EXC", STAT, "Returns the quartile of a data set, based on percentile values from 0..1, exclusive.", "array: The array or range of numeric values | quart: Which quartile to return (1, 2, or 3)"],
  ["PERCENTRANK", STAT, "Returns the rank of a value in a data set as a percentage of the data set.", "array: The array or range of numeric values | x: The value whose rank you want | [significance]: The number of significant digits for the result (default 3)"],
  ["PERCENTRANK.INC", STAT, "Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set.", "array: The array or range of numeric values | x: The value whose rank you want | [significance]: The number of significant digits for the result (default 3)"],
  ["PERCENTRANK.EXC", STAT, "Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set.", "array: The array or range of numeric values | x: The value whose rank you want | [significance]: The number of significant digits for the result (default 3)"],
  ["STDEV", STAT, "Estimates standard deviation based on a sample.", NUMBERS],
  ["STDEV.S", STAT, "Estimates standard deviation based on a sample, ignoring logical values and text.", NUMBERS],
  ["STDEVP", STAT, "Calculates standard deviation based on the entire population.", NUMBERS],
  ["STDEV.P", STAT, "Calculates standard deviation based on the entire population given as arguments, ignoring logical values and text.", NUMBERS],
  ["STDEVA", STAT, "Estimates standard deviation based on a sample, including logical values and text.", VALUES],
  ["STDEVPA", STAT, "Calculates standard deviation based on the entire population, including logical values and text.", VALUES],
  ["VAR", STAT, "Estimates variance based on a sample.", NUMBERS],
  ["VAR.S", STAT, "Estimates variance based on a sample, ignoring logical values and text.", NUMBERS],
  ["VARP", STAT, "Calculates variance based on the entire population.", NUMBERS],
  ["VAR.P", STAT, "Calculates variance based on the entire population, ignoring logical values and text.", NUMBERS],
  ["VARA", STAT, "Estimates variance based on a sample, including logical values and text.", VALUES],
  ["VARPA", STAT, "Calculates variance based on the entire population, including logical values and text.", VALUES],
  ["GEOMEAN", STAT, "Returns the geometric mean of an array or range of positive numeric data.", NUMBERS],
  ["HARMEAN", STAT, "Returns the harmonic mean of a data set of positive numbers.", NUMBERS],
  ["AVEDEV", STAT, "Returns the average of the absolute deviations of data points from their mean.", NUMBERS],
  ["DEVSQ", STAT, "Returns the sum of squares of deviations of data points from their sample mean.", NUMBERS],
  ["KURT", STAT, "Returns the kurtosis of a data set.", NUMBERS],
  ["SKEW", STAT, "Returns the skewness of a distribution.", NUMBERS],
  ["SKEW.P", STAT, "Returns the skewness of a distribution based on a population.", NUMBERS],
  ["STANDARDIZE", STAT, "Returns a normalized value from a distribution characterized by a mean and standard deviation.", "x: The value to normalize | mean: The arithmetic mean of the distribution | standard_dev: The standard deviation of the distribution"],
  ["TRIMMEAN", STAT, "Returns the mean of the interior portion of a set of data values.", "array: The array or range of values to trim and average | percent: The fraction of data points to exclude from the calculation"],
  ["FREQUENCY", STAT, "Calculates how often values occur within a range of values, and returns a vertical array of counts.", "data_array: The set of values for which you want to count frequencies | bins_array: The intervals into which to group the values", 1],
  ["CORREL", STAT, "Returns the correlation coefficient between two data sets.", TWO_ARRAYS],
  ["PEARSON", STAT, "Returns the Pearson product moment correlation coefficient.", TWO_ARRAYS],
  ["RSQ", STAT, "Returns the square of the Pearson product moment correlation coefficient.", KNOWN_YX],
  ["COVAR", STAT, "Returns covariance, the average of the products of paired deviations.", TWO_ARRAYS],
  ["COVARIANCE.P", STAT, "Returns population covariance, the average of the products of deviations for each data point pair.", TWO_ARRAYS],
  ["COVARIANCE.S", STAT, "Returns the sample covariance, the average of the products of deviations for each data point pair.", TWO_ARRAYS],
  ["SLOPE", STAT, "Returns the slope of the linear regression line through the given data points.", KNOWN_YX],
  ["INTERCEPT", STAT, "Calculates the point at which a line will intersect the y-axis by using a best-fit regression line.", KNOWN_YX],
  ["STEYX", STAT, "Returns the standard error of the predicted y-value for each x in a regression.", KNOWN_YX],
  ["FORECAST", STAT, "Calculates, or predicts, a future value along a linear trend by using existing values.", `x: The data point for which to predict a value | ${KNOWN_YX}`],
  ["FORECAST.LINEAR", STAT, "Calculates, or predicts, a future value along a linear trend by using existing values.", `x: The data point for which to predict a value | ${KNOWN_YX}`],
  ["TREND", STAT, "Returns numbers in a linear trend matching known data points, using the least squares method.", TREND_ARGS, 1],
  ["GROWTH", STAT, "Returns numbers in an exponential growth trend matching known data points.", TREND_ARGS, 1],
  ["LINEST", STAT, "Returns statistics that describe a linear trend matching known data points, fitting a straight line by the least squares method.", LINEST_ARGS, 1],
  ["LOGEST", STAT, "Returns statistics that describe an exponential curve matching known data points.", LINEST_ARGS, 1],
  ["NORM.DIST", STAT, "Returns the normal distribution for the specified mean and standard deviation.", `x: The value for which you want the distribution | ${NORMAL_PARAMS} | ${DIST_CUMULATIVE}`],
  ["NORMDIST", STAT, "Returns the normal cumulative distribution for the specified mean and standard deviation.", `x: The value for which you want the distribution | ${NORMAL_PARAMS} | ${DIST_CUMULATIVE}`],
  ["NORM.INV", STAT, "Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.", `${PROBABILITY} | ${NORMAL_PARAMS}`],
  ["NORMINV", STAT, "Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.", `${PROBABILITY} | ${NORMAL_PARAMS}`],
  ["NORM.S.DIST", STAT, "Returns the standard normal distribution (has a mean of zero and a standard deviation of one).", `z: The value for which you want the distribution | ${DIST_CUMULATIVE}`],
  ["NORMSDIST", STAT, "Returns the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).", "z: The value for which you want the distribution"],
  ["NORM.S.INV", STAT, "Returns the inverse of the standard normal cumulative distribution.", PROBABILITY],
  ["NORMSINV", STAT, "Returns the inverse of the standard normal cumulative distribution.", PROBABILITY],
  ["LOGNORM.DIST", STAT, "Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev.", `x: The value at which to evaluate the function | ${LOGNORMAL_PARAMS} | ${DIST_CUMULATIVE}`],
  ["LOGNORMDIST", STAT, "Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev.", `x: The value at which to evaluate the function | ${LOGNORMAL_PARAMS}`],
  ["LOGNORM.INV", STAT, "Returns the inverse of the lognormal cumulative distribution function.", `${PROBABILITY} | ${LOGNORMAL_PARAMS}`],
  ["LOGINV", STAT, "Returns the inverse of the lognormal cumulative distribution function.", `${PROBABILITY} | ${LOGNORMAL_PARAMS}`],
  ["T.DIST", STAT, "Returns the left-tailed Student's t-distribution.", `x: The numeric value at which to evaluate the distribution | ${DF} | ${DIST_CUMULATIVE}`],
  ["T.DIST.2T", STAT, "Returns the two-tailed Student's t-distribution.", `x: The numeric value at which to evaluate the distribution | ${DF}`],
  ["T.DIST.RT", STAT, "Returns the right-tailed Student's t-distribution.", `x: The numeric value at which to evaluate the distribution | ${DF}`],
  ["TDIST", STAT, "Returns the Student's t-distribution.", `x: The numeric value at which to evaluate the distribution | ${DF} | tails: The number of distribution tails to return (1 or 2)`],
  ["T.INV", STAT, "Returns the left-tailed inverse of the Student's t-distribution.", `probability: The probability associated with the Student's t-distribution | ${DF}`],
  ["T.INV.2T", STAT, "Returns the two-tailed inverse of the Student's t-distribution.", `probability: The probability associated with the two-tailed Student's t-distribution | ${DF}`],
  ["TINV", STAT, "Returns the two-tailed inverse of the Student's t-distribution.", `probability: The probability associated with the two-tailed Student's t-distribution | ${DF}`],
  ["T.TEST", STAT, "Returns the probability associated with a Student's t-test.", TTEST_ARGS],
  ["TTEST", STAT, "Returns the probability associated with a Student's t-test.", TTEST_ARGS],
  ["CHISQ.DIST", STAT, "Returns the left-tailed probability of the chi-squared distribution.", `x: The value at which to evaluate the distribution | ${DF} | ${DIST_CUMULATIVE}`],
  ["CHISQ.DIST.RT", STAT, "Returns the right-tailed probability of the chi-squared distribution.", `x: The value at which to evaluate the distribution | ${DF}`],
  ["CHIDIST", STAT, "Returns the right-tailed probability of the chi-squared distribution.", `x: The value at which to evaluate the distribution | ${DF}`],
  ["CHISQ.INV", STAT, "Returns the inverse of the left-tailed probability of the chi-squared distribution.", `probability: A probability associated with the chi-squared distribution | ${DF}`],
  ["CHISQ.INV.RT", STAT, "Returns the inverse of the right-tailed probability of the chi-squared distribution.", `probability: A probability associated with the chi-squared distribution | ${DF}`],
  ["CHIINV", STAT, "Returns the inverse of the right-tailed probability of the chi-squared distribution.", `probability: A probability associated with the chi-squared distribution | ${DF}`],
  ["CHISQ.TEST", STAT, "Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.", CHISQ_TEST_ARGS],
  ["CHITEST", STAT, "Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.", CHISQ_TEST_ARGS],
  ["F.DIST", STAT, "Returns the (left-tailed) F probability distribution for two data sets.", `x: The value at which to evaluate the function | ${F_DF} | ${DIST_CUMULATIVE}`],
  ["F.DIST.RT", STAT, "Returns the (right-tailed) F probability distribution for two data sets.", `x: The value at which to evaluate the function | ${F_DF}`],
  ["FDIST", STAT, "Returns the (right-tailed) F probability distribution for two data sets.", `x: The value at which to evaluate the function | ${F_DF}`],
  ["F.INV", STAT, "Returns the inverse of the (left-tailed) F probability distribution.", `probability: A probability associated with the F cumulative distribution | ${F_DF}`],
  ["F.INV.RT", STAT, "Returns the inverse of the (right-tailed) F probability distribution.", `probability: A probability associated with the F cumulative distribution | ${F_DF}`],
  ["FINV", STAT, "Returns the inverse of the (right-tailed) F probability distribution.", `probability: A probability associated with the F cumulative distribution | ${F_DF}`],
  ["F.TEST", STAT, "Returns the two-tailed probability that the variances in two data sets are not significantly different.", TWO_ARRAYS],
  ["FTEST", STAT, "Returns the two-tailed probability that the variances in two data sets are not significantly different.", TWO_ARRAYS],
  ["BINOM.DIST", STAT, "Returns the individual term binomial distribution probability.", `${BINOM_ARGS} | ${DIST_CUMULATIVE}`],
  ["BINOMDIST", STAT, "Returns the individual term binomial distribution probability.", `${BINOM_ARGS} | ${DIST_CUMULATIVE}`],
  ["BINOM.DIST.RANGE", STAT, "Returns the probability of a trial result using a binomial distribution.", "trials: The number of independent trials | probability_s: The probability of success in each trial | number_s: The number of successes in trials | [number_s2]: If given, the upper bound of the number of successes"],
  ["BINOM.INV", STAT, "Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.", "trials: The number of Bernoulli trials | probability_s: The probability of a success on each trial | alpha: The criterion value"],
  ["CRITBINOM", STAT, "Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.", "trials: The number of Bernoulli trials | probability_s: The probability of a success on each trial | alpha: The criterion value"],
  ["POISSON.DIST", STAT, "Returns the Poisson distribution.", `x: The number of events | mean: The expected numeric value | ${DIST_CUMULATIVE}`],
  ["POISSON", STAT, "Returns the Poisson distribution.", `x: The number of events | mean: The expected numeric value | ${DIST_CUMULATIVE}`],
  ["EXPON.DIST", STAT, "Returns the exponential distribution.", `x: The value of the function | lambda: The parameter value | ${DIST_CUMULATIVE}`],
  ["EXPONDIST", STAT, "Returns the exponential distribution.", `x: The value of the function | lambda: The parameter value | ${DIST_CUMULATIVE}`],
  ["GAMMA", STAT, "Returns the Gamma function value.", "x: The value for which you want Gamma"],
  ["GAMMALN", STAT, "Returns the natural logarithm of the gamma function, ln(Γ(x)).", "x: The positive value for which you want to calculate GAMMALN"],
  ["GAMMALN.PRECISE", STAT, "Returns the natural logarithm of the gamma function, ln(Γ(x)).", "x: The positive value for which you want to calculate GAMMALN.PRECISE"],
  ["GAMMA.DIST", STAT, "Returns the gamma distribution.", `x: The value at which to evaluate the distribution | ${GAMMA_PARAMS} | ${DIST_CUMULATIVE}`],
  ["GAMMADIST", STAT, "Returns the gamma distribution.", `x: The value at which to evaluate the distribution | ${GAMMA_PARAMS} | ${DIST_CUMULATIVE}`],
  ["GAMMA.INV", STAT, "Returns the inverse of the gamma cumulative distribution.", `probability: The probability associated with the gamma distribution | ${GAMMA_PARAMS}`],
  ["GAMMAINV", STAT, "Returns the inverse of the gamma cumulative distribution.", `probability: The probability associated with the gamma distribution | ${GAMMA_PARAMS}`],
  ["BETA.DIST", STAT, "Returns the beta distribution.", `x: The value between A and B at which to evaluate the function | ${BETA_PARAMS} | ${DIST_CUMULATIVE} | ${BETA_BOUNDS}`],
  ["BETADIST", STAT, "Returns the cumulative beta probability density function.", `x: The value between A and B at which to evaluate the function | ${BETA_PARAMS} | ${BETA_BOUNDS}`],
  ["BETA.INV", STAT, "Returns the inverse of the cumulative distribution function for a specified beta distribution.", `probability: A probability associated with the beta distribution | ${BETA_PARAMS} | ${BETA_BOUNDS}`],
  ["BETAINV", STAT, "Returns the inverse of the cumulative distribution function for a specified beta distribution.", `probability: A probability associated with the beta distribution | ${BETA_PARAMS} | ${BETA_BOUNDS}`],
  ["WEIBULL.DIST", STAT, "Returns the Weibull distribution.", `x: The value at which to evaluate the function | alpha: The shape parameter of the distribution | beta: The scale parameter of the distribution | ${DIST_CUMULATIVE}`],
  ["WEIBULL", STAT, "Returns the Weibull distribution.", `x: The value at which to evaluate the function | alpha: The shape parameter of the distribution | beta: The scale parameter of the distribution | ${DIST_CUMULATIVE}`],
  ["HYPGEOM.DIST", STAT, "Returns the hypergeometric distribution.", `${HYPGEOM_ARGS} | ${DIST_CUMULATIVE}`],
  ["HYPGEOMDIST", STAT, "Returns the hypergeometric distribution.", HYPGEOM_ARGS],
  ["NEGBINOM.DIST", STAT, "Returns the negative binomial distribution.", `${NEGBINOM_ARGS} | ${DIST_CUMULATIVE}`],
  ["NEGBINOMDIST", STAT, "Returns the negative binomial distribution.", NEGBINOM_ARGS],
  ["CONFIDENCE.NORM", STAT, "Returns the confidence interval for a population mean, using a normal distribution.", CONF_ARGS],
  ["CONFIDENCE.T", STAT, "Returns the confidence interval for a population mean, using a Student's t distribution.", CONF_ARGS],
  ["CONFIDENCE", STAT, "Returns the confidence interval for a population mean, using a normal distribution.", CONF_ARGS],
  ["Z.TEST", STAT, "Returns the one-tailed P-value of a z-test.", ZTEST_ARGS],
  ["ZTEST", STAT, "Returns the one-tailed P-value of a z-test.", ZTEST_ARGS],
  ["PHI", STAT, "Returns the value of the density function for a standard normal distribution.", "x: The number for which you want the density of the standard normal distribution"],
  ["GAUSS", STAT, "Returns 0.5 less than the standard normal cumulative distribution.", "z: The value for which you want the distribution"],
  ["FISHER", STAT, "Returns the Fisher transformation.", "x: A numeric value between -1 and 1, exclusive"],
  ["FISHERINV", STAT, "Returns the inverse of the Fisher transformation.", "y: The value for which you want to perform the inverse of the transformation"],
  ["PROB", STAT, "Returns the probability that values in a range are between two limits.", "x_range: The range of numeric values of x with associated probabilities | prob_range: A set of probabilities associated with the values in x_range | lower_limit: The lower bound on the value | [upper_limit]: The optional upper bound on the value"],
  ["PERMUT", STAT, "Returns the number of permutations for a given number of objects.", "number: The number of objects | number_chosen: The number of objects in each permutation"],
  ["PERMUTATIONA", STAT, "Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.", "number: The total number of objects | number_chosen: The number of objects in each permutation"],

  // Text
  ["CONCAT", TEXT, "Combines the text from multiple ranges and/or strings.", TEXTS],
  ["CONCATENATE", TEXT, "Joins several text strings into one text string.", TEXTS],
  ["TEXTJOIN", TEXT, "Combines the text from multiple ranges and/or strings, with a delimiter between each value.", "delimiter: The text inserted between each item | ignore_empty: TRUE to skip empty cells | text1: The first text item or range to join | [text2]...: Additional text items or ranges"],
  ["JOIN", TEXT, "Concatenates the elements of one or more arrays using a specified delimiter.", "delimiter: The text inserted between each value | value_or_array1: The first value or range to join | [value_or_array2]...: Additional values or ranges"],
  ["LEFT", TEXT, "Returns the specified number of characters from the start of a text string.", "text: The text string containing the characters you want to extract | [num_chars]: The number of characters to extract (default 1)"],
  ["RIGHT", TEXT, "Returns the specified number of characters from the end of a text string.", "text: The text string containing the characters you want to extract | [num_chars]: The number of characters to extract (default 1)"],
  ["MID", TEXT, "Returns the characters from the middle of a text string, given a starting position and length.", "text: The text string containing the characters you want to extract | start_num: The position of the first character to extract | num_chars: The number of characters to extract"],
  ["LEFTB", TEXT, "Returns the specified number of characters from the start of a text string (character-based here).", "text: The text string containing the characters you want to extract | [num_bytes]: The number of characters to extract (default 1)"],
  ["RIGHTB", TEXT, "Returns the specified number of characters from the end of a text string (character-based here).", "text: The text string containing the characters you want to extract | [num_bytes]: The number of characters to extract (default 1)"],
  ["MIDB", TEXT, "Returns characters from the middle of a text string, given a starting position and length (character-based here).", "text: The text string containing the characters you want to extract | start_num: The position of the first character to extract | num_bytes: The number of characters to extract"],
  ["LEN", TEXT, "Returns the number of characters in a text string.", "text: The text whose length you want to find"],
  ["LENB", TEXT, "Returns the number of characters in a text string (character-based here).", "text: The text whose length you want to find"],
  ["FIND", TEXT, "Returns the starting position of one text string within another text string; FIND is case-sensitive.", "find_text: The text you want to find | within_text: The text containing the text you want to find | [start_num]: The character at which to start the search (default 1)"],
  ["FINDB", TEXT, "Returns the starting position of one text string within another text string; case-sensitive (character-based here).", "find_text: The text you want to find | within_text: The text containing the text you want to find | [start_num]: The character at which to start the search (default 1)"],
  ["SEARCH", TEXT, "Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive).", "find_text: The text you want to find; wildcards ? and * are allowed | within_text: The text in which you want to search | [start_num]: The character at which to start the search (default 1)"],
  ["SEARCHB", TEXT, "Returns the position at which a text string is first found, not case-sensitive (character-based here).", "find_text: The text you want to find; wildcards ? and * are allowed | within_text: The text in which you want to search | [start_num]: The character at which to start the search (default 1)"],
  ["REPLACE", TEXT, "Replaces part of a text string with a different text string.", "old_text: The text in which you want to replace some characters | start_num: The position of the first character to replace | num_chars: The number of characters to replace | new_text: The text that will replace the characters"],
  ["REPLACEB", TEXT, "Replaces part of a text string with a different text string (character-based here).", "old_text: The text in which you want to replace some characters | start_num: The position of the first character to replace | num_bytes: The number of characters to replace | new_text: The text that will replace the characters"],
  ["SUBSTITUTE", TEXT, "Substitutes new text for old text in a text string.", "text: The text or reference to a cell containing text | old_text: The text you want to replace | new_text: The text you want to replace old_text with | [instance_num]: Which occurrence of old_text to replace; all occurrences if omitted"],
  ["TEXTBEFORE", TEXT, "Returns text that occurs before a given character or string.", "text: The text you are searching within | delimiter: The text that marks the point before which you want to extract | [instance_num]: Which occurrence of the delimiter to use; negative values search from the end | [match_mode]: 0 for case-sensitive (default), 1 for case-insensitive | [match_end]: 1 to treat the end of text as a delimiter | [if_not_found]: The value returned if no match is found"],
  ["TEXTAFTER", TEXT, "Returns text that occurs after a given character or string.", "text: The text you are searching within | delimiter: The text that marks the point after which you want to extract | [instance_num]: Which occurrence of the delimiter to use; negative values search from the end | [match_mode]: 0 for case-sensitive (default), 1 for case-insensitive | [match_end]: 1 to treat the end of text as a delimiter | [if_not_found]: The value returned if no match is found"],
  ["TEXTSPLIT", TEXT, "Splits text strings by using column and row delimiters.", "text: The text to split | col_delimiter: The text that marks where to spill across columns | [row_delimiter]: The text that marks where to spill down rows | [ignore_empty]: TRUE to ignore consecutive delimiters | [match_mode]: 0 for case-sensitive (default), 1 for case-insensitive | [pad_with]: The value used to pad missing results (default #N/A)", 1],
  ["SPLIT", TEXT, "Divides text around a specified character or string, and puts each fragment into a separate cell in the row.", "text: The text to divide | delimiter: The character or characters to use to split text | [split_by_each]: TRUE (default) to split around each character in delimiter | [remove_empty_text]: TRUE (default) to remove empty text from the result", 1],
  ["TRIM", TEXT, "Removes all spaces from a text string except for single spaces between words.", "text: The text from which you want spaces removed"],
  ["CLEAN", TEXT, "Removes all nonprintable characters from text.", "text: Any worksheet information from which you want to remove nonprintable characters"],
  ["UPPER", TEXT, "Converts text to uppercase.", "text: The text you want converted to uppercase"],
  ["LOWER", TEXT, "Converts all letters in a text string to lowercase.", "text: The text you want to convert to lowercase"],
  ["PROPER", TEXT, "Converts a text string to proper case: the first letter in each word to uppercase, and all other letters to lowercase.", "text: The text you want partially capitalized"],
  ["EXACT", TEXT, "Checks whether two text strings are exactly the same, and returns TRUE or FALSE; EXACT is case-sensitive.", "text1: The first text string | text2: The second text string"],
  ["REPT", TEXT, "Repeats text a given number of times.", "text: The text you want to repeat | number_times: The number of times to repeat text"],
  ["CHAR", TEXT, "Returns the character specified by the code number.", NUM("A number between 1 and 255 specifying which character you want")],
  ["CODE", TEXT, "Returns a numeric code for the first character in a text string.", "text: The text for which you want the code of the first character"],
  ["UNICHAR", TEXT, "Returns the Unicode character that is referenced by the given numeric value.", NUM("The Unicode number that represents the character")],
  ["UNICODE", TEXT, "Returns the number (code point) that corresponds to the first character of the text.", "text: The character for which you want the Unicode value"],
  ["TEXT", TEXT, "Converts a value to text in a specific number format.", "value: The numeric value to convert to text | format_text: The number format to apply, such as \"0.00\" or \"mm/dd/yyyy\""],
  ["FIXED", TEXT, "Formats a number as text with a fixed number of decimals.", "number: The number you want to round and convert to text | [decimals]: The number of digits to the right of the decimal point (default 2) | [no_commas]: TRUE to omit thousands separators"],
  ["DOLLAR", TEXT, "Converts a number to text, using currency format.", "number: A number, a reference to a cell containing a number, or a formula that evaluates to a number | [decimals]: The number of digits to the right of the decimal point (default 2)"],
  ["VALUE", TEXT, "Converts a text string that represents a number to a number.", "text: The text enclosed in quotation marks or a reference to a cell containing the text you want to convert"],
  ["NUMBERVALUE", TEXT, "Converts text to a number, in a locale-independent way.", "text: The text to convert to a number | [decimal_separator]: The character used as the decimal separator | [group_separator]: The character used to separate groups of digits"],
  ["VALUETOTEXT", TEXT, "Returns text from any specified value.", "value: The value to return as text | [format]: 0 for concise (default), 1 for strict format that can be parsed back"],
  ["ARRAYTOTEXT", TEXT, "Returns an array of text values from any specified range.", "array: The array to return as text | [format]: 0 for concise (default), 1 for strict format with braces"],
  ["T", TEXT, "Checks whether a value is text, and returns the text if it is, or returns empty text if it is not.", "value: The value you want to test"],
  ["ASC", TEXT, "Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters.", "text: The text or a reference to a cell that contains the text you want to change"],
  ["JIS", TEXT, "Changes half-width (single-byte) characters within a character string to full-width (double-byte) characters.", "text: The text or a reference to a cell that contains the text you want to change"],
  ["DBCS", TEXT, "Changes half-width (single-byte) characters within a character string to full-width (double-byte) characters.", "text: The text or a reference to a cell that contains the text you want to change"],
  ["PHONETIC", TEXT, "Extracts the phonetic (furigana) characters from a text string.", "reference: A text string or a reference to a cell containing a furigana text string"],
  ["REGEXTEST", TEXT, "Determines whether any part of text matches the pattern.", `text: The text or cell reference to test | pattern: The regular expression that describes the text pattern | ${REGEX_CASE}`],
  ["REGEXEXTRACT", TEXT, "Extracts strings within the provided text that match the pattern.", `text: The text or cell reference from which to extract strings | pattern: The regular expression that describes the text pattern | [return_mode]: 0 = first match (default), 1 = all matches as an array, 2 = capturing groups of the first match | ${REGEX_CASE}`, 1],
  ["REGEXREPLACE", TEXT, "Replaces strings within the provided text that match the pattern with replacement.", `text: The text or cell reference in which to replace strings | pattern: The regular expression that describes the text pattern | replacement: The text to replace instances of pattern; $1, $2 refer to groups | [occurrence]: Which instance to replace; 0 replaces all (default), negative counts from the end | ${REGEX_CASE}`],
  ["REGEXMATCH", TEXT, "Checks whether a piece of text matches a regular expression.", "text: The text to be tested against the regular expression | regular_expression: The regular expression to test the text against"],
  ["TO_TEXT", TEXT, "Converts a provided numeric value to a text value.", "value: The argument or reference to a cell to be converted to text"],
  ["TO_PURE_NUMBER", TEXT, "Converts a provided date/time, percentage, currency, or other formatted numeric value to a pure number without formatting.", "value: The argument or reference to a cell to be converted to a pure number"],
  ["TO_PERCENT", TEXT, "Converts a provided number to a percentage.", "value: The argument or reference to a cell to be converted to a percentage"],
  ["TO_DOLLARS", TEXT, "Converts a provided number to a dollar value.", "value: The argument or reference to a cell to be converted to a dollar value"],

  // Web
  ["ENCODEURL", WEB, "Returns a URL-encoded string.", "text: A string to be URL encoded"],

  // Date and time
  ["DATE", DATE, "Returns the serial number that represents a particular date.", "year: The year, from 1900 to 9999 | month: The month of the year, from 1 to 12; out-of-range values roll over | day: The day of the month, from 1 to 31; out-of-range values roll over"],
  ["DATEVALUE", DATE, "Converts a date in the form of text to a serial number.", "date_text: Text that represents a date"],
  ["DATEDIF", DATE, "Calculates the number of days, months, or years between two dates.", "start_date: The first, or starting, date of a given period | end_date: The last, or ending, date of the period | unit: The type of information to return: \"Y\", \"M\", \"D\", \"MD\", \"YM\", or \"YD\""],
  ["DAYS", DATE, "Returns the number of days between two dates.", "end_date: The later of the two dates | start_date: The earlier of the two dates"],
  ["DAYS360", DATE, "Returns the number of days between two dates based on a 360-day year.", "start_date: The start date of the period | end_date: The end date of the period | [method]: FALSE for the US (NASD) method (default), TRUE for the European method"],
  ["DAY", DATE, "Returns the day of a date, represented by a serial number, as an integer from 1 to 31.", "serial_number: The date of the day you are trying to find"],
  ["MONTH", DATE, "Returns the month of a date represented by a serial number, as an integer from 1 (January) to 12 (December).", "serial_number: The date of the month you are trying to find"],
  ["YEAR", DATE, "Returns the year corresponding to a date, as an integer in the range 1900-9999.", "serial_number: The date of the year you want to find"],
  ["EDATE", DATE, "Returns the serial number of the date that is the indicated number of months before or after the start date.", "start_date: The starting date | months: The number of months before (negative) or after (positive) start_date"],
  ["EOMONTH", DATE, "Returns the serial number of the last day of the month before or after a specified number of months.", "start_date: The starting date | months: The number of months before (negative) or after (positive) start_date"],
  ["HOUR", DATE, "Returns the hour of a time value, as an integer from 0 (12:00 A.M.) to 23 (11:00 P.M.).", "serial_number: The time that contains the hour you want to find"],
  ["MINUTE", DATE, "Returns the minutes of a time value, as an integer from 0 to 59.", "serial_number: The time that contains the minute you want to find"],
  ["SECOND", DATE, "Returns the seconds of a time value, as an integer from 0 to 59.", "serial_number: The time that contains the seconds you want to find"],
  ["TIME", DATE, "Returns the decimal number for a particular time.", "hour: A number from 0 to 32767 representing the hour | minute: A number from 0 to 32767 representing the minute | second: A number from 0 to 32767 representing the second"],
  ["TIMEVALUE", DATE, "Converts a time in the form of text to a serial number.", "time_text: A text string that represents a time"],
  ["TODAY", DATE, "Returns the serial number of the current date.", ""],
  ["NOW", DATE, "Returns the serial number of the current date and time.", ""],
  ["WEEKDAY", DATE, "Returns a number from 1 to 7 identifying the day of the week of a date.", "serial_number: A sequential number that represents the date of the day you are trying to find | [return_type]: A number that determines the type of return value (1 = Sunday-Saturday, 2 = Monday-Sunday, 3 = Monday 0-Sunday 6, 11-17)"],
  ["WEEKNUM", DATE, "Returns the week number of a specific date.", "serial_number: A date within the week | [return_type]: The day on which the week begins (1 = Sunday, 2 = Monday, 11-17, 21 = ISO)"],
  ["ISOWEEKNUM", DATE, "Returns the ISO week number of the year for a given date.", "date: The date-time code used for date and time calculation"],
  ["NETWORKDAYS", DATE, "Returns the number of whole workdays between two dates.", "start_date: The start date | end_date: The end date | [holidays]: A range or array of dates to exclude from the working calendar"],
  ["NETWORKDAYS.INTL", DATE, "Returns the number of whole workdays between two dates, using parameters to indicate which and how many days are weekend days.", `start_date: The start date | end_date: The end date | ${WORKDAY_WEEKEND} | [holidays]: A range or array of dates to exclude from the working calendar`],
  ["WORKDAY", DATE, "Returns the serial number of the date before or after a specified number of workdays.", "start_date: The start date | days: The number of nonweekend and nonholiday days before or after start_date | [holidays]: A range or array of dates to exclude from the working calendar"],
  ["WORKDAY.INTL", DATE, "Returns the serial number of the date before or after a specified number of workdays, with custom weekend parameters.", `start_date: The start date | days: The number of workdays before or after start_date | ${WORKDAY_WEEKEND} | [holidays]: A range or array of dates to exclude from the working calendar`],
  ["YEARFRAC", DATE, "Returns the year fraction representing the number of whole days between start_date and end_date.", `start_date: The start date | end_date: The end date | ${BASIS}`],
  ["EPOCHTODATE", DATE, "Converts a Unix epoch timestamp in seconds, milliseconds, or microseconds to a date and time in UTC.", "timestamp: A Unix epoch timestamp | [unit]: The unit of time of the timestamp (1 = seconds (default), 2 = milliseconds, 3 = microseconds)"],
  ["TO_DATE", DATE, "Converts a provided number to a date.", "value: The argument or reference to a cell to be converted to a date"],

  // Financial
  ["PMT", FIN, "Calculates the payment for a loan based on constant payments and a constant interest rate.", `${RATE_ARG} | nper: The total number of payments | pv: The present value, or the total amount that a series of future payments is worth now | [fv]: The future value you want after the last payment (default 0) | ${TYPE_ARG}`],
  ["IPMT", FIN, "Returns the interest payment for a given period for an investment based on periodic, constant payments and a constant interest rate.", `${RATE_ARG} | per: The period for which you want to find the interest | nper: The total number of payment periods | pv: The present value | [fv]: The future value (default 0) | ${TYPE_ARG}`],
  ["PPMT", FIN, "Returns the payment on the principal for a given period for an investment based on periodic, constant payments and a constant interest rate.", `${RATE_ARG} | per: The period for which you want to find the principal | nper: The total number of payment periods | pv: The present value | [fv]: The future value (default 0) | ${TYPE_ARG}`],
  ["FV", FIN, "Returns the future value of an investment based on periodic, constant payments and a constant interest rate.", `${RATE_ARG} | nper: The total number of payment periods | pmt: The payment made each period | [pv]: The present value (default 0) | ${TYPE_ARG}`],
  ["PV", FIN, "Returns the present value of an investment: the total amount that a series of future payments is worth now.", `${RATE_ARG} | nper: The total number of payment periods | pmt: The payment made each period | [fv]: The future value (default 0) | ${TYPE_ARG}`],
  ["NPER", FIN, "Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.", `${RATE_ARG} | pmt: The payment made each period | pv: The present value | [fv]: The future value (default 0) | ${TYPE_ARG}`],
  ["RATE", FIN, "Returns the interest rate per period of a loan or an annuity.", `nper: The total number of payment periods | pmt: The payment made each period | pv: The present value | [fv]: The future value (default 0) | ${TYPE_ARG} | [guess]: Your guess for what the rate will be (default 10%)`],
  ["NPV", FIN, "Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).", "rate: The rate of discount over one period | value1: The first payment or income | [value2]...: Additional payments and income"],
  ["IRR", FIN, "Returns the internal rate of return for a series of cash flows.", "values: An array or range containing the cash flows | [guess]: A number you guess is close to the result (default 10%)"],
  ["XNPV", FIN, "Returns the net present value for a schedule of cash flows that is not necessarily periodic.", "rate: The discount rate to apply to the cash flows | values: A series of cash flows that corresponds to a schedule of payments in dates | dates: A schedule of payment dates that corresponds to the cash flow payments"],
  ["XIRR", FIN, "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.", "values: A series of cash flows that corresponds to a schedule of payments in dates | dates: A schedule of payment dates that corresponds to the cash flow payments | [guess]: A number you guess is close to the result (default 10%)"],
  ["MIRR", FIN, "Returns the internal rate of return where positive and negative cash flows are financed at different rates.", "values: An array or range containing the cash flows | finance_rate: The interest rate you pay on the money used in the cash flows | reinvest_rate: The interest rate you receive on the cash flows as you reinvest them"],
  ["EFFECT", FIN, "Returns the effective annual interest rate.", "nominal_rate: The nominal interest rate | npery: The number of compounding periods per year"],
  ["NOMINAL", FIN, "Returns the annual nominal interest rate.", "effect_rate: The effective interest rate | npery: The number of compounding periods per year"],
  ["PDURATION", FIN, "Returns the number of periods required by an investment to reach a specified value.", "rate: The interest rate per period | pv: The present value of the investment | fv: The desired future value of the investment"],
  ["RRI", FIN, "Returns an equivalent interest rate for the growth of an investment.", "nper: The number of periods for the investment | pv: The present value of the investment | fv: The future value of the investment"],
  ["FVSCHEDULE", FIN, "Returns the future value of an initial principal after applying a series of compound interest rates.", "principal: The present value | schedule: An array of interest rates to apply"],
  ["CUMIPMT", FIN, "Returns the cumulative interest paid on a loan between start_period and end_period.", `${RATE_ARG} | nper: The total number of payment periods | pv: The present value | start_period: The first period in the calculation | end_period: The last period in the calculation | type: When payments are due (0 = end of period, 1 = beginning)`],
  ["CUMPRINC", FIN, "Returns the cumulative principal paid on a loan between start_period and end_period.", `${RATE_ARG} | nper: The total number of payment periods | pv: The present value | start_period: The first period in the calculation | end_period: The last period in the calculation | type: When payments are due (0 = end of period, 1 = beginning)`],
  ["ISPMT", FIN, "Calculates the interest paid during a specific period of an investment with even principal payments.", `${RATE_ARG} | per: The period for which you want to find the interest | nper: The total number of payment periods | pv: The present value`],
  ["SLN", FIN, "Returns the straight-line depreciation of an asset for one period.", DEPRECIATION],
  ["SYD", FIN, "Returns the sum-of-years' digits depreciation of an asset for a specified period.", `${DEPRECIATION} | per: The period`],
  ["DDB", FIN, "Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.", `${DEPRECIATION} | period: The period for which you want to calculate the depreciation | [factor]: The rate at which the balance declines (default 2)`],
  ["DB", FIN, "Returns the depreciation of an asset for a specified period using the fixed-declining balance method.", `${DEPRECIATION} | period: The period for which you want to calculate the depreciation | [month]: The number of months in the first year (default 12)`],
  ["VDB", FIN, "Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.", `${DEPRECIATION} | start_period: The starting period | end_period: The ending period | [factor]: The rate at which the balance declines (default 2) | [no_switch]: TRUE to never switch to straight-line depreciation`],
  ["AMORLINC", FIN, "Returns the depreciation for each accounting period (French accounting system).", `cost: The cost of the asset | date_purchased: The date of the purchase of the asset | first_period: The date of the end of the first period | salvage: The salvage value at the end of the life of the asset | period: The period | rate: The rate of depreciation | ${BASIS}`],
  ["AMORDEGRC", FIN, "Returns the depreciation for each accounting period by using a depreciation coefficient (French accounting system).", `cost: The cost of the asset | date_purchased: The date of the purchase of the asset | first_period: The date of the end of the first period | salvage: The salvage value at the end of the life of the asset | period: The period | rate: The rate of depreciation | ${BASIS}`],
  ["DOLLARDE", FIN, "Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal number.", "fractional_dollar: A number expressed as an integer part and a fraction part, separated by a decimal symbol | fraction: The integer to use in the denominator of the fraction"],
  ["DOLLARFR", FIN, "Converts a dollar price expressed as a decimal number into a dollar price expressed as a fraction.", "decimal_dollar: A decimal number | fraction: The integer to use in the denominator of a fraction"],
  ["ACCRINT", FIN, "Returns the accrued interest for a security that pays periodic interest.", `issue: The security's issue date | first_interest: The security's first interest date | ${SETTLE} | rate: The security's annual coupon rate | par: The security's par value | ${FREQUENCY} | ${BASIS} | [calc_method]: TRUE to accrue from issue (default), FALSE to accrue from first_interest`],
  ["ACCRINTM", FIN, "Returns the accrued interest for a security that pays interest at maturity.", `issue: The security's issue date | settlement: The security's maturity date | rate: The security's annual coupon rate | par: The security's par value | ${BASIS}`],
  ["COUPDAYBS", FIN, "Returns the number of days from the beginning of the coupon period to the settlement date.", COUPON_ARGS],
  ["COUPDAYS", FIN, "Returns the number of days in the coupon period that contains the settlement date.", COUPON_ARGS],
  ["COUPDAYSNC", FIN, "Returns the number of days from the settlement date to the next coupon date.", COUPON_ARGS],
  ["COUPNCD", FIN, "Returns the next coupon date after the settlement date.", COUPON_ARGS],
  ["COUPNUM", FIN, "Returns the number of coupons payable between the settlement date and maturity date.", COUPON_ARGS],
  ["COUPPCD", FIN, "Returns the previous coupon date before the settlement date.", COUPON_ARGS],
  ["PRICE", FIN, "Returns the price per $100 face value of a security that pays periodic interest.", `${SETTLE} | ${MATURE} | rate: The security's annual coupon rate | yld: The security's annual yield | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],
  ["YIELD", FIN, "Returns the yield on a security that pays periodic interest.", `${SETTLE} | ${MATURE} | rate: The security's annual coupon rate | pr: The security's price per $100 face value | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],
  ["DURATION", FIN, "Returns the annual duration of a security with periodic interest payments.", `${SETTLE} | ${MATURE} | coupon: The security's annual coupon rate | yld: The security's annual yield | ${FREQUENCY} | ${BASIS}`],
  ["MDURATION", FIN, "Returns the Macauley modified duration for a security with an assumed par value of $100.", `${SETTLE} | ${MATURE} | coupon: The security's annual coupon rate | yld: The security's annual yield | ${FREQUENCY} | ${BASIS}`],
  ["DISC", FIN, "Returns the discount rate for a security.", `${SETTLE} | ${MATURE} | pr: The security's price per $100 face value | redemption: The security's redemption value per $100 face value | ${BASIS}`],
  ["INTRATE", FIN, "Returns the interest rate for a fully invested security.", `${SETTLE} | ${MATURE} | investment: The amount invested in the security | redemption: The amount to be received at maturity | ${BASIS}`],
  ["RECEIVED", FIN, "Returns the amount received at maturity for a fully invested security.", `${SETTLE} | ${MATURE} | investment: The amount invested in the security | discount: The security's discount rate | ${BASIS}`],
  ["PRICEDISC", FIN, "Returns the price per $100 face value of a discounted security.", `${SETTLE} | ${MATURE} | discount: The security's discount rate | redemption: The security's redemption value per $100 face value | ${BASIS}`],
  ["PRICEMAT", FIN, "Returns the price per $100 face value of a security that pays interest at maturity.", `${SETTLE} | ${MATURE} | issue: The security's issue date | rate: The security's interest rate at date of issue | yld: The security's annual yield | ${BASIS}`],
  ["YIELDDISC", FIN, "Returns the annual yield for a discounted security.", `${SETTLE} | ${MATURE} | pr: The security's price per $100 face value | redemption: The security's redemption value per $100 face value | ${BASIS}`],
  ["YIELDMAT", FIN, "Returns the annual yield of a security that pays interest at maturity.", `${SETTLE} | ${MATURE} | issue: The security's issue date | rate: The security's interest rate at date of issue | pr: The security's price per $100 face value | ${BASIS}`],
  ["TBILLEQ", FIN, "Returns the bond-equivalent yield for a Treasury bill.", "settlement: The Treasury bill's settlement date | maturity: The Treasury bill's maturity date | discount: The Treasury bill's discount rate"],
  ["TBILLPRICE", FIN, "Returns the price per $100 face value for a Treasury bill.", "settlement: The Treasury bill's settlement date | maturity: The Treasury bill's maturity date | discount: The Treasury bill's discount rate"],
  ["TBILLYIELD", FIN, "Returns the yield for a Treasury bill.", "settlement: The Treasury bill's settlement date | maturity: The Treasury bill's maturity date | pr: The Treasury bill's price per $100 face value"],
  ["ODDFPRICE", FIN, "Returns the price per $100 face value of a security with an odd first period.", `${SETTLE} | ${MATURE} | issue: The security's issue date | first_coupon: The security's first coupon date | rate: The security's interest rate | yld: The security's annual yield | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],
  ["ODDFYIELD", FIN, "Returns the yield of a security with an odd first period.", `${SETTLE} | ${MATURE} | issue: The security's issue date | first_coupon: The security's first coupon date | rate: The security's interest rate | pr: The security's price | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],
  ["ODDLPRICE", FIN, "Returns the price per $100 face value of a security with an odd last period.", `${SETTLE} | ${MATURE} | last_interest: The security's last coupon date | rate: The security's interest rate | yld: The security's annual yield | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],
  ["ODDLYIELD", FIN, "Returns the yield of a security with an odd last period.", `${SETTLE} | ${MATURE} | last_interest: The security's last coupon date | rate: The security's interest rate | pr: The security's price | redemption: The security's redemption value per $100 face value | ${FREQUENCY} | ${BASIS}`],

  // Engineering
  ["BIN2DEC", ENG, "Converts a binary number to decimal.", NUM("The binary number you want to convert (up to 10 characters)")],
  ["BIN2HEX", ENG, "Converts a binary number to hexadecimal.", `${NUM("The binary number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["BIN2OCT", ENG, "Converts a binary number to octal.", `${NUM("The binary number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["DEC2BIN", ENG, "Converts a decimal number to binary.", `${NUM("The decimal integer you want to convert (-512 to 511)")} | ${BASE_PLACES}`],
  ["DEC2HEX", ENG, "Converts a decimal number to hexadecimal.", `${NUM("The decimal integer you want to convert")} | ${BASE_PLACES}`],
  ["DEC2OCT", ENG, "Converts a decimal number to octal.", `${NUM("The decimal integer you want to convert")} | ${BASE_PLACES}`],
  ["HEX2BIN", ENG, "Converts a hexadecimal number to binary.", `${NUM("The hexadecimal number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["HEX2DEC", ENG, "Converts a hexadecimal number to decimal.", NUM("The hexadecimal number you want to convert (up to 10 characters)")],
  ["HEX2OCT", ENG, "Converts a hexadecimal number to octal.", `${NUM("The hexadecimal number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["OCT2BIN", ENG, "Converts an octal number to binary.", `${NUM("The octal number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["OCT2DEC", ENG, "Converts an octal number to decimal.", NUM("The octal number you want to convert (up to 10 characters)")],
  ["OCT2HEX", ENG, "Converts an octal number to hexadecimal.", `${NUM("The octal number you want to convert (up to 10 characters)")} | ${BASE_PLACES}`],
  ["BITAND", ENG, "Returns a bitwise 'And' of two numbers.", "number1: A nonnegative integer less than 2^48 | number2: A nonnegative integer less than 2^48"],
  ["BITOR", ENG, "Returns a bitwise 'Or' of two numbers.", "number1: A nonnegative integer less than 2^48 | number2: A nonnegative integer less than 2^48"],
  ["BITXOR", ENG, "Returns a bitwise 'Exclusive Or' of two numbers.", "number1: A nonnegative integer less than 2^48 | number2: A nonnegative integer less than 2^48"],
  ["BITLSHIFT", ENG, "Returns a number shifted left by shift_amount bits.", "number: A nonnegative integer less than 2^48 | shift_amount: The number of bits to shift; negative values shift right"],
  ["BITRSHIFT", ENG, "Returns a number shifted right by shift_amount bits.", "number: A nonnegative integer less than 2^48 | shift_amount: The number of bits to shift; negative values shift left"],
  ["DELTA", ENG, "Tests whether two values are equal, returning 1 if they are and 0 otherwise.", "number1: The first number | [number2]: The second number (default 0)"],
  ["GESTEP", ENG, "Tests whether a number is greater than or equal to a threshold value, returning 1 if it is and 0 otherwise.", "number: The value to test against step | [step]: The threshold value (default 0)"],
  ["CONVERT", ENG, "Converts a number from one measurement system to another.", "number: The value in from_unit to convert | from_unit: The units for number, such as \"m\", \"lbm\", or \"C\" | to_unit: The units for the result"],
  ["ERF", ENG, "Returns the error function integrated between lower_limit and upper_limit.", "lower_limit: The lower bound for integrating ERF | [upper_limit]: The upper bound for integrating ERF; if omitted, integrates from 0 to lower_limit"],
  ["ERF.PRECISE", ENG, "Returns the error function integrated between 0 and x.", "x: The upper bound for integrating ERF.PRECISE"],
  ["ERFC", ENG, "Returns the complementary error function integrated between x and infinity.", "x: The lower bound for integrating ERFC"],
  ["ERFC.PRECISE", ENG, "Returns the complementary error function integrated between x and infinity.", "x: The lower bound for integrating ERFC.PRECISE"],
  ["BESSELI", ENG, "Returns the modified Bessel function In(x).", "x: The value at which to evaluate the function | n: The order of the Bessel function"],
  ["BESSELJ", ENG, "Returns the Bessel function Jn(x).", "x: The value at which to evaluate the function | n: The order of the Bessel function"],
  ["BESSELK", ENG, "Returns the modified Bessel function Kn(x).", "x: The value at which to evaluate the function | n: The order of the function"],
  ["BESSELY", ENG, "Returns the Bessel function Yn(x), also called the Weber function or the Neumann function.", "x: The value at which to evaluate the function | n: The order of the function"],
  ["COMPLEX", ENG, "Converts real and imaginary coefficients into a complex number.", "real_num: The real coefficient of the complex number | i_num: The imaginary coefficient of the complex number | [suffix]: The suffix for the imaginary component, \"i\" (default) or \"j\""],
  ["IMREAL", ENG, "Returns the real coefficient of a complex number.", INUMBER],
  ["IMAGINARY", ENG, "Returns the imaginary coefficient of a complex number.", INUMBER],
  ["IMABS", ENG, "Returns the absolute value (modulus) of a complex number.", INUMBER],
  ["IMARGUMENT", ENG, "Returns the argument theta, an angle expressed in radians.", INUMBER],
  ["IMCONJUGATE", ENG, "Returns the complex conjugate of a complex number.", INUMBER],
  ["IMSUM", ENG, "Returns the sum of complex numbers.", "inumber1: The first complex number to add | [inumber2]...: Additional complex numbers to add"],
  ["IMSUB", ENG, "Returns the difference between two complex numbers.", "inumber1: The complex number from which to subtract inumber2 | inumber2: The complex number to subtract from inumber1"],
  ["IMPRODUCT", ENG, "Returns the product of complex numbers.", "inumber1: The first complex number to multiply | [inumber2]...: Additional complex numbers to multiply"],
  ["IMDIV", ENG, "Returns the quotient of two complex numbers.", "inumber1: The complex numerator or dividend | inumber2: The complex denominator or divisor"],
  ["IMPOWER", ENG, "Returns a complex number raised to a power.", `${INUMBER} | number: The power to which you want to raise the complex number`],
  ["IMSQRT", ENG, "Returns the square root of a complex number.", INUMBER],
  ["IMEXP", ENG, "Returns the exponential of a complex number.", INUMBER],
  ["IMLN", ENG, "Returns the natural logarithm of a complex number.", INUMBER],
  ["IMLOG10", ENG, "Returns the base-10 logarithm of a complex number.", INUMBER],
  ["IMLOG2", ENG, "Returns the base-2 logarithm of a complex number.", INUMBER],
  ["IMSIN", ENG, "Returns the sine of a complex number.", INUMBER],
  ["IMCOS", ENG, "Returns the cosine of a complex number.", INUMBER],
  ["IMTAN", ENG, "Returns the tangent of a complex number.", INUMBER],
  ["IMSEC", ENG, "Returns the secant of a complex number.", INUMBER],
  ["IMCSC", ENG, "Returns the cosecant of a complex number.", INUMBER],
  ["IMCOT", ENG, "Returns the cotangent of a complex number.", INUMBER],
  ["IMSINH", ENG, "Returns the hyperbolic sine of a complex number.", INUMBER],
  ["IMCOSH", ENG, "Returns the hyperbolic cosine of a complex number.", INUMBER],
  ["IMSECH", ENG, "Returns the hyperbolic secant of a complex number.", INUMBER],
  ["IMCSCH", ENG, "Returns the hyperbolic cosecant of a complex number.", INUMBER],

  // Database
  ["DSUM", DB, "Adds the numbers in the field (column) of records in the database that match the conditions you specify.", DATABASE],
  ["DAVERAGE", DB, "Averages the values in a field (column) of records in a list or database that match conditions you specify.", DATABASE],
  ["DCOUNT", DB, "Counts the cells that contain numbers in a field (column) of records in a database that match the conditions you specify.", DATABASE],
  ["DCOUNTA", DB, "Counts the nonblank cells in a field (column) of records in a database that match the conditions you specify.", DATABASE],
  ["DGET", DB, "Extracts from a database a single record that matches the conditions you specify.", DATABASE],
  ["DMAX", DB, "Returns the largest number in a field (column) of records in a database that match the conditions you specify.", DATABASE],
  ["DMIN", DB, "Returns the smallest number in a field (column) of records in a database that match the conditions you specify.", DATABASE],
  ["DPRODUCT", DB, "Multiplies the values in a field (column) of records in a database that match the conditions you specify.", DATABASE],
  ["DSTDEV", DB, "Estimates the standard deviation of a population based on a sample by using the numbers in a field of records that match the conditions you specify.", DATABASE],
  ["DSTDEVP", DB, "Calculates the standard deviation of a population based on the entire population by using the numbers in a field of records that match the conditions you specify.", DATABASE],
  ["DVAR", DB, "Estimates the variance of a population based on a sample by using the numbers in a field of records that match the conditions you specify.", DATABASE],
  ["DVARP", DB, "Calculates the variance of a population based on the entire population by using the numbers in a field of records that match the conditions you specify.", DATABASE],

  // Operator functions (Google Sheets)
  ["ADD", OP, "Returns the sum of two numbers; equivalent to the + operator.", "value1: The first addend | value2: The second addend"],
  ["MINUS", OP, "Returns the difference of two numbers; equivalent to the - operator.", "value1: The minuend, or number to be subtracted from | value2: The subtrahend, or number to subtract from value1"],
  ["MULTIPLY", OP, "Returns the product of two numbers; equivalent to the * operator.", "factor1: The first multiplicand | factor2: The second multiplicand"],
  ["DIVIDE", OP, "Returns one number divided by another; equivalent to the / operator.", "dividend: The number to be divided | divisor: The number to divide by"],
  ["POW", OP, "Returns a number raised to a power; equivalent to the ^ operator.", "base: The number to raise to the exponent power | exponent: The exponent to raise base to"],
  ["EQ", OP, "Returns TRUE if two specified values are equal and FALSE otherwise; equivalent to the = operator.", OPERANDS],
  ["NE", OP, "Returns TRUE if two specified values are not equal and FALSE otherwise; equivalent to the <> operator.", OPERANDS],
  ["GT", OP, "Returns TRUE if the first argument is strictly greater than the second, and FALSE otherwise; equivalent to the > operator.", OPERANDS],
  ["GTE", OP, "Returns TRUE if the first argument is greater than or equal to the second, and FALSE otherwise; equivalent to the >= operator.", OPERANDS],
  ["LT", OP, "Returns TRUE if the first argument is strictly less than the second, and FALSE otherwise; equivalent to the < operator.", OPERANDS],
  ["LTE", OP, "Returns TRUE if the first argument is less than or equal to the second, and FALSE otherwise; equivalent to the <= operator.", OPERANDS],
  ["UMINUS", OP, "Returns a number with the sign reversed.", "value: The number to negate"],
  ["UPLUS", OP, "Returns a specified number, unchanged.", "value: The number to return"],
  ["UNARY_PERCENT", OP, "Returns a value interpreted as a percentage; that is, divided by 100.", "percentage: The value to interpret as a percentage"],
];

// ---- Parsing --------------------------------------------------------------------------------

function parseArguments(source: string): FunctionArgumentInfo[] {
  if (!source.trim()) return [];
  return source.split(" | ").map((part) => {
    const separator = part.indexOf(": ");
    let name = (separator >= 0 ? part.slice(0, separator) : part).trim();
    const description = separator >= 0 ? part.slice(separator + 2).trim() : "";
    const info: FunctionArgumentInfo = { name, description };
    if (name.endsWith("...")) {
      name = name.slice(0, -3);
      info.repeating = true;
    }
    if (name.startsWith("[") && name.endsWith("]")) {
      name = name.slice(1, -1);
      info.optional = true;
    }
    info.name = name;
    return info;
  });
}

/** Metadata for every worksheet function the engine supports, keyed by upper-case name. */
export const FUNCTION_CATALOG: Record<string, FunctionInfo> = {};
for (const [name, category, description, args, returnsArray] of ROWS) {
  const info: FunctionInfo = { name, category, description, args: parseArguments(args) };
  if (returnsArray) info.returnsArray = true;
  FUNCTION_CATALOG[name] = info;
}

// ---- Signatures -----------------------------------------------------------------------------

/** Excel-style signature, e.g. "SUMIF(range, criteria, [sum_range])" or "SUM(number1, [number2], ...)". */
export function functionSignature(info: FunctionInfo): string {
  const parts: string[] = [];
  info.args.forEach((argument, index) => {
    parts.push(argument.optional ? `[${argument.name}]` : argument.name);
    // "..." follows the last argument of a repeating group.
    if (argument.repeating && !info.args[index + 1]?.repeating) parts.push("...");
  });
  return `${info.name}(${parts.join(", ")})`;
}

// ---- Lookup and search ----------------------------------------------------------------------

function normalizeName(name: string): string {
  return String(name ?? "")
    .trim()
    .replace(/^=+/, "")
    .replace(/\(.*$/, "")
    .trim()
    .replace(/^(?:_xlfn\.)?(?:_xlws\.)?/i, "")
    .toUpperCase();
}

/** Catalog entry for a function name (case-insensitive; `_xlfn.`/`_xlws.` prefixes allowed). */
export function getFunctionInfo(name: string): FunctionInfo | undefined {
  const key = normalizeName(name);
  return Object.prototype.hasOwnProperty.call(FUNCTION_CATALOG, key) ? FUNCTION_CATALOG[key] : undefined;
}

// Most frequently used functions first; drives autocomplete ranking.
const POPULAR = [
  "SUM", "IF", "VLOOKUP", "XLOOKUP", "COUNTIF", "SUMIF", "IFERROR", "INDEX", "MATCH", "AVERAGE",
  "COUNT", "COUNTA", "MAX", "MIN", "ROUND", "CONCAT", "TEXT", "LEFT", "RIGHT", "MID", "LEN", "TRIM",
  "TODAY", "NOW", "DATE", "SUMIFS", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "FILTER", "SORT", "UNIQUE",
  "XMATCH", "TEXTJOIN", "SUBSTITUTE", "AND", "OR", "NOT", "IFS", "SWITCH", "LET", "LAMBDA", "YEAR",
  "MONTH", "DAY", "EOMONTH", "NETWORKDAYS", "WORKDAY", "PMT", "NPV", "IRR", "SUMPRODUCT", "HLOOKUP",
  "ROUNDUP", "ROUNDDOWN", "INT", "ABS", "MOD", "UPPER", "LOWER", "PROPER", "FIND", "SEARCH", "VALUE",
  "CONCATENATE", "IFNA", "ISBLANK", "ISNUMBER", "ISERROR", "ISTEXT", "COUNTBLANK", "MEDIAN", "LARGE",
  "SMALL", "RANK", "STDEV", "STDEV.S", "OFFSET", "INDIRECT", "ROW", "COLUMN", "ROWS", "COLUMNS",
  "TRANSPOSE", "SEQUENCE", "CHOOSE", "DATEDIF", "WEEKDAY", "EDATE", "HOUR", "MINUTE", "SECOND", "TIME",
  "DAYS", "FV", "PV", "RATE", "NPER", "SUBTOTAL", "AGGREGATE", "MAXIFS", "MINIFS", "TEXTSPLIT",
  "TEXTBEFORE", "TEXTAFTER", "HSTACK", "VSTACK", "TAKE", "DROP", "SORTBY", "CHOOSECOLS", "CHOOSEROWS",
  "POWER", "SQRT", "CEILING", "FLOOR", "RAND", "RANDBETWEEN", "RANDARRAY", "PRODUCT", "REPT", "EXACT",
  "REPLACE", "TRUNC", "LOOKUP", "HYPERLINK", "DATEVALUE", "TIMEVALUE", "WEEKNUM", "YEARFRAC",
  "NETWORKDAYS.INTL", "WORKDAY.INTL", "MAP", "REDUCE", "BYROW", "BYCOL", "SCAN", "MAKEARRAY",
  "TOCOL", "TOROW", "GROUPBY", "PIVOTBY", "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE", "MODE",
  "PERCENTILE", "QUARTILE", "CORREL", "FORECAST", "TREND", "FREQUENCY", "NORM.DIST", "CHAR", "CODE",
  "CLEAN", "FIXED", "DOLLAR", "NUMBERVALUE", "ADDRESS", "CELL", "ISNA", "ISLOGICAL", "N", "T", "TYPE",
  "XNPV", "XIRR", "IPMT", "PPMT", "EFFECT", "SLN", "DB", "DDB", "CONVERT", "DEC2BIN", "DEC2HEX",
  "COUNTUNIQUE", "SPLIT", "JOIN", "REGEXMATCH", "ARRAYFORMULA", "DSUM", "DCOUNT", "DGET",
];
const POPULARITY = new Map(POPULAR.map((name, index) => [name, index]));

// Excel "Compatibility" functions: ranked after their modern replacements.
const LEGACY = new Set([
  "BETADIST", "BETAINV", "BINOMDIST", "CHIDIST", "CHIINV", "CHITEST", "CONFIDENCE", "COVAR",
  "CRITBINOM", "EXPONDIST", "FDIST", "FINV", "FTEST", "GAMMADIST", "GAMMAINV", "HYPGEOMDIST",
  "LOGINV", "LOGNORMDIST", "NEGBINOMDIST", "NORMDIST", "NORMINV", "NORMSDIST", "NORMSINV",
  "PERCENTRANK", "POISSON", "TDIST", "TINV", "TTEST", "WEIBULL", "ZTEST", "STDEVP", "VAR", "VARP",
]);

function rankKey(name: string): number {
  const popularity = POPULARITY.get(name);
  if (popularity !== undefined) return popularity;
  return LEGACY.has(name) ? 20_000 : 10_000;
}

/** All entries ordered by popularity, then modern before legacy, then alphabetically. */
const RANKED: FunctionInfo[] = Object.values(FUNCTION_CATALOG).sort((left, right) => {
  const difference = rankKey(left.name) - rankKey(right.name);
  if (difference !== 0) return difference;
  return left.name < right.name ? -1 : left.name > right.name ? 1 : 0;
});

/** Whether `query` starts at a segment boundary after the first one (".DIST" in NORM.S.DIST). */
function matchesSegment(name: string, query: string): boolean {
  for (let index = 1; index < name.length; index += 1) {
    const previous = name[index - 1];
    if ((previous === "." || previous === "_") && name.startsWith(query, index)) return true;
  }
  return false;
}

/**
 * Functions matching what the user typed, for autocomplete: names starting with the text come
 * first, then names with a later segment starting with it (DIST → NORM.DIST), then other
 * substring matches; each group is ranked by popularity. An empty prefix returns the most
 * popular functions. A leading "=" and `_xlfn.`/`_xlws.` prefixes are ignored.
 */
export function searchFunctions(prefix: string, limit = 20): FunctionInfo[] {
  const count = Math.max(0, Math.floor(Number.isFinite(limit) ? limit : RANKED.length));
  if (count === 0) return [];
  const query = normalizeName(prefix);
  if (!query) return RANKED.slice(0, count);
  const starts: FunctionInfo[] = [];
  const segments: FunctionInfo[] = [];
  const contains: FunctionInfo[] = [];
  for (const info of RANKED) {
    if (info.name.startsWith(query)) starts.push(info);
    else if (matchesSegment(info.name, query)) segments.push(info);
    else if (info.name.includes(query)) contains.push(info);
  }
  return [...starts, ...segments, ...contains].slice(0, count);
}
