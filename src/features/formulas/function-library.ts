import { functionSignature, nativeFunctions, type NativeFunction } from '../../calculation/formula-functions';

export type FunctionLibraryCategory =
  | 'Financial'
  | 'Logical'
  | 'Text'
  | 'Date & Time'
  | 'Lookup & Reference'
  | 'Math & Trig'
  | 'Statistical'
  | 'Information';

export type InsertFunctionFilter = 'Most Recently Used' | 'All' | FunctionLibraryCategory;

export const topLibraryCategories: readonly FunctionLibraryCategory[] = [
  'Financial',
  'Logical',
  'Text',
  'Date & Time',
  'Lookup & Reference',
  'Math & Trig',
];

export const moreLibraryCategories: readonly FunctionLibraryCategory[] = [
  'Statistical',
  'Information',
];

export const allLibraryCategories: readonly FunctionLibraryCategory[] = [
  ...topLibraryCategories,
  ...moreLibraryCategories,
];

export const dialogCategoryOptions: readonly InsertFunctionFilter[] = [
  'Most Recently Used',
  'All',
  'Financial',
  'Date & Time',
  'Math & Trig',
  'Statistical',
  'Lookup & Reference',
  'Text',
  'Logical',
  'Information',
];

export const defaultRecentFunctions: readonly NativeFunction[] = [
  'SUM',
  'AVERAGE',
  'IF',
  'XLOOKUP',
  'VLOOKUP',
  'FILTER',
  'COUNT',
  'MAX',
  'SUMIF',
  'TEXTJOIN',
];

export const maxRecentFunctions = 10;

interface FunctionMetadata {
  readonly category: FunctionLibraryCategory;
  readonly description: string;
}

const catalog: Record<NativeFunction, FunctionMetadata> = {
  // Financial (14)
  DB: { category: 'Financial', description: 'Returns depreciation of an asset using the fixed-declining balance method.' },
  DDB: { category: 'Financial', description: 'Returns depreciation of an asset using the double-declining balance method.' },
  FV: { category: 'Financial', description: 'Returns the future value of an investment based on periodic, constant payments.' },
  IPMT: { category: 'Financial', description: 'Returns the interest payment for a given period of an investment.' },
  IRR: { category: 'Financial', description: 'Returns the internal rate of return for a series of periodic cash flows.' },
  NPER: { category: 'Financial', description: 'Returns the number of periods for an investment.' },
  NPV: { category: 'Financial', description: 'Returns the net present value of an investment from a discount rate and cash flows.' },
  PMT: { category: 'Financial', description: 'Calculates the payment for a loan based on constant payments and interest rate.' },
  PPMT: { category: 'Financial', description: 'Returns the payment on the principal for a given investment period.' },
  PV: { category: 'Financial', description: 'Returns the present value of an investment.' },
  RATE: { category: 'Financial', description: 'Returns the interest rate per period of an annuity.' },
  SLN: { category: 'Financial', description: 'Returns the straight-line depreciation of an asset for one period.' },
  XIRR: { category: 'Financial', description: 'Returns the internal rate of return for a schedule of cash flows with dates.' },
  XNPV: { category: 'Financial', description: 'Returns the net present value for a schedule of cash flows with dates.' },

  // Logical (19)
  AND: { category: 'Logical', description: 'Returns TRUE if all of its arguments evaluate to TRUE.' },
  BYCOL: { category: 'Logical', description: 'Applies a LAMBDA to each column and returns an array of results.' },
  BYROW: { category: 'Logical', description: 'Applies a LAMBDA to each row and returns an array of results.' },
  FALSE: { category: 'Logical', description: 'Returns the logical value FALSE.' },
  IF: { category: 'Logical', description: 'Checks a condition and returns one value if TRUE, another if FALSE.' },
  IFERROR: { category: 'Logical', description: 'Returns value_if_error if a formula evaluates to an error; otherwise returns the value.' },
  IFNA: { category: 'Logical', description: 'Returns value_if_na if the expression resolves to #N/A; otherwise returns the value.' },
  IFS: { category: 'Logical', description: 'Checks multiple conditions in order and returns the value for the first TRUE result.' },
  LAMBDA: { category: 'Logical', description: 'Creates a reusable custom function defined by parameter names and a calculation.' },
  LET: { category: 'Logical', description: 'Assigns names to calculation results for reuse within a formula.' },
  MAKEARRAY: { category: 'Logical', description: 'Creates an array of a specified size by applying a LAMBDA calculation.' },
  MAP: { category: 'Logical', description: 'Applies a LAMBDA to each element in one or more arrays to return a new array.' },
  NOT: { category: 'Logical', description: 'Reverses the logical value of its argument.' },
  OR: { category: 'Logical', description: 'Returns TRUE if any argument evaluates to TRUE.' },
  REDUCE: { category: 'Logical', description: 'Reduces an array to a single accumulated value using a LAMBDA.' },
  SCAN: { category: 'Logical', description: 'Scans an array and returns each intermediate accumulated value via a LAMBDA.' },
  SWITCH: { category: 'Logical', description: 'Evaluates an expression against a list of values and returns the matching result.' },
  TRUE: { category: 'Logical', description: 'Returns the logical value TRUE.' },
  XOR: { category: 'Logical', description: 'Returns TRUE if an odd number of arguments evaluate to TRUE.' },

  // Text (33)
  ARRAYTOTEXT: { category: 'Text', description: 'Returns an array of text values from any specified range.' },
  CHAR: { category: 'Text', description: 'Returns the character specified by a Windows-1252 code number.' },
  CLEAN: { category: 'Text', description: 'Removes all nonprintable control characters from text.' },
  CODE: { category: 'Text', description: 'Returns the numeric code for the first character in a text string.' },
  CONCAT: { category: 'Text', description: 'Combines text from multiple ranges and strings.' },
  CONCATENATE: { category: 'Text', description: 'Joins several text strings into one text string.' },
  DOLLAR: { category: 'Text', description: 'Formats a number as text using currency formatting.' },
  EXACT: { category: 'Text', description: 'Checks whether two text strings are identical, including case.' },
  FIND: { category: 'Text', description: 'Finds one text string within another (case-sensitive).' },
  FIXED: { category: 'Text', description: 'Formats a number as text with a fixed number of decimals.' },
  LEFT: { category: 'Text', description: 'Returns the specified number of characters from the start of a text string.' },
  LEN: { category: 'Text', description: 'Returns the number of characters in a text string.' },
  LOWER: { category: 'Text', description: 'Converts all uppercase letters in a text string to lowercase.' },
  MID: { category: 'Text', description: 'Returns characters from the middle of a text string given a starting position.' },
  NUMBERVALUE: { category: 'Text', description: 'Converts text to a number in a locale-independent way using explicit separators.' },
  PROPER: { category: 'Text', description: 'Capitalizes the first letter of each word in a text string.' },
  REPLACE: { category: 'Text', description: 'Replaces part of a text string with a different text string by position.' },
  REPT: { category: 'Text', description: 'Repeats text a given number of times.' },
  RIGHT: { category: 'Text', description: 'Returns the specified number of characters from the end of a text string.' },
  SEARCH: { category: 'Text', description: 'Finds text within another string using wildcards (case-insensitive).' },
  SUBSTITUTE: { category: 'Text', description: 'Substitutes new text for existing text in a text string.' },
  T: { category: 'Text', description: 'Returns the text referred to by a value, or empty text if not a string.' },
  TEXT: { category: 'Text', description: 'Formats a number and converts it to text according to a format code.' },
  TEXTAFTER: { category: 'Text', description: 'Returns text that occurs after a given delimiter.' },
  TEXTBEFORE: { category: 'Text', description: 'Returns text that occurs before a given delimiter.' },
  TEXTJOIN: { category: 'Text', description: 'Combines text from ranges or strings with a delimiter, optionally skipping blanks.' },
  TEXTSPLIT: { category: 'Text', description: 'Splits text strings by column and row delimiters into a spilled array.' },
  TRIM: { category: 'Text', description: 'Removes extra leading, trailing, and repeated spaces from text.' },
  UNICHAR: { category: 'Text', description: 'Returns the Unicode character referenced by a numeric code point.' },
  UNICODE: { category: 'Text', description: 'Returns the Unicode code point of the first character of text.' },
  UPPER: { category: 'Text', description: 'Converts text to uppercase.' },
  VALUE: { category: 'Text', description: 'Converts a text string that represents a number, date, or time into a number.' },
  VALUETOTEXT: { category: 'Text', description: 'Returns text representing any scalar value in concise or strict form.' },

  // Date & Time (25)
  DATE: { category: 'Date & Time', description: 'Returns the Excel serial number for a specified year, month, and day.' },
  DATEDIF: { category: 'Date & Time', description: 'Calculates the number of days, months, or years between two dates.' },
  DATEVALUE: { category: 'Date & Time', description: 'Converts a date stored as text into an Excel date serial number.' },
  DAY: { category: 'Date & Time', description: 'Returns the day of the month (1–31) from a date serial number.' },
  DAYS: { category: 'Date & Time', description: 'Returns the number of days between two dates.' },
  DAYS360: { category: 'Date & Time', description: 'Calculates the number of days between two dates based on a 360-day year.' },
  EDATE: { category: 'Date & Time', description: 'Returns the date serial number a given number of months before or after start_date.' },
  EOMONTH: { category: 'Date & Time', description: 'Returns the last day of the month a specified number of months from start_date.' },
  HOUR: { category: 'Date & Time', description: 'Returns the hour component (0–23) of a time serial value.' },
  ISOWEEKNUM: { category: 'Date & Time', description: 'Returns the ISO 8601 week number of the year for a given date.' },
  MINUTE: { category: 'Date & Time', description: 'Returns the minute component (0–59) of a time serial value.' },
  MONTH: { category: 'Date & Time', description: 'Returns the month of the year (1–12) from a date serial number.' },
  NETWORKDAYS: { category: 'Date & Time', description: 'Returns the number of whole workdays between two dates.' },
  'NETWORKDAYS.INTL': { category: 'Date & Time', description: 'Returns the number of whole workdays between two dates with custom weekends.' },
  NOW: { category: 'Date & Time', description: 'Returns the current date and time serial number.' },
  SECOND: { category: 'Date & Time', description: 'Returns the second component (0–59) of a time serial value.' },
  TIME: { category: 'Date & Time', description: 'Returns the decimal time serial for a specified hour, minute, and second.' },
  TIMEVALUE: { category: 'Date & Time', description: 'Converts a time stored as text into a decimal time serial number.' },
  TODAY: { category: 'Date & Time', description: 'Returns the current date serial number.' },
  WEEKDAY: { category: 'Date & Time', description: 'Returns a number identifying the day of the week for a date.' },
  WEEKNUM: { category: 'Date & Time', description: 'Returns the week number in the year for a given date.' },
  WORKDAY: { category: 'Date & Time', description: 'Returns the date a given number of working days before or after start_date.' },
  'WORKDAY.INTL': { category: 'Date & Time', description: 'Returns the date a number of working days away with custom weekend days.' },
  YEAR: { category: 'Date & Time', description: 'Returns the year (1900–9999) corresponding to a date serial number.' },
  YEARFRAC: { category: 'Date & Time', description: 'Calculates the fraction of a year represented by whole days between two dates.' },

  // Lookup & Reference (35)
  ADDRESS: { category: 'Lookup & Reference', description: 'Creates a cell address string from row and column numbers.' },
  AREAS: { category: 'Lookup & Reference', description: 'Returns the number of rectangular areas in a reference.' },
  CHOOSE: { category: 'Lookup & Reference', description: 'Uses index_num to select and return a value from the list of arguments.' },
  CHOOSECOLS: { category: 'Lookup & Reference', description: 'Returns the specified columns from an array or range.' },
  CHOOSEROWS: { category: 'Lookup & Reference', description: 'Returns the specified rows from an array or range.' },
  COLUMN: { category: 'Lookup & Reference', description: 'Returns the column number of a reference or the current formula cell.' },
  COLUMNS: { category: 'Lookup & Reference', description: 'Returns the number of columns in an array or reference.' },
  DROP: { category: 'Lookup & Reference', description: 'Excludes a specified number of rows or columns from the start or end of an array.' },
  EXPAND: { category: 'Lookup & Reference', description: 'Expands or pads an array to specified row and column dimensions.' },
  FILTER: { category: 'Lookup & Reference', description: 'Filters a range of data based on criteria you define.' },
  FORMULATEXT: { category: 'Lookup & Reference', description: 'Returns a formula as a text string from the referenced cell.' },
  GETPIVOTDATA: { category: 'Lookup & Reference', description: 'Returns visible aggregated data stored in a PivotTable report.' },
  HLOOKUP: { category: 'Lookup & Reference', description: 'Searches the top row of a table and returns a value in the specified row.' },
  HSTACK: { category: 'Lookup & Reference', description: 'Appends arrays horizontally into a single wider array.' },
  HYPERLINK: { category: 'Lookup & Reference', description: 'Returns friendly text or link text for a hyperlink formula.' },
  INDEX: { category: 'Lookup & Reference', description: 'Returns a cell or slice from a range or array by row and column number.' },
  INDIRECT: { category: 'Lookup & Reference', description: 'Returns the reference specified by a text string.' },
  LOOKUP: { category: 'Lookup & Reference', description: 'Looks up a value from a one-row or one-column range or array.' },
  MATCH: { category: 'Lookup & Reference', description: 'Searches for an item in a range and returns its relative position.' },
  OFFSET: { category: 'Lookup & Reference', description: 'Returns a reference offset from a starting cell or range.' },
  ROW: { category: 'Lookup & Reference', description: 'Returns the row number of a reference or the current formula cell.' },
  ROWS: { category: 'Lookup & Reference', description: 'Returns the number of rows in an array or reference.' },
  SORT: { category: 'Lookup & Reference', description: 'Sorts the contents of a range or array.' },
  SORTBY: { category: 'Lookup & Reference', description: 'Sorts a range or array based on values in corresponding arrays.' },
  TAKE: { category: 'Lookup & Reference', description: 'Returns a specified number of contiguous rows or columns from an array.' },
  TOCOL: { category: 'Lookup & Reference', description: 'Returns an array flattened into a single column.' },
  TOROW: { category: 'Lookup & Reference', description: 'Returns an array flattened into a single row.' },
  TRANSPOSE: { category: 'Lookup & Reference', description: 'Converts a vertical range of cells to a horizontal range, or vice versa.' },
  UNIQUE: { category: 'Lookup & Reference', description: 'Returns a list of unique values or rows from a range or array.' },
  VLOOKUP: { category: 'Lookup & Reference', description: 'Searches the first column of a table and returns a value in the same row.' },
  VSTACK: { category: 'Lookup & Reference', description: 'Appends arrays vertically into a single taller array.' },
  WRAPCOLS: { category: 'Lookup & Reference', description: 'Wraps a 1D row or column of values into a 2D array by columns.' },
  WRAPROWS: { category: 'Lookup & Reference', description: 'Wraps a 1D row or column of values into a 2D array by rows.' },
  XLOOKUP: { category: 'Lookup & Reference', description: 'Searches a range or array and returns the corresponding item from another range.' },
  XMATCH: { category: 'Lookup & Reference', description: 'Searches for an item and returns its relative position with flexible match modes.' },

  // Math & Trig (49)
  ABS: { category: 'Math & Trig', description: 'Returns the absolute value of a number.' },
  ACOS: { category: 'Math & Trig', description: 'Returns the arccosine of a number, in radians.' },
  AGGREGATE: { category: 'Math & Trig', description: 'Returns an aggregate calculation, optionally ignoring hidden rows or errors.' },
  ASIN: { category: 'Math & Trig', description: 'Returns the arcsine of a number, in radians.' },
  ATAN: { category: 'Math & Trig', description: 'Returns the arctangent of a number, in radians.' },
  ATAN2: { category: 'Math & Trig', description: 'Returns the arctangent from x- and y-coordinates, in radians.' },
  CEILING: { category: 'Math & Trig', description: 'Rounds a number up to the nearest multiple of significance.' },
  'CEILING.MATH': { category: 'Math & Trig', description: 'Rounds a number up to the nearest integer or multiple of significance.' },
  COMBIN: { category: 'Math & Trig', description: 'Returns the number of combinations for a given number of items.' },
  COS: { category: 'Math & Trig', description: 'Returns the cosine of an angle given in radians.' },
  DEGREES: { category: 'Math & Trig', description: 'Converts radians into degrees.' },
  EVEN: { category: 'Math & Trig', description: 'Rounds a number away from zero to the nearest even integer.' },
  EXP: { category: 'Math & Trig', description: 'Returns e raised to the power of a given number.' },
  FACT: { category: 'Math & Trig', description: 'Returns the factorial of a number.' },
  FLOOR: { category: 'Math & Trig', description: 'Rounds a number down toward zero to the nearest multiple of significance.' },
  'FLOOR.MATH': { category: 'Math & Trig', description: 'Rounds a number down to the nearest integer or multiple of significance.' },
  GCD: { category: 'Math & Trig', description: 'Returns the greatest common divisor of two or more integers.' },
  INT: { category: 'Math & Trig', description: 'Rounds a number down to the nearest integer.' },
  LCM: { category: 'Math & Trig', description: 'Returns the least common multiple of integers.' },
  LN: { category: 'Math & Trig', description: 'Returns the natural logarithm of a number.' },
  LOG: { category: 'Math & Trig', description: 'Returns the logarithm of a number to a specified base.' },
  LOG10: { category: 'Math & Trig', description: 'Returns the base-10 logarithm of a number.' },
  MMULT: { category: 'Math & Trig', description: 'Returns the matrix product of two arrays.' },
  MOD: { category: 'Math & Trig', description: 'Returns the remainder after dividing a number by a divisor.' },
  MROUND: { category: 'Math & Trig', description: 'Returns a number rounded to the desired multiple.' },
  ODD: { category: 'Math & Trig', description: 'Rounds a number away from zero to the nearest odd integer.' },
  PI: { category: 'Math & Trig', description: 'Returns the mathematical constant pi.' },
  POWER: { category: 'Math & Trig', description: 'Returns the result of a number raised to a power.' },
  PRODUCT: { category: 'Math & Trig', description: 'Multiplies all the numbers given as arguments.' },
  QUOTIENT: { category: 'Math & Trig', description: 'Returns the integer portion of a division.' },
  RADIANS: { category: 'Math & Trig', description: 'Converts degrees to radians.' },
  RAND: { category: 'Math & Trig', description: 'Returns an evenly distributed random real number greater than or equal to 0 and less than 1.' },
  RANDARRAY: { category: 'Math & Trig', description: 'Returns an array of random numbers.' },
  RANDBETWEEN: { category: 'Math & Trig', description: 'Returns a random integer between two specified numbers.' },
  ROUND: { category: 'Math & Trig', description: 'Rounds a number to a specified number of digits.' },
  ROUNDDOWN: { category: 'Math & Trig', description: 'Rounds a number down toward zero.' },
  ROUNDUP: { category: 'Math & Trig', description: 'Rounds a number up away from zero.' },
  SEQUENCE: { category: 'Math & Trig', description: 'Generates a list of sequential numbers in an array.' },
  SIGN: { category: 'Math & Trig', description: 'Returns the sign of a number: 1, 0, or -1.' },
  SIN: { category: 'Math & Trig', description: 'Returns the sine of an angle given in radians.' },
  SQRT: { category: 'Math & Trig', description: 'Returns a positive square root.' },
  SUBTOTAL: { category: 'Math & Trig', description: 'Returns a subtotal in a list or database, skipping hidden or nested subtotal rows.' },
  SUM: { category: 'Math & Trig', description: 'Adds all the numbers in the arguments or ranges.' },
  SUMIF: { category: 'Math & Trig', description: 'Adds the cells specified by a given condition or criteria.' },
  SUMIFS: { category: 'Math & Trig', description: 'Adds cells that meet multiple criteria.' },
  SUMPRODUCT: { category: 'Math & Trig', description: 'Multiplies corresponding components in arrays and returns the sum of those products.' },
  SUMSQ: { category: 'Math & Trig', description: 'Returns the sum of the squares of the arguments.' },
  TAN: { category: 'Math & Trig', description: 'Returns the tangent of an angle given in radians.' },
  TRUNC: { category: 'Math & Trig', description: 'Truncates a number to an integer or specified decimal places.' },

  // Statistical (34)
  AVERAGE: { category: 'Statistical', description: 'Returns the arithmetic mean of its arguments.' },
  AVERAGEIF: { category: 'Statistical', description: 'Returns the average of all cells in a range that meet one criterion.' },
  AVERAGEIFS: { category: 'Statistical', description: 'Returns the average of all cells that meet multiple criteria.' },
  CORREL: { category: 'Statistical', description: 'Returns the correlation coefficient between two cell ranges or arrays.' },
  COUNT: { category: 'Statistical', description: 'Counts how many numbers are in the arguments or ranges.' },
  COUNTA: { category: 'Statistical', description: 'Counts the number of non-empty cells in a range.' },
  COUNTBLANK: { category: 'Statistical', description: 'Counts empty cells in a specified range.' },
  COUNTIF: { category: 'Statistical', description: 'Counts cells within a range that meet the given criterion.' },
  COUNTIFS: { category: 'Statistical', description: 'Counts cells that meet multiple criteria across ranges.' },
  'COVARIANCE.S': { category: 'Statistical', description: 'Returns sample covariance between two data sets.' },
  'FORECAST.LINEAR': { category: 'Statistical', description: 'Calculates a future value using linear regression.' },
  INTERCEPT: { category: 'Statistical', description: 'Calculates the y-intercept of a linear regression line.' },
  LARGE: { category: 'Statistical', description: 'Returns the k-th largest value in a data set.' },
  MAX: { category: 'Statistical', description: 'Returns the largest numeric value in a set of values.' },
  MAXIFS: { category: 'Statistical', description: 'Returns the maximum value among cells specified by multiple criteria.' },
  MEDIAN: { category: 'Statistical', description: 'Returns the median of the given numbers.' },
  MIN: { category: 'Statistical', description: 'Returns the smallest numeric value in a set of values.' },
  MINIFS: { category: 'Statistical', description: 'Returns the minimum value among cells specified by multiple criteria.' },
  MODE: { category: 'Statistical', description: 'Returns the most frequently occurring value in an array or range.' },
  'MODE.SNGL': { category: 'Statistical', description: 'Returns the most frequently occurring value in an array or range.' },
  'PERCENTILE.EXC': { category: 'Statistical', description: 'Returns the k-th percentile of values in a range, exclusive of 0 and 1.' },
  'PERCENTILE.INC': { category: 'Statistical', description: 'Returns the k-th percentile of values in a range, inclusive of 0 and 1.' },
  PERMUT: { category: 'Statistical', description: 'Returns the number of permutations for a given number of objects.' },
  'QUARTILE.EXC': { category: 'Statistical', description: 'Returns the quartile of a data set, exclusive of 0 and 1.' },
  'QUARTILE.INC': { category: 'Statistical', description: 'Returns the quartile of a data set, inclusive of 0 and 1.' },
  RANK: { category: 'Statistical', description: 'Returns the rank of a number in a list of numbers.' },
  'RANK.AVG': { category: 'Statistical', description: 'Returns the rank of a number, averaging ties.' },
  'RANK.EQ': { category: 'Statistical', description: 'Returns the rank of a number, assigning top rank to ties.' },
  SLOPE: { category: 'Statistical', description: 'Returns the slope of the linear regression line.' },
  SMALL: { category: 'Statistical', description: 'Returns the k-th smallest value in a data set.' },
  'STDEV.P': { category: 'Statistical', description: 'Calculates standard deviation based on the entire population.' },
  'STDEV.S': { category: 'Statistical', description: 'Estimates standard deviation based on a sample.' },
  'VAR.P': { category: 'Statistical', description: 'Calculates variance based on the entire population.' },
  'VAR.S': { category: 'Statistical', description: 'Estimates variance based on a sample.' },

  // Information (17)
  'ERROR.TYPE': { category: 'Information', description: 'Returns a number corresponding to an Excel error value.' },
  ISBLANK: { category: 'Information', description: 'Returns TRUE if the referenced cell is empty.' },
  ISERR: { category: 'Information', description: 'Returns TRUE if the value is any error other than #N/A.' },
  ISERROR: { category: 'Information', description: 'Returns TRUE if the value is any error value.' },
  ISEVEN: { category: 'Information', description: 'Returns TRUE if the number is even.' },
  ISFORMULA: { category: 'Information', description: 'Returns TRUE if the referenced cell contains a formula.' },
  ISLOGICAL: { category: 'Information', description: 'Returns TRUE if the value is a logical boolean.' },
  ISNA: { category: 'Information', description: 'Returns TRUE if the value is the #N/A error value.' },
  ISNONTEXT: { category: 'Information', description: 'Returns TRUE if the value is not text.' },
  ISNUMBER: { category: 'Information', description: 'Returns TRUE if the value is a number.' },
  ISODD: { category: 'Information', description: 'Returns TRUE if the number is odd.' },
  ISOMITTED: { category: 'Information', description: 'Checks whether a LAMBDA parameter value is missing.' },
  ISREF: { category: 'Information', description: 'Returns TRUE if the value is a reference.' },
  ISTEXT: { category: 'Information', description: 'Returns TRUE if the value is text.' },
  N: { category: 'Information', description: 'Converts a value to a number.' },
  NA: { category: 'Information', description: 'Returns the error value #N/A.' },
  TYPE: { category: 'Information', description: 'Returns a number indicating the data type of a value.' },
};

const sortedAllFunctions: readonly NativeFunction[] = [...nativeFunctions].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));

export const functionsByCategory: Readonly<Record<FunctionLibraryCategory, readonly NativeFunction[]>> = {
  Financial: sortedAllFunctions.filter(name => catalog[name].category === 'Financial'),
  Logical: sortedAllFunctions.filter(name => catalog[name].category === 'Logical'),
  Text: sortedAllFunctions.filter(name => catalog[name].category === 'Text'),
  'Date & Time': sortedAllFunctions.filter(name => catalog[name].category === 'Date & Time'),
  'Lookup & Reference': sortedAllFunctions.filter(name => catalog[name].category === 'Lookup & Reference'),
  'Math & Trig': sortedAllFunctions.filter(name => catalog[name].category === 'Math & Trig'),
  Statistical: sortedAllFunctions.filter(name => catalog[name].category === 'Statistical'),
  Information: sortedAllFunctions.filter(name => catalog[name].category === 'Information'),
};

export function functionCategory(name: NativeFunction): FunctionLibraryCategory {
  return catalog[name].category;
}

export function functionDescription(name: NativeFunction): string {
  return catalog[name].description;
}

export function functionSignatureText(name: NativeFunction, separator: ',' | ';' = ','): string {
  const sig = functionSignature(name);
  if (!sig) return `${name}()`;
  const [params, variadic] = sig;
  const list = variadic ? [...params, '…'] : params;
  return `${name}(${list.join(separator + ' ')})`;
}

const nativeSet = new Set<string>(nativeFunctions);

export function sanitizeRecentFunctions(stored: readonly string[] | undefined): readonly NativeFunction[] {
  if (!stored || !stored.length) return defaultRecentFunctions;
  const valid = stored.filter((name): name is NativeFunction => nativeSet.has(name));
  return valid.length ? valid.slice(0, maxRecentFunctions) : defaultRecentFunctions;
}

export function recordRecentFunction(current: readonly NativeFunction[], inserted: NativeFunction): readonly NativeFunction[] {
  return [inserted, ...current.filter(name => name !== inserted)].slice(0, maxRecentFunctions);
}

export function searchFunctions(
  query: string,
  filter: InsertFunctionFilter,
  recent: readonly NativeFunction[] = defaultRecentFunctions,
): readonly NativeFunction[] {
  const trimmed = query.trim().toLowerCase();
  if (trimmed) {
    return sortedAllFunctions.filter(name =>
      name.toLowerCase().includes(trimmed) || catalog[name].description.toLowerCase().includes(trimmed));
  }
  if (filter === 'Most Recently Used') return sanitizeRecentFunctions(recent);
  if (filter === 'All') return sortedAllFunctions;
  return functionsByCategory[filter];
}

export interface FunctionArgumentSpec {
  readonly name: string;
  readonly required: boolean;
}

export function initialFunctionArguments(name: NativeFunction): readonly FunctionArgumentSpec[] {
  const sig = functionSignature(name);
  if (!sig) return [];
  return sig[0].map(raw => {
    const required = !(raw.startsWith('[') && raw.endsWith(']'));
    return { name: required ? raw : raw.slice(1, -1), required };
  });
}

export function isVariadicFunction(name: NativeFunction): boolean {
  return Boolean(functionSignature(name)?.[1]);
}

function incrementTrailingNumber(label: string, fallbackIndex: number): string {
  const match = /^(.*?)(\d+)$/u.exec(label);
  if (match) return `${match[1]}${Number(match[2]) + 1}`;
  return `${label}${fallbackIndex}`;
}

export function nextVariadicArgument(
  name: NativeFunction,
  current: readonly FunctionArgumentSpec[],
): FunctionArgumentSpec | undefined {
  if (!isVariadicFunction(name) || current.length >= 30 || current.length === 0) return undefined;
  const last = current[current.length - 1]!;
  const prev = current[current.length - 2];
  const lastMatch = /^(.*?)(\d+)$/u.exec(last.name);
  const prevMatch = prev ? /^(.*?)(\d+)$/u.exec(prev.name) : null;
  if (lastMatch && prevMatch && lastMatch[2] === prevMatch[2] && lastMatch[1] !== prevMatch[1]) {
    // Paired variadic parameters (e.g. criteria_range1, criteria1 -> criteria_range2, criteria2).
    const nextNum = Number(lastMatch[2]) + 1;
    return { name: `${prevMatch[1]}${nextNum}`, required: false };
  }
  const ante = current[current.length - 3];
  const anteMatch = ante ? /^(.*?)(\d+)$/u.exec(ante.name) : null;
  if (prevMatch && anteMatch && prevMatch[1] !== anteMatch[1] && Number(prevMatch[2]) === Number(anteMatch[2])) {
    // Second item of a newly started pair -> match the prefix of prev and index of last.
    return { name: `${prevMatch[1]}${lastMatch ? lastMatch[2] : prevMatch[2]}`, required: false };
  }
  return { name: incrementTrailingNumber(last.name, current.length + 1), required: false };
}

export function buildFunctionFormula(
  name: NativeFunction,
  values: readonly string[],
  separator: ',' | ';',
): string {
  let end = values.length;
  while (end > 0 && values[end - 1]!.trim() === '') end--;
  const args = values.slice(0, end).map(value => value.trim());
  return `=${name}(${args.join(separator)})`;
}

export function formatArgumentReference(refersToText: string, activeSheetName: string): string {
  const withoutEq = refersToText.startsWith('=') ? refersToText.slice(1) : refersToText;
  const bang = withoutEq.lastIndexOf('!');
  if (bang <= 0) return withoutEq.replaceAll('$', '');
  const rawSheet = withoutEq.slice(0, bang);
  const relativeRef = withoutEq.slice(bang + 1).replaceAll('$', '');
  const unquoted = rawSheet.startsWith("'") && rawSheet.endsWith("'")
    ? rawSheet.slice(1, -1).replaceAll("''", "'")
    : rawSheet;
  return unquoted.toLowerCase() === activeSheetName.toLowerCase()
    ? relativeRef
    : `${rawSheet}!${relativeRef}`;
}
