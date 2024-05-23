# Excel Tricks

My commonly used Excel and Google Sheets formulas and tricks

## Content

- [Excel Tricks](#excel-tricks)
  - [Content](#content)
    - [Time and Date Formulas](#time-and-date-formulas)
      - [Convert the format "Thu Oct 02 12:03:39 GMT 2014" to "10/02/2014"](#convert-the-format-thu-oct-02-120339-gmt-2014-to-10022014)
      - [Convert the format "2014-Dec-01 5:00:54 AM" to "12/01/2014"](#convert-the-format-2014-dec-01-50054-am-to-12012014)
      - [Convert EPOCH format (Unix time) to Gregorian format (mm/dd/yyyy hh:mm:ss)](#convert-epoch-format-unix-time-to-gregorian-format-mmddyyyy-hhmmss)
      - [Convert a date and time field to ISO 8601 timestamp format](#convert-a-date-and-time-field-to-iso-8601-timestamp-format)
      - [Convert a ISO 8601 timestamp format field to date and time](#convert-a-iso-8601-timestamp-format-field-to-date-and-time)
      - [Get the quarter of the year from a date](#get-the-quarter-of-the-year-from-a-date)
    - [Number Manipulation](#number-manipulation)
      - [Convert $20,000,000.00 to $20.0M](#convert-2000000000-to-200m)
    - [Text Manipulation](#text-manipulation)
      - [Extract the domain name from an email address ](#extract-only-the-domain-name-from-an-email-address)
      - [Find what is to the RIGHT of the last instances of a specific character](#find-what-is-to-the-right-of-the-last-instances-of-a-specific-character)
      - [Find if cell contains a space](#find-if-cell-contains-a-space)
      - [Extract text between two characters in a cell](#extract-text-between-two-characters-in-a-cell)
      - [Trim All Whitespace Including Nonbreaking Space Characters (nbsp)](#trim-all-whitespace-including-nonbreaking-space-characters-nbsp)
      - [VLookUp and Replace #N/A with some text](#vlookup-and-replace-na-with-some-text)
      - [Search for text within a cell and label it as X](#search-for-text-within-a-cell-and-label-it-as-x)
      - [Lookup a Value in 2 Different Columns and return the one you want](#lookup-a-value-in-2-different-columns-and-return-the-one-you-want)
      - [Get OS Short name from long Operating System name (Windows 10 Enterprise = Windows)](#get-os-short-name-from-long-operating-system-name-windows-10-enterprise--windows)
      - [Get system type from OS (Windows Serer 2012 = Server)](#get-system-type-from-os-windows-serer-2012--server)

### Time and Date Formulas

#### Convert the format "Thu Oct 02 12:03:39 GMT 2014" to "10/02/2014"

``` bash
=CONCATENATE("10/",MID(A2,9,2),"/2014")
```

#### Convert the format "2014-Dec-01 5:00:54 AM" to "12/01/2014"

- Perform a Text-to-Columns on the cells to split the date from the time information (assuming you don't need time)
- You will be left with this:

``` bash
 |__A1__|  |__B1__|
 2014-Dec-01  05:00:54 AM
```

On cell A1 rearrange the text and add in the date delimiters:

``` bash
=CONCATENATE(MID(A2,6,3)&"/"&RIGHT(A2,2)&"/"&LEFT(A2,4))
```

Result = Dec/01/2014

- Do a Find & Replace "Dec" with "12"
- Cells get automatically converted to Date/Time format
- Repeat for different months

#### Convert EPOCH format (Unix time) to Gregorian format (mm/dd/yyyy hh:mm:ss)

Unix time is the number of seconds since January 1, 1970.

``` bash
=CELL/(60*60*24)+"1/1/1970"
```

Turns 1424783916.796051000 = 02/24/2015 13:18:37

#### Convert a date and time field to [ISO 8601](https://en.wikipedia.org/wiki/ISO_8601) timestamp format

Example: 8/3/21 12:12:12 PM to 2021-08-03T12:12:12

``` bash
=TEXT(A1,"yyyy-mm-ddThh:MM:ss")
```

#### Convert a [ISO 8601](https://en.wikipedia.org/wiki/ISO_8601) timestamp format field to date and time

Example: 2021-08-03T12:12:12 to 8/3/21 12:12:12 PM

``` bash
=DATEVALUE(MID(A1,1,10))+TIMEVALUE(MID(A1,12,8))
```

#### Get the quarter of the year from a date

Example: "Monday, July 3, 2023" to "2"

``` bash
=ROUNDUP(MONTH(A2)/3,0)
```

Add a "Q" to the quarter number

``` bash
=CONCAT("Q",ROUNDUP(MONTH(A2)/3,0)
```

### Number Manipulation

#### Convert $20,000,000.00 to $20.0M

Select the cell you want to convert and add the following custom number format

``` bash
$[>=999950]0.0,,"M";[<=-999950]0.0,,"M";0.0,"K"
```

### Text Manipulation

#### Extract only the domain name from an email address

``` bash
=RIGHT(A1,LEN(A1)-FIND("@",A1))
```

#### Find what is to the RIGHT of the last instances of a specific character

Example = Drive:\Folder\SubFolder\Filename.ext (where you just want to find Filename.ext)

Find to the right of the last "\" character

``` bash
=REGEXEXTRACT(A1,"\\([^\\]*$)")
```

To find what's to the LEFT, just replace "RIGHT" with "LEFT" in the formula

Example = "First_Name Last_Name" (where you just want "First_Name")

``` bash
=REGEXEXTRACT(A1,"(^[^ ]*) ")
```

#### Find if cell contains a space

``` bash
=IF(COUNTIF(H2,"* *"),"No","Yes")
```

#### Extract text between two characters in a cell

``` bash
=REGEXEXTRACT(A1,"vip\.ce\.(.*)\.http")
```

Original = vip.ce.api-prd.website.com.http

After = api-prd.website.com

#### Trim All Whitespace Including Nonbreaking Space Characters (nbsp)

``` bash
=TRIM(SUBSTITUTE(A1, CHAR(160), " "))
```

#### VLookUp and Replace #N/A with some text

This works in both Excel and Google Sheets

``` bash
=IF(ISNA(VLOOKUP(A2,<Table Range>,1,FALSE)),"Thing not found",VLOOKUP(A2,<Table Range>,1,FALSE))
```

```XLOOKUP``` already has built in error handling for the ```#N/A``` messages, but only works in Excel at the date of publishing this.

#### Search for text within a cell and label it as X

``` bash
=IF(IFERROR(SEARCH("<word>",A2),0),"Cleaned",IF(IFERROR(SEARCH("<other word>",A2),0),"Unknown","Not Cleaned"))
```

#### Lookup a Value in 2 Different Columns and return the one you want

=Index(array, Match(value_to_lookup, lookup_array, match_type))

``` bash
=INDEX('TabName'!$A$1:$C$1000, MATCH('TabName'!A2,'TabName'!$A$1:$C$1000,0))
```

#### Get OS Short name from long Operating System name (Windows 10 Enterprise = Windows)

``` bash
=IF(IFERROR(SEARCH("Windows",C2),0),"Windows",IF(IFERROR(SEARCH("AIX",C2),0),"AIX",IF(IFERROR(SEARCH("Linux",C2),0),"Linux",IF(IFERROR(SEARCH("SunOS",C2),0),"SunOS",IF(IFERROR(SEARCH("OS X",C2),0),"Mac","Unknown")))))
```

#### Get system type from OS (Windows Serer 2012 = Server)

``` bash
=IF(IFERROR(SEARCH("Server",E2),0),"Server",IF(IFERROR(SEARCH("AIX",E2),0),"Server",IF(IFERROR(SEARCH("Linux",E2),0),"Server",IF(IFERROR(SEARCH("SunOS",E2),0),"Server",IF(IFERROR(SEARCH("Enterprise",E2),0),"Desktop",IF(IFERROR(SEARCH("Pro",E2),0),"Desktop",IF(IFERROR(SEARCH("Embedded",E2),0),"Desktop",IF(IFERROR(SEARCH("Windows 7",E2),0),"Desktop",IF(IFERROR(SEARCH("Windows 10",E2),0),"Desktop",IF(IFERROR(SEARCH("OS X",E2),0),"Desktop","Unknown"))))))))))
```
