### Usage: 
 > excel-sorter [-h|--help] [--cut|-c | --compare|-cmp] <files...> [--conf <config-blocks...>]

### Options:
-h, --help       Show help message and exit

### Modes:
(default)        Sort only — creates *_sorted.xlsx files
--cut, -c        Sort and cut — also creates *_sortcutted.xlsx with equal leading rows removed
--compare, -cmp  Sort and compare — also creates *_compared.xlsx with equal leading rows highlighted green

Files can be paired by naming convention:
prefix_old.xlsx + prefix_new.xlsx

Flags --cut and --compare are mutually exclusive.

### Configuration (--conf):
When --conf is present, HOCON config files are ignored. All configuration is provided via CLI.
Config blocks after --conf (each --sortings/--tracks/--comparisons starts a new block):

--sortings -sheet <name> -sort <asc|desc> <col-index> <type> [-sort ...]
--sortings -s <name> -o <asc|desc> <col-index> <type> [-o ...]
Defines sort configuration for a sheet. At least one -sort/-o is required.

--tracks -sheet <name> -cond <col-index> <type> [-cond ...]
--tracks -s <name> -d <col-index> <type> [-d ...]
Defines data row detection for a sheet. At least one -cond/-d is required.

--comparisons -sheet <name> -ic <col-index> [<col-index> ...]
--comparisons -s <name> -ic <col-index> [<col-index> ...]
Defines columns to ignore when comparing rows. At least one column index is required.

Sheet name (-sheet/-s) interpretation:
"default"  → default policy (fallback for all sheets)
<number>   → sheet by index (e.g. \"0\")
<string>   → sheet by name

Supported types: String, Int, Long, Double, BigDecimal, LocalDate, LocalDate(<pattern>)

### Example:
```
excel-sorter --cut file_old.xlsx file_new.xlsx --conf \
  --sortings -sheet "Sheet1" -sort asc 0 LocalDate -sort desc 2 String \
  --tracks -sheet default -cond 0 LocalDate \
  --comparisons -sheet "Sheet1" -ic 1 13
```