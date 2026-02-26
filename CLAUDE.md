# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Test Commands

This is a Scala 2.13.18 project using sbt 1.12.2.

- **Compile:** `sbt compile`
- **Run tests:** `sbt test`
- **Run a single test class:** `sbt "testOnly io.github.ssstlis.excelsorter.processor.FilePairerSpec"`
- **Run the app:** `sbt run` (requires `.xlsx` file arguments)
- **Package as distributable:** `sbt universal:packageBin` (via sbt-native-packager)

## Architecture

Excel-sorter is a CLI tool that sorts rows within Excel (.xlsx) sheets and optionally cuts or highlights paired files. It uses Apache POI for Excel manipulation. Sort/track/compare configurations are loaded from Typesafe Config (HOCON) files (`application.conf`).

### Package Structure

- `config` — CLI argument parsing, HOCON config reading, track/compare policies (`CliArgs`, `ConfigReader`, `TrackConfig`, `CompareConfig`)
- `dsl` — Sorting DSL types and parsers (`SortingDsl`, `SheetSortingConfig`, `ColumnSortConfig`)
- `processor` — Core processing logic (`SheetSorter`, `PairedSheetCutter`, `PairedSheetHighlighter`, `FilePairer`, `CellUtils`, `CompareResult`, `HighlightResult`)

### Run Modes

The app supports three mutually exclusive modes selected via CLI flags:

- **(default) SortOnly** — creates `*_sorted.xlsx` files
- **`--cut` / `-c`** — sort + cut equal leading data rows from paired files (`*_sortcutted.xlsx`)
- **`--compare` / `-cmp`** — sort + key-based comparison of paired files with three highlight colors (`*_compared.xlsx`): green (same key + same data), pale red (same key + different data), pale orange (key only in one file)

### Processing Pipeline

1. **File pairing** (`FilePairer`): Input files are grouped by naming convention — files ending in `_old.xlsx` and `_new.xlsx` with matching prefixes form pairs. Remaining files are treated as unpaired.

2. **Sheet sorting** (`SheetSorter`): Each file's sheets are sorted according to `SheetSortingConfig` loaded from config. Header rows (everything before the first detected data row) are preserved in place; only data rows are sorted. Output is written to a new `*_sorted.xlsx` file.

3. **Paired cut** (`PairedSheetCutter`): For paired files in `--cut` mode, equal leading data rows are removed from both the old and new sorted files, leaving only the differing tail.

4. **Paired highlight** (`PairedSheetHighlighter`): For paired files in `--compare` mode, data rows are matched by key (sort column indices from config) and highlighted with three colors: green for matching rows (same key + same data), pale red for changed rows (same key + different data), and pale orange for rows only in one file. Returns `HighlightResult` with per-sheet counts. A `*_compare_results.txt` file is also written.

### Configuration (HOCON)

Sort, track, and compare configs are defined in `application.conf` / `application-test.conf`:

- **`sortings`** — list of sheet sort configs. Each entry has `name` (sheet name), `sorts` (list of `{order, index, as}`). Supported `as` types: `String`, `Int`, `Long`, `Double`, `BigDecimal`, `LocalDate`, `LocalDate(<pattern>)`.
- **`tracks`** (optional) — configurable data row detection per sheet via `SheetSelector` (`null` = default, string = by name, number = by index) and `conditions` list. Falls back to checking if column 0 contains a parseable date.
- **`comparisons`** (optional) — columns to ignore when comparing rows for equality, per sheet via `SheetSelector` and `ignoreColumns`.

### CLI Configuration (`--conf`)

As an alternative to HOCON, all configuration can be provided via CLI using `--conf`. When `--conf` is present, HOCON config files are ignored entirely. When absent, HOCON is used as before.

```
excel-sorter [-h|--help] [--cut|-c | --compare|-cmp] <files...> [--conf <config-blocks...>]
```

Config blocks after `--conf` (each `--sortings`/`--tracks`/`--comparisons` starts a new block):

```
--conf \
  --sortings -sheet "Sheet Name" -sort asc 0 LocalDate -sort desc 2 String \
  --tracks -sheet "Sheet Name" -cond 0 LocalDate -cond 2 Int \
  --comparisons -sheet "Sheet Name" -ic 1 13
```

Short flag equivalents: `-s` for `-sheet`, `-o` for `-sort`, `-d` for `-cond`.

Sheet name (`-sheet`/`-s`) interpretation:
- `"default"` → `SheetSelector.Default`
- Numeric string (e.g. `"0"`) → `SheetSelector.ByIndex(0)`
- Other string → `SheetSelector.ByName(name)`

The `-h`/`--help` flag prints usage with full `--conf` syntax documentation.

### Sorting DSL (`SortingDsl`)

Sort configurations can also be built programmatically with a DSL: `sheet("name")(asc(colIndex)(parser), desc(colIndex)(parser))`. Parsers (`Parsers` object) convert cell string values to typed values for comparison — available parsers include `asLocalDateDefault`, `asString`, `asInt`, `asLong`, `asDouble`, `asBigDecimal`, and `asLocalDate(pattern)`.

### Data Row Detection (`TrackConfig`)

`TrackConfig` provides configurable data row detection. Policies are matched by sheet name, sheet index, or a default fallback. Each policy has a list of `TrackCondition`s (column index + validator). A row is considered a data row when all conditions pass. When no policy matches, the default checks if column 0 contains a parseable date (ISO, dd.MM.yyyy, dd/MM/yyyy, MM/dd/yyyy, yyyy/MM/dd).

### Default Sheet Configuration

The default config in `application-test.conf` targets Russian-language financial report sheets. Sheet names are in Cyrillic.

## Key Dependencies

- **Apache POI 5.5.1** — Excel file I/O (`poi` and `poi-ooxml`)
- **Typesafe Config 1.4.3** — HOCON configuration loading
- **ScalaTest 3.2.19** — Testing (FreeSpec style with Matchers)
- **sbt-native-packager** — Packaging for distribution
- **sbt-scoverage** — Code coverage