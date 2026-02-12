package io.github.ssstlis.excelsorter

import java.io.{File, PrintWriter}

import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.config._
import io.github.ssstlis.excelsorter.dsl.SheetSortingConfig
import io.github.ssstlis.excelsorter.processor._

import scala.util.{Failure, Success, Try}

object ExcelSorterApp extends App {

  def run(
    filePaths: Seq[String],
    configs: Seq[SheetSortingConfig],
    mode: RunMode,
    trackConfig: TrackConfig,
    compareConfig: CompareConfig
  ): Unit = {
    val sorter = SheetSorter(configs, trackConfig)
    val sheetNames = configs.map(_.sheetName).toSet
    val sortConfigsMap = configs.map(c => c.sheetName -> c).toMap

    val grouped = FilePairer.groupFiles(filePaths)

    val sortedFiles = scala.collection.mutable.Map[String, String]()

    (grouped.pairs.flatMap(p => List(p.oldFile, p.newFile)) ++ grouped.unpaired).foreach { path =>
      println(s"Sorting: $path")
      Try(sorter.sortFile(path)) match {
        case Success(outputPath) =>
          println(s"  -> Created: $outputPath")
          sortedFiles(path) = outputPath
        case Failure(ex) =>
          System.err.println(s"  -> Error: ${ex.getMessage}")
      }
    }

    mode match {
      case RunMode.SortOnly =>
        // Sort-only mode: no further processing

      case RunMode.Cut =>
        val cutter = PairedSheetCutter(sheetNames, trackConfig, compareConfig, sortConfigsMap)
        grouped.pairs.foreach { pair =>
          val sortedOld = sortedFiles.get(pair.oldFile)
          val sortedNew = sortedFiles.get(pair.newFile)

          (sortedOld, sortedNew) match {
            case (Some(oldPath), Some(newPath)) =>
              println(s"\nCutting pair: ${pair.prefix}")
              println(s"  Old: $oldPath")
              println(s"  New: $newPath")

              Try(cutter.cutEqualLeadingRows(oldPath, newPath)) match {
                case Success((oldCutPath, newCutPath, results)) =>
                  println(s"  -> Created: $oldCutPath")
                  println(s"  -> Created: $newCutPath")
                  printCompareResults(results, "removed")
                  writeResultsFile(pair, results)
                case Failure(ex) =>
                  System.err.println(s"  -> Cut error: ${ex.getMessage}")
              }

            case _ =>
              System.err.println(s"Skipping pair ${pair.prefix}: one or both files failed to sort")
          }
        }

      case RunMode.Compare =>
        val highlighter = PairedSheetHighlighter(sheetNames, trackConfig, compareConfig, sortConfigsMap)
        grouped.pairs.foreach { pair =>
          val sortedOld = sortedFiles.get(pair.oldFile)
          val sortedNew = sortedFiles.get(pair.newFile)

          (sortedOld, sortedNew) match {
            case (Some(oldPath), Some(newPath)) =>
              println(s"\nComparing pair: ${pair.prefix}")
              println(s"  Old: $oldPath")
              println(s"  New: $newPath")

              Try(highlighter.highlightPairedSheets(oldPath, newPath)) match {
                case Success((oldCmpPath, newCmpPath, results)) =>
                  println(s"  -> Created: $oldCmpPath")
                  println(s"  -> Created: $newCmpPath")
                  printHighlightResults(results)
                  writeHighlightResultsFile(pair, results)
                case Failure(ex) =>
                  System.err.println(s"  -> Compare error: ${ex.getMessage}")
              }

            case _ =>
              System.err.println(s"Skipping pair ${pair.prefix}: one or both files failed to sort")
          }
        }
    }
  }

  private def printCompareResults(results: List[CompareResult], action: String): Unit = {
    results.foreach { r =>
      if (r.equalRowCount > 0) {
        val mismatchInfo = r.firstMismatchRowNum match {
          case Some(rowNum) =>
            s", first difference at row $rowNum (key: ${r.firstMismatchKey.getOrElse("")})"
          case None => ""
        }
        println(s"  Sheet '${r.sheetName}': $action ${r.equalRowCount} equal leading rows$mismatchInfo")
      } else {
        r.firstMismatchRowNum match {
          case Some(rowNum) =>
            println(s"  Sheet '${r.sheetName}': no equal leading rows, first difference at row $rowNum (key: ${r.firstMismatchKey.getOrElse("")})")
          case None =>
            println(s"  Sheet '${r.sheetName}': all data rows are equal")
        }
      }
    }
  }

  private def writeResultsFile(pair: FilePair, results: List[CompareResult]): Unit = {
    val oldFileDir = new File(pair.oldFile).getParentFile
    val resultsPath = new File(oldFileDir, s"${pair.prefix}_compare_results.txt").getAbsolutePath

    Try {
      val writer = new PrintWriter(resultsPath)
      try {
        writer.println(s"Comparison results for: ${pair.prefix}")
        writer.println()

        results.foreach { r =>
          writer.println(s"Sheet '${r.sheetName}':")
          if (r.equalRowCount > 0) {
            writer.println(s"  Equal leading rows: ${r.equalRowCount}")
            r.firstMismatchRowNum.foreach { rowNum =>
              writer.println(s"  First difference at row $rowNum (key: ${r.firstMismatchKey.getOrElse("")})")
            }
          } else {
            r.firstMismatchRowNum match {
              case Some(rowNum) =>
                writer.println(s"  Equal leading rows: 0")
                writer.println(s"  First difference at row $rowNum (key: ${r.firstMismatchKey.getOrElse("")})")
              case None =>
                writer.println(s"  All data rows are equal (${r.equalRowCount} rows)")
            }
          }
          writer.println()
        }
      } finally {
        writer.close()
      }

      println(s"  -> Results written to: $resultsPath")
    } match {
      case Failure(ex) =>
        System.err.println(s"  -> Error writing results file: ${ex.getMessage}")
      case _ =>
    }
  }

  private def printHighlightResults(results: List[HighlightResult]): Unit = {
    results.foreach { r =>
      println(s"  Sheet '${r.sheetName}': ${r.matchedSameDataCount} matching, ${r.matchedDifferentDataCount} changed, ${r.oldOnlyCount} old-only, ${r.newOnlyCount} new-only")
      if (r.oldOnlyColumns.nonEmpty)
        println(s"    Columns only in old file: ${r.oldOnlyColumns.mkString(", ")}")
      if (r.newOnlyColumns.nonEmpty)
        println(s"    Columns only in new file: ${r.newOnlyColumns.mkString(", ")}")
      if (r.rowDiffs.nonEmpty)
        println(s"    ${r.rowDiffs.size} row(s) with cell-level differences")
    }
  }

  private def writeHighlightResultsFile(pair: FilePair, results: List[HighlightResult]): Unit = {
    val oldFileDir = new File(pair.oldFile).getParentFile
    val resultsPath = new File(oldFileDir, s"${pair.prefix}_compare_results.txt").getAbsolutePath

    Try {
      val writer = new PrintWriter(resultsPath)
      try {
        writer.println(s"Comparison results for: ${pair.prefix}")
        writer.println()

        results.foreach { r =>
          writer.println(s"Sheet '${r.sheetName}':")
          writer.println(s"  Matching rows (same key + same data): ${r.matchedSameDataCount}")
          writer.println(s"  Changed rows (same key + different data): ${r.matchedDifferentDataCount}")
          writer.println(s"  Old-only rows: ${r.oldOnlyCount}")
          writer.println(s"  New-only rows: ${r.newOnlyCount}")

          if (r.oldOnlyColumns.nonEmpty)
            writer.println(s"  Columns only in old file: ${r.oldOnlyColumns.mkString(", ")}")
          if (r.newOnlyColumns.nonEmpty)
            writer.println(s"  Columns only in new file: ${r.newOnlyColumns.mkString(", ")}")

          if (r.rowDiffs.nonEmpty) {
            writer.println()
            writer.println(s"  Row differences (${r.rowDiffs.size}):")
            r.rowDiffs.foreach { rd =>
              writer.println(s"    Key: ${rd.key} (old row ${rd.oldRowNum}, new row ${rd.newRowNum})")
              rd.cellDiffs.foreach { cd =>
                writer.println(s"      ${cd.columnName}: '${cd.oldValue}' -> '${cd.newValue}'")
              }
            }
          }

          writer.println()
        }
      } finally {
        writer.close()
      }

      println(s"  -> Results written to: $resultsPath")
    } match {
      case Failure(ex) =>
        System.err.println(s"  -> Error writing results file: ${ex.getMessage}")
      case _ =>
    }
  }

  private def printUsage(): Unit = {
    System.err.println("""Usage: excel-sorter [-h|--help] [--cut|-c | --compare|-cmp] <files...> [--conf <config-blocks...>]

    Options:
      -h, --help       Show this help message and exit

    Modes:
      (default)        Sort only — creates *_sorted.xlsx files
      --cut, -c        Sort and cut — also creates *_sortcutted.xlsx with equal leading rows removed
      --compare, -cmp  Sort and compare — also creates *_compared.xlsx with equal leading rows highlighted green

    Files can be paired by naming convention:
      prefix_old.xlsx + prefix_new.xlsx

    Flags --cut and --compare are mutually exclusive.

    Configuration (--conf):
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

      Example:
        excel-sorter --cut file_old.xlsx file_new.xlsx --conf \
          --sortings -sheet "Sheet1" -sort asc 0 LocalDate -sort desc 2 String \
          --tracks -sheet default -cond 0 LocalDate \
          --comparisons -sheet "Sheet1" -ic 1 13""")
  }

  if (args.isEmpty) {
    printUsage()
  } else {
    CliArgs.parse(args) match {
      case Left("help") =>
        printUsage()
      case Left(error) =>
        System.err.println(s"Error: $error")
        printUsage()
      case Right(cliArgs) =>
        val (sortConfigs, trackConfig, compareConfig) = cliArgs.cliConfig match {
          case Some(cc) => (cc.sortings, cc.trackConfig, cc.compareConfig)
          case None =>
            val config = ConfigFactory.load()
            (ConfigReader.fromConfig(config), ConfigReader.readTrackConfig(config), ConfigReader.readCompareConfig(config))
        }
        run(cliArgs.filePaths, sortConfigs, cliArgs.mode, trackConfig, compareConfig)
    }
  }
}
