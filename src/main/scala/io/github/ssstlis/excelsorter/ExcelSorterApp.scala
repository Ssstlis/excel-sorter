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
    System.err.println("Usage: excel-sorter [--cut|-c | --compare|-cmp] <file1.xlsx> [file2.xlsx] ...")
    System.err.println()
    System.err.println("Modes:")
    System.err.println("  (default)        Sort only — creates *_sorted.xlsx files")
    System.err.println("  --cut, -c        Sort and cut — also creates *_sortcutted.xlsx with equal leading rows removed")
    System.err.println("  --compare, -cmp  Sort and compare — also creates *_compared.xlsx with equal leading rows highlighted green")
    System.err.println()
    System.err.println("Files can be paired by naming convention:")
    System.err.println("  prefix_old.xlsx + prefix_new.xlsx")
    System.err.println()
    System.err.println("Flags --cut and --compare are mutually exclusive.")
  }

  if (args.isEmpty) {
    printUsage()
  } else {
    CliArgs.parse(args) match {
      case Left(error) =>
        System.err.println(s"Error: $error")
        printUsage()
      case Right(cliArgs) =>
        val config = ConfigFactory.load()
        val sortConfigs = ConfigReader.fromConfig(config)
        val trackConfig = ConfigReader.readTrackConfig(config)
        val compareConfig = ConfigReader.readCompareConfig(config)
        run(cliArgs.filePaths, sortConfigs, cliArgs.mode, trackConfig, compareConfig)
    }
  }
}
