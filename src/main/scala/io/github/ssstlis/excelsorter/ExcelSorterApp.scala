package io.github.ssstlis.excelsorter

import com.typesafe.config.ConfigFactory

import scala.util.{Failure, Success, Try}

object ExcelSorterApp extends App {

  def run(filePaths: Seq[String], configs: Seq[SheetSortingConfig], mode: RunMode, trackConfig: TrackConfig): Unit = {
    val sorter = SheetSorter(configs, trackConfig)
    val sheetNames = configs.map(_.sheetName).toSet

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
        val cutter = PairedSheetCutter(sheetNames, trackConfig)
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
                  if (results.isEmpty) {
                    println("  No equal leading rows found in any sheet")
                  } else {
                    results.foreach { r =>
                      println(s"  Sheet '${r.sheetName}': removed ${r.removedRowCount} equal leading rows")
                    }
                  }
                case Failure(ex) =>
                  System.err.println(s"  -> Cut error: ${ex.getMessage}")
              }

            case _ =>
              System.err.println(s"Skipping pair ${pair.prefix}: one or both files failed to sort")
          }
        }

      case RunMode.Compare =>
        val highlighter = PairedSheetHighlighter(sheetNames, trackConfig)
        grouped.pairs.foreach { pair =>
          val sortedOld = sortedFiles.get(pair.oldFile)
          val sortedNew = sortedFiles.get(pair.newFile)

          (sortedOld, sortedNew) match {
            case (Some(oldPath), Some(newPath)) =>
              println(s"\nComparing pair: ${pair.prefix}")
              println(s"  Old: $oldPath")
              println(s"  New: $newPath")

              Try(highlighter.highlightEqualLeadingRows(oldPath, newPath)) match {
                case Success((oldCmpPath, newCmpPath, results)) =>
                  println(s"  -> Created: $oldCmpPath")
                  println(s"  -> Created: $newCmpPath")
                  if (results.isEmpty) {
                    println("  No equal leading rows found in any sheet")
                  } else {
                    results.foreach { r =>
                      println(s"  Sheet '${r.sheetName}': highlighted ${r.highlightedRowCount} equal leading rows")
                    }
                  }
                case Failure(ex) =>
                  System.err.println(s"  -> Compare error: ${ex.getMessage}")
              }

            case _ =>
              System.err.println(s"Skipping pair ${pair.prefix}: one or both files failed to sort")
          }
        }
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
        run(cliArgs.filePaths, sortConfigs, cliArgs.mode, trackConfig)
    }
  }
}
