package io.github.ssstlis.excelsorter

import com.typesafe.config.ConfigFactory

import scala.util.{Failure, Success, Try}

object ExcelSorterApp extends App {

  def run(filePaths: Seq[String], configs: Seq[SheetSortingConfig]): Unit = {
    val sorter = SheetSorter(configs: _*)
    val sheetNames = configs.map(_.sheetName).toSet
    val comparer = PairedSheetComparer(sheetNames, SheetSorter.defaultDateValidator)

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

    grouped.pairs.foreach { pair =>
      val sortedOld = sortedFiles.get(pair.oldFile)
      val sortedNew = sortedFiles.get(pair.newFile)

      (sortedOld, sortedNew) match {
        case (Some(oldPath), Some(newPath)) =>
          println(s"\nComparing pair: ${pair.prefix}")
          println(s"  Old: $oldPath")
          println(s"  New: $newPath")

          Try(comparer.compareAndRemoveEqualLeadingRows(oldPath, newPath)) match {
            case Success(results) =>
              if (results.isEmpty) {
                println("  No equal leading rows found in any sheet")
              } else {
                results.foreach { r =>
                  println(s"  Sheet '${r.sheetName}': removed ${r.removedRowCount} equal leading rows")
                }
              }
            case Failure(ex) =>
              System.err.println(s"  -> Comparison error: ${ex.getMessage}")
          }

        case _ =>
          System.err.println(s"Skipping pair ${pair.prefix}: one or both files failed to sort")
      }
    }
  }

  if (args.isEmpty) {
    System.err.println("Usage: excel-sorter <file1.xlsx> [file2.xlsx] ...")
    System.err.println()
    System.err.println("Files can be paired by naming convention:")
    System.err.println("  prefix_old.xlsx + prefix_new.xlsx")
    System.err.println()
    System.err.println("For paired files, equal leading data rows will be removed from both.")
  } else {
    run(args.toSeq, ConfigReader.fromConfig(ConfigFactory.load()))
  }
}
