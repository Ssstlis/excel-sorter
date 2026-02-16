package io.github.ssstlis.excelsorter.dsl.config

import io.github.ssstlis.excelsorter.config.CliArgs.parseSheetName
import io.github.ssstlis.excelsorter.config.ConfigReader

case class SheetSortingConfig(sheetName: String, sortConfigs: List[ColumnSortConfig[_]])

object SheetSortingConfig {
  def parseSortingsBlock(args: List[String]): Either[String, SheetSortingConfig] = {
    parseSheetName(args).flatMap { case (sheetName, rest) =>
      parseSortEntries(rest).map((sheetName, _))
    } match {
      case Left(err) => Left(s"--sortings: $err")
      case Right((_, Nil)) => Left("--sortings: at least one -sort/-o entry is required.")
      case Right((sheetName, sorts)) => Right(SheetSortingConfig(sheetName, sorts))
    }
  }

  private def parseSortEntries(args: List[String]): Either[String, List[ColumnSortConfig[_]]] = {
    val sortFlags = Set("-sort", "-o")
    var remaining = args
    val result = List.newBuilder[ColumnSortConfig[_]]

    while (remaining.nonEmpty) {
      remaining match {
        case flag :: order :: idxStr :: asType :: tail if sortFlags.contains(flag) =>
          val index = try { idxStr.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index: '$idxStr'. Expected an integer.")
          }
          try {
            result += ConfigReader.resolveColumnSort(order, index, asType)
          } catch {
            case e: IllegalArgumentException => return Left(e.getMessage)
          }
          remaining = tail
        case flag :: _ if sortFlags.contains(flag) =>
          return Left(s"$flag requires 3 arguments: <order> <column-index> <type>")
        case other :: _ =>
          return Left(s"Unexpected argument: '$other'. Expected -sort or -o.")
        case Nil => // won't happen due to while condition
      }
    }

    Right(result.result())
  }
}