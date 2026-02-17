package io.github.ssstlis.excelsorter.config.sorting

import cats.instances.list._
import cats.syntax.traverse._
import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.CliArgs.parseSheetName

import scala.annotation.tailrec
import scala.jdk.CollectionConverters.ListHasAsScala

case class SheetSortingConfig(sheetName: String, sortConfigs: List[ColumnSortConfig])

object SheetSortingConfig {

  def readSortConfig(config: Config): Either[String, List[SheetSortingConfig]] = {
    config.getConfigList("sortings").asScala.toList.traverse(parseSheet)
  }

  private[config] def parseSheet(sheetConfig: Config): Either[String, SheetSortingConfig] = {
    val name = sheetConfig.getString("name")
    val sorts = sheetConfig.getConfigList("sorts").asScala.toList.traverse(ColumnSortConfig.parseSortConfig)
    sorts.map(SheetSortingConfig(name, _))
  }

  def parseSortingsBlock(args: List[String]): Either[String, SheetSortingConfig] = {
    parseSheetName(args).flatMap { case (sheetName, rest) =>
      parseSortEntries(rest).map((sheetName, _))
    } match {
      case Left(err) => Left(s"--sortings: $err")
      case Right((_, Nil)) => Left("--sortings: at least one -sort/-o entry is required.")
      case Right((sheetName, sorts)) => Right(SheetSortingConfig(sheetName, sorts))
    }
  }

  private[config] def parseSortEntries(args: List[String]): Either[String, List[ColumnSortConfig]] = {
    val sortFlags = Set("-sort", "-o")

    @tailrec
    def rec(rest: List[String], result: List[ColumnSortConfig]): Either[String, List[ColumnSortConfig]] = {
      rest match {
        case flag :: order :: idxStr :: asType :: tail if sortFlags.contains(flag) =>
          idxStr
            .toIntOption
            .toRight(s"Invalid column index: '$idxStr'. Expected an integer.")
            .flatMap { index =>
              ColumnSortConfig.resolveColumnSort(order, index, asType)
            } match {
            case Right(columnSortConfig) => rec(tail, columnSortConfig :: result)
            case Left(err) => Left(err)
          }
        case flag :: _ if sortFlags.contains(flag) =>
          Left(s"$flag requires 3 arguments: <order> <column-index> <type>")
        case other :: _ => Left(s"Unexpected argument: '$other'. Expected -sort or -o.")
        case Nil => Right(result.reverse)
      }
    }

    rec(args, Nil)
  }
}