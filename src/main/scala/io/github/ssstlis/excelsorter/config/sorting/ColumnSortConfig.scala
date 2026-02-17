package io.github.ssstlis.excelsorter.config.sorting

import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.dsl.SortingDsl.Parsers.LocalDatePattern
import io.github.ssstlis.excelsorter.dsl.{SortOrder, SortingDsl}

import scala.util.Try

case class ColumnSortConfig(columnIndex: Int, order: SortOrder, compare: (String, String) => Int)

object ColumnSortConfig {

  def parseSortConfig(sortConfig: Config): Either[String, ColumnSortConfig] = {
    val orderStr = sortConfig.getString("order")
    val index = sortConfig.getInt("index")
    val as = sortConfig.getString("as")
    resolveColumnSort(orderStr, index, as)
  }

  def create[T](columnIndex: Int, order: SortOrder)(parser: String => T)(implicit O: Ordering[T]): ColumnSortConfig = {
    val cmp = (a: String, b: String) => {
      val result = for {
        parsedA <- Try(parser(a))
        parsedB <- Try(parser(b))
      } yield O.compare(parsedA, parsedB)

      val cmp = result.getOrElse(a.compareTo(b))
      order match {
        case SortOrder.Asc => cmp
        case SortOrder.Desc => -cmp
      }
    }
    new ColumnSortConfig(columnIndex, order, cmp)
  }

  def resolveColumnSort(orderStr: String, index: Int, asType: String): Either[String, ColumnSortConfig] = {
    val order = orderStr.toLowerCase match {
      case "asc"  => Right(SortOrder.Asc)
      case "desc" => Right(SortOrder.Desc)
      case other  => Left(s"Unknown sort order: '$other'. Expected 'asc' or 'desc'.")
    }

    order.flatMap { order =>
      asType match {
        case "String" => Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asString))
        case "Int" => Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asInt))
        case "Long" => Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLong))
        case "Double" => Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asDouble))
        case "BigDecimal" => Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asBigDecimal))
        case "LocalDate" =>
          Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLocalDateDefault))
        case LocalDatePattern(pattern) =>
          Right(ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLocalDate(pattern)))
        case other =>
          Left(
            s"Unknown parser type: '$other'. Expected one of: String, Int, Long, Double, BigDecimal, LocalDate, LocalDate(<pattern>)."
          )
      }
    }
  }
}
