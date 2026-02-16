package io.github.ssstlis.excelsorter.dsl.config

import io.github.ssstlis.excelsorter.config.{ConfigReader => LocalDatePattern}
import io.github.ssstlis.excelsorter.dsl.{SortOrder, SortingDsl}

import scala.util.Try

case class ColumnSortConfig[T](
                                columnIndex: Int,
                                order: SortOrder,
                                parser: String => T
                              )(implicit val ordering: Ordering[T]) {

  def compare(a: String, b: String): Int = {
    val result = for {
      parsedA <- Try(parser(a))
      parsedB <- Try(parser(b))
    } yield ordering.compare(parsedA, parsedB)

    val cmp = result.getOrElse(a.compareTo(b))
    order match {
      case SortOrder.Asc => cmp
      case SortOrder.Desc => -cmp
    }
  }
}

object ColumnSortConfig {
  final val LocalDatePattern = """LocalDate\((.+)\)""".r

  def create[T: Ordering](columnIndex: Int, order: SortOrder)(parser: String => T): ColumnSortConfig[T] =
    new ColumnSortConfig(columnIndex, order, parser)

  def resolveColumnSort(orderStr: String, index: Int, asType: String): ColumnSortConfig[_] = {
    val order = orderStr.toLowerCase match {
      case "asc"  => SortOrder.Asc
      case "desc" => SortOrder.Desc
      case other  => throw new IllegalArgumentException(s"Unknown sort order: '$other'. Expected 'asc' or 'desc'.")
    }

    asType match {
      case "String"     => ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asString)
      case "Int"        => ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asInt)
      case "Long"       => ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLong)
      case "Double"     => ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asDouble)
      case "BigDecimal" => ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asBigDecimal)
      case "LocalDate" =>
        ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLocalDateDefault)
      case LocalDatePattern(pattern) =>
        ColumnSortConfig.create(index, order)(SortingDsl.Parsers.asLocalDate(pattern))
      case other =>
        throw new IllegalArgumentException(
          s"Unknown parser type: '$other'. Expected one of: String, Int, Long, Double, BigDecimal, LocalDate, LocalDate(<pattern>)."
        )
    }
  }
}
