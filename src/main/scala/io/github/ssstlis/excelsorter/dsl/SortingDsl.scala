package io.github.ssstlis.excelsorter.dsl

import io.github.ssstlis.excelsorter.dsl.config.{ColumnSortConfig, SheetSortingConfig}

sealed trait SortOrder
object SortOrder {
  case object Asc extends SortOrder
  case object Desc extends SortOrder
}

object SortingDsl {
  import SortOrder._

  def asc[T: Ordering](columnIndex: Int)(parser: String => T): ColumnSortConfig[T] =
    new ColumnSortConfig(columnIndex, Asc, parser)

  def desc[T: Ordering](columnIndex: Int)(parser: String => T): ColumnSortConfig[T] =
    new ColumnSortConfig(columnIndex, Desc, parser)

  def sheet(name: String)(configs: ColumnSortConfig[_]*): SheetSortingConfig =
    SheetSortingConfig(name, configs.toList)

  object Parsers {
    import java.time.LocalDate
    import java.time.format.DateTimeFormatter

    implicit val localDateOrdering: Ordering[LocalDate] = Ordering.fromLessThan(_ isBefore _)

    val asString: String => String = identity
    val asInt: String => Int = _.toInt
    val asLong: String => Long = _.toLong
    val asDouble: String => Double = _.toDouble
    val asBigDecimal: String => BigDecimal = BigDecimal(_)

    def asLocalDate(pattern: String): String => LocalDate = { s =>
      LocalDate.parse(s, DateTimeFormatter.ofPattern(pattern))
    }

    val asLocalDateDefault: String => LocalDate = asLocalDate("yyyy-MM-dd")
  }
}
