package io.github.ssstlis.excelsorter.dsl

import scala.util.Try

sealed trait SortOrder
object SortOrder {
  case object Asc extends SortOrder
  case object Desc extends SortOrder
}

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
      case SortOrder.Asc  => cmp
      case SortOrder.Desc => -cmp
    }
  }
}

object ColumnSortConfig {
  def create[T: Ordering](columnIndex: Int, order: SortOrder)(parser: String => T): ColumnSortConfig[T] =
    new ColumnSortConfig(columnIndex, order, parser)
}

case class SheetSortingConfig(
  sheetName: String,
  sortConfigs: List[ColumnSortConfig[_]]
)

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
