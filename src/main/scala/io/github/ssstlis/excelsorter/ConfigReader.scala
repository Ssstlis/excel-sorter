package io.github.ssstlis.excelsorter

import com.typesafe.config.Config

import scala.jdk.CollectionConverters._

object ConfigReader {
  implicit val ord: Ordering[java.time.LocalDate] = SortingDsl.Parsers.localDateOrdering

  private val LocalDatePattern = """LocalDate\((.+)\)""".r

  def fromConfig(config: Config): Seq[SheetSortingConfig] = {
    config.getConfigList("sortings").asScala.map(parseSheet).toSeq
  }

  private def parseSheet(sheetConfig: Config): SheetSortingConfig = {
    val name = sheetConfig.getString("name")
    val sorts = sheetConfig.getConfigList("sorts").asScala.map(parseSortConfig).toList
    SheetSortingConfig(name, sorts)
  }

  private def parseSortConfig(sortConfig: Config): ColumnSortConfig[_] = {
    val order = sortConfig.getString("order").toLowerCase match {
      case "asc"  => SortOrder.Asc
      case "desc" => SortOrder.Desc
      case other  => throw new IllegalArgumentException(s"Unknown sort order: '$other'. Expected 'asc' or 'desc'.")
    }
    val index = sortConfig.getInt("index")
    val as = sortConfig.getString("as")

    as match {
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
