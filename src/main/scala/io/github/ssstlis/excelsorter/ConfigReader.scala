package io.github.ssstlis.excelsorter

import java.time.LocalDate
import java.time.format.DateTimeFormatter

import com.typesafe.config.{Config, ConfigValueType}

import scala.jdk.CollectionConverters._
import scala.util.Try

object ConfigReader {
  implicit val ord: Ordering[java.time.LocalDate] = SortingDsl.Parsers.localDateOrdering

  private val LocalDatePattern = """LocalDate\((.+)\)""".r

  def fromConfig(config: Config): Seq[SheetSortingConfig] = {
    config.getConfigList("sortings").asScala.map(parseSheet).toSeq
  }

  def readTrackConfig(config: Config): TrackConfig = {
    if (!config.hasPath("tracks")) return TrackConfig.empty

    val trackList = config.getConfigList("tracks").asScala.toList
    val policies = trackList.map(parseTrackPolicy)
    TrackConfig(policies)
  }

  private def parseTrackPolicy(trackConfig: Config): TrackPolicy = {
    val sheetSelector = if (trackConfig.getIsNull("sheet")) {
      SheetSelector.Default
    } else {
      trackConfig.getValue("sheet").valueType() match {
        case ConfigValueType.NUMBER =>
          SheetSelector.ByIndex(trackConfig.getInt("sheet"))
        case ConfigValueType.STRING =>
          SheetSelector.ByName(trackConfig.getString("sheet"))
        case other =>
          throw new IllegalArgumentException(
            s"Invalid sheet selector type: $other. Expected null, string, or number."
          )
      }
    }

    val conditions = trackConfig.getConfigList("conditions").asScala.toList.map(parseTrackCondition)
    TrackPolicy(sheetSelector, conditions)
  }

  private def parseTrackCondition(condConfig: Config): TrackCondition = {
    val index = condConfig.getInt("index")
    val as = condConfig.getString("as")

    val validator: String => Boolean = as match {
      case "String" =>
        s => s != null && s.nonEmpty
      case "Int" =>
        s => Try(s.toInt).isSuccess
      case "Long" =>
        s => Try(s.toLong).isSuccess
      case "Double" =>
        s => Try(s.toDouble).isSuccess
      case "BigDecimal" =>
        s => Try(BigDecimal(s)).isSuccess
      case "LocalDate" =>
        s => SheetSorter.defaultDateValidator(s)
      case LocalDatePattern(pattern) =>
        val fmt = DateTimeFormatter.ofPattern(pattern)
        s => s != null && s.trim.nonEmpty && Try(LocalDate.parse(s.trim, fmt)).isSuccess
      case other =>
        throw new IllegalArgumentException(
          s"Unknown track condition type: '$other'. Expected one of: String, Int, Long, Double, BigDecimal, LocalDate, LocalDate(<pattern>)."
        )
    }

    TrackCondition(index, validator)
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
