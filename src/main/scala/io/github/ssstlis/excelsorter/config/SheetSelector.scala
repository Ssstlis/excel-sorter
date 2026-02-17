package io.github.ssstlis.excelsorter.config

import com.typesafe.config.{Config, ConfigValueType}

sealed trait SheetSelector
object SheetSelector {
  case object Default extends SheetSelector
  case class ByName(name: String) extends SheetSelector
  case class ByIndex(index: Int) extends SheetSelector

  def parseSheetSelector(value: String): SheetSelector = {
    if (value == "default") {
      SheetSelector.Default
    } else {
      try {
        SheetSelector.ByIndex(value.toInt)
      } catch {
        case _: NumberFormatException => SheetSelector.ByName(value)
      }
    }
  }

  def parseSheetSelector(config: Config): Either[String, SheetSelector] = {
    if (config.getIsNull("sheet")) {
      Right(SheetSelector.Default)
    } else {
      config.getValue("sheet").valueType() match {
        case ConfigValueType.NUMBER =>
          Right(SheetSelector.ByIndex(config.getInt("sheet")))
        case ConfigValueType.STRING =>
          Right(SheetSelector.ByName(config.getString("sheet")))
        case other =>
          Left(
            s"Invalid sheet selector type: $other. Expected null, string, or number."
          )
      }
    }
  }
}