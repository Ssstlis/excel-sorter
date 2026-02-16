package io.github.ssstlis.excelsorter.config

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
}