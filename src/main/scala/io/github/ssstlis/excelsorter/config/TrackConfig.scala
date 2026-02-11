package io.github.ssstlis.excelsorter.config

import java.time.LocalDate
import java.time.format.DateTimeFormatter

import org.apache.poi.ss.usermodel.Row

import scala.util.Try

sealed trait SheetSelector
object SheetSelector {
  case object Default extends SheetSelector
  case class ByName(name: String) extends SheetSelector
  case class ByIndex(index: Int) extends SheetSelector
}

case class TrackCondition(columnIndex: Int, validator: String => Boolean)

case class TrackPolicy(sheetSelector: SheetSelector, conditions: List[TrackCondition])

case class TrackConfig(policies: List[TrackPolicy]) {

  def dataRowDetector(sheetName: String, sheetIndex: Int, getCellValue: (Row, Int) => String): Row => Boolean = {
    val matchingPolicy = policies.find { policy =>
      policy.sheetSelector match {
        case SheetSelector.ByName(name) => name == sheetName
        case SheetSelector.ByIndex(idx) => idx == sheetIndex
        case SheetSelector.Default => false
      }
    }.orElse {
      policies.find(_.sheetSelector == SheetSelector.Default)
    }

    matchingPolicy match {
      case Some(policy) =>
        row => policy.conditions.forall { cond =>
          val value = getCellValue(row, cond.columnIndex)
          cond.validator(value)
        }
      case None =>
        row => {
          val value = getCellValue(row, 0)
          TrackConfig.defaultDateValidator(value)
        }
    }
  }
}

object TrackConfig {
  val empty: TrackConfig = TrackConfig(Nil)

  private val defaultDatePatterns = List(
    DateTimeFormatter.ISO_LOCAL_DATE,
    DateTimeFormatter.ofPattern("dd.MM.yyyy"),
    DateTimeFormatter.ofPattern("dd/MM/yyyy"),
    DateTimeFormatter.ofPattern("MM/dd/yyyy"),
    DateTimeFormatter.ofPattern("yyyy/MM/dd")
  )

  val defaultDateValidator: String => Boolean = { s =>
    if (s == null || s.trim.isEmpty) false
    else {
      defaultDatePatterns.exists { fmt =>
        Try(LocalDate.parse(s.trim, fmt)).isSuccess
      }
    }
  }
}
