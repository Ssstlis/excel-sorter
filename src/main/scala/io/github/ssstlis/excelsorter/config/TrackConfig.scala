package io.github.ssstlis.excelsorter.config

import io.github.ssstlis.excelsorter.config.CliArgs.parseSheetName
import io.github.ssstlis.excelsorter.config.SheetSelector.parseSheetSelector

import java.time.LocalDate
import java.time.format.DateTimeFormatter
import org.apache.poi.ss.usermodel.Row

import scala.util.Try

case class TrackCondition(columnIndex: Int, validator: String => Boolean)

case class TrackPolicy(sheetSelector: SheetSelector, conditions: List[TrackCondition])

object TrackPolicy {
  def parseTracksBlock(args: List[String]): Either[String, TrackPolicy] = {
    parseSheetName(args).flatMap { case (sheetNameOrDefault, rest) =>
      val selector = SheetSelector.parseSheetSelector(sheetNameOrDefault)
      parseCondEntries(rest).map((selector, _))
    } match {
      case Left(err) => Left(s"--tracks: $err")
      case Right((_, Nil)) => Left("--tracks: at least one -cond/-d entry is required.")
      case Right((selector, conditions)) => Right(TrackPolicy(selector, conditions))
    }
  }

  private def parseCondEntries(args: List[String]): Either[String, List[TrackCondition]] = {
    val condFlags = Set("-cond", "-d")
    var remaining = args
    val result = List.newBuilder[TrackCondition]

    while (remaining.nonEmpty) {
      remaining match {
        case flag :: idxStr :: asType :: tail if condFlags.contains(flag) =>
          val index = try { idxStr.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index: '$idxStr'. Expected an integer.")
          }
          val validator = try {
            ConfigReader.resolveTrackValidator(asType)
          } catch {
            case e: IllegalArgumentException => return Left(e.getMessage)
          }
          result += TrackCondition(index, validator)
          remaining = tail
        case flag :: _ if condFlags.contains(flag) =>
          return Left(s"$flag requires 2 arguments: <column-index> <type>")
        case other :: _ =>
          return Left(s"Unexpected argument: '$other'. Expected -cond or -d.")
        case Nil =>
      }
    }

    Right(result.result())
  }
}

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
