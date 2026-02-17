package io.github.ssstlis.excelsorter.config.track

import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.dsl.SortingDsl.Parsers.LocalDatePattern

import java.time.LocalDate
import java.time.format.DateTimeFormatter
import scala.annotation.tailrec
import scala.util.Try

case class TrackCondition(columnIndex: Int, validator: String => Boolean)

object TrackCondition {
  def parseTrackCondition(condConfig: Config): Either[String, TrackCondition] = {
    val index = condConfig.getInt("index")
    val as = condConfig.getString("as")
    resolveTrackValidator(as).map(TrackCondition(index, _))
  }

  def resolveTrackValidator(asType: String): Either[String, String => Boolean] = asType match {
    case "String" =>
      Right(s => s != null && s.nonEmpty)
    case "Int" =>
      Right(s => Try(s.toInt).isSuccess)
    case "Long" =>
      Right(s => Try(s.toLong).isSuccess)
    case "Double" =>
      Right(s => Try(s.toDouble).isSuccess)
    case "BigDecimal" =>
      Right(s => Try(BigDecimal(s)).isSuccess)
    case "LocalDate" =>
      Right(s => TrackConfig.defaultDateValidator(s))
    case LocalDatePattern(pattern) =>
      val fmt = DateTimeFormatter.ofPattern(pattern)
      Right(s => s != null && s.trim.nonEmpty && Try(LocalDate.parse(s.trim, fmt)).isSuccess)
    case other =>
      Left(
        s"Unknown track condition type: '$other'. Expected one of: String, Int, Long, Double, BigDecimal, LocalDate, LocalDate(<pattern>)."
      )
  }

  def parseCondEntries(args: List[String]): Either[String, List[TrackCondition]] = {
    val condFlags = Set("-cond", "-d")

    @tailrec
    def rec(rest: List[String], result: List[TrackCondition]): Either[String, List[TrackCondition]] = {
      rest match {
        case flag :: idxStr :: asType :: tail if condFlags.contains(flag) =>
          idxStr
            .toIntOption
            .toRight(s"Invalid column index: '$idxStr'. Expected an integer.")
            .flatMap { index =>
              resolveTrackValidator(asType).map { validator =>
                TrackCondition(index, validator)
              }
            } match {
            case Right(trackCondition) => rec(tail, trackCondition :: result)
            case Left(err) => Left(err)
          }
        case flag :: _ if condFlags.contains(flag) =>
          Left(s"$flag requires 2 arguments: <column-index> <type>")
        case other :: _ => Left(s"Unexpected argument: '$other'. Expected -cond or -d.")
        case Nil => Right(result.reverse)
      }
    }

    rec(args, Nil)
  }
}