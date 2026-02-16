package io.github.ssstlis.excelsorter.config

import io.github.ssstlis.excelsorter.config.CliArgs.parseSheetName

case class ComparePolicy(sheetSelector: SheetSelector, ignoreColumns: Set[Int])

object ComparePolicy {
  def parseComparisonsBlock(args: List[String]): Either[String, ComparePolicy] = {
    parseSheetName(args).flatMap { case (sheetNameOrDefault, rest) =>
      val selector = SheetSelector.parseSheetSelector(sheetNameOrDefault)
      parseIgnoreColumns(rest).flatMap { ignoreColumns =>
        Either.cond(ignoreColumns.nonEmpty, ComparePolicy(selector, ignoreColumns), "at least one column index after -ic is required.")
      }
    }.left.map(err => s"--comparisons: $err")
  }

  private def parseIgnoreColumns(args: List[String]): Either[String, Set[Int]] = {
    args match {
      case "-ic" :: tail =>
        if (tail.isEmpty) return Left("-ic requires at least one column index.")
        val indices = tail.map { s =>
          try { s.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index after -ic: '$s'. Expected an integer.")
          }
        }
        Right(indices.toSet)
      case Nil =>
        Left("Expected -ic flag.")
      case other :: _ =>
        Left(s"Unexpected argument: '$other'. Expected -ic.")
    }
  }
}

case class CompareConfig(policies: List[ComparePolicy]) {

  def ignoredColumns(sheetName: String, sheetIndex: Int): Set[Int] = {
    val matchingPolicy = policies.find { policy =>
      policy.sheetSelector match {
        case SheetSelector.ByName(name) => name == sheetName
        case SheetSelector.ByIndex(idx) => idx == sheetIndex
        case SheetSelector.Default      => false
      }
    }.orElse {
      policies.find(_.sheetSelector == SheetSelector.Default)
    }

    matchingPolicy.map(_.ignoreColumns).getOrElse(Set.empty)
  }
}

object CompareConfig {
  val empty: CompareConfig = CompareConfig(Nil)
}
