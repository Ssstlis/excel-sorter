package io.github.ssstlis.excelsorter.config.compare

import cats.syntax.traverse._
import cats.instances.list._
import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.CliArgs.parseSheetName
import io.github.ssstlis.excelsorter.config.{SelectedPolicy, SheetSelector}

import scala.jdk.CollectionConverters.ListHasAsScala

case class ComparePolicy(sheetSelector: SheetSelector, ignoreColumns: Set[Int]) extends SelectedPolicy

object ComparePolicy {
  def parseComparePolicy(compareConfig: Config): Either[String, ComparePolicy] = {
    SheetSelector.parseSheetSelector(compareConfig).map { sheetSelector =>
      val ignoreColumns = compareConfig.getIntList("ignoreColumns").asScala.map(_.intValue()).toSet
      ComparePolicy(sheetSelector, ignoreColumns)
    }
  }

  def parseComparisonsBlock(args: List[String]): Either[String, ComparePolicy] = {
    parseSheetName(args).flatMap { case (sheet, rest) =>
      val selector = SheetSelector.parseSheetSelector(sheet)
      parseIgnoreColumns(rest).flatMap { ignoreColumns =>
        Either.cond(ignoreColumns.nonEmpty, ComparePolicy(selector, ignoreColumns), "at least one column index after -ic is required.")
      }
    }.left.map(err => s"--comparisons: $err")
  }

  private def parseIgnoreColumns(args: List[String]): Either[String, Set[Int]] = {
    args match {
      case "-ic" :: Nil =>
        Left("-ic requires at least one column index.")
      case "-ic" :: tail =>
        tail.traverse { s =>
          s.toIntOption.toRight(s"Invalid column index after -ic: '$s'. Expected an integer.")
        }.map(_.toSet)
      case Nil =>
        Left("Expected -ic flag.")
      case other :: _ =>
        Left(s"Unexpected argument: '$other'. Expected -ic.")
    }
  }
}