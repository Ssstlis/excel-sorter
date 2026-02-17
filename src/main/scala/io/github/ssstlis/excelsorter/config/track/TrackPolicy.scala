package io.github.ssstlis.excelsorter.config.track

import cats.data.NonEmptyList
import cats.instances.list._
import cats.syntax.traverse._
import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.{CliArgs, SelectedPolicy, SheetSelector}

import scala.jdk.CollectionConverters.ListHasAsScala

case class TrackPolicy(sheetSelector: SheetSelector, conditions: NonEmptyList[TrackCondition]) extends SelectedPolicy

object TrackPolicy {
  def parseTrackPolicy(trackConfig: Config): Either[String, Option[TrackPolicy]] = {
    SheetSelector.parseSheetSelector(trackConfig).flatMap { sheetSelector =>
      trackConfig
        .getConfigList("conditions")
        .asScala
        .toList
        .traverse(TrackCondition.parseTrackCondition)
        .map(c => NonEmptyList.fromList(c).map(TrackPolicy(sheetSelector, _)))
    }
  }

  def parseTracksBlock(args: List[String]): Either[String, TrackPolicy] = {
    CliArgs.parseSheetName(args).flatMap { case (sheetName, rest) =>
      val selector = SheetSelector.parseSheetSelector(sheetName)
      TrackCondition.parseCondEntries(rest).map(list => (selector, NonEmptyList.fromList(list)))
    } match {
      case Left(err)                           => Left(s"--tracks: $err")
      case Right((_, None))                    => Left("--tracks: at least one -cond/-d entry is required.")
      case Right((selector, Some(conditions))) => Right(TrackPolicy(selector, conditions))
    }
  }
}
