package io.github.ssstlis.excelsorter.config.track

import cats.instances.list._
import cats.syntax.traverse._
import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.SelectedPolicies
import org.apache.poi.ss.usermodel.Row

import java.time.LocalDate
import java.time.format.DateTimeFormatter
import scala.jdk.CollectionConverters.ListHasAsScala
import scala.util.Try

case class TrackConfig(policies: List[TrackPolicy]) extends SelectedPolicies[TrackPolicy] {

  def dataRowDetector(sheetName: String, sheetIndex: Int, getCellValue: (Row, Int) => String): Row => Boolean = {
    row =>
      matchingPolicy(sheetName, sheetIndex) match {
        case Some(policy) =>
          policy.conditions.forall { cond =>
            val value = getCellValue(row, cond.columnIndex)
            cond.validator(value)
          }
        case None =>
          val value = getCellValue(row, 0)
          TrackConfig.defaultDateValidator(value)
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

  def readTrackConfig(config: Config): Either[String, TrackConfig] = {
    if (!config.hasPath("tracks")) {
      Right(TrackConfig.empty)
    } else {
      val trackList = config.getConfigList("tracks").asScala.toList
      val policies = trackList.traverse(TrackPolicy.parseTrackPolicy)
      policies.map(p => TrackConfig(p.flatten))
    }
  }
}