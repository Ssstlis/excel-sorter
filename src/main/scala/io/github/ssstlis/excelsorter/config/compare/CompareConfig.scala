package io.github.ssstlis.excelsorter.config.compare

import cats.syntax.traverse._
import cats.instances.list._
import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.SelectedPolicies

import scala.jdk.CollectionConverters.ListHasAsScala

case class CompareConfig(policies: List[ComparePolicy]) extends SelectedPolicies[ComparePolicy] {

  def ignoredColumns(sheetName: String, sheetIndex: Int): Set[Int] = {
    matchingPolicy(sheetName, sheetIndex).map(_.ignoreColumns).getOrElse(Set.empty)
  }
}

object CompareConfig {
  val empty: CompareConfig = CompareConfig(Nil)

  def readCompareConfig(config: Config): Either[String, CompareConfig] = {
    if (!config.hasPath("comparisons")) {
      Right(CompareConfig.empty)
    } else {
      config.getConfigList("comparisons")
        .asScala
        .toList
        .traverse(ComparePolicy.parseComparePolicy)
        .map(CompareConfig(_))
    }
  }
}