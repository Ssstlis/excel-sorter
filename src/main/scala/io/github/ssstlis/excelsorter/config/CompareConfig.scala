package io.github.ssstlis.excelsorter.config

case class ComparePolicy(sheetSelector: SheetSelector, ignoreColumns: Set[Int])

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
