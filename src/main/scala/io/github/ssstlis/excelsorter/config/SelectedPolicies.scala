package io.github.ssstlis.excelsorter.config

trait SelectedPolicies[T <: SelectedPolicy] {
  def policies: List[T]

  def matchingPolicy(sheetName: String, sheetIndex: Int): Option[T] = {
    policies.find { policy =>
      policy.sheetSelector match {
        case SheetSelector.ByName(name) => name == sheetName
        case SheetSelector.ByIndex(idx) => idx == sheetIndex
        case SheetSelector.Default      => false
      }
    }.orElse {
      policies.find(_.sheetSelector == SheetSelector.Default)
    }
  }

}
