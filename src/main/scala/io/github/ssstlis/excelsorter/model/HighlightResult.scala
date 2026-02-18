package io.github.ssstlis.excelsorter.model

case class HighlightResult(
  sheetName: String,
  matchedSameDataCount: Int,
  matchedDifferentDataCount: Int,
  oldOnlyCount: Int,
  newOnlyCount: Int,
  rowDiffs: List[RowDiff] = Nil,
  oldOnlyColumns: List[String] = Nil,
  newOnlyColumns: List[String] = Nil
)
