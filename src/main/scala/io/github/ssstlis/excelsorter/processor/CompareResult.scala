package io.github.ssstlis.excelsorter.processor

case class CompareResult(
  sheetName: String,
  equalRowCount: Int,
  firstMismatchRowNum: Option[Int],
  firstMismatchKey: Option[String]
)

case class ColumnMapping(
  commonColumns: List[(Int, Int)],
  oldOnlyColumns: List[(Int, String)],
  newOnlyColumns: List[(Int, String)],
  oldKeyIndices: List[Int],
  newKeyIndices: List[Int]
)

case class CellDiff(
  columnName: String,
  oldColumnIndex: Int,
  newColumnIndex: Int,
  oldValue: String,
  newValue: String
)

case class RowDiff(
  key: String,
  oldRowNum: Int,
  newRowNum: Int,
  cellDiffs: List[CellDiff]
)

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
