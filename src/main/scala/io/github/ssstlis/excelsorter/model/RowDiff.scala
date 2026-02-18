package io.github.ssstlis.excelsorter.model

case class CellDiff(columnName: String, oldColumnIndex: Int, newColumnIndex: Int, oldValue: String, newValue: String)

case class RowDiff(key: String, oldRowNum: Int, newRowNum: Int, cellDiffs: List[CellDiff])
