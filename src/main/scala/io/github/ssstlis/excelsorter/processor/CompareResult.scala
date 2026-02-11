package io.github.ssstlis.excelsorter.processor

case class CompareResult(
  sheetName: String,
  equalRowCount: Int,
  firstMismatchRowNum: Option[Int],
  firstMismatchKey: Option[String]
)

case class HighlightResult(
  sheetName: String,
  matchedSameDataCount: Int,
  matchedDifferentDataCount: Int,
  oldOnlyCount: Int,
  newOnlyCount: Int
)
