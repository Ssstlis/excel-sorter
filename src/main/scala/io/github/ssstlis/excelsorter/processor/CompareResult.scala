package io.github.ssstlis.excelsorter.processor

case class CompareResult(
  sheetName: String,
  equalRowCount: Int,
  firstMismatchRowNum: Option[Int],
  firstMismatchKey: Option[String]
)
