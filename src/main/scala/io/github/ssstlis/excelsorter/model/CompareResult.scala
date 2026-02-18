package io.github.ssstlis.excelsorter.model

case class CompareResult(
  sheetName: String,
  equalRowCount: Int,
  firstMismatchRowNum: Option[Int],
  firstMismatchKey: Option[String]
)
