package io.github.ssstlis.excelsorter.model

case class ColumnMapping(
  commonColumns: List[(Int, Int)],
  oldOnlyColumns: List[(Int, String)],
  newOnlyColumns: List[(Int, String)],
  oldKeyIndices: List[Int],
  newKeyIndices: List[Int]
)
