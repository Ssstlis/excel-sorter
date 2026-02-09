package io.github.ssstlis.excelsorter

import org.apache.poi.ss.usermodel._

import scala.jdk.CollectionConverters._

class PairedSheetCutter(
  sheetConfigs: Set[String],
  trackConfig: TrackConfig = TrackConfig.empty
) {

  case class CutResult(
    sheetName: String,
    removedRowCount: Int
  )

  def cutEqualLeadingRows(oldSortedPath: String, newSortedPath: String): (String, String, List[CutResult]) = {
    val oldCutPath = buildCutPath(oldSortedPath)
    val newCutPath = buildCutPath(newSortedPath)

    CellUtils.copyFile(oldSortedPath, oldCutPath)
    CellUtils.copyFile(newSortedPath, newCutPath)

    val oldWorkbook = CellUtils.loadWorkbook(oldCutPath)
    val newWorkbook = CellUtils.loadWorkbook(newCutPath)

    try {
      val results = sheetConfigs.toList.sorted.flatMap { sheetName =>
        val oldSheet = Option(oldWorkbook.getSheet(sheetName))
        val newSheet = Option(newWorkbook.getSheet(sheetName))

        (oldSheet, newSheet) match {
          case (Some(os), Some(ns)) =>
            val sheetIndex = oldWorkbook.getSheetIndex(os)
            val removed = removeEqualLeadingRows(os, ns, sheetName, sheetIndex)
            if (removed > 0) Some(CutResult(sheetName, removed)) else None
          case _ => None
        }
      }

      CellUtils.writeWorkbook(oldWorkbook, oldCutPath)
      CellUtils.writeWorkbook(newWorkbook, newCutPath)

      (oldCutPath, newCutPath, results)
    } finally {
      oldWorkbook.close()
      newWorkbook.close()
    }
  }

  private def buildCutPath(sortedPath: String): String = {
    sortedPath.replace("_sorted.xlsx", "_sortcutted.xlsx")
  }

  private def removeEqualLeadingRows(oldSheet: Sheet, newSheet: Sheet, sheetName: String, sheetIndex: Int): Int = {
    val oldRows = oldSheet.iterator().asScala.toList
    val newRows = newSheet.iterator().asScala.toList

    if (oldRows.isEmpty || newRows.isEmpty) return 0

    val isDataRow = trackConfig.dataRowDetector(sheetName, sheetIndex, CellUtils.getRowCellValue)

    val oldDataStartIdx = oldRows.indexWhere(isDataRow)
    val newDataStartIdx = newRows.indexWhere(isDataRow)

    if (oldDataStartIdx < 0 || newDataStartIdx < 0) return 0

    val oldDataRows = oldRows.drop(oldDataStartIdx)
    val newDataRows = newRows.drop(newDataStartIdx)

    val equalCount = oldDataRows.zip(newDataRows).takeWhile { case (oldRow, newRow) =>
      CellUtils.rowsAreEqual(oldRow, newRow)
    }.size

    if (equalCount > 0) {
      removeRows(oldSheet, oldDataStartIdx, equalCount)
      removeRows(newSheet, newDataStartIdx, equalCount)
    }

    equalCount
  }

  private def removeRows(sheet: Sheet, startIdx: Int, count: Int): Unit = {
    val lastRowNum = sheet.getLastRowNum

    (0 until count).foreach { i =>
      val row = sheet.getRow(startIdx + i)
      if (row != null) {
        sheet.removeRow(row)
      }
    }

    val remainingDataStart = startIdx + count
    if (remainingDataStart <= lastRowNum) {
      sheet.shiftRows(remainingDataStart, lastRowNum, -count)
    }
  }
}

object PairedSheetCutter {
  def apply(sheetNames: Set[String], trackConfig: TrackConfig = TrackConfig.empty): PairedSheetCutter =
    new PairedSheetCutter(sheetNames, trackConfig)
}
