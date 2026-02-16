package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.{CompareConfig, TrackConfig}
import io.github.ssstlis.excelsorter.dsl.config.SheetSortingConfig
import org.apache.poi.ss.usermodel._

import scala.jdk.CollectionConverters._

class PairedSheetCutter(
  sheetConfigs: Set[String],
  trackConfig: TrackConfig = TrackConfig.empty,
  compareConfig: CompareConfig = CompareConfig.empty,
  sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
) {

  def cutEqualLeadingRows(oldSortedPath: String, newSortedPath: String): (String, String, List[CompareResult]) = {
    val oldCutPath = buildCutPath(oldSortedPath)
    val newCutPath = buildCutPath(newSortedPath)

    CellUtils.copyFile(oldSortedPath, oldCutPath)
    CellUtils.copyFile(newSortedPath, newCutPath)

    val oldWorkbook = CellUtils.loadWorkbook(oldCutPath)
    val newWorkbook = CellUtils.loadWorkbook(newCutPath)

    try {
      val results = sheetConfigs.toList.sorted.map { sheetName =>
        val oldSheet = Option(oldWorkbook.getSheet(sheetName))
        val newSheet = Option(newWorkbook.getSheet(sheetName))

        (oldSheet, newSheet) match {
          case (Some(os), Some(ns)) =>
            val sheetIndex = oldWorkbook.getSheetIndex(os)
            processSheet(os, ns, sheetName, sheetIndex)
          case _ =>
            CompareResult(sheetName, 0, None, None)
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

  private def processSheet(oldSheet: Sheet, newSheet: Sheet, sheetName: String, sheetIndex: Int): CompareResult = {
    val oldRows = oldSheet.iterator().asScala.toList
    val newRows = newSheet.iterator().asScala.toList

    if (oldRows.isEmpty || newRows.isEmpty) return CompareResult(sheetName, 0, None, None)

    val isDataRow = trackConfig.dataRowDetector(sheetName, sheetIndex, CellUtils.getRowCellValue)
    val ignoredCols = compareConfig.ignoredColumns(sheetName, sheetIndex)

    val oldDataStartIdx = oldRows.indexWhere(isDataRow)
    val newDataStartIdx = newRows.indexWhere(isDataRow)

    if (oldDataStartIdx < 0 || newDataStartIdx < 0) return CompareResult(sheetName, 0, None, None)

    val oldDataRows = oldRows.drop(oldDataStartIdx)
    val newDataRows = newRows.drop(newDataStartIdx)

    val equalCount = oldDataRows.zip(newDataRows).takeWhile { case (oldRow, newRow) =>
      CellUtils.rowsAreEqual(oldRow, newRow, ignoredCols)
    }.size

    val (mismatchRowNum, mismatchKey) = findFirstMismatch(
      oldDataRows, newDataRows, equalCount, oldDataStartIdx, sheetName
    )

    if (equalCount > 0) {
      removeRows(oldSheet, oldDataStartIdx, equalCount)
      removeRows(newSheet, newDataStartIdx, equalCount)
    }

    CompareResult(sheetName, equalCount, mismatchRowNum, mismatchKey)
  }

  private def findFirstMismatch(
    oldDataRows: List[Row],
    newDataRows: List[Row],
    equalCount: Int,
    dataStartIdx: Int,
    sheetName: String
  ): (Option[Int], Option[String]) = {
    if (equalCount < oldDataRows.size && equalCount < newDataRows.size) {
      val mismatchRow = oldDataRows(equalCount)
      val excelRowNum = dataStartIdx + equalCount + 1
      val key = extractKey(mismatchRow, sheetName)
      (Some(excelRowNum), Some(key))
    } else {
      (None, None)
    }
  }

  private def extractKey(row: Row, sheetName: String): String = {
    val keyColumns = sortConfigsMap.get(sheetName) match {
      case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
      case _ => List(0)
    }
    keyColumns.map(col => CellUtils.getRowCellValue(row, col)).mkString(", ")
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
  def apply(
    sheetNames: Set[String],
    trackConfig: TrackConfig = TrackConfig.empty,
    compareConfig: CompareConfig = CompareConfig.empty,
    sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
  ): PairedSheetCutter =
    new PairedSheetCutter(sheetNames, trackConfig, compareConfig, sortConfigsMap)
}
