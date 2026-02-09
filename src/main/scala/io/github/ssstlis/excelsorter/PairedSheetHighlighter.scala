package io.github.ssstlis.excelsorter

import org.apache.poi.ss.usermodel._

import scala.collection.mutable
import scala.jdk.CollectionConverters._

class PairedSheetHighlighter(
  sheetConfigs: Set[String],
  trackConfig: TrackConfig = TrackConfig.empty
) {

  case class HighlightResult(
    sheetName: String,
    highlightedRowCount: Int
  )

  def highlightEqualLeadingRows(oldSortedPath: String, newSortedPath: String): (String, String, List[HighlightResult]) = {
    val oldCmpPath = buildComparePath(oldSortedPath)
    val newCmpPath = buildComparePath(newSortedPath)

    CellUtils.copyFile(oldSortedPath, oldCmpPath)
    CellUtils.copyFile(newSortedPath, newCmpPath)

    val oldWorkbook = CellUtils.loadWorkbook(oldCmpPath)
    val newWorkbook = CellUtils.loadWorkbook(newCmpPath)

    try {
      val results = sheetConfigs.toList.sorted.flatMap { sheetName =>
        val oldSheet = Option(oldWorkbook.getSheet(sheetName))
        val newSheet = Option(newWorkbook.getSheet(sheetName))

        (oldSheet, newSheet) match {
          case (Some(os), Some(ns)) =>
            val sheetIndex = oldWorkbook.getSheetIndex(os)
            val count = highlightSheet(os, ns, oldWorkbook, newWorkbook, sheetName, sheetIndex)
            if (count > 0) Some(HighlightResult(sheetName, count)) else None
          case _ => None
        }
      }

      CellUtils.writeWorkbook(oldWorkbook, oldCmpPath)
      CellUtils.writeWorkbook(newWorkbook, newCmpPath)

      (oldCmpPath, newCmpPath, results)
    } finally {
      oldWorkbook.close()
      newWorkbook.close()
    }
  }

  private def buildComparePath(sortedPath: String): String = {
    sortedPath.replace("_sorted.xlsx", "_compared.xlsx")
  }

  private def highlightSheet(
    oldSheet: Sheet,
    newSheet: Sheet,
    oldWorkbook: Workbook,
    newWorkbook: Workbook,
    sheetName: String,
    sheetIndex: Int
  ): Int = {
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
      val oldStyleCache = mutable.Map[Short, CellStyle]()
      val newStyleCache = mutable.Map[Short, CellStyle]()

      (0 until equalCount).foreach { i =>
        applyGreenBackground(oldSheet.getRow(oldDataStartIdx + i), oldWorkbook, oldStyleCache)
        applyGreenBackground(newSheet.getRow(newDataStartIdx + i), newWorkbook, newStyleCache)
      }
    }

    equalCount
  }

  private def applyGreenBackground(row: Row, workbook: Workbook, styleCache: mutable.Map[Short, CellStyle]): Unit = {
    if (row == null) return

    val lastCellNum = row.getLastCellNum
    if (lastCellNum < 0) return

    (0 until lastCellNum).foreach { colIdx =>
      val cell = row.getCell(colIdx)
      if (cell != null) {
        val originalStyle = cell.getCellStyle
        val key = originalStyle.getIndex
        val greenStyle = styleCache.getOrElseUpdate(key, {
          val newStyle = workbook.createCellStyle()
          newStyle.cloneStyleFrom(originalStyle)
          newStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex)
          newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
          newStyle
        })
        cell.setCellStyle(greenStyle)
      }
    }
  }
}

object PairedSheetHighlighter {
  def apply(sheetNames: Set[String], trackConfig: TrackConfig = TrackConfig.empty): PairedSheetHighlighter =
    new PairedSheetHighlighter(sheetNames, trackConfig)
}
