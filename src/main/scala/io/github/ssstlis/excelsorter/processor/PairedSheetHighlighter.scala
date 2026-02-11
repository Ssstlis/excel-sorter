package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.{CompareConfig, TrackConfig}
import io.github.ssstlis.excelsorter.dsl.SheetSortingConfig
import org.apache.poi.ss.usermodel._

import scala.collection.mutable
import scala.jdk.CollectionConverters._

class PairedSheetHighlighter(
  sheetConfigs: Set[String],
  trackConfig: TrackConfig = TrackConfig.empty,
  compareConfig: CompareConfig = CompareConfig.empty,
  sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
) {

  def highlightEqualLeadingRows(oldSortedPath: String, newSortedPath: String): (String, String, List[CompareResult]) = {
    val oldCmpPath = buildComparePath(oldSortedPath)
    val newCmpPath = buildComparePath(newSortedPath)

    CellUtils.copyFile(oldSortedPath, oldCmpPath)
    CellUtils.copyFile(newSortedPath, newCmpPath)

    val oldWorkbook = CellUtils.loadWorkbook(oldCmpPath)
    val newWorkbook = CellUtils.loadWorkbook(newCmpPath)

    try {
      val results = sheetConfigs.toList.sorted.map { sheetName =>
        val oldSheet = Option(oldWorkbook.getSheet(sheetName))
        val newSheet = Option(newWorkbook.getSheet(sheetName))

        (oldSheet, newSheet) match {
          case (Some(os), Some(ns)) =>
            val sheetIndex = oldWorkbook.getSheetIndex(os)
            highlightSheet(os, ns, oldWorkbook, newWorkbook, sheetName, sheetIndex)
          case _ =>
            CompareResult(sheetName, 0, None, None)
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
  ): CompareResult = {
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
      val oldStyleCache = mutable.Map[Short, CellStyle]()
      val newStyleCache = mutable.Map[Short, CellStyle]()

      (0 until equalCount).foreach { i =>
        applyGreenBackground(oldSheet.getRow(oldDataStartIdx + i), oldWorkbook, oldStyleCache)
        applyGreenBackground(newSheet.getRow(newDataStartIdx + i), newWorkbook, newStyleCache)
      }
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
  def apply(
    sheetNames: Set[String],
    trackConfig: TrackConfig = TrackConfig.empty,
    compareConfig: CompareConfig = CompareConfig.empty,
    sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
  ): PairedSheetHighlighter =
    new PairedSheetHighlighter(sheetNames, trackConfig, compareConfig, sortConfigsMap)
}
