package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.{CompareConfig, TrackConfig}
import io.github.ssstlis.excelsorter.dsl.SheetSortingConfig
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor}

import scala.collection.mutable
import scala.jdk.CollectionConverters._

class PairedSheetHighlighter(
  sheetConfigs: Set[String],
  trackConfig: TrackConfig = TrackConfig.empty,
  compareConfig: CompareConfig = CompareConfig.empty,
  sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
) {

  private sealed trait HighlightColor
  private case object Green extends HighlightColor
  private case object PaleRed extends HighlightColor
  private case object PaleOrange extends HighlightColor

  def highlightPairedSheets(oldSortedPath: String, newSortedPath: String): (String, String, List[HighlightResult]) = {
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
            HighlightResult(sheetName, 0, 0, 0, 0)
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
  ): HighlightResult = {
    val oldRows = oldSheet.iterator().asScala.toList
    val newRows = newSheet.iterator().asScala.toList

    if (oldRows.isEmpty || newRows.isEmpty) return HighlightResult(sheetName, 0, 0, 0, 0)

    val isDataRow = trackConfig.dataRowDetector(sheetName, sheetIndex, CellUtils.getRowCellValue)
    val ignoredCols = compareConfig.ignoredColumns(sheetName, sheetIndex)

    val oldDataStartIdx = oldRows.indexWhere(isDataRow)
    val newDataStartIdx = newRows.indexWhere(isDataRow)

    if (oldDataStartIdx < 0 || newDataStartIdx < 0) return HighlightResult(sheetName, 0, 0, 0, 0)

    val oldDataRows = oldRows.drop(oldDataStartIdx)
    val newDataRows = newRows.drop(newDataStartIdx)

    val oldByKey: Map[String, Row] = oldDataRows.map(r => extractKey(r, sheetName) -> r).toMap
    val newByKey: Map[String, Row] = newDataRows.map(r => extractKey(r, sheetName) -> r).toMap

    val allKeys = oldByKey.keySet ++ newByKey.keySet

    val oldStyleCache = mutable.Map[(Short, HighlightColor), CellStyle]()
    val newStyleCache = mutable.Map[(Short, HighlightColor), CellStyle]()

    var matchedSameData = 0
    var matchedDiffData = 0
    var oldOnly = 0
    var newOnly = 0

    allKeys.foreach { key =>
      (oldByKey.get(key), newByKey.get(key)) match {
        case (Some(oldRow), Some(newRow)) =>
          if (CellUtils.rowsAreEqual(oldRow, newRow, ignoredCols)) {
            applyBackground(oldRow, oldWorkbook, oldStyleCache, Green)
            applyBackground(newRow, newWorkbook, newStyleCache, Green)
            matchedSameData += 1
          } else {
            applyBackground(oldRow, oldWorkbook, oldStyleCache, PaleRed)
            applyBackground(newRow, newWorkbook, newStyleCache, PaleRed)
            matchedDiffData += 1
          }
        case (Some(oldRow), None) =>
          applyBackground(oldRow, oldWorkbook, oldStyleCache, PaleOrange)
          oldOnly += 1
        case (None, Some(newRow)) =>
          applyBackground(newRow, newWorkbook, newStyleCache, PaleOrange)
          newOnly += 1
        case _ => // impossible
      }
    }

    HighlightResult(sheetName, matchedSameData, matchedDiffData, oldOnly, newOnly)
  }

  private def extractKey(row: Row, sheetName: String): String = {
    val keyColumns = sortConfigsMap.get(sheetName) match {
      case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
      case _ => List(0)
    }
    keyColumns.map(col => CellUtils.getRowCellValue(row, col)).mkString(", ")
  }

  private def applyBackground(row: Row, workbook: Workbook, styleCache: mutable.Map[(Short, HighlightColor), CellStyle], color: HighlightColor): Unit = {
    if (row == null) return

    val lastCellNum = row.getLastCellNum
    if (lastCellNum < 0) return

    (0 until lastCellNum).foreach { colIdx =>
      val cell = row.getCell(colIdx)
      if (cell != null) {
        val originalStyle = cell.getCellStyle
        val key = (originalStyle.getIndex, color)
        val coloredStyle = styleCache.getOrElseUpdate(key, {
          val newStyle = workbook.createCellStyle()
          newStyle.cloneStyleFrom(originalStyle)
          color match {
            case Green =>
              newStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex)
            case PaleRed =>
              newStyle.asInstanceOf[XSSFCellStyle].setFillForegroundColor(
                new XSSFColor(Array[Byte](0xFF.toByte, 0xCC.toByte, 0xCC.toByte), null)
              )
            case PaleOrange =>
              newStyle.asInstanceOf[XSSFCellStyle].setFillForegroundColor(
                new XSSFColor(Array[Byte](0xFF.toByte, 0xE5.toByte, 0xCC.toByte), null)
              )
          }
          newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
          newStyle
        })
        cell.setCellStyle(coloredStyle)
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
