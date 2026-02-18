package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.compare.CompareConfig
import io.github.ssstlis.excelsorter.config.track.TrackConfig
import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig
import io.github.ssstlis.excelsorter.model._
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

  private[processor] def buildColumnMapping(
    oldSheet: Sheet,
    newSheet: Sheet,
    oldDataStartIdx: Int,
    newDataStartIdx: Int,
    sheetName: String
  ): ColumnMapping =
    CellUtils.buildColumnMapping(oldSheet, newSheet, oldDataStartIdx, newDataStartIdx, sheetName, sortConfigsMap)

  private def extractKey(row: Row, keyColumnIndices: List[Int]): String = {
    keyColumnIndices.map(col => CellUtils.getRowCellValue(row, col)).mkString(", ")
  }

  private case class ProcessRowsState(
    matchedSameData: Int = 0,
    matchedDiffData: Int = 0,
    oldOnly: Int = 0,
    newOnly: Int = 0,
    rowDiffs: List[RowDiff] = Nil
  )

  //format: off
  private def processRows(
    key: String,
    oldByKey: Map[String, Row], newByKey: Map[String, Row],
    mapping: ColumnMapping, headerNames: Map[Int, String], ignoredCols: Set[Int],
    oldWorkbook: Workbook, oldStyleCache: mutable.Map[(Short, HighlightColor), CellStyle],
    newWorkbook: Workbook, newStyleCache: mutable.Map[(Short, HighlightColor), CellStyle],
    processRowsState: ProcessRowsState
  ): ProcessRowsState = {
    //format: on
    (oldByKey.get(key), newByKey.get(key)) match {
      case (Some(oldRow), Some(newRow)) =>
        val diffs =
          CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.from(mapping.commonColumns), headerNames, ignoredCols)
        val oldDiffCols = diffs.map(_.oldColumnIndex).toSet ++ mapping.oldOnlyColumns.map(_._1).toSet
        val newDiffCols = diffs.map(_.newColumnIndex).toSet ++ mapping.newOnlyColumns.map(_._1).toSet
        applyBackground(oldRow, oldWorkbook, oldStyleCache, oldDiffCols, Green)
        applyBackground(newRow, newWorkbook, newStyleCache, newDiffCols, Green)
        if (diffs.isEmpty) {
          processRowsState.copy(matchedSameData = processRowsState.matchedSameData + 1)
        } else {
          processRowsState.copy(
            matchedDiffData = processRowsState.matchedDiffData + 1,
            rowDiffs = RowDiff(key, oldRow.getRowNum + 1, newRow.getRowNum + 1, diffs) :: processRowsState.rowDiffs
          )
        }
      case (Some(oldRow), None) =>
        applyBackground(oldRow, oldWorkbook, oldStyleCache, Set.empty, PaleOrange)
        processRowsState.copy(oldOnly = processRowsState.oldOnly + 1)
      case (None, Some(newRow)) =>
        applyBackground(newRow, newWorkbook, newStyleCache, Set.empty, PaleOrange)
        processRowsState.copy(newOnly = processRowsState.newOnly + 1)
      case _ => processRowsState
    }
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

    if (oldRows.isEmpty || newRows.isEmpty) {
      HighlightResult(sheetName, 0, 0, 0, 0)
    } else {

      val isDataRow   = trackConfig.dataRowDetector(sheetName, sheetIndex, CellUtils.getRowCellValue)
      val ignoredCols = compareConfig.ignoredColumns(sheetName, sheetIndex)

      val oldDataStartIdx = oldRows.indexWhere(isDataRow)
      val newDataStartIdx = newRows.indexWhere(isDataRow)

      if (oldDataStartIdx < 0 || newDataStartIdx < 0) {
        HighlightResult(sheetName, 0, 0, 0, 0)
      } else {
        val mapping = buildColumnMapping(oldSheet, newSheet, oldDataStartIdx, newDataStartIdx, sheetName)

        val headerNames: Map[Int, String] = {
          val headerRowIdx = oldDataStartIdx - 1
          if (headerRowIdx >= 0) {
            Option(oldSheet.getRow(headerRowIdx)) match {
              case Some(hr) => CellUtils.extractHeaders(hr).toMap
              case None     => Map.empty[Int, String]
            }
          } else Map.empty[Int, String]
        }

        val oldDataRows = oldRows.drop(oldDataStartIdx)
        val newDataRows = newRows.drop(newDataStartIdx)

        val oldByKey: Map[String, Row] = oldDataRows.map(r => extractKey(r, mapping.oldKeyIndices) -> r).toMap
        val newByKey: Map[String, Row] = newDataRows.map(r => extractKey(r, mapping.newKeyIndices) -> r).toMap

        val allKeys = oldByKey.keySet ++ newByKey.keySet

        val oldStyleCache = mutable.Map[(Short, HighlightColor), CellStyle]()
        val newStyleCache = mutable.Map[(Short, HighlightColor), CellStyle]()

        //format: off
        val processRowsState = allKeys.foldLeft(ProcessRowsState()) { case (state, key) =>
          processRows(
            key,
            oldByKey, newByKey,
            mapping, headerNames, ignoredCols,
            oldWorkbook, oldStyleCache,
            newWorkbook, newStyleCache,
            state
          )
        }
        //format: on

        HighlightResult(
          sheetName = sheetName,
          matchedSameDataCount = processRowsState.matchedSameData,
          matchedDifferentDataCount = processRowsState.matchedDiffData,
          oldOnlyCount = processRowsState.oldOnly,
          newOnlyCount = processRowsState.newOnly,
          rowDiffs = processRowsState.rowDiffs,
          oldOnlyColumns = mapping.oldOnlyColumns.map(_._2),
          newOnlyColumns = mapping.newOnlyColumns.map(_._2)
        )
      }
    }
  }

  private def buildXssfColor(r: Int, g: Int, b: Int) = new XSSFColor(Array[Byte](r.toByte, g.toByte, b.toByte), null)

  private val xssfGreenColor      = buildXssfColor(0xe1, 0xfa, 0xe1)
  private val xssfPaleRedColor    = buildXssfColor(0xff, 0xcc, 0xcc)
  private val xssfPaleOrangeColor = buildXssfColor(0xf5, 0xe7, 0x9a)

  private def styleSetFillForegroundColor(style: XSSFCellStyle, color: HighlightColor): Unit = {
    color match {
      case Green      => style.setFillForegroundColor(xssfGreenColor)
      case PaleRed    => style.setFillForegroundColor(xssfPaleRedColor)
      case PaleOrange => style.setFillForegroundColor(xssfPaleOrangeColor)
    }
  }

  private def applyCellBackgroundStyle(
    cell: Cell,
    workbook: Workbook,
    styleCache: mutable.Map[(Short, HighlightColor), CellStyle],
    color: HighlightColor
  ): Unit = {
    val originalStyle = cell.getCellStyle
    val key           = (originalStyle.getIndex, color)
    val coloredStyle  = styleCache.getOrElseUpdate(
      key, {
        val newStyle = workbook.createCellStyle()
        newStyle.cloneStyleFrom(originalStyle)
        newStyle.setBorderTop(BorderStyle.THIN)
        newStyle.setBorderBottom(BorderStyle.THIN)
        newStyle.setBorderLeft(BorderStyle.THIN)
        newStyle.setBorderRight(BorderStyle.THIN)
        newStyle.setTopBorderColor(IndexedColors.BLACK.getIndex)
        newStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex)
        newStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex)
        newStyle.setRightBorderColor(IndexedColors.BLACK.getIndex)
        val xssfNew = newStyle.asInstanceOf[XSSFCellStyle]
        styleSetFillForegroundColor(xssfNew, color)
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        newStyle
      }
    )
    cell.setCellStyle(coloredStyle)
  }

  private def applyBackground(
    row: Row,
    workbook: Workbook,
    styleCache: mutable.Map[(Short, HighlightColor), CellStyle],
    diffColumnIndices: Set[Int],
    color: HighlightColor
  ): Unit = {
    Option(row).foreach { row =>
      (0 until row.getLastCellNum).foreach { colIdx =>
        Option(row.getCell(colIdx)).foreach { cell =>
          val colorToApply = if (diffColumnIndices.contains(colIdx)) PaleRed else color
          applyCellBackgroundStyle(cell, workbook, styleCache, colorToApply)
        }
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
