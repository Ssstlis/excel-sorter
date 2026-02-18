package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.compare.CompareConfig
import io.github.ssstlis.excelsorter.config.track.TrackConfig
import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig
import io.github.ssstlis.excelsorter.model._
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFColor}

import java.text.Normalizer
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

  //format: off
  private[processor] def buildColumnMapping(
    oldSheet: Sheet, newSheet: Sheet,
    oldDataStartIdx: Int, newDataStartIdx: Int,
    sheetName: String
  ): ColumnMapping = {
  //format: on
    val oldHeaderRowIdx = oldDataStartIdx - 1
    val newHeaderRowIdx = newDataStartIdx - 1

    val oldHeaderRow = if (oldHeaderRowIdx >= 0) Option(oldSheet.getRow(oldHeaderRowIdx)) else None
    val newHeaderRow = if (newHeaderRowIdx >= 0) Option(newSheet.getRow(newHeaderRowIdx)) else None

    (oldHeaderRow, newHeaderRow) match {
      case (Some(ohr), Some(nhr)) =>
        val oldHeaders = extractHeaders(ohr)
        val newHeaders = extractHeaders(nhr)

        val newHeadersByName = mutable.Map[String, mutable.Queue[Int]]()
        newHeaders.foreach { case (idx, name) =>
          newHeadersByName.getOrElseUpdate(name, mutable.Queue[Int]()) += idx
        }

        val usedNewIndices = mutable.Set[Int]()
        val commonColumns  = mutable.ListBuffer[(Int, Int)]()
        val oldOnlyCols    = mutable.ListBuffer[(Int, String)]()

        oldHeaders.foreach { case (oldIdx, name) =>
          newHeadersByName.get(name).flatMap(q => if (q.nonEmpty) Some(q.dequeue()) else None) match {
            case Some(newIdx) =>
              commonColumns += ((oldIdx, newIdx))
              usedNewIndices += newIdx
            case None =>
              oldOnlyCols += ((oldIdx, name))
          }
        }

        val newOnlyCols = newHeaders.filterNot { case (idx, _) => usedNewIndices.contains(idx) }

        val sortColumnIndices = sortConfigsMap.get(sheetName) match {
          case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
          case _                                     => List(0)
        }

        val oldKeyIndices = sortColumnIndices
        val newKeyIndices = sortColumnIndices.map { oldIdx =>
          commonColumns.find(_._1 == oldIdx).map(_._2).getOrElse(oldIdx)
        }

        ColumnMapping(commonColumns.toList, oldOnlyCols.toList, newOnlyCols, oldKeyIndices, newKeyIndices)

      case _ =>
        // Fallback: positional mapping
        val oldLastCell = Option(oldSheet.getRow(oldDataStartIdx)).map(_.getLastCellNum.toInt).getOrElse(0)
        val newLastCell = Option(newSheet.getRow(newDataStartIdx)).map(_.getLastCellNum.toInt).getOrElse(0)
        val maxCols     = math.max(oldLastCell, newLastCell)
        val positional  = (0 until maxCols).map(i => (i, i)).toList

        val sortColumnIndices = sortConfigsMap.get(sheetName) match {
          case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
          case _                                     => List(0)
        }

        ColumnMapping(positional, Nil, Nil, sortColumnIndices, sortColumnIndices)
    }
  }

  private def normalizeHeader(s: String): String = {
    val normalized = Normalizer.normalize(s, Normalizer.Form.NFKC)
    normalized.replaceAll("[\\p{Cf}]", "").trim
  }

  private def extractHeaders(row: Row): List[(Int, String)] = {
    val lastCell = row.getLastCellNum
    if (lastCell < 0) Nil
    else
      (0 until lastCell).map { i =>
        i -> normalizeHeader(Option(row.getCell(i)).map(CellUtils.getCellValueAsString).getOrElse(""))
      }.toList
  }

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
        if (CellUtils.rowsAreEqualMapped(oldRow, newRow, mapping.commonColumns, ignoredCols)) {
          applyBackground(oldRow, oldWorkbook, oldStyleCache, Green)
          applyBackground(newRow, newWorkbook, newStyleCache, Green)
          processRowsState.copy(matchedSameData = processRowsState.matchedSameData + 1)
        } else {
          applyBackground(oldRow, oldWorkbook, oldStyleCache, PaleRed)
          applyBackground(newRow, newWorkbook, newStyleCache, PaleRed)
          val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, mapping.commonColumns, headerNames, ignoredCols)
          processRowsState.copy(
            matchedDiffData = processRowsState.matchedDiffData + 1,
            rowDiffs = RowDiff(key, oldRow.getRowNum + 1, newRow.getRowNum + 1, diffs) :: processRowsState.rowDiffs
          )
        }
      case (Some(oldRow), None) =>
        applyBackground(oldRow, oldWorkbook, oldStyleCache, PaleOrange)
        processRowsState.copy(oldOnly = processRowsState.oldOnly + 1)
      case (None, Some(newRow)) =>
        applyBackground(newRow, newWorkbook, newStyleCache, PaleOrange)
        processRowsState.copy(newOnly = processRowsState.newOnly + 1)
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
              case Some(hr) => extractHeaders(hr).toMap
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
          rowDiffs = processRowsState.rowDiffs.reverse,
          oldOnlyColumns = mapping.oldOnlyColumns.map(_._2),
          newOnlyColumns = mapping.newOnlyColumns.map(_._2)
        )
      }
    }
  }

  private def applyBackground(
    row: Row,
    workbook: Workbook,
    styleCache: mutable.Map[(Short, HighlightColor), CellStyle],
    color: HighlightColor
  ): Unit = {
    Option(row).foreach { row =>
      val lastCellNum = row.getLastCellNum
      (if (lastCellNum < 0) None else Some(lastCellNum)).foreach { lastCellNum =>
        (0 until lastCellNum).foreach { colIdx =>
          Option(row.getCell(colIdx)).foreach { cell =>
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
                color match {
                  case Green =>
                    xssfNew.setFillForegroundColor(
                      new XSSFColor(Array[Byte](0xe1.toByte, 0xfa.toByte, 0xe1.toByte), null)
                    )
                  case PaleRed =>
                    xssfNew.setFillForegroundColor(
                      new XSSFColor(Array[Byte](0xff.toByte, 0xcc.toByte, 0xcc.toByte), null)
                    )
                  case PaleOrange =>
                    xssfNew.setFillForegroundColor(
                      new XSSFColor(Array[Byte](0xf5.toByte, 0xe7.toByte, 0x9a.toByte), null)
                    )
                }
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
                newStyle
              }
            )
            cell.setCellStyle(coloredStyle)
          }
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
