package io.github.ssstlis.excelsorter

import java.io.{FileInputStream, FileOutputStream}
import java.time.format.DateTimeFormatter

import org.apache.poi.ss.usermodel._

import scala.jdk.CollectionConverters._
import scala.util.Try

class PairedSheetComparer(
  sheetConfigs: Set[String],
  dateValidator: String => Boolean = SheetSorter.defaultDateValidator
) {

  case class CompareResult(
    sheetName: String,
    removedRowCount: Int
  )

  def compareAndRemoveEqualLeadingRows(oldPath: String, newPath: String): List[CompareResult] = {
    val oldWorkbook = loadWorkbook(oldPath)
    val newWorkbook = loadWorkbook(newPath)

    try {
      val results = sheetConfigs.toList.flatMap { sheetName =>
        val oldSheet = Option(oldWorkbook.getSheet(sheetName))
        val newSheet = Option(newWorkbook.getSheet(sheetName))

        (oldSheet, newSheet) match {
          case (Some(os), Some(ns)) =>
            val removed = removeEqualLeadingRows(os, ns)
            if (removed > 0) Some(CompareResult(sheetName, removed)) else None
          case _ => None
        }
      }

      writeWorkbook(oldWorkbook, oldPath)
      writeWorkbook(newWorkbook, newPath)

      results
    } finally {
      oldWorkbook.close()
      newWorkbook.close()
    }
  }

  private def loadWorkbook(path: String): Workbook = {
    val fis = new FileInputStream(path)
    try {
      WorkbookFactory.create(fis)
    } finally {
      fis.close()
    }
  }

  private def writeWorkbook(workbook: Workbook, path: String): Unit = {
    val fos = new FileOutputStream(path)
    try {
      workbook.write(fos)
    } finally {
      fos.close()
    }
  }

  private def removeEqualLeadingRows(oldSheet: Sheet, newSheet: Sheet): Int = {
    val oldRows = oldSheet.iterator().asScala.toList
    val newRows = newSheet.iterator().asScala.toList

    if (oldRows.isEmpty || newRows.isEmpty) return 0

    val oldDataStartIdx = findDataStartIndex(oldRows)
    val newDataStartIdx = findDataStartIndex(newRows)

    if (oldDataStartIdx < 0 || newDataStartIdx < 0) return 0

    val oldDataRows = oldRows.drop(oldDataStartIdx)
    val newDataRows = newRows.drop(newDataStartIdx)

    val equalCount = countEqualLeadingRows(oldDataRows, newDataRows)

    if (equalCount > 0) {
      removeRows(oldSheet, oldDataStartIdx, equalCount)
      removeRows(newSheet, newDataStartIdx, equalCount)
    }

    equalCount
  }

  private def findDataStartIndex(rows: List[Row]): Int = {
    rows.indexWhere { row =>
      val firstCell = Option(row.getCell(0))
      val cellValue = firstCell.map(getCellValueAsString).getOrElse("")
      dateValidator(cellValue)
    }
  }

  private def countEqualLeadingRows(oldRows: List[Row], newRows: List[Row]): Int = {
    oldRows.zip(newRows).takeWhile { case (oldRow, newRow) =>
      rowsAreEqual(oldRow, newRow)
    }.size
  }

  private def rowsAreEqual(row1: Row, row2: Row): Boolean = {
    val maxCells = math.max(
      Option(row1).map(_.getLastCellNum.toInt).getOrElse(0),
      Option(row2).map(_.getLastCellNum.toInt).getOrElse(0)
    )

    if (maxCells <= 0) return true

    (0 until maxCells).forall { colIdx =>
      val cell1 = Option(row1.getCell(colIdx))
      val cell2 = Option(row2.getCell(colIdx))

      val val1 = cell1.map(getCellValueAsString).getOrElse("")
      val val2 = cell2.map(getCellValueAsString).getOrElse("")

      val1 == val2
    }
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

  private def getCellValueAsString(cell: Cell): String = {
    cell.getCellType match {
      case CellType.NUMERIC =>
        if (DateUtil.isCellDateFormatted(cell)) {
          Try {
            val date = cell.getLocalDateTimeCellValue.toLocalDate
            date.format(DateTimeFormatter.ISO_LOCAL_DATE)
          }.getOrElse {
            val date = cell.getDateCellValue
            date.toString
          }
        } else {
          val num = cell.getNumericCellValue
          if (num == num.toLong) num.toLong.toString
          else num.toString
        }
      case CellType.STRING => cell.getStringCellValue
      case CellType.BOOLEAN => cell.getBooleanCellValue.toString
      case CellType.FORMULA =>
        Try {
          val evaluator = cell.getSheet.getWorkbook.getCreationHelper.createFormulaEvaluator()
          val evaluated = evaluator.evaluate(cell)
          evaluated.getCellType match {
            case CellType.NUMERIC => evaluated.getNumberValue.toString
            case CellType.STRING => evaluated.getStringValue
            case CellType.BOOLEAN => evaluated.getBooleanValue.toString
            case _ => ""
          }
        }.getOrElse(cell.getCellFormula)
      case CellType.BLANK => ""
      case _ => ""
    }
  }
}

object PairedSheetComparer {
  def apply(sheetNames: String*): PairedSheetComparer =
    new PairedSheetComparer(sheetNames.toSet)

  def apply(sheetNames: Set[String], dateValidator: String => Boolean): PairedSheetComparer =
    new PairedSheetComparer(sheetNames, dateValidator)
}