package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.model.CellDiff

import java.io.{FileInputStream, FileOutputStream}
import java.nio.file.{Files, Paths, StandardCopyOption}
import java.time.format.DateTimeFormatter

import org.apache.poi.ss.usermodel._

import scala.util.Try

object CellUtils {

  def getCellValueAsString(cell: Cell): String = {
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
      case CellType.STRING  => cell.getStringCellValue
      case CellType.BOOLEAN => cell.getBooleanCellValue.toString
      case CellType.FORMULA =>
        Try {
          val evaluator = cell.getSheet.getWorkbook.getCreationHelper.createFormulaEvaluator()
          val evaluated = evaluator.evaluate(cell)
          evaluated.getCellType match {
            case CellType.NUMERIC => evaluated.getNumberValue.toString
            case CellType.STRING  => evaluated.getStringValue
            case CellType.BOOLEAN => evaluated.getBooleanValue.toString
            case _                => ""
          }
        }.getOrElse(cell.getCellFormula)
      case CellType.BLANK => ""
      case _              => ""
    }
  }

  def getRowCellValue(row: Row, colIndex: Int): String = {
    Option(row.getCell(colIndex)).map(getCellValueAsString).getOrElse("")
  }

  def loadWorkbook(path: String): Workbook = {
    val fis = new FileInputStream(path)
    try {
      WorkbookFactory.create(fis)
    } finally {
      fis.close()
    }
  }

  def writeWorkbook(workbook: Workbook, path: String): Unit = {
    val fos = new FileOutputStream(path)
    try {
      workbook.write(fos)
    } finally {
      fos.close()
    }
  }

  def copyFile(sourcePath: String, destPath: String): Unit = {
    Files.copy(Paths.get(sourcePath), Paths.get(destPath), StandardCopyOption.REPLACE_EXISTING)
  }

  def rowsAreEqualMapped(
    oldRow: Row,
    newRow: Row,
    columnMapping: List[(Int, Int)],
    ignoredOldColumns: Set[Int] = Set.empty
  ): Boolean = {
    columnMapping.forall { case (oldIdx, newIdx) =>
      ignoredOldColumns.contains(oldIdx) || {
        val oldVal = getRowCellValue(oldRow, oldIdx)
        val newVal = getRowCellValue(newRow, newIdx)
        oldVal == newVal
      }
    }
  }

  def findCellDiffsMapped(
    oldRow: Row,
    newRow: Row,
    columnMapping: List[(Int, Int)],
    headerNames: Map[Int, String],
    ignoredOldColumns: Set[Int] = Set.empty
  ): List[CellDiff] = {
    columnMapping.flatMap { case (oldIdx, newIdx) =>
      if (ignoredOldColumns.contains(oldIdx)) {
        None
      } else {
        val oldVal = getRowCellValue(oldRow, oldIdx)
        val newVal = getRowCellValue(newRow, newIdx)
        if (oldVal == newVal) None
        else Some(CellDiff(headerNames.getOrElse(oldIdx, s"Column $oldIdx"), oldIdx, newIdx, oldVal, newVal))
      }
    }
  }

  def rowsAreEqual(row1: Row, row2: Row, ignoredColumns: Set[Int] = Set.empty): Boolean = {
    val maxCells = math.max(
      Option(row1).map(_.getLastCellNum.toInt).getOrElse(0),
      Option(row2).map(_.getLastCellNum.toInt).getOrElse(0)
    )

    (maxCells <= 0) || {
      (0 until maxCells).forall { colIdx =>
        ignoredColumns.contains(colIdx) || {
          val val1 = getRowCellValue(row1, colIdx)
          val val2 = getRowCellValue(row2, colIdx)
          val1 == val2
        }
      }
    }
  }
}
