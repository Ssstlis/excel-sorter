package io.github.ssstlis.excelsorter.processor

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

  def rowsAreEqual(row1: Row, row2: Row, ignoredColumns: Set[Int] = Set.empty): Boolean = {
    val maxCells = math.max(
      Option(row1).map(_.getLastCellNum.toInt).getOrElse(0),
      Option(row2).map(_.getLastCellNum.toInt).getOrElse(0)
    )

    if (maxCells <= 0) return true

    (0 until maxCells).forall { colIdx =>
      ignoredColumns.contains(colIdx) || {
        val val1 = Option(row1.getCell(colIdx)).map(getCellValueAsString).getOrElse("")
        val val2 = Option(row2.getCell(colIdx)).map(getCellValueAsString).getOrElse("")
        val1 == val2
      }
    }
  }
}
