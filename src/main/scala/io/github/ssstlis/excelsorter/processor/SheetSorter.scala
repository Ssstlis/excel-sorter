package io.github.ssstlis.excelsorter.processor

import java.io.{File, FileOutputStream}

import io.github.ssstlis.excelsorter.config.TrackConfig
import io.github.ssstlis.excelsorter.dsl.{ColumnSortConfig, SheetSortingConfig}
import org.apache.poi.ss.usermodel._

import scala.jdk.CollectionConverters._

class SheetSorter(
  sheetConfigs: Map[String, SheetSortingConfig],
  trackConfig: TrackConfig = TrackConfig.empty
) {

  def sortFile(inputPath: String): String = {
    val inputFile = new File(inputPath)
    val outputPath = buildOutputPath(inputPath)

    val workbook = WorkbookFactory.create(inputFile)
    try {
      sortWorkbook(workbook)
      val fos = new FileOutputStream(outputPath)
      try {
        workbook.write(fos)
      } finally {
        fos.close()
      }
    } finally {
      workbook.close()
    }

    outputPath
  }

  private def buildOutputPath(inputPath: String): String = {
    val lastDot = inputPath.lastIndexOf('.')
    if (lastDot > 0) {
      val base = inputPath.substring(0, lastDot)
      val ext = inputPath.substring(lastDot)
      s"${base}_sorted$ext"
    } else {
      s"${inputPath}_sorted"
    }
  }

  private def sortWorkbook(workbook: Workbook): Unit = {
    for (i <- 0 until workbook.getNumberOfSheets) {
      val sheet = workbook.getSheetAt(i)
      val sheetName = sheet.getSheetName
      sheetConfigs.get(sheetName) match {
        case Some(config) => sortSheet(sheet, config, i)
        case _ => System.err.println(s"No config for $sheetName")
      }
    }
  }

  private def sortSheet(sheet: Sheet, config: SheetSortingConfig, sheetIndex: Int): Unit = {
    val rows = sheet.iterator().asScala.toList
    if (rows.isEmpty) return

    val isDataRow = trackConfig.dataRowDetector(sheet.getSheetName, sheetIndex, CellUtils.getRowCellValue)

    val (headerRows, dataRows) = rows.span(row => !isDataRow(row))

    if (dataRows.isEmpty) return

    val sortedDataRows = dataRows.sortWith { (rowA, rowB) =>
      compareRows(rowA, rowB, config.sortConfigs) < 0
    }

    val dataStartIndex = headerRows.size
    val allRowData = sortedDataRows.map(extractRowData)

    sortedDataRows.indices.foreach { i =>
      val targetRowIndex = dataStartIndex + i
      val existingRow = sheet.getRow(targetRowIndex)
      if (existingRow != null) {
        sheet.removeRow(existingRow)
      }
    }

    allRowData.zipWithIndex.foreach { case (rowData, i) =>
      val targetRowIndex = dataStartIndex + i
      val newRow = sheet.createRow(targetRowIndex)
      rowData.zipWithIndex.foreach { case ((value, cellType, style), colIdx) =>
        val cell = newRow.createCell(colIdx)
        setCellValue(cell, value, cellType)
        if (style != null) {
          cell.setCellStyle(style)
        }
      }
    }
  }

  private def extractRowData(row: Row): List[(Any, CellType, CellStyle)] = {
    val lastCellNum = row.getLastCellNum
    if (lastCellNum < 0) return Nil

    (0 until lastCellNum).map { colIdx =>
      val cell = row.getCell(colIdx)
      if (cell == null) {
        (null, CellType.BLANK, null)
      } else {
        val value = cell.getCellType match {
          case CellType.NUMERIC =>
            if (DateUtil.isCellDateFormatted(cell)) cell.getDateCellValue
            else cell.getNumericCellValue
          case CellType.STRING => cell.getStringCellValue
          case CellType.BOOLEAN => cell.getBooleanCellValue
          case CellType.FORMULA => cell.getCellFormula
          case CellType.BLANK => null
          case CellType.ERROR => cell.getErrorCellValue
          case _ => null
        }
        (value, cell.getCellType, cell.getCellStyle)
      }
    }.toList
  }

  private def setCellValue(cell: Cell, value: Any, originalType: CellType): Unit = {
    if (value == null) {
      cell.setBlank()
    } else {
      (value, originalType) match {
        case (d: java.util.Date, _) => cell.setCellValue(d)
        case (e: java.lang.Byte, CellType.ERROR) => cell.setCellErrorValue(e)
        case (n: Number, _) => cell.setCellValue(n.doubleValue())
        case (b: java.lang.Boolean, _) => cell.setCellValue(b)
        case (s: String, CellType.FORMULA) => cell.setCellFormula(s)
        case (s: String, _) => cell.setCellValue(s)
        case (other, _) => cell.setCellValue(other.toString)
      }
    }
  }

  private def compareRows(rowA: Row, rowB: Row, configs: List[ColumnSortConfig[_]]): Int = {
    configs.foldLeft(0) { (acc, config) =>
      if (acc != 0) acc
      else {
        val valueA = CellUtils.getRowCellValue(rowA, config.columnIndex)
        val valueB = CellUtils.getRowCellValue(rowB, config.columnIndex)
        config.compare(valueA, valueB)
      }
    }
  }
}

object SheetSorter {
  def apply(configs: SheetSortingConfig*): SheetSorter = {
    val configMap = configs.map(c => c.sheetName -> c).toMap
    new SheetSorter(configMap)
  }

  def apply(configs: Seq[SheetSortingConfig], trackConfig: TrackConfig): SheetSorter = {
    val configMap = configs.map(c => c.sheetName -> c).toMap
    new SheetSorter(configMap, trackConfig)
  }
}
