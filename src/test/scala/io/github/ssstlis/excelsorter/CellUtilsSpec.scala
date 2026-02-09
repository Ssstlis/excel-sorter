package io.github.ssstlis.excelsorter

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CellUtilsSpec extends AnyFreeSpec with Matchers {

  "CellUtils" - {

    "getCellValueAsString" - {

      "should handle string cells" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          val cell = row.createCell(0)
          cell.setCellValue("hello")

          CellUtils.getCellValueAsString(cell) shouldBe "hello"
        } finally {
          wb.close()
        }
      }

      "should handle numeric cells with integer value" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          val cell = row.createCell(0)
          cell.setCellValue(42.0)

          CellUtils.getCellValueAsString(cell) shouldBe "42"
        } finally {
          wb.close()
        }
      }

      "should handle numeric cells with decimal value" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          val cell = row.createCell(0)
          cell.setCellValue(3.14)

          CellUtils.getCellValueAsString(cell) shouldBe "3.14"
        } finally {
          wb.close()
        }
      }

      "should handle boolean cells" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          val cell = row.createCell(0)
          cell.setCellValue(true)

          CellUtils.getCellValueAsString(cell) shouldBe "true"
        } finally {
          wb.close()
        }
      }

      "should handle blank cells" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          val cell = row.createCell(0)
          cell.setBlank()

          CellUtils.getCellValueAsString(cell) shouldBe ""
        } finally {
          wb.close()
        }
      }
    }

    "getRowCellValue" - {

      "should return empty string for null cell" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          row.createCell(0).setCellValue("exists")

          CellUtils.getRowCellValue(row, 5) shouldBe ""
        } finally {
          wb.close()
        }
      }

      "should return cell value for existing cell" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row = sheet.createRow(0)
          row.createCell(0).setCellValue("value")

          CellUtils.getRowCellValue(row, 0) shouldBe "value"
        } finally {
          wb.close()
        }
      }
    }

    "rowsAreEqual" - {

      "should return true for identical rows" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1 = sheet.createRow(0)
          row1.createCell(0).setCellValue("a")
          row1.createCell(1).setCellValue("b")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("a")
          row2.createCell(1).setCellValue("b")

          CellUtils.rowsAreEqual(row1, row2) shouldBe true
        } finally {
          wb.close()
        }
      }

      "should return false for different rows" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1 = sheet.createRow(0)
          row1.createCell(0).setCellValue("a")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("b")

          CellUtils.rowsAreEqual(row1, row2) shouldBe false
        } finally {
          wb.close()
        }
      }

      "should handle rows with different column counts" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1 = sheet.createRow(0)
          row1.createCell(0).setCellValue("a")
          row1.createCell(1).setCellValue("b")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("a")

          CellUtils.rowsAreEqual(row1, row2) shouldBe false
        } finally {
          wb.close()
        }
      }

      "should return true for two empty rows" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1 = sheet.createRow(0)
          val row2 = sheet.createRow(1)

          CellUtils.rowsAreEqual(row1, row2) shouldBe true
        } finally {
          wb.close()
        }
      }
    }
  }
}
