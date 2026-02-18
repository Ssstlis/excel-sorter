package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.model.CellDiff
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
          val row   = sheet.createRow(0)
          val cell  = row.createCell(0)
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
          val row   = sheet.createRow(0)
          val cell  = row.createCell(0)
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
          val row   = sheet.createRow(0)
          val cell  = row.createCell(0)
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
          val row   = sheet.createRow(0)
          val cell  = row.createCell(0)
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
          val row   = sheet.createRow(0)
          val cell  = row.createCell(0)
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
          val row   = sheet.createRow(0)
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
          val row   = sheet.createRow(0)
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
          val row1  = sheet.createRow(0)
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
          val row1  = sheet.createRow(0)
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
          val row1  = sheet.createRow(0)
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
          val row1  = sheet.createRow(0)
          val row2  = sheet.createRow(1)

          CellUtils.rowsAreEqual(row1, row2) shouldBe true
        } finally {
          wb.close()
        }
      }

      "should return true when differing columns are ignored" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1  = sheet.createRow(0)
          row1.createCell(0).setCellValue("same")
          row1.createCell(1).setCellValue("diff-old")
          row1.createCell(2).setCellValue("same")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("same")
          row2.createCell(1).setCellValue("diff-new")
          row2.createCell(2).setCellValue("same")

          CellUtils.rowsAreEqual(row1, row2) shouldBe false
          CellUtils.rowsAreEqual(row1, row2, Set(1)) shouldBe true
        } finally {
          wb.close()
        }
      }

      "should return false when non-ignored columns differ" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1  = sheet.createRow(0)
          row1.createCell(0).setCellValue("a")
          row1.createCell(1).setCellValue("same")
          row1.createCell(2).setCellValue("c")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("x")
          row2.createCell(1).setCellValue("same")
          row2.createCell(2).setCellValue("c")

          CellUtils.rowsAreEqual(row1, row2, Set(1)) shouldBe false
        } finally {
          wb.close()
        }
      }

      "should ignore multiple columns" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          val row1  = sheet.createRow(0)
          row1.createCell(0).setCellValue("same")
          row1.createCell(1).setCellValue("diff1")
          row1.createCell(2).setCellValue("diff2")

          val row2 = sheet.createRow(1)
          row2.createCell(0).setCellValue("same")
          row2.createCell(1).setCellValue("other1")
          row2.createCell(2).setCellValue("other2")

          CellUtils.rowsAreEqual(row1, row2, Set(1, 2)) shouldBe true
        } finally {
          wb.close()
        }
      }
    }

    "rowsAreEqualMapped" - {

      "should compare only mapped column pairs" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet = wb.createSheet("Test")
          // old row: col0=A, col1=B, col2=C
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("A")
          oldRow.createCell(1).setCellValue("B")
          oldRow.createCell(2).setCellValue("C")

          // new row: col0=X, col1=A, col2=B, col3=C
          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("X")
          newRow.createCell(1).setCellValue("A")
          newRow.createCell(2).setCellValue("B")
          newRow.createCell(3).setCellValue("C")

          // Map old col0->new col1, old col1->new col2, old col2->new col3
          val mapping = List((0, 1), (1, 2), (2, 3))
          CellUtils.rowsAreEqualMapped(oldRow, newRow, mapping) shouldBe true
        } finally {
          wb.close()
        }
      }

      "should return false when mapped columns differ" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("A")
          oldRow.createCell(1).setCellValue("B")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("A")
          newRow.createCell(1).setCellValue("Z")

          val mapping = List((0, 0), (1, 1))
          CellUtils.rowsAreEqualMapped(oldRow, newRow, mapping) shouldBe false
        } finally {
          wb.close()
        }
      }

      "should respect ignoredColumns" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("same")
          oldRow.createCell(1).setCellValue("old-val")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("same")
          newRow.createCell(1).setCellValue("new-val")

          val mapping = List((0, 0), (1, 1))
          CellUtils.rowsAreEqualMapped(oldRow, newRow, mapping) shouldBe false
          CellUtils.rowsAreEqualMapped(oldRow, newRow, mapping, Set(1)) shouldBe true
        } finally {
          wb.close()
        }
      }

      "should return true for empty mapping" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("A")
          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("B")

          CellUtils.rowsAreEqualMapped(oldRow, newRow, Nil) shouldBe true
        } finally {
          wb.close()
        }
      }
    }

    "findCellDiffsMapped" - {

      "should return correct diffs for differing cells" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("same")
          oldRow.createCell(1).setCellValue("old-val")
          oldRow.createCell(2).setCellValue("old-val2")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("same")
          newRow.createCell(1).setCellValue("new-val")
          newRow.createCell(2).setCellValue("new-val2")

          val mapping = List((0, 0), (1, 1), (2, 2))
          val headers = Map(0 -> "Date", 1 -> "Amount", 2 -> "Note")

          val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, mapping, headers)
          diffs should have size 2
          diffs(0) shouldBe CellDiff("Amount", 1, 1, "old-val", "new-val")
          diffs(1) shouldBe CellDiff("Note", 2, 2, "old-val2", "new-val2")
        } finally {
          wb.close()
        }
      }

      "should return empty list for equal rows" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("same")
          oldRow.createCell(1).setCellValue("same2")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("same")
          newRow.createCell(1).setCellValue("same2")

          val mapping = List((0, 0), (1, 1))
          val headers = Map(0 -> "A", 1 -> "B")

          CellUtils.findCellDiffsMapped(oldRow, newRow, mapping, headers) shouldBe empty
        } finally {
          wb.close()
        }
      }

      "should skip ignored columns" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("same")
          oldRow.createCell(1).setCellValue("old-val")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("same")
          newRow.createCell(1).setCellValue("new-val")

          val mapping = List((0, 0), (1, 1))
          val headers = Map(0 -> "A", 1 -> "B")

          CellUtils.findCellDiffsMapped(oldRow, newRow, mapping, headers, Set(1)) shouldBe empty
        } finally {
          wb.close()
        }
      }

      "should handle remapped columns" in {
        val wb = new XSSFWorkbook()
        try {
          val sheet  = wb.createSheet("Test")
          val oldRow = sheet.createRow(0)
          oldRow.createCell(0).setCellValue("A")
          oldRow.createCell(1).setCellValue("B")

          val newRow = sheet.createRow(1)
          newRow.createCell(0).setCellValue("X")
          newRow.createCell(1).setCellValue("A")
          newRow.createCell(2).setCellValue("changed-B")

          // old col0 -> new col1, old col1 -> new col2
          val mapping = List((0, 1), (1, 2))
          val headers = Map(0 -> "First", 1 -> "Second")

          val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, mapping, headers)
          diffs should have size 1
          diffs.head shouldBe CellDiff("Second", 1, 2, "B", "changed-B")
        } finally {
          wb.close()
        }
      }
    }
  }
}
