package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.model.{CellDiff, Mapping}
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

    "findCellDiffsMapped" - {

      "Mapping.from" - {

        "should return empty for identical rows" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet = wb.createSheet("Test")
            val row1  = sheet.createRow(0)
            row1.createCell(0).setCellValue("a")
            row1.createCell(1).setCellValue("b")

            val row2 = sheet.createRow(1)
            row2.createCell(0).setCellValue("a")
            row2.createCell(1).setCellValue("b")

            CellUtils.findCellDiffsMapped(row1, row2, Mapping.Identity) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return diffs for different rows" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet = wb.createSheet("Test")
            val row1  = sheet.createRow(0)
            row1.createCell(0).setCellValue("a")

            val row2 = sheet.createRow(1)
            row2.createCell(0).setCellValue("b")

            CellUtils.findCellDiffsMapped(row1, row2, Mapping.Identity) should not be empty
          } finally {
            wb.close()
          }
        }

        "should compare only the declared pairs, ignoring other columns" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet = wb.createSheet("Test")
            // old row: col0=A, col1=B, col2=C
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")
            oldRow.createCell(1).setCellValue("B")
            oldRow.createCell(2).setCellValue("C")

            // new row: col0=X, col1=A, col2=B, col3=C  (shifted by 1)
            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("X")
            newRow.createCell(1).setCellValue("A")
            newRow.createCell(2).setCellValue("B")
            newRow.createCell(3).setCellValue("C")

            // Map old col0->new col1, old col1->new col2, old col2->new col3
            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.from((0, 1), (1, 2), (2, 3))) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return empty for two empty rows" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet = wb.createSheet("Test")
            val row1  = sheet.createRow(0)
            val row2  = sheet.createRow(1)

            CellUtils.findCellDiffsMapped(row1, row2, Mapping.Identity) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return empty for Mapping.from(Nil) regardless of row content" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")
            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("B")

            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.from()) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return empty when differing column is ignored" in {
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

            val mapping = Mapping.Identity
            CellUtils.findCellDiffsMapped(row1, row2, mapping) should not be empty
            CellUtils.findCellDiffsMapped(row1, row2, mapping, ignoredColumns = Set(1)) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return diffs when non-ignored column differs" in {
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

            CellUtils.findCellDiffsMapped(row1, row2, Mapping.Identity, ignoredColumns = Set(1)) should not be empty
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

            CellUtils.findCellDiffsMapped(row1, row2, Mapping.Identity, ignoredColumns = Set(1, 2)) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return correct CellDiff values" in {
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

            val headers = Map(0 -> "Date", 1 -> "Amount", 2 -> "Note")

            val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.Identity, headers)
            diffs should have size 2
            diffs(0) shouldBe CellDiff("Amount", 1, 1, "old-val", "new-val")
            diffs(1) shouldBe CellDiff("Note", 2, 2, "old-val2", "new-val2")
          } finally {
            wb.close()
          }
        }

        "should skip ignored columns in diff result" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("same")
            oldRow.createCell(1).setCellValue("old-val")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("same")
            newRow.createCell(1).setCellValue("new-val")

            val headers = Map(0 -> "A", 1 -> "B")

            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.Identity, headers, Set(1)) shouldBe empty
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
            val mapping = Mapping.from((0, 1), (1, 2))
            val headers = Map(0 -> "First", 1 -> "Second")

            val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, mapping, headers)
            diffs should have size 1
            diffs.head shouldBe CellDiff("Second", 1, 2, "B", "changed-B")
          } finally {
            wb.close()
          }
        }
      }

      "Mapping.identity" - {

        "should detect diffs in all columns when rows differ" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")
            oldRow.createCell(1).setCellValue("B")
            oldRow.createCell(2).setCellValue("same")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("X")
            newRow.createCell(1).setCellValue("Y")
            newRow.createCell(2).setCellValue("same")

            val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity)
            diffs should have size 2
            diffs(0).oldColumnIndex shouldBe 0
            diffs(1).oldColumnIndex shouldBe 1
          } finally {
            wb.close()
          }
        }

        "should return empty when rows are identical" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")
            oldRow.createCell(1).setCellValue("B")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("A")
            newRow.createCell(1).setCellValue("B")

            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should return empty for two empty rows" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            val newRow = sheet.createRow(1)

            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity) shouldBe empty
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
            oldRow.createCell(2).setCellValue("same2")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("same")
            newRow.createCell(1).setCellValue("new-val")
            newRow.createCell(2).setCellValue("same2")

            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity) should not be empty
            CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity, ignoredColumns = Set(1)) shouldBe empty
          } finally {
            wb.close()
          }
        }

        "should detect diff in extra column when new row is wider" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("A")
            newRow.createCell(1).setCellValue("extra") // absent in old → "" vs "extra"

            val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow, Mapping.identity)
            diffs should have size 1
            diffs.head.oldColumnIndex shouldBe 1
            diffs.head.oldValue shouldBe ""
            diffs.head.newValue shouldBe "extra"
          } finally {
            wb.close()
          }
        }

        "should be the default parameter value" in {
          val wb = new XSSFWorkbook()
          try {
            val sheet  = wb.createSheet("Test")
            val oldRow = sheet.createRow(0)
            oldRow.createCell(0).setCellValue("A")
            oldRow.createCell(1).setCellValue("B")

            val newRow = sheet.createRow(1)
            newRow.createCell(0).setCellValue("A")
            newRow.createCell(1).setCellValue("X")

            // no mapping argument — uses Mapping.identity by default
            val diffs = CellUtils.findCellDiffsMapped(oldRow, newRow)
            diffs should have size 1
            diffs.head.oldColumnIndex shouldBe 1
          } finally {
            wb.close()
          }
        }
      }
    }
  }
}
