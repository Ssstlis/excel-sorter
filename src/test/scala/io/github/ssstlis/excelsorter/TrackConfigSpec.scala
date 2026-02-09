package io.github.ssstlis.excelsorter

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class TrackConfigSpec extends AnyFreeSpec with Matchers {

  private def createRow(workbook: XSSFWorkbook, sheetName: String, values: String*): Row = {
    val sheet = Option(workbook.getSheet(sheetName)).getOrElse(workbook.createSheet(sheetName))
    val nextRowIdx = if (sheet.getPhysicalNumberOfRows == 0) 0 else sheet.getLastRowNum + 1
    val row = sheet.createRow(nextRowIdx)
    values.zipWithIndex.foreach { case (v, i) =>
      row.createCell(i).setCellValue(v)
    }
    row
  }

  "TrackConfig" - {

    "should fall back to defaultDateValidator when no policies" in {
      val config = TrackConfig.empty
      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("Sheet1", 0, CellUtils.getRowCellValue)

        val dateRow = createRow(wb, "Sheet1", "2024-01-15", "data")
        val headerRow = createRow(wb, "Sheet1", "Header", "stuff")

        detector(dateRow) shouldBe true
        detector(headerRow) shouldBe false
      } finally {
        wb.close()
      }
    }

    "should match by sheet name" in {
      val config = TrackConfig(List(
        TrackPolicy(SheetSelector.ByName("Target"), List(
          TrackCondition(1, s => s == "MATCH")
        ))
      ))

      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("Target", 0, CellUtils.getRowCellValue)

        val matchRow = createRow(wb, "Target", "anything", "MATCH")
        val noMatchRow = createRow(wb, "Target", "anything", "NOPE")

        detector(matchRow) shouldBe true
        detector(noMatchRow) shouldBe false
      } finally {
        wb.close()
      }
    }

    "should match by sheet index" in {
      val config = TrackConfig(List(
        TrackPolicy(SheetSelector.ByIndex(2), List(
          TrackCondition(0, s => s.startsWith("DATA"))
        ))
      ))

      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("AnyName", 2, CellUtils.getRowCellValue)

        val matchRow = createRow(wb, "AnyName", "DATA-row")
        val noMatchRow = createRow(wb, "AnyName", "header-row")

        detector(matchRow) shouldBe true
        detector(noMatchRow) shouldBe false
      } finally {
        wb.close()
      }
    }

    "should use Default policy as fallback for unconfigured sheets" in {
      val config = TrackConfig(List(
        TrackPolicy(SheetSelector.ByName("Specific"), List(
          TrackCondition(0, _ == "X")
        )),
        TrackPolicy(SheetSelector.Default, List(
          TrackCondition(0, _ == "DEFAULT")
        ))
      ))

      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("Unconfigured", 5, CellUtils.getRowCellValue)

        val matchRow = createRow(wb, "Unconfigured", "DEFAULT")
        val noMatchRow = createRow(wb, "Unconfigured", "other")

        detector(matchRow) shouldBe true
        detector(noMatchRow) shouldBe false
      } finally {
        wb.close()
      }
    }

    "should require all conditions to match (AND semantics)" in {
      val config = TrackConfig(List(
        TrackPolicy(SheetSelector.ByName("Multi"), List(
          TrackCondition(0, s => scala.util.Try(s.toInt).isSuccess),
          TrackCondition(1, s => s.nonEmpty)
        ))
      ))

      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("Multi", 0, CellUtils.getRowCellValue)

        val bothMatch = createRow(wb, "Multi", "42", "text")
        val firstOnly = createRow(wb, "Multi", "42", "")
        val secondOnly = createRow(wb, "Multi", "abc", "text")
        val neitherMatch = createRow(wb, "Multi", "abc", "")

        detector(bothMatch) shouldBe true
        detector(firstOnly) shouldBe false
        detector(secondOnly) shouldBe false
        detector(neitherMatch) shouldBe false
      } finally {
        wb.close()
      }
    }

    "should prefer specific name match over Default" in {
      val config = TrackConfig(List(
        TrackPolicy(SheetSelector.Default, List(
          TrackCondition(0, _ == "DEFAULT")
        )),
        TrackPolicy(SheetSelector.ByName("Named"), List(
          TrackCondition(0, _ == "NAMED")
        ))
      ))

      val wb = new XSSFWorkbook()
      try {
        val detector = config.dataRowDetector("Named", 0, CellUtils.getRowCellValue)

        val namedRow = createRow(wb, "Named", "NAMED")
        val defaultRow = createRow(wb, "Named", "DEFAULT")

        detector(namedRow) shouldBe true
        detector(defaultRow) shouldBe false
      } finally {
        wb.close()
      }
    }
  }
}
