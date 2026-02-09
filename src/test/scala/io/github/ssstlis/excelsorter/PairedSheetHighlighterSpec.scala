package io.github.ssstlis.excelsorter

import java.io.{File, FileOutputStream}
import java.nio.file.Files

import org.apache.poi.ss.usermodel.{FillPatternType, IndexedColors}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

import scala.jdk.CollectionConverters._

class PairedSheetHighlighterSpec extends AnyFreeSpec with Matchers {

  private def createTestWorkbook(path: String, sheetName: String, headers: List[String], dataRows: List[List[String]]): Unit = {
    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)

    val headerRow = sheet.createRow(0)
    headers.zipWithIndex.foreach { case (h, i) => headerRow.createCell(i).setCellValue(h) }

    dataRows.zipWithIndex.foreach { case (values, rowIdx) =>
      val row = sheet.createRow(rowIdx + 1)
      values.zipWithIndex.foreach { case (v, colIdx) => row.createCell(colIdx).setCellValue(v) }
    }

    val fos = new FileOutputStream(path)
    try { wb.write(fos) } finally { fos.close(); wb.close() }
  }

  private def readSheetRows(path: String, sheetName: String): List[List[String]] = {
    val wb = CellUtils.loadWorkbook(path)
    try {
      val sheet = wb.getSheet(sheetName)
      if (sheet == null) return Nil
      sheet.iterator().asScala.toList.map { row =>
        val lastCell = row.getLastCellNum
        if (lastCell < 0) Nil
        else (0 until lastCell).map(i => CellUtils.getRowCellValue(row, i)).toList
      }
    } finally {
      wb.close()
    }
  }

  private def isRowHighlighted(path: String, sheetName: String, rowIndex: Int): Boolean = {
    val wb = CellUtils.loadWorkbook(path)
    try {
      val sheet = wb.getSheet(sheetName)
      val row = sheet.getRow(rowIndex)
      if (row == null || row.getLastCellNum < 0) return false
      val cell = row.getCell(0)
      if (cell == null) return false
      val style = cell.getCellStyle
      style.getFillPattern == FillPatternType.SOLID_FOREGROUND &&
        style.getFillForegroundColor == IndexedColors.LIGHT_GREEN.getIndex
    } finally {
      wb.close()
    }
  }

  "PairedSheetHighlighter" - {

    "should create _compared files" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-02")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-03")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightEqualLeadingRows(oldPath, newPath)

      oldCmpPath should endWith("_compared.xlsx")
      newCmpPath should endWith("_compared.xlsx")
      new File(oldCmpPath).exists() shouldBe true
      new File(newCmpPath).exists() shouldBe true
    }

    "should not modify _sorted files" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-02")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-03")))

      val oldOriginalRows = readSheetRows(oldPath, "Sheet1")
      val newOriginalRows = readSheetRows(newPath, "Sheet1")

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      highlighter.highlightEqualLeadingRows(oldPath, newPath)

      readSheetRows(oldPath, "Sheet1") shouldBe oldOriginalRows
      readSheetRows(newPath, "Sheet1") shouldBe newOriginalRows
    }

    "should highlight equal leading rows with green" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-03", "old-only")))
      createTestWorkbook(newPath, "Sheet1", List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-04", "new-only")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightEqualLeadingRows(oldPath, newPath)

      results should have size 1
      results.head.highlightedRowCount shouldBe 2

      // Data rows 1 and 2 (indices 1, 2) should be highlighted
      isRowHighlighted(oldCmpPath, "Sheet1", 1) shouldBe true
      isRowHighlighted(oldCmpPath, "Sheet1", 2) shouldBe true
      // Data row 3 (index 3) should NOT be highlighted
      isRowHighlighted(oldCmpPath, "Sheet1", 3) shouldBe false

      isRowHighlighted(newCmpPath, "Sheet1", 1) shouldBe true
      isRowHighlighted(newCmpPath, "Sheet1", 2) shouldBe true
      isRowHighlighted(newCmpPath, "Sheet1", 3) shouldBe false
    }

    "should preserve all rows (no removal)" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "diff-old")))
      createTestWorkbook(newPath, "Sheet1", List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "diff-new")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightEqualLeadingRows(oldPath, newPath)

      readSheetRows(oldCmpPath, "Sheet1") should have size 3 // header + 2 data
      readSheetRows(newCmpPath, "Sheet1") should have size 3
    }

    "should not highlight when no equal leading rows" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01", "old")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-02", "new")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, _, results) = highlighter.highlightEqualLeadingRows(oldPath, newPath)

      results shouldBe empty
      isRowHighlighted(oldCmpPath, "Sheet1", 1) shouldBe false
    }

    "should use trackConfig for data row detection" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Title"), List(
        List("SubHeader"),
        List("DATA-1", "same"),
        List("DATA-2", "old-only")
      ))
      createTestWorkbook(newPath, "Sheet1", List("Title"), List(
        List("SubHeader"),
        List("DATA-1", "same"),
        List("DATA-3", "new-only")
      ))

      val track = TrackConfig(List(
        TrackPolicy(SheetSelector.Default, List(
          TrackCondition(0, _.startsWith("DATA"))
        ))
      ))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"), track)
      val (oldCmpPath, _, results) = highlighter.highlightEqualLeadingRows(oldPath, newPath)

      results should have size 1
      results.head.highlightedRowCount shouldBe 1

      // Row index 2 is the first data row (DATA-1) - should be highlighted
      isRowHighlighted(oldCmpPath, "Sheet1", 2) shouldBe true
      // Row index 3 is the second data row (DATA-2) - should NOT be highlighted
      isRowHighlighted(oldCmpPath, "Sheet1", 3) shouldBe false
    }
  }
}
