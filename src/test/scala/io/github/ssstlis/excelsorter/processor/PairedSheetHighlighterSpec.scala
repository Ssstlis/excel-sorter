package io.github.ssstlis.excelsorter.processor

import java.io.{File, FileOutputStream}
import java.nio.file.Files

import io.github.ssstlis.excelsorter.config._
import io.github.ssstlis.excelsorter.dsl._
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, IndexedColors}
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFWorkbook}
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

import scala.jdk.CollectionConverters._

class PairedSheetHighlighterSpec extends AnyFreeSpec with Matchers {

  private sealed trait DetectedColor
  private case object Green extends DetectedColor
  private case object PaleRed extends DetectedColor
  private case object PaleOrange extends DetectedColor
  private case object NoHighlight extends DetectedColor

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

  private def detectRowColor(path: String, sheetName: String, rowIndex: Int): DetectedColor = {
    val wb = CellUtils.loadWorkbook(path)
    try {
      val sheet = wb.getSheet(sheetName)
      val row = sheet.getRow(rowIndex)
      if (row == null || row.getLastCellNum < 0) return NoHighlight
      val cell = row.getCell(0)
      if (cell == null) return NoHighlight
      val style = cell.getCellStyle
      if (style.getFillPattern != FillPatternType.SOLID_FOREGROUND) return NoHighlight

      style match {
        case xssf: XSSFCellStyle =>
          val color = xssf.getFillForegroundXSSFColor
          if (color == null) return NoHighlight
          val rgb = color.getRGB
          if (rgb == null) return NoHighlight
          val r = rgb(0) & 0xFF
          val g = rgb(1) & 0xFF
          val b = rgb(2) & 0xFF
          if (r == 0xE1 && g == 0xFA && b == 0xE1) Green         // LIGHT_GREEN indexed color maps to this RGB
          else if (r == 0xFF && g == 0xCC && b == 0xCC) PaleRed
          else if (r == 0xF5 && g == 0xE7 && b == 0x9A) PaleOrange
          else {
            // LIGHT_GREEN can also appear as indexed color
            val indexed = xssf.getFillForegroundColor
            if (indexed == org.apache.poi.ss.usermodel.IndexedColors.LIGHT_GREEN.getIndex) Green
            else NoHighlight
          }
        case _ => NoHighlight
      }
    } finally {
      wb.close()
    }
  }

  private def assertThinBlackBorders(path: String, sheetName: String, rowIndex: Int, colCount: Int): Unit = {
    val wb = CellUtils.loadWorkbook(path)
    try {
      val sheet = wb.getSheet(sheetName)
      val row = sheet.getRow(rowIndex)
      row should not be null
      (0 until colCount).foreach { colIdx =>
        val cell = row.getCell(colIdx)
        cell should not be null
        val style = cell.getCellStyle
        withClue(s"row=$rowIndex col=$colIdx borderTop: ") {
          style.getBorderTop shouldBe BorderStyle.THIN
        }
        withClue(s"row=$rowIndex col=$colIdx borderBottom: ") {
          style.getBorderBottom shouldBe BorderStyle.THIN
        }
        withClue(s"row=$rowIndex col=$colIdx borderLeft: ") {
          style.getBorderLeft shouldBe BorderStyle.THIN
        }
        withClue(s"row=$rowIndex col=$colIdx borderRight: ") {
          style.getBorderRight shouldBe BorderStyle.THIN
        }
        withClue(s"row=$rowIndex col=$colIdx topBorderColor: ") {
          style.getTopBorderColor shouldBe IndexedColors.BLACK.getIndex
        }
        withClue(s"row=$rowIndex col=$colIdx bottomBorderColor: ") {
          style.getBottomBorderColor shouldBe IndexedColors.BLACK.getIndex
        }
        withClue(s"row=$rowIndex col=$colIdx leftBorderColor: ") {
          style.getLeftBorderColor shouldBe IndexedColors.BLACK.getIndex
        }
        withClue(s"row=$rowIndex col=$colIdx rightBorderColor: ") {
          style.getRightBorderColor shouldBe IndexedColors.BLACK.getIndex
        }
      }
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
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightPairedSheets(oldPath, newPath)

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
      highlighter.highlightPairedSheets(oldPath, newPath)

      readSheetRows(oldPath, "Sheet1") shouldBe oldOriginalRows
      readSheetRows(newPath, "Sheet1") shouldBe newOriginalRows
    }

    "should highlight same key + same data rows green" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 2

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 2) shouldBe Green
    }

    "should highlight same key + different data rows pale red" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "old-val"), List("2024-01-02", "old-val2")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "new-val"), List("2024-01-02", "new-val2")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedDifferentDataCount shouldBe 2

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe PaleRed
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe PaleRed
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe PaleRed
      detectRowColor(newCmpPath, "Sheet1", 2) shouldBe PaleRed
    }

    "should highlight key only in old file as pale orange" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "val1"), List("2024-01-02", "val2")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "val1")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 1
      results.head.oldOnlyCount shouldBe 1

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe PaleOrange
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe Green
    }

    "should highlight key only in new file as pale orange" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "val1")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "val1"), List("2024-01-03", "val3")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 1
      results.head.newOnlyCount shouldBe 1

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 2) shouldBe PaleOrange
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
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      readSheetRows(oldCmpPath, "Sheet1") should have size 3
      readSheetRows(newCmpPath, "Sheet1") should have size 3
    }

    "should return correct HighlightResult counts" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(
          List("2024-01-01", "same"),      // same key + same data -> green
          List("2024-01-02", "old-val"),    // same key + diff data -> pale red
          List("2024-01-03", "old-only")    // old only -> pale orange
        ))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(
          List("2024-01-01", "same"),      // same key + same data -> green
          List("2024-01-02", "new-val"),    // same key + diff data -> pale red
          List("2024-01-04", "new-only")    // new only -> pale orange
        ))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (_, _, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      val r = results.head
      r.matchedSameDataCount shouldBe 1
      r.matchedDifferentDataCount shouldBe 1
      r.oldOnlyCount shouldBe 1
      r.newOnlyCount shouldBe 1
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
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 1
      results.head.oldOnlyCount shouldBe 1
      results.head.newOnlyCount shouldBe 1

      // Header rows should not be highlighted
      detectRowColor(oldCmpPath, "Sheet1", 0) shouldBe NoHighlight
      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe NoHighlight
      // Data rows should be highlighted
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe Green
      detectRowColor(oldCmpPath, "Sheet1", 3) shouldBe PaleOrange
    }

    "should ignore specified columns when comparing (compareConfig)" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Amount", "Note"),
        List(
          List("2024-01-01", "100", "old-note"),
          List("2024-01-02", "200", "old-diff")
        ))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Amount", "Note"),
        List(
          List("2024-01-01", "100", "new-note"),
          List("2024-01-02", "200", "new-diff")
        ))

      val compare = CompareConfig(List(
        ComparePolicy(SheetSelector.Default, Set(2))
      ))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"), compareConfig = compare)
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 2
      results.head.matchedDifferentDataCount shouldBe 0

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe Green
      detectRowColor(newCmpPath, "Sheet1", 2) shouldBe Green
    }

    "should use sortConfigsMap columns as keys" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      // Column 1 is the key (not default column 0)
      createTestWorkbook(oldPath, "Sheet1", List("Date", "ID", "Value"),
        List(
          List("2024-01-01", "A", "100"),
          List("2024-01-02", "B", "200")
        ))
      createTestWorkbook(newPath, "Sheet1", List("Date", "ID", "Value"),
        List(
          List("2024-01-03", "A", "100"),
          List("2024-01-04", "C", "300")
        ))

      val sortConfigs = Map(
        "Sheet1" -> SheetSortingConfig("Sheet1", List(
          SortingDsl.asc(1)(SortingDsl.Parsers.asString)
        ))
      )

      val highlighter = PairedSheetHighlighter(Set("Sheet1"), sortConfigsMap = sortConfigs)
      val (_, _, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      val r = results.head
      // Key "A" in both: dates differ -> pale red (unless we also ignore date column)
      // Key "B" only in old -> pale orange
      // Key "C" only in new -> pale orange
      r.matchedDifferentDataCount shouldBe 1
      r.oldOnlyCount shouldBe 1
      r.newOnlyCount shouldBe 1
    }

    "should highlight all rows green when all data is the same" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "a"), List("2024-01-02", "b"), List("2024-01-03", "c")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "a"), List("2024-01-02", "b"), List("2024-01-03", "c")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 3
      results.head.matchedDifferentDataCount shouldBe 0
      results.head.oldOnlyCount shouldBe 0
      results.head.newOnlyCount shouldBe 0

      (1 to 3).foreach { i =>
        detectRowColor(oldCmpPath, "Sheet1", i) shouldBe Green
        detectRowColor(newCmpPath, "Sheet1", i) shouldBe Green
      }
    }

    "should apply thin black borders to all highlighted cells (green)" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, _, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      assertThinBlackBorders(oldCmpPath, "Sheet1", rowIndex = 1, colCount = 2)
    }

    "should apply thin black borders to all highlighted cells (pale red)" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "old")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "new")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      assertThinBlackBorders(oldCmpPath, "Sheet1", rowIndex = 1, colCount = 2)
      assertThinBlackBorders(newCmpPath, "Sheet1", rowIndex = 1, colCount = 2)
    }

    "should apply thin black borders to all highlighted cells (pale orange)" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "only-old")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-02", "only-new")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      assertThinBlackBorders(oldCmpPath, "Sheet1", rowIndex = 1, colCount = 2)
      assertThinBlackBorders(newCmpPath, "Sheet1", rowIndex = 1, colCount = 2)
    }

    "should not apply borders to non-highlighted header rows" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "same")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, _, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      val wb = CellUtils.loadWorkbook(oldCmpPath)
      try {
        val sheet = wb.getSheet("Sheet1")
        val headerRow = sheet.getRow(0)
        (0 until headerRow.getLastCellNum).foreach { colIdx =>
          val style = headerRow.getCell(colIdx).getCellStyle
          style.getBorderTop shouldBe BorderStyle.NONE
          style.getBorderBottom shouldBe BorderStyle.NONE
          style.getBorderLeft shouldBe BorderStyle.NONE
          style.getBorderRight shouldBe BorderStyle.NONE
        }
      } finally {
        wb.close()
      }
    }

    "should apply thin black borders to every cell in a multi-column highlighted row" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("A", "B", "C", "D", "E"),
        List(List("2024-01-01", "b", "c", "d", "e")))
      createTestWorkbook(newPath, "Sheet1", List("A", "B", "C", "D", "E"),
        List(List("2024-01-01", "b", "c", "d", "e")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, _, _) = highlighter.highlightPairedSheets(oldPath, newPath)

      assertThinBlackBorders(oldCmpPath, "Sheet1", rowIndex = 1, colCount = 5)
    }

    "should highlight all rows pale orange when no matching keys" in {
      val tmpDir = Files.createTempDirectory("highlight-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-01", "a"), List("2024-01-02", "b")))
      createTestWorkbook(newPath, "Sheet1", List("Date", "Value"),
        List(List("2024-01-03", "c"), List("2024-01-04", "d")))

      val highlighter = PairedSheetHighlighter(Set("Sheet1"))
      val (oldCmpPath, newCmpPath, results) = highlighter.highlightPairedSheets(oldPath, newPath)

      results should have size 1
      results.head.matchedSameDataCount shouldBe 0
      results.head.matchedDifferentDataCount shouldBe 0
      results.head.oldOnlyCount shouldBe 2
      results.head.newOnlyCount shouldBe 2

      detectRowColor(oldCmpPath, "Sheet1", 1) shouldBe PaleOrange
      detectRowColor(oldCmpPath, "Sheet1", 2) shouldBe PaleOrange
      detectRowColor(newCmpPath, "Sheet1", 1) shouldBe PaleOrange
      detectRowColor(newCmpPath, "Sheet1", 2) shouldBe PaleOrange
    }
  }
}
