package io.github.ssstlis.excelsorter.processor

import cats.data.NonEmptyList

import java.io.{File, FileOutputStream}
import java.nio.file.Files
import io.github.ssstlis.excelsorter.config._
import io.github.ssstlis.excelsorter.config.compare._
import io.github.ssstlis.excelsorter.config.track._
import io.github.ssstlis.excelsorter.dsl._
import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.Checkpoints
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

import scala.jdk.CollectionConverters._

class PairedSheetCutterSpec extends AnyFreeSpec with Matchers with Checkpoints {

  private def createTestWorkbook(
    path: String,
    sheetName: String,
    headers: List[String],
    dataRows: List[List[String]]
  ): Unit = {
    val wb    = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)

    val headerRow = sheet.createRow(0)
    headers.zipWithIndex.foreach { case (h, i) => headerRow.createCell(i).setCellValue(h) }

    dataRows.zipWithIndex.foreach { case (values, rowIdx) =>
      val row = sheet.createRow(rowIdx + 1)
      values.zipWithIndex.foreach { case (v, colIdx) => row.createCell(colIdx).setCellValue(v) }
    }

    val fos = new FileOutputStream(path)
    try { wb.write(fos) }
    finally { fos.close(); wb.close() }
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

  "PairedSheetCutter" - {

    "should create _sortcutted files" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-02")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-03")))

      val cutter                      = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, _) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { oldCutPath should endWith("_sortcutted.xlsx") }
      cp { newCutPath should endWith("_sortcutted.xlsx") }
      cp { new File(oldCutPath).exists() shouldBe true }
      cp { new File(newCutPath).exists() shouldBe true }
      cp.reportAll()
    }

    "should not modify _sorted files" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-02")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-01"), List("2024-01-03")))

      val oldOriginalRows = readSheetRows(oldPath, "Sheet1")
      val newOriginalRows = readSheetRows(newPath, "Sheet1")

      val cutter = PairedSheetCutter(Set("Sheet1"))
      cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { readSheetRows(oldPath, "Sheet1") shouldBe oldOriginalRows }
      cp { readSheetRows(newPath, "Sheet1") shouldBe newOriginalRows }
      cp.reportAll()
    }

    "should remove equal leading data rows from _sortcutted files" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-03", "old-only"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-04", "new-only"))
      )

      val cutter                            = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")
      val newCutRows = readSheetRows(newCutPath, "Sheet1")

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { oldCutRows should have size 2 }
      cp { oldCutRows(1) shouldBe List("2024-01-03", "old-only") }
      cp { newCutRows should have size 2 }
      cp { newCutRows(1) shouldBe List("2024-01-04", "new-only") }
      cp.reportAll()
    }

    "should preserve headers" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-02", "diff"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-03", "diff"))
      )

      val cutter                      = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, _) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { readSheetRows(oldCutPath, "Sheet1").head shouldBe List("Date", "Value") }
      cp { readSheetRows(newCutPath, "Sheet1").head shouldBe List("Date", "Value") }
      cp.reportAll()
    }

    "should handle no equal leading rows" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01", "old")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-02", "new")))

      val cutter                            = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 0 }
      cp { readSheetRows(oldCutPath, "Sheet1") should have size 2 }
      cp { readSheetRows(newCutPath, "Sheet1") should have size 2 }
      cp.reportAll()
    }

    "should use trackConfig for data row detection" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Title"),
        List(List("SubHeader"), List("DATA-1", "same"), List("DATA-2", "old-only"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Title"),
        List(List("SubHeader"), List("DATA-1", "same"), List("DATA-3", "new-only"))
      )

      val track =
        TrackConfig(List(TrackPolicy(SheetSelector.Default, NonEmptyList.of(TrackCondition(0, _.startsWith("DATA"))))))

      val cutter                   = PairedSheetCutter(Set("Sheet1"), track)
      val (oldCutPath, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 1 }
      cp { oldCutRows should have size 3 }
      cp { oldCutRows(2) shouldBe List("DATA-2", "old-only") }
      cp.reportAll()
    }

    "should ignore specified columns when comparing" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Amount", "Note"),
        List(List("2024-01-01", "100", "old-note"), List("2024-01-02", "200", "same"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Amount", "Note"),
        List(List("2024-01-01", "100", "new-note"), List("2024-01-02", "200", "same"))
      )

      val compare = CompareConfig(List(ComparePolicy(SheetSelector.Default, Set(2))))

      val cutter                   = PairedSheetCutter(Set("Sheet1"), compareConfig = compare)
      val (oldCutPath, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { readSheetRows(oldCutPath, "Sheet1") should have size 1 }
      cp.reportAll()
    }

    "should report first mismatch row and key" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-03", "old-only"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Value"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"), List("2024-01-04", "new-only"))
      )

      val sortConfigs =
        Map("Sheet1" -> SheetSortingConfig("Sheet1", List(SortingDsl.asc(0)(SortingDsl.Parsers.asLocalDateDefault))))

      val cutter          = PairedSheetCutter(Set("Sheet1"), sortConfigsMap = sortConfigs)
      val (_, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { results.head.firstMismatchRowNum shouldBe Some(4) }
      cp { results.head.firstMismatchKey shouldBe Some("2024-01-03") }
      cp.reportAll()
    }

    "should treat reordered columns as equal when cutting" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      // Old: Date | Amount | Note
      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Amount", "Note"),
        List(
          List("2024-01-01", "100", "note1"),
          List("2024-01-02", "200", "note2"),
          List("2024-01-03", "300", "old-only")
        )
      )
      // New: Date | Note | Amount  (Note and Amount swapped)
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Note", "Amount"),
        List(
          List("2024-01-01", "note1", "100"),
          List("2024-01-02", "note2", "200"),
          List("2024-01-04", "new-only", "400")
        )
      )

      val cutter                            = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")
      val newCutRows = readSheetRows(newCutPath, "Sheet1")

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { oldCutRows should have size 2 }
      cp { oldCutRows(1) shouldBe List("2024-01-03", "300", "old-only") }
      cp { newCutRows should have size 2 }
      cp { newCutRows(1) shouldBe List("2024-01-04", "new-only", "400") }
      cp.reportAll()
    }

    "should treat reordered columns with different values as unequal" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      // Old: Date | Amount | Note  — Amount = 100
      createTestWorkbook(oldPath, "Sheet1", List("Date", "Amount", "Note"), List(List("2024-01-01", "100", "note1")))
      // New: Date | Note | Amount  — Amount = 999 (different)
      createTestWorkbook(newPath, "Sheet1", List("Date", "Note", "Amount"), List(List("2024-01-01", "note1", "999")))

      val cutter          = PairedSheetCutter(Set("Sheet1"))
      val (_, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 0 }
      cp.reportAll()
    }

    "should ignore new-only column and cut equal rows on common columns" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      // Old: Date | Amount  (no Note column)
      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Amount"),
        List(List("2024-01-01", "100"), List("2024-01-02", "200"))
      )
      // New: Date | Amount | Note  (extra Note column)
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Amount", "Note"),
        List(List("2024-01-01", "100", "extra1"), List("2024-01-02", "200", "extra2"))
      )

      val cutter                   = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { readSheetRows(oldCutPath, "Sheet1") should have size 1 }
      cp.reportAll()
    }

    "should combine column mapping with ignored columns" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      // Old: Date | Amount | Note  (Note at old-index 2)
      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Date", "Amount", "Note"),
        List(List("2024-01-01", "100", "old-note"), List("2024-01-02", "200", "old-note2"))
      )
      // New: Date | Note | Amount  (Note and Amount swapped; Note values differ from old)
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Date", "Note", "Amount"),
        List(List("2024-01-01", "new-note", "100"), List("2024-01-02", "new-note2", "200"))
      )

      // Ignore old-column-index 2 (Note in old file)
      val compare = CompareConfig(List(ComparePolicy(SheetSelector.Default, Set(2))))

      val cutter                   = PairedSheetCutter(Set("Sheet1"), compareConfig = compare)
      val (oldCutPath, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { readSheetRows(oldCutPath, "Sheet1") should have size 1 }
      cp.reportAll()
    }

    "should report no mismatch when all data rows are equal" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(
        oldPath,
        "Sheet1",
        List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"))
      )
      createTestWorkbook(
        newPath,
        "Sheet1",
        List("Header"),
        List(List("2024-01-01", "same"), List("2024-01-02", "same"))
      )

      val cutter          = PairedSheetCutter(Set("Sheet1"))
      val (_, _, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      val cp = new Checkpoint
      cp { results should have size 1 }
      cp { results.head.equalRowCount shouldBe 2 }
      cp { results.head.firstMismatchRowNum shouldBe None }
      cp { results.head.firstMismatchKey shouldBe None }
      cp.reportAll()
    }
  }
}
