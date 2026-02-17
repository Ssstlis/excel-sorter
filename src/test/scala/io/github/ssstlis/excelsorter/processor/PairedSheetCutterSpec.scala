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
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

import scala.jdk.CollectionConverters._

class PairedSheetCutterSpec extends AnyFreeSpec with Matchers {

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

      oldCutPath should endWith("_sortcutted.xlsx")
      newCutPath should endWith("_sortcutted.xlsx")
      new File(oldCutPath).exists() shouldBe true
      new File(newCutPath).exists() shouldBe true
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

      readSheetRows(oldPath, "Sheet1") shouldBe oldOriginalRows
      readSheetRows(newPath, "Sheet1") shouldBe newOriginalRows
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

      results should have size 1
      results.head.equalRowCount shouldBe 2

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")
      oldCutRows should have size 2 // header + 1 data row
      oldCutRows(1) shouldBe List("2024-01-03", "old-only")

      val newCutRows = readSheetRows(newCutPath, "Sheet1")
      newCutRows should have size 2
      newCutRows(1) shouldBe List("2024-01-04", "new-only")
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

      readSheetRows(oldCutPath, "Sheet1").head shouldBe List("Date", "Value")
      readSheetRows(newCutPath, "Sheet1").head shouldBe List("Date", "Value")
    }

    "should handle no equal leading rows" in {
      val tmpDir  = Files.createTempDirectory("cutter-test").toFile
      val oldPath = new File(tmpDir, "test_old_sorted.xlsx").getAbsolutePath
      val newPath = new File(tmpDir, "test_new_sorted.xlsx").getAbsolutePath

      createTestWorkbook(oldPath, "Sheet1", List("Header"), List(List("2024-01-01", "old")))
      createTestWorkbook(newPath, "Sheet1", List("Header"), List(List("2024-01-02", "new")))

      val cutter                            = PairedSheetCutter(Set("Sheet1"))
      val (oldCutPath, newCutPath, results) = cutter.cutEqualLeadingRows(oldPath, newPath)

      results should have size 1
      results.head.equalRowCount shouldBe 0

      readSheetRows(oldCutPath, "Sheet1") should have size 2
      readSheetRows(newCutPath, "Sheet1") should have size 2
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

      results should have size 1
      results.head.equalRowCount shouldBe 1

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")
      oldCutRows should have size 3 // Title + SubHeader + 1 data
      oldCutRows(2) shouldBe List("DATA-2", "old-only")
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

      results should have size 1
      results.head.equalRowCount shouldBe 2

      val oldCutRows = readSheetRows(oldCutPath, "Sheet1")
      oldCutRows should have size 1
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

      results should have size 1
      results.head.equalRowCount shouldBe 2
      results.head.firstMismatchRowNum shouldBe Some(4)
      results.head.firstMismatchKey shouldBe Some("2024-01-03")
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

      results should have size 1
      results.head.equalRowCount shouldBe 2
      results.head.firstMismatchRowNum shouldBe None
      results.head.firstMismatchKey shouldBe None
    }
  }
}
