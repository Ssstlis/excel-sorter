package io.github.ssstlis.excelsorter.config.track

import cats.data.NonEmptyList
import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.config.SheetSelector
import io.github.ssstlis.excelsorter.processor.CellUtils
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class TrackConfigSpec extends AnyFreeSpec with Matchers  {
  "TrackConfig.readTrackConfig" - {

    "should return empty TrackConfig when tracks key is absent" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: []
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      trackConfig shouldBe TrackConfig.empty
    }

    "should parse null sheet as Default selector" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: null
          |    conditions: [
          |      {index: 0, as: "LocalDate"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      trackConfig.policies should have size 1
      trackConfig.policies.head.sheetSelector shouldBe SheetSelector.Default
    }

    "should parse named sheet selector" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: "MySheet"
          |    conditions: [
          |      {index: 0, as: "String"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      trackConfig.policies should have size 1
      trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("MySheet")
    }

    "should parse indexed sheet selector" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: 2
          |    conditions: [
          |      {index: 0, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      trackConfig.policies should have size 1
      trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByIndex(2)
    }

    "should parse multiple conditions" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: null
          |    conditions: [
          |      {index: 0, as: "LocalDate"},
          |      {index: 1, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      trackConfig.policies.head.conditions should have size 2
      trackConfig.policies.head.conditions.toList(0).columnIndex shouldBe 0
      trackConfig.policies.head.conditions.toList(1).columnIndex shouldBe 1
    }

    "should create validators for all supported parser types" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: null
          |    conditions: [
          |      {index: 0, as: "String"},
          |      {index: 1, as: "Int"},
          |      {index: 2, as: "Long"},
          |      {index: 3, as: "Double"},
          |      {index: 4, as: "BigDecimal"},
          |      {index: 5, as: "LocalDate"},
          |      {index: 6, as: "LocalDate(dd.MM.yyyy)"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val preTrackConfig = TrackConfig.readTrackConfig(config)
      preTrackConfig shouldBe a[Right[_, _]]
      val trackConfig = preTrackConfig.toOption.get
      val conditions = trackConfig.policies.head.conditions

      conditions.toList(0).validator("hello") shouldBe true
      conditions.toList(0).validator("") shouldBe false
      conditions.toList(1).validator("42") shouldBe true
      conditions.toList(1).validator("abc") shouldBe false
      conditions.toList(2).validator("123456789") shouldBe true
      conditions.toList(3).validator("3.14") shouldBe true
      conditions.toList(4).validator("99.99") shouldBe true
      conditions.toList(5).validator("2024-01-15") shouldBe true
      conditions.toList(5).validator("not-a-date") shouldBe false
      conditions.toList(6).validator("15.01.2024") shouldBe true
      conditions.toList(6).validator("2024-01-15") shouldBe false
    }

    "should throw on unknown track condition type" in {
      val config = ConfigFactory.parseString(
        """
          |tracks: [
          |  {
          |    sheet: null
          |    conditions: [
          |      {index: 0, as: "Unknown"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = TrackConfig.readTrackConfig(config)
      result shouldBe a[Left[_, _]]
      result.swap.toOption.get should include("Unknown")
    }
  }

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
        TrackPolicy(SheetSelector.ByName("Target"), NonEmptyList.of(
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
        TrackPolicy(SheetSelector.ByIndex(2), NonEmptyList.of(
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
        TrackPolicy(SheetSelector.ByName("Specific"), NonEmptyList.of(
          TrackCondition(0, _ == "X")
        )),
        TrackPolicy(SheetSelector.Default, NonEmptyList.of(
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
      val config = TrackConfig(
        List(
          TrackPolicy(
            SheetSelector.ByName("Multi"),
            NonEmptyList.of(
              TrackCondition(0, s => scala.util.Try(s.toInt).isSuccess),
              TrackCondition(1, s => s.nonEmpty)
            )
          )
        )
      )

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
      val config = TrackConfig(
        List(
          TrackPolicy(
            SheetSelector.Default,
            NonEmptyList.of(TrackCondition(0, _ == "DEFAULT"))
          ),
          TrackPolicy(
            SheetSelector.ByName("Named"),
            NonEmptyList.of(TrackCondition(0, _ == "NAMED"))
          )
        )
      )

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
