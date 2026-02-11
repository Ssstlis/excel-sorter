package io.github.ssstlis.excelsorter.config

import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.dsl.SortOrder
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class ConfigReaderSpec extends AnyFreeSpec with Matchers {

  "ConfigReader" - {

    "should parse a simple config with one sheet and one sort" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Sheet1"
          |    sorts: [
          |      {order: "asc", index: 0, as: "String"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)

      result should have size 1
      result.head.sheetName shouldBe "Sheet1"
      result.head.sortConfigs should have size 1
      result.head.sortConfigs.head.columnIndex shouldBe 0
      result.head.sortConfigs.head.order shouldBe SortOrder.Asc
    }

    "should parse desc order" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Sheet1"
          |    sorts: [
          |      {order: "desc", index: 3, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)

      result.head.sortConfigs.head.order shouldBe SortOrder.Desc
      result.head.sortConfigs.head.columnIndex shouldBe 3
    }

    "should parse all supported parser types" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "AllTypes"
          |    sorts: [
          |      {order: "asc", index: 0, as: "String"},
          |      {order: "asc", index: 1, as: "Int"},
          |      {order: "asc", index: 2, as: "Long"},
          |      {order: "asc", index: 3, as: "Double"},
          |      {order: "asc", index: 4, as: "BigDecimal"},
          |      {order: "asc", index: 5, as: "LocalDate"},
          |      {order: "asc", index: 6, as: "LocalDate(dd.MM.yyyy)"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)

      result should have size 1
      result.head.sortConfigs should have size 7

      result.head.sortConfigs(0).columnIndex shouldBe 0
      result.head.sortConfigs(1).columnIndex shouldBe 1
      result.head.sortConfigs(2).columnIndex shouldBe 2
      result.head.sortConfigs(3).columnIndex shouldBe 3
      result.head.sortConfigs(4).columnIndex shouldBe 4
      result.head.sortConfigs(5).columnIndex shouldBe 5
      result.head.sortConfigs(6).columnIndex shouldBe 6
    }

    "should parse multiple sheets" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "First"
          |    sorts: [
          |      {order: "asc", index: 0, as: "LocalDate"}
          |    ]
          |  },
          |  {
          |    name: "Second"
          |    sorts: [
          |      {order: "desc", index: 1, as: "String"},
          |      {order: "asc", index: 2, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)

      result should have size 2
      result(0).sheetName shouldBe "First"
      result(0).sortConfigs should have size 1
      result(1).sheetName shouldBe "Second"
      result(1).sortConfigs should have size 2
    }

    "should throw on unknown parser type" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Bad"
          |    sorts: [
          |      {order: "asc", index: 0, as: "Unknown"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val ex = the[IllegalArgumentException] thrownBy ConfigReader.fromConfig(config)
      ex.getMessage should include("Unknown")
    }

    "should throw on unknown sort order" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Bad"
          |    sorts: [
          |      {order: "sideways", index: 0, as: "String"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val ex = the[IllegalArgumentException] thrownBy ConfigReader.fromConfig(config)
      ex.getMessage should include("sideways")
    }

    "should correctly parse parsers that can compare values" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Sheet1"
          |    sorts: [
          |      {order: "asc", index: 0, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("1", "2") should be < 0
      sortConfig.compare("10", "2") should be > 0
      sortConfig.compare("5", "5") shouldBe 0
    }

    "should correctly apply desc ordering" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Sheet1"
          |    sorts: [
          |      {order: "desc", index: 0, as: "Int"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("1", "2") should be > 0
      sortConfig.compare("10", "2") should be < 0
    }

    "should parse LocalDate with custom pattern and compare correctly" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: [
          |  {
          |    name: "Sheet1"
          |    sorts: [
          |      {order: "asc", index: 0, as: "LocalDate(dd.MM.yyyy)"}
          |    ]
          |  }
          |]
          |""".stripMargin)

      val result = ConfigReader.fromConfig(config)
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("01.01.2024", "15.06.2024") should be < 0
      sortConfig.compare("31.12.2024", "01.01.2024") should be > 0
    }
  }

  "ConfigReader.readTrackConfig" - {

    "should return empty TrackConfig when tracks key is absent" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: []
          |""".stripMargin)

      val trackConfig = ConfigReader.readTrackConfig(config)
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

      val trackConfig = ConfigReader.readTrackConfig(config)
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

      val trackConfig = ConfigReader.readTrackConfig(config)
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

      val trackConfig = ConfigReader.readTrackConfig(config)
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

      val trackConfig = ConfigReader.readTrackConfig(config)
      trackConfig.policies.head.conditions should have size 2
      trackConfig.policies.head.conditions(0).columnIndex shouldBe 0
      trackConfig.policies.head.conditions(1).columnIndex shouldBe 1
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

      val trackConfig = ConfigReader.readTrackConfig(config)
      val conditions = trackConfig.policies.head.conditions

      conditions(0).validator("hello") shouldBe true
      conditions(0).validator("") shouldBe false
      conditions(1).validator("42") shouldBe true
      conditions(1).validator("abc") shouldBe false
      conditions(2).validator("123456789") shouldBe true
      conditions(3).validator("3.14") shouldBe true
      conditions(4).validator("99.99") shouldBe true
      conditions(5).validator("2024-01-15") shouldBe true
      conditions(5).validator("not-a-date") shouldBe false
      conditions(6).validator("15.01.2024") shouldBe true
      conditions(6).validator("2024-01-15") shouldBe false
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

      val ex = the[IllegalArgumentException] thrownBy ConfigReader.readTrackConfig(config)
      ex.getMessage should include("Unknown")
    }
  }

  "ConfigReader.readCompareConfig" - {

    "should return empty CompareConfig when comparisons key is absent" in {
      val config = ConfigFactory.parseString(
        """
          |sortings: []
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig shouldBe CompareConfig.empty
    }

    "should parse null sheet as Default selector" in {
      val config = ConfigFactory.parseString(
        """
          |comparisons: [
          |  {
          |    sheet: null
          |    ignoreColumns: [3, 5]
          |  }
          |]
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig.policies should have size 1
      compareConfig.policies.head.sheetSelector shouldBe SheetSelector.Default
      compareConfig.policies.head.ignoreColumns shouldBe Set(3, 5)
    }

    "should parse named sheet selector" in {
      val config = ConfigFactory.parseString(
        """
          |comparisons: [
          |  {
          |    sheet: "MySheet"
          |    ignoreColumns: [1, 4, 7]
          |  }
          |]
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig.policies should have size 1
      compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("MySheet")
      compareConfig.policies.head.ignoreColumns shouldBe Set(1, 4, 7)
    }

    "should parse indexed sheet selector" in {
      val config = ConfigFactory.parseString(
        """
          |comparisons: [
          |  {
          |    sheet: 0
          |    ignoreColumns: [2]
          |  }
          |]
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig.policies should have size 1
      compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByIndex(0)
      compareConfig.policies.head.ignoreColumns shouldBe Set(2)
    }

    "should parse multiple policies" in {
      val config = ConfigFactory.parseString(
        """
          |comparisons: [
          |  {
          |    sheet: null
          |    ignoreColumns: [3]
          |  },
          |  {
          |    sheet: "Sheet1"
          |    ignoreColumns: [1, 2]
          |  }
          |]
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig.policies should have size 2
    }

    "should parse empty ignoreColumns list" in {
      val config = ConfigFactory.parseString(
        """
          |comparisons: [
          |  {
          |    sheet: null
          |    ignoreColumns: []
          |  }
          |]
          |""".stripMargin)

      val compareConfig = ConfigReader.readCompareConfig(config)
      compareConfig.policies.head.ignoreColumns shouldBe Set.empty
    }
  }
}
