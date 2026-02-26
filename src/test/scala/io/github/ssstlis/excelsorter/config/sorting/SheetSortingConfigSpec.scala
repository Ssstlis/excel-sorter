package io.github.ssstlis.excelsorter.config.sorting

import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.dsl.SortOrder
import org.scalatest.Checkpoints
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class SheetSortingConfigSpec extends AnyFreeSpec with Matchers with Checkpoints {

  "SheetSortingConfig.readSortConfig" - {

    "should parse a simple config with one sheet and one sort" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Sheet1"
                                               |    sorts: [
                                               |      {order: "asc", index: 0, as: "String"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val result = preResult.toOption.get

      val cp = new Checkpoint
      cp { result should have size 1 }
      cp { result.head.sheetName shouldBe "Sheet1" }
      cp { result.head.sortConfigs should have size 1 }
      cp { result.head.sortConfigs.head.columnIndex shouldBe 0 }
      cp { result.head.sortConfigs.head.order shouldBe SortOrder.Asc }
      cp.reportAll()
    }

    "should parse desc order" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Sheet1"
                                               |    sorts: [
                                               |      {order: "desc", index: 3, as: "Int"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val result = preResult.toOption.get

      val cp = new Checkpoint
      cp { result.head.sortConfigs.head.order shouldBe SortOrder.Desc }
      cp { result.head.sortConfigs.head.columnIndex shouldBe 3 }
      cp.reportAll()
    }

    "should parse all supported parser types" in {
      val config = ConfigFactory.parseString("""
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

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val result = preResult.toOption.get

      val cp = new Checkpoint
      cp { result should have size 1 }
      cp { result.head.sortConfigs should have size 7 }
      cp { result.head.sortConfigs(0).columnIndex shouldBe 0 }
      cp { result.head.sortConfigs(1).columnIndex shouldBe 1 }
      cp { result.head.sortConfigs(2).columnIndex shouldBe 2 }
      cp { result.head.sortConfigs(3).columnIndex shouldBe 3 }
      cp { result.head.sortConfigs(4).columnIndex shouldBe 4 }
      cp { result.head.sortConfigs(5).columnIndex shouldBe 5 }
      cp { result.head.sortConfigs(6).columnIndex shouldBe 6 }
      cp.reportAll()
    }

    "should parse multiple sheets" in {
      val config = ConfigFactory.parseString("""
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

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val result = preResult.toOption.get

      val cp = new Checkpoint
      cp { result should have size 2 }
      cp { result(0).sheetName shouldBe "First" }
      cp { result(0).sortConfigs should have size 1 }
      cp { result(1).sheetName shouldBe "Second" }
      cp { result(1).sortConfigs should have size 2 }
      cp.reportAll()
    }

    "should throw on unknown parser type" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Bad"
                                               |    sorts: [
                                               |      {order: "asc", index: 0, as: "Unknown"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val result = SheetSortingConfig.readSortConfig(config)
      val cp     = new Checkpoint
      cp { result shouldBe a[Left[_, _]] }
      cp { result.left.getOrElse("") should include("Unknown") }
      cp.reportAll()
    }

    "should throw on unknown sort order" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Bad"
                                               |    sorts: [
                                               |      {order: "sideways", index: 0, as: "String"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val result = SheetSortingConfig.readSortConfig(config)
      val cp     = new Checkpoint
      cp { result shouldBe a[Left[_, _]] }
      cp { result.left.getOrElse("") should include("sideways") }
      cp.reportAll()
    }

    "should correctly parse parsers that can compare values" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Sheet1"
                                               |    sorts: [
                                               |      {order: "asc", index: 0, as: "Int"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val sortConfig = preResult.toOption.get.head.sortConfigs.head

      val cp = new Checkpoint
      cp { sortConfig.compare("1", "2") should be < 0 }
      cp { sortConfig.compare("10", "2") should be > 0 }
      cp { sortConfig.compare("5", "5") shouldBe 0 }
      cp.reportAll()
    }

    "should correctly apply desc ordering" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Sheet1"
                                               |    sorts: [
                                               |      {order: "desc", index: 0, as: "Int"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val sortConfig = preResult.toOption.get.head.sortConfigs.head

      val cp = new Checkpoint
      cp { sortConfig.compare("1", "2") should be > 0 }
      cp { sortConfig.compare("10", "2") should be < 0 }
      cp.reportAll()
    }

    "should parse LocalDate with custom pattern and compare correctly" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: [
                                               |  {
                                               |    name: "Sheet1"
                                               |    sorts: [
                                               |      {order: "asc", index: 0, as: "LocalDate(dd.MM.yyyy)"}
                                               |    ]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preResult = SheetSortingConfig.readSortConfig(config)
      preResult shouldBe a[Right[_, _]]
      val sortConfig = preResult.toOption.get.head.sortConfigs.head

      val cp = new Checkpoint
      cp { sortConfig.compare("01.01.2024", "15.06.2024") should be < 0 }
      cp { sortConfig.compare("31.12.2024", "01.01.2024") should be > 0 }
      cp.reportAll()
    }
  }
}
