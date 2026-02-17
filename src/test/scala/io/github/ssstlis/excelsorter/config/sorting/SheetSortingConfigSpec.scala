package io.github.ssstlis.excelsorter.config.sorting

import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.dsl.SortOrder
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class SheetSortingConfigSpec extends AnyFreeSpec with Matchers {

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

      result should have size 1
      result.head.sheetName shouldBe "Sheet1"
      result.head.sortConfigs should have size 1
      result.head.sortConfigs.head.columnIndex shouldBe 0
      result.head.sortConfigs.head.order shouldBe SortOrder.Asc
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

      result.head.sortConfigs.head.order shouldBe SortOrder.Desc
      result.head.sortConfigs.head.columnIndex shouldBe 3
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

      result should have size 2
      result(0).sheetName shouldBe "First"
      result(0).sortConfigs should have size 1
      result(1).sheetName shouldBe "Second"
      result(1).sortConfigs should have size 2
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
      result shouldBe a[Left[_, _]]
      result.swap.toOption.get should include("Unknown")
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
      result shouldBe a[Left[_, _]]
      result.swap.toOption.get should include("sideways")
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
      val result     = preResult.toOption.get
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("1", "2") should be < 0
      sortConfig.compare("10", "2") should be > 0
      sortConfig.compare("5", "5") shouldBe 0
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
      val result     = preResult.toOption.get
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("1", "2") should be > 0
      sortConfig.compare("10", "2") should be < 0
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
      val result     = preResult.toOption.get
      val sortConfig = result.head.sortConfigs.head

      sortConfig.compare("01.01.2024", "15.06.2024") should be < 0
      sortConfig.compare("31.12.2024", "01.01.2024") should be > 0
    }
  }
}
