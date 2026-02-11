package io.github.ssstlis.excelsorter.config

import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CompareConfigSpec extends AnyFreeSpec with Matchers {

  "CompareConfig" - {

    "should return empty set for empty config" in {
      val config = CompareConfig.empty
      config.ignoredColumns("Sheet1", 0) shouldBe Set.empty
    }

    "should return ignored columns by sheet name" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.ByName("Sheet1"), Set(1, 3)),
        ComparePolicy(SheetSelector.ByName("Sheet2"), Set(5))
      ))

      config.ignoredColumns("Sheet1", 0) shouldBe Set(1, 3)
      config.ignoredColumns("Sheet2", 1) shouldBe Set(5)
    }

    "should return ignored columns by sheet index" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.ByIndex(0), Set(2, 4))
      ))

      config.ignoredColumns("AnyName", 0) shouldBe Set(2, 4)
    }

    "should fall back to Default when no name/index match" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.ByName("SpecificSheet"), Set(1)),
        ComparePolicy(SheetSelector.Default, Set(3, 5))
      ))

      config.ignoredColumns("UnknownSheet", 99) shouldBe Set(3, 5)
    }

    "should prefer name match over default" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.Default, Set(0)),
        ComparePolicy(SheetSelector.ByName("Sheet1"), Set(1, 2))
      ))

      config.ignoredColumns("Sheet1", 0) shouldBe Set(1, 2)
    }

    "should prefer index match over default" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.Default, Set(0)),
        ComparePolicy(SheetSelector.ByIndex(0), Set(3))
      ))

      config.ignoredColumns("Sheet1", 0) shouldBe Set(3)
    }

    "should return empty set when no matching policy and no default" in {
      val config = CompareConfig(List(
        ComparePolicy(SheetSelector.ByName("Other"), Set(1))
      ))

      config.ignoredColumns("Sheet1", 0) shouldBe Set.empty
    }
  }
}
