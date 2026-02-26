package io.github.ssstlis.excelsorter.config.compare

import com.typesafe.config.ConfigFactory
import io.github.ssstlis.excelsorter.config.SheetSelector
import org.scalatest.Checkpoints
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CompareConfigSpec extends AnyFreeSpec with Matchers with Checkpoints {

  "CompareConfig.readCompareConfig" - {

    "should return empty CompareConfig when comparisons key is absent" in {
      val config = ConfigFactory.parseString("""
                                               |sortings: []
                                               |""".stripMargin)

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      compareConfig shouldBe CompareConfig.empty
    }

    "should parse null sheet as Default selector" in {
      val config = ConfigFactory.parseString("""
                                               |comparisons: [
                                               |  {
                                               |    sheet: null
                                               |    ignoreColumns: [3, 5]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      val cp            = new Checkpoint
      cp { compareConfig.policies should have size 1 }
      cp { compareConfig.policies.head.sheetSelector shouldBe SheetSelector.Default }
      cp { compareConfig.policies.head.ignoreColumns shouldBe Set(3, 5) }
      cp.reportAll()
    }

    "should parse named sheet selector" in {
      val config = ConfigFactory.parseString("""
                                               |comparisons: [
                                               |  {
                                               |    sheet: "MySheet"
                                               |    ignoreColumns: [1, 4, 7]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      val cp            = new Checkpoint
      cp { compareConfig.policies should have size 1 }
      cp { compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("MySheet") }
      cp { compareConfig.policies.head.ignoreColumns shouldBe Set(1, 4, 7) }
      cp.reportAll()
    }

    "should parse indexed sheet selector" in {
      val config = ConfigFactory.parseString("""
                                               |comparisons: [
                                               |  {
                                               |    sheet: 0
                                               |    ignoreColumns: [2]
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      val cp            = new Checkpoint
      cp { compareConfig.policies should have size 1 }
      cp { compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByIndex(0) }
      cp { compareConfig.policies.head.ignoreColumns shouldBe Set(2) }
      cp.reportAll()
    }

    "should parse multiple policies" in {
      val config = ConfigFactory.parseString("""
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

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      compareConfig.policies should have size 2
    }

    "should parse empty ignoreColumns list" in {
      val config = ConfigFactory.parseString("""
                                               |comparisons: [
                                               |  {
                                               |    sheet: null
                                               |    ignoreColumns: []
                                               |  }
                                               |]
                                               |""".stripMargin)

      val preCompareConfig = CompareConfig.readCompareConfig(config)
      preCompareConfig shouldBe a[Right[_, _]]
      val compareConfig = preCompareConfig.toOption.get
      compareConfig.policies.head.ignoreColumns shouldBe Set.empty
    }
  }

  "CompareConfig" - {

    "should return empty set for empty config" in {
      val config = CompareConfig.empty
      config.ignoredColumns("Sheet1", 0) shouldBe Set.empty
    }

    "should return ignored columns by sheet name" in {
      val config = CompareConfig(
        List(
          ComparePolicy(SheetSelector.ByName("Sheet1"), Set(1, 3)),
          ComparePolicy(SheetSelector.ByName("Sheet2"), Set(5))
        )
      )

      val cp = new Checkpoint
      cp { config.ignoredColumns("Sheet1", 0) shouldBe Set(1, 3) }
      cp { config.ignoredColumns("Sheet2", 1) shouldBe Set(5) }
      cp.reportAll()
    }

    "should return ignored columns by sheet index" in {
      val config = CompareConfig(List(ComparePolicy(SheetSelector.ByIndex(0), Set(2, 4))))

      config.ignoredColumns("AnyName", 0) shouldBe Set(2, 4)
    }

    "should fall back to Default when no name/index match" in {
      val config = CompareConfig(
        List(
          ComparePolicy(SheetSelector.ByName("SpecificSheet"), Set(1)),
          ComparePolicy(SheetSelector.Default, Set(3, 5))
        )
      )

      config.ignoredColumns("UnknownSheet", 99) shouldBe Set(3, 5)
    }

    "should prefer name match over default" in {
      val config = CompareConfig(
        List(ComparePolicy(SheetSelector.Default, Set(0)), ComparePolicy(SheetSelector.ByName("Sheet1"), Set(1, 2)))
      )

      config.ignoredColumns("Sheet1", 0) shouldBe Set(1, 2)
    }

    "should prefer index match over default" in {
      val config = CompareConfig(
        List(ComparePolicy(SheetSelector.Default, Set(0)), ComparePolicy(SheetSelector.ByIndex(0), Set(3)))
      )

      config.ignoredColumns("Sheet1", 0) shouldBe Set(3)
    }

    "should return empty set when no matching policy and no default" in {
      val config = CompareConfig(List(ComparePolicy(SheetSelector.ByName("Other"), Set(1))))

      config.ignoredColumns("Sheet1", 0) shouldBe Set.empty
    }
  }
}
