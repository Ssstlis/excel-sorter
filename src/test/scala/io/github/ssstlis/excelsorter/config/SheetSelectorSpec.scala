package io.github.ssstlis.excelsorter.config

import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class SheetSelectorSpec extends AnyFreeSpec with Matchers {
  "SheetSelector.parseSheetSelector" - {

    "parses 'default' as Default" in {
      SheetSelector.parseSheetSelector("default") shouldBe SheetSelector.Default
    }

    "parses numeric string as ByIndex" in {
      SheetSelector.parseSheetSelector("0") shouldBe SheetSelector.ByIndex(0)
      SheetSelector.parseSheetSelector("3") shouldBe SheetSelector.ByIndex(3)
    }

    "parses non-numeric string as ByName" in {
      SheetSelector.parseSheetSelector("MySheet") shouldBe SheetSelector.ByName("MySheet")
    }
  }
}
