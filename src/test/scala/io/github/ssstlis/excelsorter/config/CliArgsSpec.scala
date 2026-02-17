package io.github.ssstlis.excelsorter.config

import io.github.ssstlis.excelsorter.dsl.SortOrder
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CliArgsSpec extends AnyFreeSpec with Matchers {

  "CliArgs.parse" - {

    "should parse SortOnly mode when no flags given" in {
      val result = CliArgs.parse(Array("file1.xlsx", "file2.xlsx"))

      result shouldBe Right(CliArgs(RunMode.SortOnly, Seq("file1.xlsx", "file2.xlsx"), None))
    }

    "should parse Cut mode with --cut flag" in {
      val result = CliArgs.parse(Array("--cut", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx"), None))
    }

    "should parse Cut mode with -c flag" in {
      val result = CliArgs.parse(Array("-c", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx"), None))
    }

    "should parse Compare mode with --compare flag" in {
      val result = CliArgs.parse(Array("--compare", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Compare, Seq("file1.xlsx"), None))
    }

    "should parse Compare mode with -cmp flag" in {
      val result = CliArgs.parse(Array("-cmp", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Compare, Seq("file1.xlsx"), None))
    }

    "should reject mutually exclusive --cut and --compare" in {
      val result = CliArgs.parse(Array("--cut", "--compare", "file1.xlsx"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("mutually exclusive")
    }

    "should reject mutually exclusive -c and -cmp" in {
      val result = CliArgs.parse(Array("-c", "-cmp", "file1.xlsx"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("mutually exclusive")
    }

    "should reject unknown flags" in {
      val result = CliArgs.parse(Array("--unknown", "file1.xlsx"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Unknown flag")
    }

    "should reject empty files" in {
      val result = CliArgs.parse(Array("--cut"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("No input files")
    }

    "should handle flag after files" in {
      val result = CliArgs.parse(Array("file1.xlsx", "--cut"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx"), None))
    }

    "should handle multiple files with flag" in {
      val result = CliArgs.parse(Array("-c", "a.xlsx", "b.xlsx", "c.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("a.xlsx", "b.xlsx", "c.xlsx"), None))
    }

    // Help flag tests

    "-h returns help" in {
      val result = CliArgs.parse(Array("-h"))

      result shouldBe Left("help")
    }

    "--help returns help" in {
      val result = CliArgs.parse(Array("--help"))

      result shouldBe Left("help")
    }

    "-h with other args still returns help" in {
      val result = CliArgs.parse(Array("-c", "file.xlsx", "-h"))

      result shouldBe Left("help")
    }

    // --conf with --comparisons

    "--conf with --comparisons block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--comparisons", "-sheet", "Sheet1", "-ic", "1", "13"))

      result.isRight shouldBe true
      val cliArgs = result.toOption.get
      cliArgs.filePaths shouldBe Seq("file.xlsx")
      cliArgs.appConfig shouldBe defined

      val cc = cliArgs.appConfig.get
      cc.compareConfig.policies should have size 1
      cc.compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1")
      cc.compareConfig.policies.head.ignoreColumns shouldBe Set(1, 13)
    }

    // --conf with --sortings

    "--conf with --sortings block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "Sheet1", "-sort", "asc", "0", "LocalDate"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.sortConfig should have size 1
      cc.sortConfig.head.sheetName shouldBe "Sheet1"
      cc.sortConfig.head.sortConfigs should have size 1
      cc.sortConfig.head.sortConfigs.head.columnIndex shouldBe 0
      cc.sortConfig.head.sortConfigs.head.order shouldBe SortOrder.Asc
    }

    "--conf with --sortings multiple sorts" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf",
        "--sortings", "-sheet", "Sheet1", "-sort", "asc", "0", "LocalDate", "-sort", "desc", "2", "String"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.sortConfig.head.sortConfigs should have size 2
      cc.sortConfig.head.sortConfigs(0).order shouldBe SortOrder.Asc
      cc.sortConfig.head.sortConfigs(0).columnIndex shouldBe 0
      cc.sortConfig.head.sortConfigs(1).order shouldBe SortOrder.Desc
      cc.sortConfig.head.sortConfigs(1).columnIndex shouldBe 2
    }

    // --conf with --tracks

    "--conf with --tracks block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "Sheet1", "-cond", "0", "LocalDate"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.trackConfig.policies should have size 1
      cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1")
      cc.trackConfig.policies.head.conditions should have size 1
      cc.trackConfig.policies.head.conditions.head.columnIndex shouldBe 0
    }

    // Multiple blocks of different types

    "--conf with multiple blocks of different types" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf",
        "--sortings", "-sheet", "Sheet1", "-sort", "asc", "0", "String",
        "--tracks", "-sheet", "default", "-cond", "0", "LocalDate",
        "--comparisons", "-sheet", "Sheet1", "-ic", "5"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.sortConfig should have size 1
      cc.trackConfig.policies should have size 1
      cc.compareConfig.policies should have size 1
    }

    // Short flags

    "short flags -s and -o for sortings" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-s", "Sheet1", "-o", "desc", "1", "Int"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.sortConfig.head.sheetName shouldBe "Sheet1"
      cc.sortConfig.head.sortConfigs.head.columnIndex shouldBe 1
      cc.sortConfig.head.sortConfigs.head.order shouldBe SortOrder.Desc
    }

    "short flags -s and -d for tracks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-s", "Sheet1", "-d", "2", "Int"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1")
      cc.trackConfig.policies.head.conditions.head.columnIndex shouldBe 2
    }

    // Sheet selector parsing

    "-sheet default produces Default selector" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "default", "-cond", "0", "String"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.Default
    }

    "-sheet 0 produces ByIndex(0) selector" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "0", "-cond", "0", "String"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByIndex(0)
    }

    // Mode flag + files + --conf combined

    "mode flag + files + --conf combined" in {
      val result = CliArgs.parse(Array("-c", "old.xlsx", "new.xlsx", "--conf",
        "--sortings", "-sheet", "Data", "-sort", "asc", "0", "String"))

      result.isRight shouldBe true
      val cliArgs = result.toOption.get
      cliArgs.mode shouldBe RunMode.Cut
      cliArgs.filePaths shouldBe Seq("old.xlsx", "new.xlsx")
      cliArgs.appConfig shouldBe defined
      cliArgs.appConfig.get.sortConfig.head.sheetName shouldBe "Data"
    }

    // Error cases

    "error: unknown block type after --conf" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--unknown", "-sheet", "X"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Unknown config block type")
    }

    "error: missing -sheet/-s in sortings block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sort", "asc", "0", "String"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Expected -sheet/-s")
    }

    "error: invalid sort order" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S", "-sort", "invalid", "0", "String"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Unknown sort order")
    }

    "error: invalid parser type in sortings" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S", "-sort", "asc", "0", "BadType"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Unknown parser type")
    }

    "error: invalid parser type in tracks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "S", "-cond", "0", "BadType"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Unknown track condition type")
    }

    "error: missing sort entries in sortings block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("at least one -sort/-o entry")
    }

    "error: missing cond entries in tracks block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "S"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("at least one -cond/-d entry")
    }

    "error: missing -ic in comparisons block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--comparisons", "-sheet", "S"))

      result.isLeft shouldBe true
      result.left.getOrElse("") should include("Expected -ic")
    }

    "error: --conf with no blocks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf"))

      result shouldBe Right(CliArgs(RunMode.SortOnly, Seq("file.xlsx"), None))
    }
  }
}
