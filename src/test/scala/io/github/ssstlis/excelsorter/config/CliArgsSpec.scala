package io.github.ssstlis.excelsorter.config

import io.github.ssstlis.excelsorter.dsl.SortOrder
import org.scalatest.Checkpoints
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CliArgsSpec extends AnyFreeSpec with Matchers with Checkpoints {

  "CliArgs.splitIntoBlocks" - {

    "empty list produces empty result" in {
      CliArgs.splitIntoBlocks(Nil) shouldBe Nil
    }

    "single block with no args" in {
      CliArgs.splitIntoBlocks(List("--sortings")) shouldBe
        List("--sortings" -> Nil)
    }

    "single block with args" in {
      CliArgs.splitIntoBlocks(List("--sortings", "a", "b")) shouldBe
        List("--sortings" -> List("a", "b"))
    }

    "two adjacent blocks with no args" in {
      CliArgs.splitIntoBlocks(List("--sortings", "--tracks")) shouldBe
        List("--sortings" -> Nil, "--tracks" -> Nil)
    }

    "two blocks with args each" in {
      CliArgs.splitIntoBlocks(List("--sortings", "a", "b", "--tracks", "c")) shouldBe
        List("--sortings" -> List("a", "b"), "--tracks" -> List("c"))
    }

    "three blocks preserve order" in {
      CliArgs.splitIntoBlocks(List("--sortings", "s1", "--tracks", "t1", "--comparisons", "c1")) shouldBe
        List("--sortings" -> List("s1"), "--tracks" -> List("t1"), "--comparisons" -> List("c1"))
    }

    "non-starter as first arg produces __unknown__" in {
      CliArgs.splitIntoBlocks(List("unexpected", "--sortings", "a")) shouldBe
        List("__unknown__" -> List("unexpected"))
    }

    "single non-starter arg produces __unknown__" in {
      CliArgs.splitIntoBlocks(List("unexpected")) shouldBe
        List("__unknown__" -> List("unexpected"))
    }
  }

  "CliArgs.parseConfSection" - {

    "empty list returns error" in {
      val result = CliArgs.parseConfSection(Nil)
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("--conf requires at least one") }
      cp.reportAll()
    }

    "valid --sortings block" in {
      val result = CliArgs.parseConfSection(List("--sortings", "-sheet", "Sheet1", "-sort", "asc", "0", "String"))
      result.isRight shouldBe true
      val cfg = result.toOption.get
      val cp  = new Checkpoint
      cp { cfg.sortConfig should have size 1 }
      cp { cfg.sortConfig.head.sheetName shouldBe "Sheet1" }
      cp { cfg.trackConfig.policies shouldBe Nil }
      cp { cfg.compareConfig.policies shouldBe Nil }
      cp.reportAll()
    }

    "valid --tracks block" in {
      val result = CliArgs.parseConfSection(List("--tracks", "-sheet", "default", "-cond", "0", "LocalDate"))
      result.isRight shouldBe true
      val cfg = result.toOption.get
      val cp  = new Checkpoint
      cp { cfg.trackConfig.policies should have size 1 }
      cp { cfg.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.Default }
      cp { cfg.sortConfig shouldBe Nil }
      cp { cfg.compareConfig.policies shouldBe Nil }
      cp.reportAll()
    }

    "valid --comparisons block" in {
      val result = CliArgs.parseConfSection(List("--comparisons", "-sheet", "Sheet1", "-ic", "3", "7"))
      result.isRight shouldBe true
      val cfg = result.toOption.get
      val cp  = new Checkpoint
      cp { cfg.compareConfig.policies should have size 1 }
      cp { cfg.compareConfig.policies.head.ignoreColumns shouldBe Set(3, 7) }
      cp { cfg.sortConfig shouldBe Nil }
      cp { cfg.trackConfig.policies shouldBe Nil }
      cp.reportAll()
    }

    "multiple blocks of each type accumulate in order" in {
      val result = CliArgs.parseConfSection(
        List(
          "--sortings",
          "-sheet",
          "Sheet1",
          "-sort",
          "asc",
          "0",
          "String",
          "--tracks",
          "-sheet",
          "Sheet1",
          "-cond",
          "0",
          "LocalDate",
          "--comparisons",
          "-sheet",
          "Sheet1",
          "-ic",
          "1"
        )
      )
      result.isRight shouldBe true
      val cfg = result.toOption.get
      val cp  = new Checkpoint
      cp { cfg.sortConfig should have size 1 }
      cp { cfg.trackConfig.policies should have size 1 }
      cp { cfg.compareConfig.policies should have size 1 }
      cp.reportAll()
    }

    "two --sortings blocks accumulate both" in {
      val result = CliArgs.parseConfSection(
        List(
          "--sortings",
          "-sheet",
          "A",
          "-sort",
          "asc",
          "0",
          "String",
          "--sortings",
          "-sheet",
          "B",
          "-sort",
          "desc",
          "1",
          "Int"
        )
      )
      result.isRight shouldBe true
      val cfg = result.toOption.get
      val cp  = new Checkpoint
      cp { cfg.sortConfig should have size 2 }
      cp { cfg.sortConfig.map(_.sheetName) shouldBe List("A", "B") }
      cp.reportAll()
    }

    "unknown block type returns error" in {
      val result = CliArgs.parseConfSection(List("--unknown", "-sheet", "X"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown config block type") }
      cp { result.left.getOrElse("") should include("__unknown__") }
      cp.reportAll()
    }

    "non-starter first arg triggers __unknown__ and returns error" in {
      val result = CliArgs.parseConfSection(List("garbage", "--sortings", "-sheet", "S", "-sort", "asc", "0", "String"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown config block type") }
      cp.reportAll()
    }

    "invalid -sort order inside sortings propagates error" in {
      val result = CliArgs.parseConfSection(List("--sortings", "-sheet", "S", "-sort", "sideways", "0", "String"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown sort order") }
      cp.reportAll()
    }

    "error in second block still returns Left" in {
      val result = CliArgs.parseConfSection(
        List(
          "--sortings",
          "-sheet",
          "Good",
          "-sort",
          "asc",
          "0",
          "String",
          "--tracks",
          "-sheet",
          "Bad",
          "-cond",
          "0",
          "NotAType"
        )
      )
      val cp = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown track condition type") }
      cp.reportAll()
    }
  }

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
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("mutually exclusive") }
      cp.reportAll()
    }

    "should reject mutually exclusive -c and -cmp" in {
      val result = CliArgs.parse(Array("-c", "-cmp", "file1.xlsx"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("mutually exclusive") }
      cp.reportAll()
    }

    "should reject unknown flags" in {
      val result = CliArgs.parse(Array("--unknown", "file1.xlsx"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown flag") }
      cp.reportAll()
    }

    "should reject empty files" in {
      val result = CliArgs.parse(Array("--cut"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("No input files") }
      cp.reportAll()
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
      cliArgs.appConfig shouldBe defined
      val cc = cliArgs.appConfig.get
      val cp = new Checkpoint
      cp { cliArgs.filePaths shouldBe Seq("file.xlsx") }
      cp { cc.compareConfig.policies should have size 1 }
      cp { cc.compareConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1") }
      cp { cc.compareConfig.policies.head.ignoreColumns shouldBe Set(1, 13) }
      cp.reportAll()
    }

    // --conf with --sortings

    "--conf with --sortings block" in {
      val result =
        CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "Sheet1", "-sort", "asc", "0", "LocalDate"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.sortConfig should have size 1 }
      cp { cc.sortConfig.head.sheetName shouldBe "Sheet1" }
      cp { cc.sortConfig.head.sortConfigs should have size 1 }
      cp { cc.sortConfig.head.sortConfigs.head.columnIndex shouldBe 0 }
      cp { cc.sortConfig.head.sortConfigs.head.order shouldBe SortOrder.Asc }
      cp.reportAll()
    }

    "--conf with --sortings multiple sorts" in {
      val result = CliArgs.parse(
        Array(
          "file.xlsx",
          "--conf",
          "--sortings",
          "-sheet",
          "Sheet1",
          "-sort",
          "asc",
          "0",
          "LocalDate",
          "-sort",
          "desc",
          "2",
          "String"
        )
      )

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.sortConfig.head.sortConfigs should have size 2 }
      cp { cc.sortConfig.head.sortConfigs(0).order shouldBe SortOrder.Asc }
      cp { cc.sortConfig.head.sortConfigs(0).columnIndex shouldBe 0 }
      cp { cc.sortConfig.head.sortConfigs(1).order shouldBe SortOrder.Desc }
      cp { cc.sortConfig.head.sortConfigs(1).columnIndex shouldBe 2 }
      cp.reportAll()
    }

    // --conf with --tracks

    "--conf with --tracks block" in {
      val result =
        CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "Sheet1", "-cond", "0", "LocalDate"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.trackConfig.policies should have size 1 }
      cp { cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1") }
      cp { cc.trackConfig.policies.head.conditions should have size 1 }
      cp { cc.trackConfig.policies.head.conditions.head.columnIndex shouldBe 0 }
      cp.reportAll()
    }

    // Multiple blocks of different types

    "--conf with multiple blocks of different types" in {
      val result = CliArgs.parse(
        Array(
          "file.xlsx",
          "--conf",
          "--sortings",
          "-sheet",
          "Sheet1",
          "-sort",
          "asc",
          "0",
          "String",
          "--tracks",
          "-sheet",
          "default",
          "-cond",
          "0",
          "LocalDate",
          "--comparisons",
          "-sheet",
          "Sheet1",
          "-ic",
          "5"
        )
      )

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.sortConfig should have size 1 }
      cp { cc.trackConfig.policies should have size 1 }
      cp { cc.compareConfig.policies should have size 1 }
      cp.reportAll()
    }

    // Short flags

    "short flags -s and -o for sortings" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-s", "Sheet1", "-o", "desc", "1", "Int"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.sortConfig.head.sheetName shouldBe "Sheet1" }
      cp { cc.sortConfig.head.sortConfigs.head.columnIndex shouldBe 1 }
      cp { cc.sortConfig.head.sortConfigs.head.order shouldBe SortOrder.Desc }
      cp.reportAll()
    }

    "short flags -s and -d for tracks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-s", "Sheet1", "-d", "2", "Int"))

      result.isRight shouldBe true
      val cc = result.toOption.get.appConfig.get
      val cp = new Checkpoint
      cp { cc.trackConfig.policies.head.sheetSelector shouldBe SheetSelector.ByName("Sheet1") }
      cp { cc.trackConfig.policies.head.conditions.head.columnIndex shouldBe 2 }
      cp.reportAll()
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
      val result = CliArgs.parse(
        Array("-c", "old.xlsx", "new.xlsx", "--conf", "--sortings", "-sheet", "Data", "-sort", "asc", "0", "String")
      )

      result.isRight shouldBe true
      val cliArgs = result.toOption.get
      val cp      = new Checkpoint
      cp { cliArgs.mode shouldBe RunMode.Cut }
      cp { cliArgs.filePaths shouldBe Seq("old.xlsx", "new.xlsx") }
      cp { cliArgs.appConfig shouldBe defined }
      cp { cliArgs.appConfig.get.sortConfig.head.sheetName shouldBe "Data" }
      cp.reportAll()
    }

    // Error cases

    "error: unknown block type after --conf" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--unknown", "-sheet", "X"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown config block type") }
      cp.reportAll()
    }

    "error: missing -sheet/-s in sortings block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sort", "asc", "0", "String"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Expected -sheet/-s") }
      cp.reportAll()
    }

    "error: invalid sort order" in {
      val result =
        CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S", "-sort", "invalid", "0", "String"))
      val cp = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown sort order") }
      cp.reportAll()
    }

    "error: invalid parser type in sortings" in {
      val result =
        CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S", "-sort", "asc", "0", "BadType"))
      val cp = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown parser type") }
      cp.reportAll()
    }

    "error: invalid parser type in tracks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "S", "-cond", "0", "BadType"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Unknown track condition type") }
      cp.reportAll()
    }

    "error: missing sort entries in sortings block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--sortings", "-sheet", "S"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("at least one -sort/-o entry") }
      cp.reportAll()
    }

    "error: missing cond entries in tracks block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--tracks", "-sheet", "S"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("at least one -cond/-d entry") }
      cp.reportAll()
    }

    "error: missing -ic in comparisons block" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf", "--comparisons", "-sheet", "S"))
      val cp     = new Checkpoint
      cp { result.isLeft shouldBe true }
      cp { result.left.getOrElse("") should include("Expected -ic") }
      cp.reportAll()
    }

    "error: --conf with no blocks" in {
      val result = CliArgs.parse(Array("file.xlsx", "--conf"))

      result shouldBe Right(CliArgs(RunMode.SortOnly, Seq("file.xlsx"), None))
    }
  }
}
