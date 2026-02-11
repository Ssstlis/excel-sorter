package io.github.ssstlis.excelsorter.config

import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class CliArgsSpec extends AnyFreeSpec with Matchers {

  "CliArgs.parse" - {

    "should parse SortOnly mode when no flags given" in {
      val result = CliArgs.parse(Array("file1.xlsx", "file2.xlsx"))

      result shouldBe Right(CliArgs(RunMode.SortOnly, Seq("file1.xlsx", "file2.xlsx")))
    }

    "should parse Cut mode with --cut flag" in {
      val result = CliArgs.parse(Array("--cut", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx")))
    }

    "should parse Cut mode with -c flag" in {
      val result = CliArgs.parse(Array("-c", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx")))
    }

    "should parse Compare mode with --compare flag" in {
      val result = CliArgs.parse(Array("--compare", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Compare, Seq("file1.xlsx")))
    }

    "should parse Compare mode with -cmp flag" in {
      val result = CliArgs.parse(Array("-cmp", "file1.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Compare, Seq("file1.xlsx")))
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

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("file1.xlsx")))
    }

    "should handle multiple files with flag" in {
      val result = CliArgs.parse(Array("-c", "a.xlsx", "b.xlsx", "c.xlsx"))

      result shouldBe Right(CliArgs(RunMode.Cut, Seq("a.xlsx", "b.xlsx", "c.xlsx")))
    }
  }
}
