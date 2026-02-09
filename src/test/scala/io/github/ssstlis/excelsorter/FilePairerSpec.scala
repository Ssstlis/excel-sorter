package io.github.ssstlis.excelsorter

import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class FilePairerSpec extends AnyFreeSpec with Matchers {

  "FilePairer.groupFiles" - {

    "should pair files from different directories by basename" in {
      val oldPath = "/Users/ssstlis/projects/aton/file_sharing/feebase-psm-reports_old/21912703-RUB-2025-ITD_old.xlsx"
      val newPath = "/Users/ssstlis/projects/aton/file_sharing/feebase-psm-reports_new/21912703-RUB-2025-ITD_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))

      result.pairs should have size 1
      result.unpaired shouldBe empty

      val pair = result.pairs.head
      pair.prefix shouldBe "21912703-RUB-2025-ITD"
      pair.oldFile shouldBe oldPath
      pair.newFile shouldBe newPath
    }

    "should pair files from same directory" in {
      val oldPath = "/some/path/report_old.xlsx"
      val newPath = "/some/path/report_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))

      result.pairs should have size 1
      result.unpaired shouldBe empty

      val pair = result.pairs.head
      pair.prefix shouldBe "report"
      pair.oldFile shouldBe oldPath
      pair.newFile shouldBe newPath
    }

    "should handle multiple pairs" in {
      val paths = Seq(
        "/dir1/client1_old.xlsx",
        "/dir2/client1_new.xlsx",
        "/dir3/client2_old.xlsx",
        "/dir4/client2_new.xlsx"
      )

      val result = FilePairer.groupFiles(paths)

      result.pairs should have size 2
      result.unpaired shouldBe empty

      result.pairs.map(_.prefix) should contain theSameElementsAs List("client1", "client2")
    }

    "should leave unpaired files when only _old exists" in {
      val oldPath = "/path/report_old.xlsx"
      val otherPath = "/path/other.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, otherPath))

      result.pairs shouldBe empty
      result.unpaired should contain theSameElementsAs List(oldPath, otherPath)
    }

    "should leave unpaired files when only _new exists" in {
      val newPath = "/path/report_new.xlsx"

      val result = FilePairer.groupFiles(Seq(newPath))

      result.pairs shouldBe empty
      result.unpaired should contain(newPath)
    }

    "should handle mixed paired and unpaired files" in {
      val paths = Seq(
        "/dir1/paired_old.xlsx",
        "/dir2/paired_new.xlsx",
        "/dir3/orphan_old.xlsx",
        "/dir4/regular.xlsx"
      )

      val result = FilePairer.groupFiles(paths)

      result.pairs should have size 1
      result.pairs.head.prefix shouldBe "paired"

      result.unpaired should have size 2
      result.unpaired should contain("/dir3/orphan_old.xlsx")
      result.unpaired should contain("/dir4/regular.xlsx")
    }

    "should handle empty input" in {
      val result = FilePairer.groupFiles(Seq.empty)

      result.pairs shouldBe empty
      result.unpaired shouldBe empty
    }

    "should handle files with complex prefixes containing underscores and dashes" in {
      val oldPath = "/a/21912703-RUB-2025-ITD_old.xlsx"
      val newPath = "/b/21912703-RUB-2025-ITD_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))

      result.pairs should have size 1
      result.pairs.head.prefix shouldBe "21912703-RUB-2025-ITD"
    }
  }
}