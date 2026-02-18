package io.github.ssstlis.excelsorter.processor

import org.scalatest.Checkpoints
import org.scalatest.freespec.AnyFreeSpec
import org.scalatest.matchers.should.Matchers

class FilePairerSpec extends AnyFreeSpec with Matchers with Checkpoints {

  "FilePairer.groupFiles" - {

    "should pair files from different directories by basename" in {
      val oldPath = "/Users/ssstlis/projects/aton/file_sharing/feebase-psm-reports_old/21912703-RUB-2025-ITD_old.xlsx"
      val newPath = "/Users/ssstlis/projects/aton/file_sharing/feebase-psm-reports_new/21912703-RUB-2025-ITD_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))
      val pair   = result.pairs.head

      val cp = new Checkpoint
      cp { result.pairs should have size 1 }
      cp { result.unpaired shouldBe empty }
      cp { pair.prefix shouldBe "21912703-RUB-2025-ITD" }
      cp { pair.oldFile shouldBe oldPath }
      cp { pair.newFile shouldBe newPath }
      cp.reportAll()
    }

    "should pair files from same directory" in {
      val oldPath = "/some/path/report_old.xlsx"
      val newPath = "/some/path/report_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))
      val pair   = result.pairs.head

      val cp = new Checkpoint
      cp { result.pairs should have size 1 }
      cp { result.unpaired shouldBe empty }
      cp { pair.prefix shouldBe "report" }
      cp { pair.oldFile shouldBe oldPath }
      cp { pair.newFile shouldBe newPath }
      cp.reportAll()
    }

    "should handle multiple pairs" in {
      val paths =
        Seq("/dir1/client1_old.xlsx", "/dir2/client1_new.xlsx", "/dir3/client2_old.xlsx", "/dir4/client2_new.xlsx")

      val result = FilePairer.groupFiles(paths)

      val cp = new Checkpoint
      cp { result.pairs should have size 2 }
      cp { result.unpaired shouldBe empty }
      cp { result.pairs.map(_.prefix) should contain theSameElementsAs List("client1", "client2") }
      cp.reportAll()
    }

    "should leave unpaired files when only _old exists" in {
      val oldPath   = "/path/report_old.xlsx"
      val otherPath = "/path/other.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, otherPath))

      val cp = new Checkpoint
      cp { result.pairs shouldBe empty }
      cp { result.unpaired should contain theSameElementsAs List(oldPath, otherPath) }
      cp.reportAll()
    }

    "should leave unpaired files when only _new exists" in {
      val newPath = "/path/report_new.xlsx"

      val result = FilePairer.groupFiles(Seq(newPath))

      val cp = new Checkpoint
      cp { result.pairs shouldBe empty }
      cp { result.unpaired should contain(newPath) }
      cp.reportAll()
    }

    "should handle mixed paired and unpaired files" in {
      val paths = Seq("/dir1/paired_old.xlsx", "/dir2/paired_new.xlsx", "/dir3/orphan_old.xlsx", "/dir4/regular.xlsx")

      val result = FilePairer.groupFiles(paths)

      val cp = new Checkpoint
      cp { result.pairs should have size 1 }
      cp { result.pairs.head.prefix shouldBe "paired" }
      cp { result.unpaired should have size 2 }
      cp { result.unpaired should contain("/dir3/orphan_old.xlsx") }
      cp { result.unpaired should contain("/dir4/regular.xlsx") }
      cp.reportAll()
    }

    "should handle empty input" in {
      val result = FilePairer.groupFiles(Seq.empty)

      val cp = new Checkpoint
      cp { result.pairs shouldBe empty }
      cp { result.unpaired shouldBe empty }
      cp.reportAll()
    }

    "should handle files with complex prefixes containing underscores and dashes" in {
      val oldPath = "/a/21912703-RUB-2025-ITD_old.xlsx"
      val newPath = "/b/21912703-RUB-2025-ITD_new.xlsx"

      val result = FilePairer.groupFiles(Seq(oldPath, newPath))

      val cp = new Checkpoint
      cp { result.pairs should have size 1 }
      cp { result.pairs.head.prefix shouldBe "21912703-RUB-2025-ITD" }
      cp.reportAll()
    }
  }
}
