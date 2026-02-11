package io.github.ssstlis.excelsorter.processor

import java.io.File

case class FilePair(prefix: String, oldFile: String, newFile: String)

object FilePairer {
  private val OldPattern = """(.+)_old\.xlsx$""".r
  private val NewPattern = """(.+)_new\.xlsx$""".r

  case class GroupedFiles(
    pairs: List[FilePair],
    unpaired: List[String]
  )

  private def basename(path: String): String = new File(path).getName

  private def extractPrefix(path: String, pattern: scala.util.matching.Regex): Option[String] = {
    val name = basename(path)
    pattern.unapplySeq(name).flatMap(_.headOption)
  }

  def groupFiles(paths: Seq[String]): GroupedFiles = {
    val oldFiles: Map[String, String] = paths.flatMap { p =>
      extractPrefix(p, OldPattern).map(prefix => prefix -> p)
    }.toMap

    val newFiles: Map[String, String] = paths.flatMap { p =>
      extractPrefix(p, NewPattern).map(prefix => prefix -> p)
    }.toMap

    val pairedPrefixes = oldFiles.keySet.intersect(newFiles.keySet)

    val pairs = pairedPrefixes.toList.sorted.map { prefix =>
      FilePair(prefix, oldFiles(prefix), newFiles(prefix))
    }

    val pairedPaths = pairs.flatMap(p => List(p.oldFile, p.newFile)).toSet
    val unpaired = paths.filterNot(pairedPaths.contains).toList

    GroupedFiles(pairs, unpaired)
  }
}
