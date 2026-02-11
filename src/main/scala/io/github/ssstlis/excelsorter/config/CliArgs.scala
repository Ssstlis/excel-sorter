package io.github.ssstlis.excelsorter.config

sealed trait RunMode
object RunMode {
  case object SortOnly extends RunMode
  case object Cut extends RunMode
  case object Compare extends RunMode
}

case class CliArgs(mode: RunMode, filePaths: Seq[String])

object CliArgs {

  private val cutFlags = Set("--cut", "-c")
  private val compareFlags = Set("--compare", "-cmp")
  private val allFlags = cutFlags ++ compareFlags

  def parse(args: Array[String]): Either[String, CliArgs] = {
    val (flags, files) = args.partition(_.startsWith("-"))

    val unknownFlags = flags.filterNot(allFlags.contains)
    if (unknownFlags.nonEmpty) {
      return Left(s"Unknown flag(s): ${unknownFlags.mkString(", ")}. Supported: --cut/-c, --compare/-cmp")
    }

    val hasCut = flags.exists(cutFlags.contains)
    val hasCompare = flags.exists(compareFlags.contains)

    if (hasCut && hasCompare) {
      return Left("Flags --cut and --compare are mutually exclusive. Use only one.")
    }

    if (files.isEmpty) {
      return Left("No input files specified.")
    }

    val mode = if (hasCut) RunMode.Cut
               else if (hasCompare) RunMode.Compare
               else RunMode.SortOnly

    Right(CliArgs(mode, files.toSeq))
  }
}
