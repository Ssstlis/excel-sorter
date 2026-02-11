package io.github.ssstlis.excelsorter.config

import io.github.ssstlis.excelsorter.dsl.SheetSortingConfig

sealed trait RunMode
object RunMode {
  case object SortOnly extends RunMode
  case object Cut extends RunMode
  case object Compare extends RunMode
}

case class CliConfig(
  sortings: Seq[SheetSortingConfig],
  trackConfig: TrackConfig,
  compareConfig: CompareConfig
)

case class CliArgs(mode: RunMode, filePaths: Seq[String], cliConfig: Option[CliConfig])

object CliArgs {

  private val cutFlags = Set("--cut", "-c")
  private val compareFlags = Set("--compare", "-cmp")
  private val helpFlags = Set("-h", "--help")
  private val modeFlags = cutFlags ++ compareFlags
  private val confFlag = "--conf"

  def parse(args: Array[String]): Either[String, CliArgs] = {
    val argList = args.toList

    if (argList.exists(helpFlags.contains)) {
      return Left("help")
    }

    val confIndex = argList.indexOf(confFlag)
    val (mainArgs, confArgs) = if (confIndex >= 0) {
      (argList.take(confIndex), argList.drop(confIndex + 1))
    } else {
      (argList, Nil)
    }

    val (flags, files) = mainArgs.partition(_.startsWith("-"))

    val unknownFlags = flags.filterNot(modeFlags.contains)
    if (unknownFlags.nonEmpty) {
      return Left(s"Unknown flag(s): ${unknownFlags.mkString(", ")}. Supported: --cut/-c, --compare/-cmp, --conf, -h/--help")
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

    val cliConfig = if (confArgs.nonEmpty) {
      parseConfSection(confArgs) match {
        case Left(err) => return Left(err)
        case Right(cc) => Some(cc)
      }
    } else {
      None
    }

    Right(CliArgs(mode, files, cliConfig))
  }

  private def parseConfSection(args: List[String]): Either[String, CliConfig] = {
    val blocks = splitIntoBlocks(args)

    if (blocks.isEmpty) {
      return Left("--conf requires at least one configuration block (--sortings, --tracks, or --comparisons).")
    }

    var sortings = List.empty[SheetSortingConfig]
    var trackPolicies = List.empty[TrackPolicy]
    var comparePolicies = List.empty[ComparePolicy]

    for ((blockType, blockArgs) <- blocks) {
      blockType match {
        case "--sortings" =>
          parseSortingsBlock(blockArgs) match {
            case Left(err)  => return Left(err)
            case Right(cfg) => sortings = sortings :+ cfg
          }
        case "--tracks" =>
          parseTracksBlock(blockArgs) match {
            case Left(err)     => return Left(err)
            case Right(policy) => trackPolicies = trackPolicies :+ policy
          }
        case "--comparisons" =>
          parseComparisonsBlock(blockArgs) match {
            case Left(err)     => return Left(err)
            case Right(policy) => comparePolicies = comparePolicies :+ policy
          }
        case other =>
          return Left(s"Unknown config block type: '$other'. Expected --sortings, --tracks, or --comparisons.")
      }
    }

    Right(CliConfig(
      sortings,
      TrackConfig(trackPolicies),
      CompareConfig(comparePolicies)
    ))
  }

  private def splitIntoBlocks(args: List[String]): List[(String, List[String])] = {
    val blockStarters = Set("--sortings", "--tracks", "--comparisons")
    val result = List.newBuilder[(String, List[String])]
    var currentType: String = null
    var currentArgs = List.newBuilder[String]

    for (arg <- args) {
      if (blockStarters.contains(arg)) {
        if (currentType != null) {
          result += ((currentType, currentArgs.result()))
        }
        currentType = arg
        currentArgs = List.newBuilder[String]
      } else if (currentType == null) {
        return List(("__unknown__" -> List(arg)))
      } else {
        currentArgs += arg
      }
    }

    if (currentType != null) {
      result += ((currentType, currentArgs.result()))
    }

    result.result()
  }

  private def parseSortingsBlock(args: List[String]): Either[String, SheetSortingConfig] = {
    val (sheetName, rest) = parseSheetName(args) match {
      case Left(err) => return Left(s"--sortings: $err")
      case Right(v)  => v
    }

    val sorts = parseSortEntries(rest) match {
      case Left(err) => return Left(s"--sortings: $err")
      case Right(v)  => v
    }

    if (sorts.isEmpty) {
      return Left("--sortings: at least one -sort/-o entry is required.")
    }

    Right(SheetSortingConfig(sheetName, sorts))
  }

  private def parseSortEntries(args: List[String]): Either[String, List[io.github.ssstlis.excelsorter.dsl.ColumnSortConfig[_]]] = {
    val sortFlags = Set("-sort", "-o")
    var remaining = args
    val result = List.newBuilder[io.github.ssstlis.excelsorter.dsl.ColumnSortConfig[_]]

    while (remaining.nonEmpty) {
      remaining match {
        case flag :: order :: idxStr :: asType :: tail if sortFlags.contains(flag) =>
          val index = try { idxStr.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index: '$idxStr'. Expected an integer.")
          }
          try {
            result += ConfigReader.resolveColumnSort(order, index, asType)
          } catch {
            case e: IllegalArgumentException => return Left(e.getMessage)
          }
          remaining = tail
        case flag :: _ if sortFlags.contains(flag) =>
          return Left(s"$flag requires 3 arguments: <order> <column-index> <type>")
        case other :: _ =>
          return Left(s"Unexpected argument: '$other'. Expected -sort or -o.")
        case Nil => // won't happen due to while condition
      }
    }

    Right(result.result())
  }

  private def parseTracksBlock(args: List[String]): Either[String, TrackPolicy] = {
    val (sheetNameOrDefault, rest) = parseSheetName(args) match {
      case Left(err) => return Left(s"--tracks: $err")
      case Right(v)  => v
    }

    val selector = parseSheetSelector(sheetNameOrDefault)

    val conditions = parseCondEntries(rest) match {
      case Left(err) => return Left(s"--tracks: $err")
      case Right(v)  => v
    }

    if (conditions.isEmpty) {
      return Left("--tracks: at least one -cond/-d entry is required.")
    }

    Right(TrackPolicy(selector, conditions))
  }

  private def parseCondEntries(args: List[String]): Either[String, List[TrackCondition]] = {
    val condFlags = Set("-cond", "-d")
    var remaining = args
    val result = List.newBuilder[TrackCondition]

    while (remaining.nonEmpty) {
      remaining match {
        case flag :: idxStr :: asType :: tail if condFlags.contains(flag) =>
          val index = try { idxStr.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index: '$idxStr'. Expected an integer.")
          }
          val validator = try {
            ConfigReader.resolveTrackValidator(asType)
          } catch {
            case e: IllegalArgumentException => return Left(e.getMessage)
          }
          result += TrackCondition(index, validator)
          remaining = tail
        case flag :: _ if condFlags.contains(flag) =>
          return Left(s"$flag requires 2 arguments: <column-index> <type>")
        case other :: _ =>
          return Left(s"Unexpected argument: '$other'. Expected -cond or -d.")
        case Nil =>
      }
    }

    Right(result.result())
  }

  private def parseComparisonsBlock(args: List[String]): Either[String, ComparePolicy] = {
    val (sheetNameOrDefault, rest) = parseSheetName(args) match {
      case Left(err) => return Left(s"--comparisons: $err")
      case Right(v)  => v
    }

    val selector = parseSheetSelector(sheetNameOrDefault)

    val ignoreColumns = parseIgnoreColumns(rest) match {
      case Left(err) => return Left(s"--comparisons: $err")
      case Right(v)  => v
    }

    if (ignoreColumns.isEmpty) {
      return Left("--comparisons: at least one column index after -ic is required.")
    }

    Right(ComparePolicy(selector, ignoreColumns))
  }

  private def parseIgnoreColumns(args: List[String]): Either[String, Set[Int]] = {
    args match {
      case "-ic" :: tail =>
        if (tail.isEmpty) return Left("-ic requires at least one column index.")
        val indices = tail.map { s =>
          try { s.toInt } catch {
            case _: NumberFormatException => return Left(s"Invalid column index after -ic: '$s'. Expected an integer.")
          }
        }
        Right(indices.toSet)
      case Nil =>
        Left("Expected -ic flag.")
      case other :: _ =>
        Left(s"Unexpected argument: '$other'. Expected -ic.")
    }
  }

  private def parseSheetName(args: List[String]): Either[String, (String, List[String])] = {
    args match {
      case ("-sheet" | "-s") :: name :: tail => Right((name, tail))
      case ("-sheet" | "-s") :: Nil          => Left("Missing sheet name after -sheet/-s.")
      case _                                 => Left("Expected -sheet/-s as first argument.")
    }
  }

  private[config] def parseSheetSelector(value: String): SheetSelector = {
    if (value == "default") {
      SheetSelector.Default
    } else {
      try {
        SheetSelector.ByIndex(value.toInt)
      } catch {
        case _: NumberFormatException => SheetSelector.ByName(value)
      }
    }
  }
}
