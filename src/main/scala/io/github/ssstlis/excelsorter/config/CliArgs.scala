package io.github.ssstlis.excelsorter.config

import cats.syntax.either._
import io.github.ssstlis.excelsorter.config.compare.{CompareConfig, ComparePolicy}
import io.github.ssstlis.excelsorter.config.track.{TrackConfig, TrackPolicy}
import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig

sealed trait RunMode

object RunMode {
  case object SortOnly extends RunMode

  case object Cut extends RunMode

  case object Compare extends RunMode
}

case class CliArgs(mode: RunMode, filePaths: Seq[String], appConfig: Option[AppConfig])

object CliArgs {

  private val cutFlags = Set("--cut", "-c")
  private val compareFlags = Set("--compare", "-cmp")
  private val helpFlags = Set("-h", "--help")
  private val modeFlags = cutFlags ++ compareFlags
  private val confFlag = "--conf"
  private val blockStarters = Set("--sortings", "--tracks", "--comparisons")

  def parse(args: Array[String]): Either[String, CliArgs] = {
    val argList = args.toList

    if (argList.exists(helpFlags.contains)) {
      Left("help")
    } else {

      val confIndex = argList.indexOf(confFlag)
      val (mainArgs, confArgs) = if (confIndex >= 0) {
        (argList.take(confIndex), argList.drop(confIndex + 1))
      } else {
        (argList, Nil)
      }

      val (flags, files) = mainArgs.partition(_.startsWith("-"))

      val unknownFlags = flags.filterNot(modeFlags.contains)
      if (unknownFlags.nonEmpty) {
        Left(s"Unknown flag(s): ${unknownFlags.mkString(", ")}. Supported: --cut/-c, --compare/-cmp, --conf, -h/--help")
      } else {

        val hasCut = flags.exists(cutFlags.contains)
        val hasCompare = flags.exists(compareFlags.contains)

        if (hasCut && hasCompare) {
          Left("Flags --cut and --compare are mutually exclusive. Use only one.")
        } else {

          if (files.isEmpty) {
            Left("No input files specified.")
          } else {

            val mode = {
              if (hasCut) RunMode.Cut
              else if (hasCompare) RunMode.Compare
              else RunMode.SortOnly
            }

            val args = if (confArgs.nonEmpty) {
              parseConfSection(confArgs).map(Some(_))
            } else {
              Right(None)
            }

            args.map(CliArgs(mode, files, _))
          }
        }
      }
    }
  }

  private def parseConfSection(args: List[String]): Either[String, AppConfig] = {
    val blocks = splitIntoBlocks(args)

    if (blocks.isEmpty) {
      Left(s"--conf requires at least one configuration block (${blockStarters.mkString(", ")}).")
    } else {
      blocks.foldLeft(Either.right[String, (List[SheetSortingConfig], List[TrackPolicy], List[ComparePolicy])](Nil, Nil, Nil)) {
        case (acc, (blockType, blockArgs)) =>
          acc.flatMap { case acc@(sortings, trackPolicies, comparePolicies) =>
            blockType match {
              case "--sortings" =>
                SheetSortingConfig.parseSortingsBlock(blockArgs).map { cfg =>
                  acc.copy(_1 = cfg :: sortings)
                }
              case "--tracks" =>
                TrackPolicy.parseTracksBlock(blockArgs).map { policy =>
                  acc.copy(_2 = policy :: trackPolicies)
                }
              case "--comparisons" =>
                ComparePolicy.parseComparisonsBlock(blockArgs).map { policy =>
                  acc.copy(_3 = policy :: comparePolicies)
                }
              case other => Left(s"Unknown config block type: '$other'. Expected one of (${blockStarters.mkString("'", "', '", "'")}).")
            }
          }
      }.map { case (sortings, trackPolicies, comparePolicies) =>
        AppConfig(
          sortings.reverse,
          TrackConfig(trackPolicies.reverse),
          CompareConfig(comparePolicies.reverse)
        )
      }
    }
  }

  private def splitIntoBlocks(args: List[String]): List[(String, List[String])] = {
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

  def parseSheetName(args: List[String]): Either[String, (String, List[String])] = {
    args match {
      case ("-sheet" | "-s") :: name :: tail => Right((name, tail))
      case ("-sheet" | "-s") :: Nil          => Left("Missing sheet name after -sheet/-s.")
      case _                                 => Left("Expected -sheet/-s as first argument.")
    }
  }
}
