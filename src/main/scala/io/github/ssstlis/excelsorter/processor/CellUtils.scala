package io.github.ssstlis.excelsorter.processor

import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig
import io.github.ssstlis.excelsorter.model.{CellDiff, ColumnMapping, Mapping}

import java.io.{FileInputStream, FileOutputStream}
import java.nio.file.{Files, Paths, StandardCopyOption}
import java.text.Normalizer
import java.time.format.DateTimeFormatter

import org.apache.poi.ss.usermodel._

import scala.collection.mutable
import scala.util.Try

object CellUtils {

  def getCellValueAsString(cell: Cell): String = {
    cell.getCellType match {
      case CellType.NUMERIC =>
        if (DateUtil.isCellDateFormatted(cell)) {
          Try {
            val date = cell.getLocalDateTimeCellValue.toLocalDate
            date.format(DateTimeFormatter.ISO_LOCAL_DATE)
          }.getOrElse {
            val date = cell.getDateCellValue
            date.toString
          }
        } else {
          val num = cell.getNumericCellValue
          if (num == num.toLong) num.toLong.toString
          else num.toString
        }
      case CellType.STRING  => cell.getStringCellValue
      case CellType.BOOLEAN => cell.getBooleanCellValue.toString
      case CellType.FORMULA =>
        Try {
          val evaluator = cell.getSheet.getWorkbook.getCreationHelper.createFormulaEvaluator()
          val evaluated = evaluator.evaluate(cell)
          evaluated.getCellType match {
            case CellType.NUMERIC => evaluated.getNumberValue.toString
            case CellType.STRING  => evaluated.getStringValue
            case CellType.BOOLEAN => evaluated.getBooleanValue.toString
            case _                => ""
          }
        }.getOrElse(cell.getCellFormula)
      case CellType.BLANK => ""
      case _              => ""
    }
  }

  def getRowCellValue(row: Row, colIndex: Int): String = {
    Option(row.getCell(colIndex)).map(getCellValueAsString).getOrElse("")
  }

  def loadWorkbook(path: String): Workbook = {
    val fis = new FileInputStream(path)
    try {
      WorkbookFactory.create(fis)
    } finally {
      fis.close()
    }
  }

  def writeWorkbook(workbook: Workbook, path: String): Unit = {
    val fos = new FileOutputStream(path)
    try {
      workbook.write(fos)
    } finally {
      fos.close()
    }
  }

  def copyFile(sourcePath: String, destPath: String): Unit = {
    Files.copy(Paths.get(sourcePath), Paths.get(destPath), StandardCopyOption.REPLACE_EXISTING)
  }

  def findCellDiffsMapped(
    oldRow: Row,
    newRow: Row,
    mapping: Mapping = Mapping.identity,
    headerNames: Map[Int, String] = Map.empty,
    ignoredColumns: Set[Int] = Set.empty
  ): List[CellDiff] = {
    val pairs: List[(Int, Int)] = mapping match {
      case Mapping.Identity =>
        val maxCols = math.max(oldRow.getLastCellNum.toInt, newRow.getLastCellNum.toInt)
        (0 until maxCols).map(i => i -> i).toList
      case Mapping.Explicit(ps) => ps
    }
    pairs.flatMap { case (oldIdx, newIdx) =>
      if (ignoredColumns.contains(oldIdx)) {
        None
      } else {
        val oldVal = getRowCellValue(oldRow, oldIdx)
        val newVal = getRowCellValue(newRow, newIdx)
        if (oldVal == newVal) None
        else Some(CellDiff(headerNames.getOrElse(oldIdx, s"Column $oldIdx"), oldIdx, newIdx, oldVal, newVal))
      }
    }
  }

  def normalizeHeader(s: String): String = {
    val normalized = Normalizer.normalize(s, Normalizer.Form.NFKC)
    normalized.replaceAll("[\\p{Cf}]", "").trim
  }

  def extractHeaders(row: Row): List[(Int, String)] = {
    val lastCell = row.getLastCellNum
    (0 until lastCell).map { i =>
      i -> normalizeHeader(Option(row.getCell(i)).map(getCellValueAsString).getOrElse(""))
    }.toList
  }

  //format: off
  def buildColumnMapping(
    oldSheet: Sheet, newSheet: Sheet,
    oldDataStartIdx: Int, newDataStartIdx: Int,
    sheetName: String,
    sortConfigsMap: Map[String, SheetSortingConfig] = Map.empty
  ): ColumnMapping = {
  //format: on
    val oldHeaderRowIdx = oldDataStartIdx - 1
    val newHeaderRowIdx = newDataStartIdx - 1

    val oldHeaderRow = if (oldHeaderRowIdx >= 0) Option(oldSheet.getRow(oldHeaderRowIdx)) else None
    val newHeaderRow = if (newHeaderRowIdx >= 0) Option(newSheet.getRow(newHeaderRowIdx)) else None

    (oldHeaderRow, newHeaderRow) match {
      case (Some(ohr), Some(nhr)) =>
        val oldHeaders = extractHeaders(ohr)
        val newHeaders = extractHeaders(nhr)

        val newHeadersByName = mutable.Map[String, mutable.Queue[Int]]()
        newHeaders.foreach { case (idx, name) =>
          newHeadersByName.getOrElseUpdate(name, mutable.Queue[Int]()) += idx
        }

        val usedNewIndices = mutable.Set[Int]()
        val commonColumns  = mutable.ListBuffer[(Int, Int)]()
        val oldOnlyCols    = mutable.ListBuffer[(Int, String)]()

        oldHeaders.foreach { case (oldIdx, name) =>
          newHeadersByName.get(name).flatMap(q => if (q.nonEmpty) Some(q.dequeue()) else None) match {
            case Some(newIdx) =>
              commonColumns += (oldIdx -> newIdx)
              usedNewIndices += newIdx
            case None =>
              oldOnlyCols += (oldIdx -> name)
          }
        }

        val newOnlyCols = newHeaders.filterNot { case (idx, _) => usedNewIndices.contains(idx) }

        val sortColumnIndices = sortConfigsMap.get(sheetName) match {
          case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
          case _                                     => List(0)
        }

        val oldKeyIndices = sortColumnIndices
        val newKeyIndices = sortColumnIndices.map { oldIdx =>
          commonColumns.find(_._1 == oldIdx).map(_._2).getOrElse(oldIdx)
        }

        ColumnMapping(commonColumns.toList, oldOnlyCols.toList, newOnlyCols, oldKeyIndices, newKeyIndices)

      case _ =>
        val oldLastCell = Option(oldSheet.getRow(oldDataStartIdx)).map(_.getLastCellNum.toInt).getOrElse(0)
        val newLastCell = Option(newSheet.getRow(newDataStartIdx)).map(_.getLastCellNum.toInt).getOrElse(0)
        val maxCols     = math.max(oldLastCell, newLastCell)
        val positional  = (0 until maxCols).map(i => (i, i)).toList

        val sortColumnIndices = sortConfigsMap.get(sheetName) match {
          case Some(cfg) if cfg.sortConfigs.nonEmpty => cfg.sortConfigs.map(_.columnIndex)
          case _                                     => List(0)
        }

        ColumnMapping(positional, Nil, Nil, sortColumnIndices, sortColumnIndices)
    }
  }
}
