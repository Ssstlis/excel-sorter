package io.github.ssstlis.excelsorter.model

sealed trait Mapping extends Product with Serializable

object Mapping {
  case object Identity                               extends Mapping
  final case class Explicit(pairs: List[(Int, Int)]) extends Mapping

  val identity: Mapping = Identity

  def from(pair0: (Int, Int), pairs: (Int, Int)*): Mapping = Explicit(pair0 :: pairs.toList)
  def from(pairs: List[(Int, Int)] = Nil): Mapping         = Explicit(pairs)
}
