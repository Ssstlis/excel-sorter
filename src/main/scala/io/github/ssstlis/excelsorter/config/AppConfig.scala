package io.github.ssstlis.excelsorter.config

import com.typesafe.config.Config
import io.github.ssstlis.excelsorter.config.compare.CompareConfig
import io.github.ssstlis.excelsorter.config.track.TrackConfig
import io.github.ssstlis.excelsorter.config.sorting.SheetSortingConfig

case class AppConfig(sortConfig: List[SheetSortingConfig], trackConfig: TrackConfig, compareConfig: CompareConfig)

object AppConfig {
  def readConfigFromFile(config: Config): Either[String, AppConfig] = {
    val errBuilder = List.newBuilder[String]

    val sortConfig = SheetSortingConfig.readSortConfig(config).fold({
      err => errBuilder += err; None;
    }, Some(_))

    val trackConfig = TrackConfig.readTrackConfig(config).fold({
      err => errBuilder += err; None;
    }, Some(_))

    val compareConfig = CompareConfig.readCompareConfig(config).fold({
      err => errBuilder += err; None;
    }, Some(_))

    (for {
      s <- sortConfig
      t <- trackConfig
      c <- compareConfig
    } yield AppConfig(s, t, c))
      .toRight {
        val errors = errBuilder.result()
        val header = if (errors.lengthCompare(2) < 0) {
          "There is one error while loading config from file:"
        } else {
          "There is few errors while loading config from file:"
        }
        header + errors.mkString("\n - ", "\n - ", "")
      }
  }
}
