import sbt.{SettingKey, taskKey}

import java.time.Instant

object Key {
  val buildBranch = SettingKey[String]("buildBranch", "Git branch.")
  val buildCommit = SettingKey[String]("buildCommit", "Git commit.")
  val buildNumber = SettingKey[String]("buildNumber", "Project current build version")
  val buildTime   = taskKey[Instant]("Time of this build")
}
