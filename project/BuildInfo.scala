import Key.*
import com.typesafe.sbt.SbtGit.git
import sbt.Keys.{name, version}
import sbt.Project
import sbtbuildinfo.BuildInfoKeys.{buildInfoKeys, buildInfoOptions, buildInfoPackage}
import sbtbuildinfo.{BuildInfoKey, BuildInfoOption, BuildInfoPlugin}

import java.time.Instant

object BuildInfo {

  implicit class BuildInfoOps(val project: Project) extends AnyVal {
    def withBuildInfo: Project =
      project
        .enablePlugins(BuildInfoPlugin)
        .settings(
          buildCommit := git.gitHeadCommit.value.getOrElse("unknown"),
          buildBranch := git.gitCurrentBranch.value,
//          buildTime := Instant.now,
          buildNumber := sys.props.getOrElse("BUILD_NUMBER", "0"),
          buildInfoKeys := {
            Seq[BuildInfoKey](name, version, buildCommit, buildBranch, buildTime, buildNumber)
          },
          buildInfoPackage := "io.github.ssstlis." + name.value.replace('-', '_'),
//          wrong until 0.10.0, written `buildTime := Instant.now` to substitute
          buildInfoOptions += BuildInfoOption.BuildTime
        )
  }
}