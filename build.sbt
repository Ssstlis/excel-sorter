import sbt.librarymanagement.ModuleID

Global / onChangedBuildSource := ReloadOnSourceChanges
ThisBuild / parallelExecution := false

ThisBuild / homepage := Some(url("https://github.com/Ssstlis/excel-sorter"))
ThisBuild / description := "Excel sorting and compare util."
ThisBuild / licenses += ("Apache-2.0", url("http://www.apache.org/licenses/LICENSE-2.0"))
ThisBuild / developers := List(
  Developer("Ssstlis", "Ivan Aristov", "ssstlis@pm.me", url("https://github.com/ssstlis"))
)
ThisBuild / scmInfo := Some(
  ScmInfo(
    url("https://github.com/Ssstlis/excel-sorter"),
    "git@github.com:Ssstlis/excel-sorter.git"
  )
)

ThisBuild / version := {
  val envVersion = sys.env.getOrElse("APP_VERSION", "-SNAPSHOT")
  if (envVersion.endsWith("SNAPSHOT")) {
    git.gitHeadCommit.value.getOrElse("").take(8) + envVersion
  } else envVersion
}
ThisBuild / versionScheme := Some("pvp")

ThisBuild / scalaVersion := "2.13.18"

lazy val buildSettings = Seq(
  organization := "io.github.ssstlis",
  scalaVersion := "2.13.18",
  scalacOptions := Seq(
    "-g:vars",
    "-unchecked",
    "-deprecation",
    "-feature",
    "-language:implicitConversions",
    "-language:existentials",
    "-language:higherKinds",
    "-language:postfixOps",
    "-Xlint:infer-any",
    "-Xlint:missing-interpolator",
    "-Wunused:imports",
    "-encoding", "utf8")
)

lazy val `excel-sorter` = project.in(file("."))
  .enablePlugins(JavaAppPackaging, UniversalPlugin)
  .settings(buildSettings: _*)
  .settings(
    name := "excel-sorter",
    Compile / mainClass := Some("io.github.ssstlis.excelsorter.ExcelSorterApp"),
    libraryDependencies ++= Seq(
      "org.apache.poi" % "poi" % "5.5.1",
      "org.apache.poi" % "poi-ooxml" % "5.5.1",
      "com.typesafe" % "config" % "1.4.5",
      "org.typelevel" %% "cats-core" % "2.9.0",
      "org.scalatest" %% "scalatest" % "3.2.19" % Test
    )
  )