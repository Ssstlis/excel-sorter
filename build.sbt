ThisBuild / version := "0.1.0"

ThisBuild / scalaVersion := "2.13.18"

ThisBuild / parallelExecution := false

ThisBuild / Compile / packageSrc / publishArtifact  := false

lazy val buildSettings = Seq(
  organization := "io.github.ssstlis",
  scalaVersion := "2.13.18",
  scalacOptions := Seq(
    "-g:vars",
    "-unchecked",
    "-deprecation",
    //    "-Ymacro-debug-lite",
    //    "-Xprint:typer",          // Print AST after typer phase (shows macro expansions)
    //    "-Yshow-trees-compact",   // More compact tree output
    //    "-Yshow-trees-stringified", // Shows as Scala code when possible
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
      "com.typesafe" % "config" % "1.4.3",
      "org.scalatest" %% "scalatest" % "3.2.19" % Test
    )
  )