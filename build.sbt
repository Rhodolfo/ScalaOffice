lazy val settings = Seq(
  organization := "com.arena",
  version := "1.0",
  scalaVersion := "2.12.6",
  scalacOptions += "-deprecation",
  scalacOptions += "-unchecked",
  fork := true
)

lazy val libs = Seq(
  libraryDependencies += "org.scalatest" %% "scalatest" % "3.0.1" % "test",
  libraryDependencies += "org.apache.poi" % "poi" % "3.17",
  libraryDependencies += "org.apache.poi" % "poi-ooxml" % "3.17"
)

lazy val office = (project in file("."))
  .settings(settings: _*)
  .settings(libs: _*)
