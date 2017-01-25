lazy val root = (project in file(".")).
    settings(
        name := "hello",
        version := "1.0",
        scalaVersion := "2.11.6",
        ensimeScalaVersion in ThisBuild := "2.11.6",
        libraryDependencies ++= Seq(
            "org.apache.poi" % "poi" % "3.9",
            "org.apache.poi" % "poi-ooxml" % "3.9"
        )
    )
