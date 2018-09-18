package com.arena.testing

object auxRead { 

  case class ReadStr(cadena: String, entero: String, doble: String) 
  case class ReadMix(cadena: String, entero: Int   , doble: Double) 

  val cadenaCol: List[String] = List("abc","dce","asd","qqq","tru")
  val enteroCol: List[Int] = List(2, 4, 6, 8, 10)
  val realesCol: List[Double] = List(2.2, 4.4, 6.8, 9, 11.2)

  import scala.util.matching.Regex
  def removeDot(s: String): String = "\\.0".r.replaceAllIn(s,"")

  val dataStr = (cadenaCol, enteroCol.map(_.toString), realesCol.map(x => removeDot(x.toString))).zipped
    .toList.map(x => ReadStr(x._1,x._2,x._3))
  val dataMix = (cadenaCol, enteroCol, realesCol).zipped
    .toList.map(x => ReadMix(x._1,x._2,x._3))

}
