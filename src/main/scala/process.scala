package com.arena.office

/** Functions for preprocessing data. */
object process {

  import scala.util.matching.Regex

  private def spaceUnderscore(string: String): String = {
    def aux(s: String, l: List[String]): String = {
      if (l.isEmpty) s
      else aux(s.replaceAll(l.head,"_"), l.tail)
    }
    val list = List("\\s{1,}","/","-")
    aux(string, list)
  }
  private def removeSpecial(string: String): String = string.replaceAll("\\W","")
  private def normalizeString(string: String): String = {
    import java.text.Normalizer._
    normalize(string,Form.NFD).replaceAll("\\p{InCombiningDiacriticalMarks}+","").toLowerCase
  }

  /** Removes spcaces, diacritical marks and non-word characters  (\W in RegEx) */
  def prepareString(string: String): String = removeSpecial(spaceUnderscore(normalizeString(string.trim)))

}
