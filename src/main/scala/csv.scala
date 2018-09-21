package com.arena.office

/** Functions for reading Excel generated CSV from Scala.
  *
  * This is meant for large files, just use Totoshi's CSV reader directly if it's a small file.
  */
object csv {

  import scala.reflect.ClassTag
  import com.github.tototoshi.csv.CSVReader
  import com.arena.office.process.prepareString
  import com.arena.office.reflect.mapRowToClass

  /** Reads CSV to a Seq[T], where T is a case class.
    *
    * @tparam T Case class to extract.
    * @param file Path to workbook.
    * @param condition Condition to add an isntance of T to Seq.
    * @return Excel rows as Seq[T].
    */
  def readCSVintoClass[T:ClassTag](file: String, condition: T => Boolean): Seq[T] = {
    import scala.collection.mutable.ArrayBuffer
    val reader = CSVReader.open(file)
    val iterator = reader.iterator
    val headers = if (iterator.hasNext) {
      iterator.next.map(prepareString) match {
        case (s: Seq[String]) => ((0 until s.size) zip s).filterNot(_._2.isEmpty)
      }
    } else throw new Error("Empty file: "+file)
    val buffer = ArrayBuffer[T]()
    while (iterator.hasNext) {
      val row = mapRowToClass[T](headers, iterator.next)
      if (condition(row)) buffer += row
    }
    buffer.toArray.toSeq
  }

}
