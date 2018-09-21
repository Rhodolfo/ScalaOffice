package com.arena.office

/** Functions for reading Excel generated CSV from Scala.
  *
  * This is meant for large files, just use Totoshi's CSV reader directly if it's a small file.
  */
object csv {

  import scala.reflect.ClassTag
  import com.arena.office.process.prepareColumn
  import com.arena.office.reflect.mapRowToClass
  import com.arena.office.excel.{ExcelHeader,ExcelRow}

  private def getIterator(file: String): Iterator[ExcelRow] = {
    com.github.tototoshi.csv.CSVReader.open(file).iterator
  }

  private def getHeader(file: String, iterator: Iterator[ExcelRow]): ExcelHeader = {
    if (iterator.hasNext) {
      iterator.next.map(prepareColumn) match {
        case (s: Seq[String]) => ((0 until s.size) zip s).filterNot(_._2.isEmpty)
      }
    } else throw new Error("Empty file: "+file)
  }





  /** Reads CSV to a (Headers, Data) pair.
    *
    * Headers are stored as index+name pairs Seq[(Int,String)].
    * Data is stored as Seq[Seq[String]].
    *
    * @param file Path to CSV file.
    * @return (Headers, Data) pair.
    */
  def readCSV(file: String, condition: ExcelRow => Boolean = x => true): (ExcelHeader,Seq[ExcelRow]) = {
    val iterator = getIterator(file)
    val headers  = getHeader(file, iterator)
    val buffer = scala.collection.mutable.ArrayBuffer[ExcelRow]()
    while (iterator.hasNext) {
      iterator.next match {
        case (row: ExcelRow) => if (condition(row)) buffer += row
      }
    }
    (headers,buffer.toArray.toSeq)
  }





  /** Reads CSV to a Seq[T], where T is a case class.
    *
    * @tparam T Case class to extract.
    * @param file Path to workbook.
    * @param condition Condition to add an isntance of T to Seq.
    * @return Excel rows as Seq[T].
    */
  def readCSVintoClass[T:ClassTag](file: String, condition: T => Boolean = (x:T) => true): Seq[T] = {
    val iterator = getIterator(file)
    val headers  = getHeader(file, iterator)
    val buffer   = scala.collection.mutable.ArrayBuffer[T]()
    while (iterator.hasNext) {
      val row = mapRowToClass[T](headers, iterator.next)
      if (condition(row)) buffer += row
    }
    buffer.toArray.toSeq
  }

}
