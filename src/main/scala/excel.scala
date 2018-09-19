package com.arena.office

/** Functions for reading and writing to Excel from Scala.
  *
  * In escence, this is an Apache POI wrapper for Excel.
  * It's possible to map data to Seq[Seq[String]] or directly to a case class T through Seq[T].
  */
object excel {

  import scala.util.matching.Regex
  import scala.reflect.{ClassTag,classTag}
  import org.apache.poi.ss.usermodel.Workbook

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

  /** Type alias for an Excel row, rows are stored as Seq[String]. */
  type ExcelRow = Seq[String]
  /** Type alias for table headers, index and name are stored as a tuple.  */
  type ExcelHeader = Seq[(Int,String)]
  /** Type alias for an Excel data table, stored as Seq[ExcelRow], which translates to Seq[Seq[String]]. */
  type ExcelData = Seq[ExcelRow]



  /** Reads Excel to a (Headers, Data) pair.
    *
    * Headers are stored as index+name pairs Seq[(Int,String)].
    * Data is stored as Seq[Seq[String]].
    *
    * @param file Path to workbook.
    * @param sheet Index of sheet to read in workbook.
    * @return (Headers, Data) pair.
    */
  def readExcel(file: String, sheet: Int): (ExcelHeader,ExcelData) = {
    import org.apache.poi.ss.usermodel.{DataFormatter, WorkbookFactory, Row}
    import java.io.File
    import collection.JavaConverters._
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter = new DataFormatter()
    val file2 = new File(file)
    val workbook = WorkbookFactory.create(file2)
    val worksheet = workbook.getSheetAt(sheet).asScala
    val (headRow,dataRows) = (worksheet.head.asScala,worksheet.tail)
    val headers = for (cell<-headRow) yield (cell.getColumnIndex,prepareString(cell.getStringCellValue))
    val columns = headers.map(_._1)
    val data = for {
      row <- dataRows
      values = for {col <- columns} yield Option(formatter.formatCellValue(row.getCell(col, blankPolicy)))
      strings = values.map(e => e match { case Some(x) => x.trim; case None => throw new Error("NULL") })
    } yield strings.toSeq
    (headers.toSeq,data.toSeq)
  }



  /** Given a class T and headers, return a constructor that can be used with said headers. */
  def constructorOf[T:ClassTag](headers: ExcelHeader) = {
    val columns = headers.map(_._2).toSeq
    val constructors = classTag[T].runtimeClass.getDeclaredConstructors
    val candidates = constructors.filter(small => small.getParameters.map(_.getName).foldLeft[Boolean](true)(_ && columns.contains(_)))
    if (candidates.isEmpty) throw new Error("No valid constructor found for given class, maybe there's missing columns")
    candidates.sortWith(_.getParameterCount > _.getParameterCount).head
  }

  /** Given a case class T fetches companion class. */
  def companionOf[T](implicit tag: ClassTag[T]): Class[_] = {
    Class.forName(classTag[T](tag).runtimeClass.getName+"$", true, classTag[T](tag).runtimeClass.getClassLoader)
  }

  /** Given a class T and headers, return an apply method that can be used with said headers. */
  def applyOf[T: ClassTag](headers: ExcelHeader) = {
    val columns = headers.map(_._2).toSeq
    val candidates = companionOf[T].getMethods
      .filter(m => m.getName=="apply" && m.getParameters.map(_.getName).foldLeft[Boolean](true)(_ && columns.contains(_)))
    if (candidates.isEmpty) {
      throw new Error(
        "No valid apply method found, "+
        "there could be missing columns or "+
        "the companion object needs an apply method to implement a constructor"
      )
    } else candidates.head
  }

  /** Given a case class T and a header list (index,name) gives back an Seq 
    * of indices corresponding to case class constructor fields.
    * This is useful to map a row of data to a case class, ignoring order and extra fields.
    *
    * @return Indices of headers that correspond to class constructor parameters.
    */
  def orderedIndices[T: ClassTag](headers: ExcelHeader): Seq[Int] = {
    applyOf[T](headers).getParameters.map(field => headers.filter(pair => field.getName==pair._2))
    .map(ls => if (ls.size==1) ls.head._1 else throw new Error("Multiple fields with same name or missing field: "+ls+" vs "+headers))
  }

  /** Given the headers, transforms an Excel row to a case class T.
    *
    * Method is useful to transform ExcelRow into arbitrary case classes
    * or when one needs to validate a row before extraction.
    *
    * @tparam T Class to extract.
    * @param headers Excel headers in index+name tuples.
    * @param row Excel row in Seq[String] form.
    */
  def mapRowToClass[T: ClassTag](headers: ExcelHeader, row: ExcelRow): T = {
    val indices = orderedIndices[T](headers)
    val orderedData = indices.map(row(_))
    val companionClass = companionOf[T]
    val companionApply = applyOf[T](headers)
    val transform = companionApply.getParameters.map(_.getParameterizedType.toString).map(z => 
      if (z=="class java.lang.String") ((x:String) => x.trim)
      else if (z=="int") ((x:String) => x.toInt)
      else if (z=="double") ((x:String) => x.toDouble)
      else ((x:String) => x.trim)
    )
    if (indices.size!=transform.size) throw new Error("Apply method parameter count is not consistent with indices count")
    companionApply.invoke(companionClass.newInstance(), (orderedData zip transform).map(x => x._2(x._1).asInstanceOf[AnyRef]): _*)
      .asInstanceOf[T] 
  }

  /** Reads Excel to a Seq[T], where T is a case class.
    *
    * @tparam T Case class to extract.
    * @param file Path to workbook.
    * @param sheet Sheet index in workbook.
    * @return Excel rows as Seq[T].
    */
  def readExcelintoClass[T: ClassTag](file: String, sheet: Int): Seq[T] = {
    val (headers,data) = readExcel(file, sheet)
    val excelVars = headers.map(_._2)
    val classVars = applyOf[T](headers).getParameters.map(_.getName)
    val missingFields = classVars.filterNot(x => excelVars contains x)
    if (!missingFields.isEmpty) {
      throw new Error("Missing fields: "+missingFields.foldLeft("")((a,b) => if (a.isEmpty) b else a+","+b))
    } else data.map(mapRowToClass[T](headers,_))
  }



  /** Checks that data has same size as headers.
    *
    * @return Array of delinquent rows.
    */
  private def checkDataWithHeaders(headers: Seq[String], data: Seq[Seq[Any]]): Array[String] = {
    val hsize = headers.size
    val dsize = data.zipWithIndex.map(p => (p._1.size==hsize,p._1,p._2)).filterNot(_._1).map(p => (p._2,p._3))
    dsize.map(p => p match {
      case (row,index) => "Row "+index+" has "+row.size+" entries, should have "+headers.size
    }).toArray
  }

  /** Writes data to a new XLSX or XLS workbook.
    *
    * @param file Path to workbook, must have an .xlsx or .xls extension.
    * @param sheet Name of worksheet to save data in.
    * @param headers Headers.
    * @param data Data.
    * @param filetype Type of Excel workbook, XLSX by default. Must be XLSX or XLS (ignores case).
    * @return Writes to file.
    */
  def writeExcel(file: String, sheet: String, headers: Seq[String], data: Seq[Seq[Any]], filetype: String = "XLSX"): Unit = {
    import org.apache.poi.ss.usermodel.{DataFormatter, WorkbookFactory, Row}
    import java.io.{File,FileOutputStream}
    import collection.JavaConverters._
    import org.apache.poi.xssf.usermodel.{XSSFWorkbook,XSSFSheet,XSSFRow,XSSFCell}
    import org.apache.poi.hssf.usermodel.{HSSFWorkbook,HSSFSheet,HSSFRow,HSSFCell}
    // Checks file type
    val etype = Map("XLSX"->".xlsx","XLS"->".xls")
    val ftype = filetype.toUpperCase.trim
    val types = etype.keys.toSeq
    if (!types.contains(ftype)) throw new Error("File type must be one of the following: "+types.reduceLeft(_+", "+_))
    // Check extension
    val extension = etype(ftype)
    if (!file.endsWith(extension)) throw new Error ("File type "+ftype+" must have "+extension+" extension")
    // Start writting to Excel
    val workbook  = {
      if (filetype=="XLSX") new XSSFWorkbook() 
      else if (filetype=="XLS") new HSSFWorkbook()
      else throw new Error("Invalid file type "+ftype)
    }
    writeToWorkbook(workbook, sheet, headers, data)
    // Save to file
    val fileOut = new FileOutputStream(file)
    try {workbook.write(fileOut)} finally {fileOut.close()}
  }

  /** Writes data to existing Excel workbook. Autodetects workbook type.
    *
    * @param file Path to workbook.
    * @param sheet Name of new worksheet.
    * @param headers Headers.
    * @param data Data.
    * @return Writes to existing file.
    */
  def writeExcelNewSheet(file: String, sheet: String, headers: Seq[String], data: Seq[Seq[Any]]): Unit = {
    import org.apache.poi.ss.usermodel.{DataFormatter, WorkbookFactory, Row}
    import java.io.{File,FileInputStream,FileOutputStream}
    // Start writting to Excel
    val fileIn    = new FileInputStream(file)
    val workbook  = try {WorkbookFactory.create(fileIn)} finally {fileIn.close()}
    writeToWorkbook(workbook, sheet, headers, data)
    // Save to file
    val fileOut   = new FileOutputStream(file)
    try {workbook.write(fileOut)} finally {fileOut.close()}
  }

  /** Writes to Excel workbook, allows abstraction of XLSX and XLS workbook types. */
  private def writeToWorkbook(workbook: Workbook, sheet: String, headers: Seq[String], data: Seq[Seq[Any]]): Unit = {
    import org.apache.poi.ss.usermodel.{Sheet,Row,Cell}
    val worksheet = workbook.createSheet(sheet)
    // Check if rows are the right size
    val checkRows = checkDataWithHeaders(headers,data)
    if (!checkRows.isEmpty) {
      throw new Error(checkRows.reduceLeft(_+"\n"+_))
    }
    // Write headers
    def headWrite() = {
      val row = worksheet.createRow(0)
      for (c <- headers.indices) {
        val cell = row.createCell(c)
        cell.setCellValue(headers(c))
      }
    }
    // Write data
    def dataWrite() = {
      for (r <- data.indices) {
        val row = worksheet.createRow(r+1)
        for (c <- headers.indices) {
          val cell = row.createCell(c)
          cell.setCellValue(data(r)(c).toString)
        }
      }
    }
    // Perform the writing
    headWrite()
    dataWrite()
  }

}
