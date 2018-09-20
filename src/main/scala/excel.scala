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
  import com.arena.office.process.prepareString
  import com.arena.office.reflect.{applyOf,mapRowToClass}

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
