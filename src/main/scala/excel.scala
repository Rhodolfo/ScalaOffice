package com.arena.office

/** Functions for reading and writing to Excel from Scala.
  *
  * In escence, this is an Apache POI wrapper for Excel.
  * It's possible to map data to Seq[Seq[String]] or directly to a case class T through Seq[T].
  */
object excel {

  import scala.util.matching.Regex
  import scala.reflect.{ClassTag,classTag}
  import org.apache.poi.ss.usermodel.{Workbook,Row, Cell}
  import com.arena.office.process.prepareColumn
  import com.arena.office.reflect.{applyOf,mapRowToClass}

  /** Type alias for an Excel row, rows are stored as Seq[String]. */
  type ExcelRow = Seq[String]
  /** Type alias for table headers, index and name are stored as a tuple.  */
  type ExcelHeader = Seq[(Int,String)]
  /** Type alias for an Excel data table, stored as Seq[ExcelRow], which translates to Seq[Seq[String]]. */
  type ExcelData = Seq[ExcelRow]

  private def extractString(cell: Cell): String = {
    def sqlDate(date: java.util.Date): Int = {
      import java.util.Calendar
      val calendar = Calendar.getInstance
      calendar.setTime(date)
      val year  = calendar.get(Calendar.YEAR)
      val month = calendar.get(Calendar.MONTH)+1
      val day   = calendar.get(Calendar.DAY_OF_MONTH)
      year*10000+month*100+day
    }
    import org.apache.poi.ss.usermodel.{CellType, DateUtil}
    val t = cell.getCellTypeEnum()
    def intIt(s: String): String = "\\.0+?$".r.replaceAllIn(s,"")
    if (t == CellType.STRING) cell.getStringCellValue()
    else if (t == CellType.NUMERIC) {
      if (DateUtil.isCellDateFormatted(cell)) sqlDate(cell.getDateCellValue()).toString
      else intIt(cell.getNumericCellValue().toString)
    }
    else cell.getStringCellValue()
  }

  private def getWorkbook(file: String, method: String, fromResource: Boolean): Workbook = {
    import java.io.File
    import org.apache.poi.ss.usermodel.WorkbookFactory
    import java.io.FileInputStream
    import com.monitorjbl.xlsx.StreamingReader
    method match {
      case "default" => {
        val fileObj = {
          if (fromResource) new File(getClass.getResource(file).getFile())
          else new File(file)
        }
        WorkbookFactory.create(fileObj)
      }
      case "stream" => {
        val streamObj = {
          if (fromResource) throw new Error("Excel streaming from resources is unsupported")
          else new FileInputStream(file)
        }
        StreamingReader.builder()
          .rowCacheSize(100)
          .bufferSize(4096)
          .open(streamObj)
      }
      case _ => throw new Error("Unsupported method for getWorkbook")
    }
  }

  private def getIterator(file: String, sheet: Int, method: String, fromResource: Boolean): Iterator[Row] = {
    import collection.JavaConverters._
    getWorkbook(file, method, fromResource).getSheetAt(sheet).iterator.asScala
  }

  private def getHeaders(file: String, sheet: Int, iterator: Iterator[Row]): ExcelHeader = {
    import collection.JavaConverters._
    val headRow = {
      if (iterator.hasNext) iterator.next.asScala
      else throw new Error("Empty "+sheet+" sheet: "+file)
    }
    (for (cell<-headRow) yield (cell.getColumnIndex,prepareColumn(extractString(cell)))).toSeq
  }

  

  /** Sometimes the max index for the colection of Excel sheets in a workbook is needed.
    *
    * @param file Path to workbook.
    * @param method "default" loads file into memory, "stream" to process large files as a stream
    * @return Max sheet index.
    */
  def readExcelSheetMaxIndex(file: String, method: String): Int = {
    import collection.JavaConverters._
    import org.apache.poi.ss.usermodel.Sheet
    val workbook = getWorkbook(file, method, false)
    def aux(iterator: Iterator[Sheet], index: Int): Int = {
      if (iterator.hasNext) {
        iterator.next 
        aux(iterator, index+1)
      } else index-1
    }
    aux(workbook.iterator.asScala, 0)
  }



  /** Reads Excel to a Seq[String].
    *
    * @param file Path to workbook.
    * @param sheet Index of sheet to read in workbook.
    * @return Data from Excel file.
    */
  def readExcelRaw(file: String, sheet: Int): Seq[Seq[String]] = {
    import org.apache.poi.ss.usermodel.DataFormatter
    import collection.JavaConverters._
    // Formats
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter   = new DataFormatter()
    // Extract the iterator for this Excel sheet, then the headers
    val iterator = getIterator(file, sheet, "default", false)
    // Iterate over and read everything
    val buffer  = scala.collection.mutable.ArrayBuffer[ExcelRow]()
    while (iterator.hasNext) {
      val row = {
        val rowObj  = iterator.next
        val indices = 0 until rowObj.getLastCellNum
        val values  = for {index <- indices} yield Option(formatter.formatCellValue(rowObj.getCell(index, blankPolicy)))
        values.map(e => e match {case Some(x) => x.trim; case None => throw new Error("NULL")}).toSeq
      }
      buffer += row
    }
    buffer.toArray.toSeq
  }



  /** Reads Excel to a (Headers, Data) pair.
    *
    * Headers are stored as index+name pairs Seq[(Int,String)].
    * Data is stored as Seq[Seq[String]].
    *
    * @param file Path to workbook.
    * @param sheet Index of sheet to read in workbook.
    * @param condition Anonymous function, rows for which this function returns false are exluded from result.
    * @param method "default" loads file into memory, "stream" to process large files as a stream
    * @param fromResource If true, pulls the file from resources instead of the regular host filesystem.
    * @return (Headers, Data) pair.
    */
  def readExcel(
    file: String, 
    sheet: Int, 
    condition: ExcelRow => Boolean = x => true, 
    method: String = "default",
    fromResource: Boolean = false
  ): (ExcelHeader,ExcelData) = {
    if (method=="default") readExcelDefault(file, sheet, condition, fromResource)
    else if (method=="stream") readExcelStream(file, sheet, condition, fromResource)
    else throw new Error("Unsupported method for readExcel")
  }

  /** Default Excel reader using Apache POI. */
  private def readExcelDefault(
    file: String, 
    sheet: Int, 
    condition: ExcelRow => Boolean, 
    fromResource: Boolean
  ): (ExcelHeader, ExcelData) = {
    import org.apache.poi.ss.usermodel.DataFormatter
    import collection.JavaConverters._
    // Formats
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter   = new DataFormatter()
    // Extract the iterator for this Excel sheet, then the headers
    val iterator = getIterator(file, sheet, "default", fromResource)
    val headers  = getHeaders(file, sheet, iterator)
    // Let's extract the data now
    val indices = headers.map(_._1)
    val buffer  = scala.collection.mutable.ArrayBuffer[ExcelRow]()
    while (iterator.hasNext) {
      val row = {
        val rowObj  = iterator.next
        val values  = for {index <- indices} yield Option(formatter.formatCellValue(rowObj.getCell(index, blankPolicy)))
        values.map(e => e match {case Some(x) => x.trim; case None => throw new Error("NULL")}).toSeq
      }
      if (condition(row)) buffer += row
    }
    (headers.toSeq,buffer.toArray.toSeq)
  }

  /** Low memory footprint Excel reader. */
  private def readExcelStream(
    file: String, 
    sheet: Int, 
    condition: ExcelRow => Boolean, 
    fromResource: Boolean
  ): (ExcelHeader,ExcelData) = {
    import org.apache.poi.ss.usermodel.DataFormatter
    import collection.JavaConverters._
    // Formats
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter   = new DataFormatter()
    // Extract the iterator for this Excel sheet, then the headers
    val iterator = getIterator(file, sheet, "stream", fromResource)
    val headers  = getHeaders(file, sheet, iterator)
    // Extracting data
    val indices = headers.map(_._1)
    val buffer  = scala.collection.mutable.ArrayBuffer[ExcelRow]()
    while (iterator.hasNext) {
      val row = {
        val rowObj: Seq[Option[String]] = iterator.next.asScala.map(cell => Some(extractString(cell))).toSeq
        val values  = for {index <- indices} yield rowObj(index)
        values.map(e => e match {case Some(x) => x.trim; case None => throw new Error("NULL")}).toSeq
      }
      if (condition(row)) buffer += row
    }
    (headers,buffer.toArray.toSeq)
  }



  /** Reads Excel to a Seq[T], where T is a case class.
    *
    * @tparam T Case class to extract.
    * @param file Path to workbook.
    * @param sheet Sheet index in workbook.
    * @param condition Anonymous function, rows for which this function returns false are exluded from result.
    * @param method "default" loads file into memory, "stream" to process large files as a stream
    * @param fromResource If true, pulls the file from resources instead of the regular host filesystem.
    * @return Excel rows as Seq[T].
    */
  def readExcelintoClass[T: ClassTag](
    file: String, 
    sheet: Int, 
    condition: T => Boolean = (x:T) => true, 
    method: String = "default",
    fromResource: Boolean = false
  ): Seq[T] = {
    if (method=="default") readExcelClassDefault[T](file, sheet, condition, fromResource)
    else if (method=="stream") readExcelClassStream[T](file, sheet, condition, fromResource)
    else throw new Error("Unsupported method for readExcelintoClass")
  }

  private def isRowNonEmpty(s: Seq[String]): Boolean = s.foldLeft[Boolean](false)((a,b) => a || !b.trim.isEmpty)

  private def readExcelClassDefault[T: ClassTag](
    file: String, 
    sheet: Int, 
    condition: T => Boolean, 
    fromResource: Boolean
  ): Seq[T] = {
    import org.apache.poi.ss.usermodel.DataFormatter
    import collection.JavaConverters._
    // Formats
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter   = new DataFormatter()
    // Extract the iterator for this Excel sheet, then the headers
    val iterator = getIterator(file, sheet, "default", fromResource)
    val headers  = getHeaders(file, sheet, iterator)
    // Checking headers with the class
    val excelVars = headers.map(_._2)
    val classVars = applyOf[T](headers).getParameters.map(_.getName)
    val missingFields = classVars.filterNot(x => excelVars contains x)
    val missingThrown = missingFields.foldLeft("")((a,b) => if (a.isEmpty) b else a+","+b)
    if (!missingFields.isEmpty) throw new Error("Missing fields: "+missingThrown)
    // Extracting data
    val indices = headers.map(_._1)
    val buffer  = scala.collection.mutable.ArrayBuffer[T]()
    while (iterator.hasNext) {
      val row = {
        val rowObj  = iterator.next
        val values  = for {index <- indices} yield Option(formatter.formatCellValue(rowObj.getCell(index, blankPolicy)))
        values.map(e => e match {case Some(x) => x.trim; case None => throw new Error("NULL")}).toSeq
      }
      if (isRowNonEmpty(row)) {
        mapRowToClass[T](headers, row) match {
          case (seq:T) => if (condition(seq)) buffer += seq
        }
      }
    }
    buffer.toArray.toSeq
  }

  private def readExcelClassStream[T: ClassTag](
    file: String, 
    sheet: Int, 
    condition: T => Boolean, 
    fromResource: Boolean
  ): Seq[T] = {
    import org.apache.poi.ss.usermodel.DataFormatter
    import collection.JavaConverters._
    // Formats
    val blankPolicy = Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
    val formatter   = new DataFormatter()
    // Extract the iterator for this Excel sheet, then the headers
    val iterator = getIterator(file, sheet, "stream", fromResource)
    val headers  = getHeaders(file, sheet, iterator)
    // Checking headers with the class
    val excelVars = headers.map(_._2)
    val classVars = applyOf[T](headers).getParameters.map(_.getName)
    val missingFields = classVars.filterNot(x => excelVars contains x)
    val missingThrown = missingFields.foldLeft("")((a,b) => if (a.isEmpty) b else a+","+b)
    if (!missingFields.isEmpty) throw new Error("Missing fields: "+missingThrown)
    // Extracting data
    val indices = headers.map(_._1)
    val buffer  = scala.collection.mutable.ArrayBuffer[T]()
    while (iterator.hasNext) {
      val row = {
        val rowObj: Seq[Option[String]] = iterator.next.asScala.map(x => Some(extractString(x))).toSeq
        val values  = for {index <- indices} yield rowObj.applyOrElse(index, (x:Int)=>Some(""))
        values.map(e => e match {case Some(x) => x.trim; case None => throw new Error("NULL")}).toSeq
      }
      if (isRowNonEmpty(row)) {
        mapRowToClass[T](headers, row) match {
          case (seq:T) => if (condition(seq)) buffer += seq
        }
      }
    }
    buffer.toArray.toSeq
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





  /** Writes data to a new XLSX or XLS workbook, no headers.
    *
    * @param file Path to workbook, must have an .xlsx or .xls extension.
    * @param sheet Name of worksheet to save data in.
    * @param data Data.
    * @param filetype Type of Excel workbook, XLSX by default. Must be XLSX or XLS (ignores case).
    * @return Writes to file.
    */
  def writeExcelRaw(file: String, sheet: String, data: Seq[Seq[Any]], filetype: String = "XLSX"): Unit = {
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
    writeToWorkbookRaw(workbook, sheet, data)
    // Save to file
    val fileOut = new FileOutputStream(file)
    try {workbook.write(fileOut)} finally {fileOut.close()}
  }

  /** Writes data to existing Excel workbook. Autodetects workbook type. No headers.
    *
    * @param file Path to workbook.
    * @param sheet Name of new worksheet.
    * @param headers Headers.
    * @param data Data.
    * @return Writes to existing file.
    */
  def writeExcelNewSheetRaw(file: String, sheet: String, data: Seq[Seq[Any]]): Unit = {
    import org.apache.poi.ss.usermodel.{DataFormatter, WorkbookFactory, Row}
    import java.io.{File,FileInputStream,FileOutputStream}
    // Start writting to Excel
    val fileIn    = new FileInputStream(file)
    val workbook  = try {WorkbookFactory.create(fileIn)} finally {fileIn.close()}
    writeToWorkbookRaw(workbook, sheet, data)
    // Save to file
    val fileOut   = new FileOutputStream(file)
    try {workbook.write(fileOut)} finally {fileOut.close()}
  }

  /** Writes to Excel workbook, allows abstraction of XLSX and XLS workbook types. */
  private def writeToWorkbookRaw(workbook: Workbook, sheet: String, data: Seq[Seq[Any]]): Unit = {
    import org.apache.poi.ss.usermodel.{Sheet,Row,Cell}
    val worksheet = workbook.createSheet(sheet)
    // Write data
    def dataWrite() = {
      for (r <- data.indices) {
        val row = worksheet.createRow(r)
        for (c <- data(r).indices) {
          val cell = row.createCell(c)
          cell.setCellValue(data(r)(c).toString)
        }
      }
    }
    // Perform the writing
    dataWrite()
  }



}
