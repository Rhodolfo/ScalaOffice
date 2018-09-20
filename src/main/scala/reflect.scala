package com.arena.office

/** Contains auxiliary functions having to do with extracting class metadata and data mapping to case classes. */
object reflect {

  import scala.reflect.{ClassTag,classTag}
  import com.arena.office.excel.{ExcelHeader,ExcelRow}

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

}
