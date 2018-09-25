package com.arena

/** Contiene clases referentes a fechas. */
object tiempo {

  import java.util.Calendar
  import java.text.SimpleDateFormat

  /** Transforma una fecha mm/dd/yyyy a su representación entera. */
  private def sqlFecha(date: String): Int = {
    val calendar = Calendar.getInstance
    calendar.setTime((new SimpleDateFormat("yyyyMMdd")).parse(date))
    val year  = calendar.get(Calendar.YEAR)
    val month = calendar.get(Calendar.MONTH)+1
    val day   = calendar.get(Calendar.DAY_OF_MONTH)
    if (year < 100) (2000+year)*10000+month*100+day
    else year*10000+month*100+day
  }

  /** Clase que representa a una fecha. */
  case class Fecha(sqlFecha: Int) {
    def calendar = {
      val inter = Calendar.getInstance
      inter.setTime((new SimpleDateFormat("yyyyMMdd")).parse(sqlFecha.toString))
      inter
    }
    def week: Int = calendar.get(Calendar.WEEK_OF_YEAR)
    def month: Int = calendar.get(Calendar.MONTH)+1
    def year: Int = calendar.get(Calendar.YEAR)
    def yearWeek: Int = year*100+week
    def this(stringFecha: String) = this(sqlFecha(stringFecha))
  }

  /** Objecto acompañante de Fecha. */
  object Fecha {
    def apply(stringFecha: String) = new Fecha(stringFecha)
  }

}
