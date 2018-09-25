package com.arena

object ventas {

  import com.arena.tiempo.Fecha

  /** Trait para ventas que tienen semana. */
  sealed trait Semanal {
    def region: String
    def canal: String
    def producto: String
    def semana: Int
    def venta: Int 
  }
  /** Trait para ventas que tienen mes. */
  sealed trait Mensual {
    def region: String
    def canal: String
    def producto: String
    def mes: Int
    def venta: Int 
  }
  /** Trait para ventas que tienen clave única. */
  sealed trait Clave {
    def cve_unica: String
  }
  /** Trait para ventas que tienen semana y clave única. */
  sealed trait SemanalClave extends Semanal with Clave
  /** Trait para ventas que tienen mes y clave única. */
  sealed trait MensualClave extends Mensual with Clave

  private def parseAsInt(s: String): Int = {
    ",".r.replaceAllIn(s,"").mkString.trim match {
      case (t: String) => {
        if (t.contains("-")) 0
        else t.toDouble.round.toInt
      }
    }
  }


  /** Case class para guardar ventas a nivel día. */
  case class Venta(region: String, canal: String, cve_unica: String, producto: String, fecha: Fecha, venta: Int) 
    extends MensualClave with SemanalClave {
    /** Mes en formato yyyyMM. */
    def mes(): Int = fecha.year*10+fecha.month 
    /** Semana en formato yyyySS. */
    def semana(): Int = fecha.yearWeek
    /** Constructor diseñado para lectura de datos. */
    def this(region: String, canal: String, cve_unica_pdv: String, producto: String, fecha_inar: String, valor_ga: String) = {
      this(region, canal, cve_unica_pdv, producto, Fecha(fecha_inar), parseAsInt(valor_ga))
    }
  }
  /** Objeto acompañante de [[Venta]]. */
  object Venta {
    /** Constructor diseñado para lectura de datos. */
    def apply(region: String, canal: String, cve_unica_pdv: String, producto: String, fecha_inar: String, valor_ga: String) = {
      new Venta(region, canal, cve_unica_pdv, producto, fecha_inar, valor_ga)
    }
  }





  /** Case class para guardar ventas a nivel semanal. */
  case class VentaSemanal(region: String, canal: String, cve_unica: String, producto: String, semana: Int, venta: Int) 
    extends SemanalClave {
    def this(region: String, canal: String, cve_unica: String, producto: String, semana: Int, ventaSeq: Seq[SemanalClave]) {
      this(region, canal, cve_unica, producto, semana, 
        ventaSeq
          .filter(v => v.region==region && v.canal==canal && v.cve_unica==cve_unica && v.producto == producto && v.semana == semana)
          .foldLeft[Int](0)((acc,venta) => acc + venta.venta)
      )
    }
  }
  /** Objeto acompañante de [[VentaSemanal]]. */
  object VentaSemanal {
    def apply(region: String, canal: String, cve_unica: String, producto: String, semana: Int, ventaSeq: Seq[SemanalClave]) {
      this(region, canal, cve_unica, producto, semana, ventaSeq) 
    }
  }





  /** Case class para guardar ventas a nivel mensual. */
  case class VentaMensual(region: String, canal: String, cve_unica: String, producto: String, mes: Int, venta: Int) 
    extends MensualClave {
    /** Filtra una colección de ventas y las suma. */
    def this(region: String, canal: String, cve_unica: String, producto: String, mes: Int, ventaSeq: Seq[MensualClave]) = {
      this(region, canal, cve_unica, producto, mes, 
        ventaSeq
          .filter(v => v.region==region && v.canal==canal && v.cve_unica==cve_unica && v.producto == producto && v.mes == mes)
          .foldLeft[Int](0)((acc,venta) => acc + venta.venta)
      )
    }
  }
  /** Objeto acompañante de [[VentaMensual]]. */
  object VentaMensual {
    /** Filtra una colección de ventas y las suma. */
    def apply(region: String, canal: String, cve_unica: String, producto: String, mes: Int, ventaSeq: Seq[MensualClave]) = {
      new VentaMensual(region, canal, cve_unica, producto, mes, ventaSeq)
    }
  }





  /** Case class para guardar ventas a nivel semanal, sumadas por todos los PDVs. */
  case class VentaSemanalAgg(region: String, canal: String, producto: String, semana: Int, venta: Int) extends Semanal {
    def this(region: String, canal: String, producto: String, semana: Int, ventas: Seq[Semanal]) = {
      this(region, canal, producto, semana, 
        ventas
          .filter(v => v.region==region && v.canal==canal && v.producto == producto && v.semana == semana)
          .foldLeft[Int](0)((acc,venta) => acc + venta.venta)
      )
    }
  }
  /** Objeto acompañante de [[VentaSemanalAgg]]. */
  object VentaSemanalAgg {
    def apply(region: String, canal: String, producto: String, semana: Int, ventas: Seq[Semanal]) = {
      new VentaSemanalAgg(region, canal, producto, semana, ventas)
    }
  }





  /** Case class para guardar ventas a nivel mensual, sumadas por todos los PDVs. */
  case class VentaMensualAgg(region: String, canal: String, producto: String, mes: Int, venta: Int) extends Mensual {
    def this(region: String, canal: String, producto: String, mes: Int, ventas: Seq[Mensual]) = {
      this(region, canal, producto, mes, 
        ventas
          .filter(v => v.region==region && v.canal==canal && v.producto == producto && v.mes == mes)
          .foldLeft[Int](0)((acc,venta) => acc + venta.venta)
      )
    }
  }
  /** Objeto acompañante de [[VentaMensualAgg]]. */
  object VentaMensualAgg {
    def apply(region: String, canal: String, producto: String, mes: Int, ventas: Seq[Mensual]) = {
      new VentaMensualAgg(region, canal, producto, mes, ventas)
    }
  }





  /** Suma sobre todos los PDVs de ventas mensuales. */
  def sumaVentaMensual(ventas: Seq[Mensual]): Seq[VentaMensualAgg] = {
    ventas.map(x => (x.region,x.canal,x.producto,x.mes)).distinct
      .map(p => p match {
        case (region, canal, producto, mes) => VentaMensualAgg(region, canal, producto, mes, ventas)
      })
  }



  /** Suma sobre todos los PDVs de ventas semanales. */
  def sumaVentaSemanal(ventas: Seq[Semanal]): Seq[VentaSemanalAgg] = {
    ventas.map(x => (x.region,x.canal,x.producto,x.semana)).distinct
      .map(p => p match {
        case (region, canal, producto, semana) => VentaSemanalAgg(region, canal, producto, semana, ventas)
      })
  }

}
