import org.scalatest.FunSuite

class ReadExcelResource extends FunSuite {

  import com.arena.office.excel.{readExcel,readExcelintoClass}
  import com.arena.testing.auxRead.{ReadMix,dataMix}
  import com.arena.testing.auxRead.{cadenaCol,enteroCol,realesCol}

  // Prueba de readExcel
  val (headers,data) = readExcel("/tipos.xlsx", 0, fromResource=true)
  test("Headers must be correct for readExcel") {
    assert(headers.size === 3)
    assert(headers(0) === (0, "cadena"))
    assert(headers(1) === (1, "entero"))
    assert(headers(2) === (2, "doble"))
  }
  test("Data of Cadena column for readExcel") {
    (data.map(_(0)) zip cadenaCol).foreach(p => assert(p._1===p._2))
  }
  test("Data of Entero column for readExcel") {
    (data.map(x => Integer.parseInt(x(1))) zip enteroCol).foreach(p => assert(p._1===p._2))
  }
  test("Data of Doble column for readExcel") {
    (data.map(_(2).trim.toDouble) zip realesCol).foreach(p => assert(p._1===p._2))
  }

  // Prueba de readExcelintoClass
  val rawMix = readExcelintoClass[ReadMix]("/tipos.xlsx", 0, fromResource=true)
  test("Rows must equal preset values for readExcelintoClass") {
    assert(rawMix.size === 5)
    assert(rawMix(0) === dataMix(0))
    assert(rawMix(1) === dataMix(1))
    assert(rawMix(2) === dataMix(2))
    assert(rawMix(3) === dataMix(3))
    assert(rawMix(4) === dataMix(4))
  }


}
