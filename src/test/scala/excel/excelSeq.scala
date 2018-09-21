import org.scalatest.FunSuite

class ReadExcelSeq extends FunSuite {

  import com.arena.office.excel.readExcel
  import com.arena.testing.auxRead.{cadenaCol, enteroCol, realesCol}

  val (headers,data) = readExcel("src/test/resources/tipos.xlsx", 0)

  test("Headers must be correct") {
    assert(headers.size === 3)
    assert(headers(0) === (0, "cadena"))
    assert(headers(1) === (1, "entero"))
    assert(headers(2) === (2, "doble"))
  }

  test("Data of Cadena column") {
    (data.map(_(0)) zip cadenaCol).foreach(p => assert(p._1===p._2))
  }

  test("Data of Entero column") {
    (data.map(x => Integer.parseInt(x(1))) zip enteroCol).foreach(p => assert(p._1===p._2))
  }

  test("Data of Doble column") {
    (data.map(_(2).trim.toDouble) zip realesCol).foreach(p => assert(p._1===p._2))
  }

}
