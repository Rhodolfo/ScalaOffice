import org.scalatest.FunSuite

class ReadExcelClassStream extends FunSuite {

  import com.arena.office.excel.{readExcel,readExcelintoClass}
  import com.arena.testing.auxRead.{ReadMix,cadenaCol,enteroCol,realesCol,dataMix}

  val (rawHead,rawStr) = readExcel("src/test/resources/tipos.xlsx", 0, method="stream")
  val rawMix = readExcelintoClass[ReadMix]("src/test/resources/tipos.xlsx", 0, method="stream")

  test("Headers must be correct") {
    assert(rawHead.size === 3)
    assert(rawHead(0) === (0, "cadena"))
    assert(rawHead(1) === (1, "entero"))
    assert(rawHead(2) === (2, "doble"))
  }

  test("Data of Cadena column") {
    (rawStr.map(_(0)) zip cadenaCol).foreach(p => assert(p._1===p._2))
  }

  test("Data of Entero column") {
    (rawStr.map(x => Integer.parseInt(x(1))) zip enteroCol).foreach(p => assert(p._1===p._2))
  }

  test("Data of Doble column") {
    (rawStr.map(_(2).trim.toDouble) zip realesCol).foreach(p => assert(p._1===p._2))
  }

  test("Class instances must equal preset values, detecting types") {
    assert(rawMix.size === 5)
    assert(rawMix(0) === dataMix(0))
    assert(rawMix(1) === dataMix(1))
    assert(rawMix(2) === dataMix(2))
    assert(rawMix(3) === dataMix(3))
    assert(rawMix(4) === dataMix(4))
  }

}
