import org.scalatest.FunSuite

class ReadExcelClassMix extends FunSuite {

  import com.arena.office.excel.readExcelintoClass
  import com.arena.testing.auxRead.{ReadMix,dataMix}

  val rawMix = readExcelintoClass[ReadMix]("src/test/resources/tipos.xlsx", 0)

  test("Rows must equal preset values, detecting types") {
    assert(rawMix.size === 5)
    assert(rawMix(0) === dataMix(0))
    assert(rawMix(1) === dataMix(1))
    assert(rawMix(2) === dataMix(2))
    assert(rawMix(3) === dataMix(3))
    assert(rawMix(4) === dataMix(4))
  }

}
