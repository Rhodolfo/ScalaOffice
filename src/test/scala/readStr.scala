import org.scalatest.FunSuite

class ExcelReadStr extends FunSuite {

  import com.arena.office.excel.readExcelintoClass
  import com.arena.testing.auxRead.{ReadStr,dataStr}

  val data = readExcelintoClass[ReadStr]("src/test/resources/tipos.xlsx", 0)

  test("Rows must equal preset values") {
    assert(data.size === 5)
    assert(data(0) === dataStr(0))
    assert(data(1) === dataStr(1))
    assert(data(2) === dataStr(2))
    assert(data(3) === dataStr(3))
    assert(data(4) === dataStr(4))
  }

}
