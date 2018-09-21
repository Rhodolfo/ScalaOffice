import org.scalatest.FunSuite

class ReadCSVClassStr extends FunSuite {

  import com.arena.office.csv.readCSVintoClass
  import com.arena.testing.auxRead.{ReadStr,dataStr,ReadTrs,dataTrs,ReadRst,dataRst}

  val rawStr = readCSVintoClass[ReadStr]("src/test/resources/tipos.csv")
  val rawTrs = readCSVintoClass[ReadTrs]("src/test/resources/tipos.csv")
  val rawRst = readCSVintoClass[ReadRst]("src/test/resources/tipos.csv")

  test("Rows must equal preset values, ABC permutation") {
    assert(rawStr.size === 5)
    assert(rawStr(0) === dataStr(0))
    assert(rawStr(1) === dataStr(1))
    assert(rawStr(2) === dataStr(2))
    assert(rawStr(3) === dataStr(3))
    assert(rawStr(4) === dataStr(4))
  }

  test("Rows must equal preset values, BCA permutation") {
    assert(rawTrs.size === 5)
    assert(rawTrs(0) === dataTrs(0))
    assert(rawTrs(1) === dataTrs(1))
    assert(rawTrs(2) === dataTrs(2))
    assert(rawTrs(3) === dataTrs(3))
    assert(rawTrs(4) === dataTrs(4))
  }

  test("Rows must equal preset values, CAB permutation") {
    assert(rawRst.size === 5)
    assert(rawRst(0) === dataRst(0))
    assert(rawRst(1) === dataRst(1))
    assert(rawRst(2) === dataRst(2))
    assert(rawRst(3) === dataRst(3))
    assert(rawRst(4) === dataRst(4))
  }

}
