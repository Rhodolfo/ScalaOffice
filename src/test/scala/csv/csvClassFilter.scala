import org.scalatest.FunSuite

class ReadCSVFilterMix extends FunSuite {

  import com.arena.office.csv.readCSVintoClass
  import com.arena.testing.auxRead.{ReadMix,dataMix}

  def filtFunc(x: ReadMix): Boolean = x.entero % 2 == 0 && x.entero % 4 != 0
  val rawMix = readCSVintoClass[ReadMix]("src/test/resources/tipos.csv", filtFunc)

  test("Filtered data must contain even Ints which are not divisible by 4") {
    assert(rawMix.size === 3)
    assert(rawMix(0) === dataMix(0))
    assert(rawMix(1) === dataMix(2))
    assert(rawMix(2) === dataMix(4))
  }

}
