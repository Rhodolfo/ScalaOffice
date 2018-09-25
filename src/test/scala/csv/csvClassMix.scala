import org.scalatest.FunSuite

class ReadCSVClassMix extends FunSuite {

  import com.arena.office.csv.readCSVintoClass
  import com.arena.testing.auxRead.{ReadMix,dataMix}

  def roundIt(d: Double) = BigDecimal(d).setScale(2, BigDecimal.RoundingMode.HALF_UP).toDouble

  val rawMix = readCSVintoClass[ReadMix]("src/test/resources/tipos.csv")
    .map(_ match {case ReadMix(a,b,c) => ReadMix(a,b,roundIt(c))})

  test("Rows must equal preset values, detecting types") {
    assert(rawMix.size === 5)
    assert(rawMix(0) === dataMix(0))
    assert(rawMix(1) === dataMix(1))
    assert(rawMix(2) === dataMix(2))
    assert(rawMix(3) === dataMix(3))
    assert(rawMix(4) === dataMix(4))
  }

}
