package lin.com.readfromexcel

sealed class ReadRange(val from: Int?, val to: Int?) {
    class All : ReadRange(null, null)
    class Range(fromLine: Int?, toLine: Int?) : ReadRange(fromLine, toLine)
}