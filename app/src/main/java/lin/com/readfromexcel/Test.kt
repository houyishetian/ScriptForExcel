package lin.com.readfromexcel

object Test {
    fun test() {
        val path = "C:\\Users\\lisonglin\\Desktop\\test.xlsx"

        val originalData = ReadUtil.getSheetContentFromExcel(path, "Sheet")
        val rulesData = ParseUtil.readParseRules(path, "转换规则")
        originalData?.let {
//            ParseUtil.parseData(it, rulesData).let {
//
//            }
            val result = ParseUtil.parseSheetData(originalData,0){
                "test"+it
            }
            result.forEach {

            }
        }
    }

}