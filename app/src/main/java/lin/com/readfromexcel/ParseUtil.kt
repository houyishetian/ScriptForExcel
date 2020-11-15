package lin.com.readfromexcel

import java.lang.IllegalArgumentException
import java.lang.IllegalStateException

object ParseUtil {

    fun parseData(
        originalData: List<Map<String, String?>>,
        keyRulesData: Map<String, String?>
    ): List<Map<String, String>> {
        val afterParsedKey = parseKeysToNewKeys(originalData, keyRulesData)
        val afterParseValue = parseValuesToNewValues(afterParsedKey)
        return afterParseValue
    }

    // 将读取到的 originalData 按照 rulesData 的规则，将key替换掉，这样新的data就可以用来生成json了
    private fun parseKeysToNewKeys(
        originalData: List<Map<String, String?>>,
        rulesData: Map<String, String?>
    ): List<Map<String, String?>> {
        val result = mutableListOf<Map<String, String?>>()
        originalData.forEach {
            val newDataMap = linkedMapOf<String, String?>()
            it.forEach {
                val newKey = rulesData[it.key] ?: it.key
                newDataMap.put(newKey, it.value)
            }
            result.add(newDataMap)
        }
        return result
    }

    private fun parseValuesToNewValues(originalData: List<Map<String, String?>>): List<Map<String, String>> {
        val result = mutableListOf<Map<String, String>>()
        originalData.forEach {
            val newDataMap = linkedMapOf<String, String>()
            it.forEach {
                val newValue = it.value.parseNull().parseNA()
                newDataMap.put(it.key, newValue)
            }
            result.add(newDataMap)
        }
        return result
    }

    private fun String?.parseNull(): String = this ?: ""

    private fun String.parseNA(): String = this.takeUnless {
        it.equals("NA", ignoreCase = true) || it.equals(
            "N/A",
            ignoreCase = true
        )
    } ?: ""

    fun readParseRules(filePath: String, sheetName: String): Map<String, String?> {
        val readResult = ReadUtil.getSheetContentFromExcel(filePath, sheetName)
        // 这2个key必须有
        val fromKey = "From"
        val toKey = "To"
        readResult?.firstOrNull()
            ?.takeIf { it.size >= 2 && it.containsKey(fromKey) && it.containsKey(toKey) }?.let {
                val result = linkedMapOf<String, String?>()
                readResult.forEach {
                    result.put(it[fromKey]!!, it[toKey])
                }
                return result
            }
            ?: throw IllegalStateException("parseRule excel is empty or has a wrong format, the $sheetName must have columns of \"$fromKey\" and \"$toKey\"!")
    }

    // 将sheet中每行的数据转换成一下，相当于给每行一个name
    fun parseSheetData(
        originalData: List<Map<String, String?>>,
        usedColumn: Int = 0,
        parseFunctionForColumn: (String?) -> String
    ): Map<String, Map<String, String?>> {
        val result = linkedMapOf<String, Map<String, String?>>()
        originalData.forEach {
            it.values.withIndex().find { it.index == usedColumn }?.apply {
                val parsedValue = parseFunctionForColumn(this.value)
                result.put(parsedValue, it)
            } ?: throw IllegalArgumentException("usedColumn is illegal!")
        }
        return result
    }
}