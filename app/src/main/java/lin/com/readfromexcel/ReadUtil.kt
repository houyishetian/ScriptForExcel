package lin.com.readfromexcel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.lang.IllegalArgumentException
import java.lang.Integer.min
import java.text.NumberFormat

object ReadUtil {

    fun getSheetContentFromExcel(
        filePath: String,
        sheetName: String,
        readRange: ReadRange? = null
    ): List<Map<String, String?>>? {
        return readExcel(filePath)?.let { getSheetByName(it, sheetName) }
            ?.let { getSheetContent(it, readRange = readRange) }
    }

    private fun readExcel(filePath: String): Workbook? {
        return File(filePath).takeIf { it.exists() && it.isFile }?.let {
            it.inputStream().let { stream ->
                val result = when (it.extension) {
                    "xls" -> HSSFWorkbook(stream)
                    "xlsx" -> XSSFWorkbook(stream)
                    else -> null
                }
                stream.close()
                result
            }
        }
    }

    private fun getSheetByName(workbook: Workbook, sheetName: String): Sheet? {
        workbook.sheetIterator().forEach {
            if (it.sheetName == sheetName) return it
        }
        return null
    }

    private fun getCellValue(cell: Cell): String = when (cell.cellType) {
        CellType.NUMERIC -> NumberFormat.getInstance().let {
            it.isGroupingUsed = false
            it.format(cell.numericCellValue)
        }
        CellType.STRING -> cell.richStringCellValue.string
        else -> ""
    }

    private fun getSheetContent(
        sheet: Sheet,
        labelRow: Int = 0,
        readRange: ReadRange? = null
    ): List<Map<String, String?>> {
        //注意，对于一个 3行5列 的表格，lastRowNum=2，直接用..;但 lastCellNum需要用 until
        //获取最大列数，因为有可能有些cell是空的，poi 读取时只读取到最后一个不为空的值，但为空的依然需要遍历到并处理，所以需要先取到最大值
        val maxColumn =
            (0..sheet.lastRowNum).map { sheet.getRow(it)?.lastCellNum?.toInt() ?: 0 }.max() ?: 0
        // 获取 label 行信息，读取到的信息用来作为返回值的各个cell的key
        if (labelRow !in 0..sheet.lastRowNum) throw IllegalArgumentException("labelRow out of index!") // labelRow 越界
        val labelKeys: List<String> = sheet.getRow(labelRow)?.let getLabels@{ row ->
            // 读取 label 行所有 cell，遇到第一个为空的 cell 为止
            val keys = mutableListOf<String>()
            (0 until maxColumn).map { row.getCell(it) }.forEach { cell ->
                val value = cell?.let { getCellValue(it) }?.takeIf { it.isNotEmpty() }
                value?.let { keys.add(it) } ?: return@getLabels keys
            }
            return@getLabels keys
        }
            ?: throw IllegalArgumentException("labelRow does not has any label!") // labelRow 所有cell都为空，说明此表格没有任何label信息，读取没有意义

        val result = mutableListOf<Map<String, String?>>()
        // 获取真实的读取range，如果from不在范围就从头开始，如果to不在范围就读到末尾，如果都不在范围就读完
        val realReadRange = readRange?.takeIf { !(it.from == null && it.to == null) }?.let {
            val realFrom = it.from?.let {
                if (it in 0..sheet.lastRowNum) it else 0
            } ?: 0
            val realTo = it.to?.let {
                if (it in 0..sheet.lastRowNum) it else sheet.lastRowNum
            } ?: 0
            if (realFrom >= realTo) throw IllegalArgumentException("illegal range! from index is bigger than to index!")
            (realFrom..realTo)
        } ?: (0..sheet.lastRowNum)
        // label 行跳过
        val skipIndex = realReadRange.withIndex().find { it.value == labelRow }?.index ?: -1
        realReadRange.map { sheet.getRow(it) }.withIndex().forEach allRows@{ rowItem ->
            if (rowItem.index == skipIndex) return@allRows // label 行需要跳过
            val currentRowInfos = linkedMapOf<String, String?>()
            labelKeys.indices.forEach { columnIndex ->
                currentRowInfos.put(
                    labelKeys[columnIndex],
                    rowItem.value?.getCell(columnIndex)?.let { getCellValue(it) })
            }
            result.add(currentRowInfos)
        }
        return result
    }
}