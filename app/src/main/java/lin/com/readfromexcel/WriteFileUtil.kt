package lin.com.readfromexcel

import java.io.File

object WriteFileUtil {
    fun writeToFile(data: String, filePath: String) {
        val file = File(filePath).apply {
            if (!exists()) {
                this.parentFile?.mkdirs()
                this.createNewFile()
            }
        }
        file.appendText(data)
    }
}