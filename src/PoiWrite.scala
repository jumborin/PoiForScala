import java.io.FileOutputStream
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import java.util.Date
import java.text.SimpleDateFormat
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.poifs.filesystem.POIFSFileSystem
class PoiWrite {
  def poiWrite = {
    try {
      val outPutFile: FileOutputStream = new FileOutputStream("D:/poiTest.xls")
      var workbook: HSSFWorkbook = new HSSFWorkbook()
      val sheet: HSSFSheet = workbook.createSheet("test")
      workbook.setSheetName(0, "出力シート名")
      val row: HSSFRow = sheet.createRow(0)
      val cellA1: HSSFCell = row.createCell(0)
      cellA1.setCellValue("A1日本語文字列")
      val cellA2: HSSFCell = row.createCell(1)
      cellA2.setCellValue(new Date)
      val cellA3: HSSFCell = row.createCell(2)
      cellA3.setCellValue(2.2)
      workbook.write(outPutFile)
      outPutFile.close()
      println("正常終了")
    } catch {
      case e: Exception =>
        e.printStackTrace()
        println("処理失敗")
    }
  }
}