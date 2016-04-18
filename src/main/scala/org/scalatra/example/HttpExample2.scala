package org.scalatra.example

import java.io.OutputStream

import org.apache.poi.hssf.usermodel.{HSSFSheet, HSSFWorkbook}
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatra._

class HttpExample2 extends ScalatraServlet with GZipSupport {

  get("/delta-report") {
    Ok {
      println("in xls delta report api.... - set content type to openxmlformats")
      response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      response.setHeader("Content-Disposition", "attachment; filename=test.xls")
      val wb = new HSSFWorkbook()
      new DeltaReport().export(response.getOutputStream)
    }
  }

  get("/delta-report-xlsx") {
    Ok {
      println("in xlsx delta report api....")
      response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      response.setHeader("Content-Disposition", "attachment; filename=test.xlsx")
      val wb = new XSSFWorkbook()
      new DeltaReport().exportXlsx(response.getOutputStream, wb)
    }
  }
}

case class CellData(data: Any, dataType: Int = Cell.CELL_TYPE_STRING)

class StructuredWorksheet(name: String) {

  private var nextRow = 0
  private val workbook: HSSFWorkbook = new HSSFWorkbook()
  private val sheet = workbook.createSheet(name)

  def addHeader(values: CellData*) = addRow(values: _*)

  def addSection() = {
    if (nextRow > 1) addRow()
  }

  def addRow(values: CellData*) = {
    val row = sheet.createRow(nextRow)

    values.zipWithIndex.foreach({ case (cellData: CellData, index: Int) =>
      cellData.dataType match {
        case Cell.CELL_TYPE_NUMERIC => row.createCell(index, Cell.CELL_TYPE_NUMERIC).setCellValue(cellData.data match {
          case x: BigDecimal => x.toDouble
          case _ => 0l
        })
        case _ => row.createCell(index, Cell.CELL_TYPE_STRING).setCellValue(cellData.data.toString)
      }
      row.createCell(index, cellData.dataType).setCellValue(cellData.data.toString)
    })
    nextRow += 1
  }
  def write(output: OutputStream) = {
    workbook.write(output)
  }
}


class DeltaReport {

  implicit def tooCellData(data: Any): CellData = CellData(data)

  def export(output: OutputStream): Unit = {
    val worksheet = new StructuredWorksheet("Delta Report")

    worksheet.addHeader("BAML or Client", "Trade Ref", "Channel", "Trade Date", "Settlement Date", "Buy/Sell",
      "Instrument", "RIC", "Currency", "Quantity", "Price", "Gross Amount", "Comm Amount", "Comm Rate", "Gross Difference",
      "Commission Difference", "Mismatch Reason", "Link ID")

    for (a <- 1 to 100)
     addRow(worksheet)

    worksheet.write(output)
  }

  def addRow(worksheet: StructuredWorksheet): Unit = {
    worksheet.addRow("BAML", "MVAMTSTS91_121312", "DSA", "2016-04-18", "2016-04-21", "Sell", "KANEKA CORPORATION", "4118.T", "JPY",
      2065,
      933.4375,
      1927448.4234,
      193.1923,
//      CellData(2065, Cell.CELL_TYPE_NUMERIC),
//      CellData(933.4375, Cell.CELL_TYPE_NUMERIC),
//      CellData(1927448.4234, Cell.CELL_TYPE_NUMERIC),
//      CellData(193.1923, Cell.CELL_TYPE_NUMERIC),
      "1 bps",
      CellData(0.0000, Cell.CELL_TYPE_NUMERIC),
      CellData(-770.8077, Cell.CELL_TYPE_NUMERIC),
      "CD",
      "CR-file-upload-8512312312312123"
    )
  }

  def exportXlsx(output: OutputStream, wb:XSSFWorkbook): Unit = {
    val createHelper = wb.getCreationHelper()
    val worksheet = wb.createSheet("Test Sheet")
    val row = worksheet.createRow(0)
    row.createCell(1).setCellValue(1.2)
    row.createCell(2).setCellValue(
      createHelper.createRichTextString("This is a string"))
    row.createCell(3).setCellValue(true)
    wb.write(output)
  }
}