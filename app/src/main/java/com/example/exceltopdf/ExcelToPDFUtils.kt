package com.example.exceltopdf

import com.itextpdf.text.Document
import com.itextpdf.text.PageSize
import com.itextpdf.text.Paragraph
import com.itextpdf.text.Phrase
import com.itextpdf.text.pdf.PdfPCell
import com.itextpdf.text.pdf.PdfPTable
import com.itextpdf.text.pdf.PdfWriter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream
import java.io.InputStream

object ExcelToPDFUtils {

    fun convertExcelToPDF(excelFileStream: InputStream, pdfFilePath: File) {
        // Code to convert Excel to PDF

        val workbook = WorkbookFactory.create(excelFileStream)
        val sheet = workbook.getSheetAt(0)

        // 创建 PDF 文档
        val document = Document(PageSize.A4.rotate())
        PdfWriter.getInstance(document, FileOutputStream(pdfFilePath))
        document.open()

        // 创建 PDF 表格，列数与 Excel 表格相同
        var table = PdfPTable(sheet.getRow(0).lastCellNum.toInt())
        // 读取 Excel 数据并将其添加到 PDF 表格中
        for (row in sheet) {
            if (row.lastCellNum.toInt() == -1) {
                document.add(table)
                table = if (sheet.getRow(row.rowNum + 1).getCell(0).stringCellValue.contains("Filled in ")) {
                    PdfPTable(sheet.getRow(row.rowNum + 2).lastCellNum.toInt())
                } else {
                    PdfPTable(sheet.getRow(row.rowNum + 1).lastCellNum.toInt())
                }
                val spacer = Paragraph(" ")
                spacer.spacingAfter = 20f // 设置空行的间距（例如20点）
                document.add(spacer)
                continue
            }
            for (cell in row) {
                val cellValue = when (cell.cellType) {
                    org.apache.poi.ss.usermodel.CellType.STRING -> cell.stringCellValue
                    org.apache.poi.ss.usermodel.CellType.NUMERIC -> cell.numericCellValue.toString()
                    else -> ""
                }
                if (cellValue.contains("Filled in")) {
                    table.addCell(PdfPCell(Phrase(cellValue)))
                    table.completeRow()
                } else {
                    table.addCell(PdfPCell(Phrase(cellValue)))
                }
            }
        }
        document.add(table)
        // 将表格添加到 PDF 文档
        document.close()
        println("PDF 已生成：$pdfFilePath")
    }
}