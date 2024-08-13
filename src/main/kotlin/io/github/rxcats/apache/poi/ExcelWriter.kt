package io.github.rxcats.apache.poi

import io.github.rxcats.apache.poi.extensions.getFormat
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Calendar
import java.util.Date
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.KVisibility
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.jvm.jvmErasure

class ExcelWriter(private val dataClass: KClass<*>) {

    private val sheetName: String = sheetNameOf(this.dataClass)

    private val sortedMemberProperties: List<KProperty1<*, *>> = sortedMemberPropertiesOf(this.dataClass)

    init {
        validate(dataClass, sortedMemberProperties)
    }

    private fun createSheet(wb: XSSFWorkbook, name: String): XSSFSheet = wb.createSheet(name)

    private fun applyCellColor(wb: XSSFWorkbook, cell: XSSFCell, color: IndexedColors) {
        val style = wb.createCellStyle()
        style.fillForegroundColor = color.index
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        cell.cellStyle = style
    }

    private fun adjustColumnSize(sheet: XSSFSheet, index: Int, length: Int) {
        val defaultColumnWidth = 2048
        val maxLength = 70
        val characterWidth = 256
        val lengthWeight = 1.14388
        val len = if (length > maxLength) maxLength else length
        val adjust = ((len * lengthWeight).toInt() * characterWidth) + defaultColumnWidth

        val width = sheet.getColumnWidth(index)

        if (width > adjust) return

        val maxWidth = 255 * characterWidth

        if (adjust > maxWidth) {
            sheet.setColumnWidth(index, maxWidth)
        } else {
            sheet.setColumnWidth(index, adjust)
        }
    }

    private fun writeHeader(wb: XSSFWorkbook, sheet: XSSFSheet) {
        val titleRow = sheet.createRow(0)

        sortedMemberProperties.forEachIndexed { i, property ->
            val header = property.findAnnotation<ExcelCellHeader>()?.header ?: property.name
            val bg = property.findAnnotation<ExcelCellHeader>()?.cellColor ?: IndexedColors.GREY_25_PERCENT
            val cell = titleRow.createCell(i)
            cell.setCellValue(header)
            applyCellColor(wb, cell, bg)
            adjustColumnSize(sheet, i, header.length)
        }
    }

    private fun <T> getPropertyValueOrNull(property: KProperty1<*, *>, data: T): Any? =
        if (property.visibility == KVisibility.PUBLIC) {
            property.getter.call(data)
        } else {
            null
        }

    private fun writeStringCell(wb: XSSFWorkbook, sheet: XSSFSheet, col: Int, cell: XSSFCell, value: String, type: CellStyleType) {
        cell.cellType = CellType.STRING
        cell.setCellValue(wb.creationHelper.createRichTextString(value))
        cell.cellStyle = wb.createCellStyle().apply {
            this.dataFormat = wb.createDataFormat().getFormat(type)
        }
        adjustColumnSize(sheet, col, value.length)
    }

    private fun writeNumberCell(wb: XSSFWorkbook, sheet: XSSFSheet, col: Int, cell: XSSFCell, value: Any, type: CellStyleType) {
        value as Number
        cell.cellType = CellType.NUMERIC
        cell.setCellValue(value.toDouble())
        cell.cellStyle = wb.createCellStyle().apply {
            this.dataFormat = wb.createDataFormat().getFormat(type)
        }
        adjustColumnSize(sheet, col, value.toString().length)
    }

    private fun writeDateTimeCell(wb: XSSFWorkbook, sheet: XSSFSheet, col: Int, cell: XSSFCell, value: Any, type: CellStyleType) {
        cell.cellType = CellType.NUMERIC
        if (value::class == LocalDate::class) {
            cell.setCellValue(value as LocalDate)
            cell.cellStyle = wb.createCellStyle().apply {
                if (type == CellStyleType.DEFAULT) {
                    this.dataFormat = wb.createDataFormat().getFormat(CellStyleType.DATE)
                } else {
                    this.dataFormat = wb.createDataFormat().getFormat(type)
                }
            }
        } else {
            when (value::class) {
                LocalDateTime::class -> {
                    cell.setCellValue(value as LocalDateTime)
                }

                Date::class -> {
                    cell.setCellValue(value as Date)
                }
            }

            cell.cellStyle = wb.createCellStyle().apply {
                if (type == CellStyleType.DEFAULT) {
                    this.dataFormat = wb.createDataFormat().getFormat(CellStyleType.DATETIME)
                } else {
                    this.dataFormat = wb.createDataFormat().getFormat(type)
                }
            }
        }

        adjustColumnSize(sheet, col, value.toString().length)
    }

    private fun <T> writeBody(wb: XSSFWorkbook, sheet: XSSFSheet, dataList: List<T>) {
        dataList.forEachIndexed { i, data ->
            val row = sheet.createRow(i + 1)

            sortedMemberProperties
                .forEachIndexed { col, property ->
                    val format = property.findAnnotation<ExcelCellType>()?.type ?: CellStyleType.DEFAULT
                    val value: Any? = getPropertyValueOrNull(property, data)
                    val cell = row.createCell(col)

                    if (value == null) {
                        cell.cellType = CellType.BLANK
                        cell.setBlank()
                    } else {
                        when (property.returnType.jvmErasure) {
                            String::class -> {
                                writeStringCell(wb, sheet, col, cell, value as String, format)
                            }

                            Boolean::class -> {
                                cell.cellType = CellType.BOOLEAN
                                cell.setCellValue(value as Boolean)
                                adjustColumnSize(sheet, col, value.toString().length)
                            }

                            Double::class, Float::class, Int::class, Long::class, BigDecimal::class -> {
                                writeNumberCell(wb, sheet, col, cell, value, format)
                            }

                            Calendar::class -> {
                                value as Calendar
                                writeDateTimeCell(wb, sheet, col, cell, value.time, format)
                            }

                            LocalDate::class, LocalDateTime::class, Date::class -> {
                                writeDateTimeCell(wb, sheet, col, cell, value, format)
                            }
                        }
                    }
                }
        }
    }

    fun <T> toByteArray(data: List<T>): ByteArray {
        require(data.isNotEmpty()) { "data must not be empty" }

        return XSSFWorkbook().use { wb ->
            val sheet = createSheet(wb, sheetName)
            writeHeader(wb, sheet)
            writeBody(wb, sheet, data)

            ByteArrayOutputStream().use { os ->
                wb.write(os)
                os.toByteArray()
            }
        }
    }
}
