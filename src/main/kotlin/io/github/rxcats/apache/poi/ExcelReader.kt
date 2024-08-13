package io.github.rxcats.apache.poi

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.File
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Calendar
import java.util.Date
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.primaryConstructor
import kotlin.reflect.jvm.jvmErasure

class ExcelReader(private val dataClass: KClass<*>) {

    private val sheetName: String = sheetNameOf(this.dataClass)

    private val sortedMemberProperties: List<KProperty1<*, *>> = sortedMemberPropertiesOf(this.dataClass)

    init {
        validate(dataClass, sortedMemberProperties)
    }

    fun <T> parse(file: File): List<T> {
        return parse(file.readBytes())
    }

    fun <T> parse(bytes: ByteArray): List<T> {
        val constructor = this.dataClass.primaryConstructor!!
        val constructorParameters = constructor.parameters

        return ByteArrayInputStream(bytes).use {
            XSSFWorkbook(it).use { wb ->
                val sheet = wb.getSheet(sheetName)

                val result = mutableListOf<T>()

                sheet.rowIterator().asSequence().forEachIndexed rowLoop@{ i, row ->
                    if (i == 0) return@rowLoop

                    val constructorParams = arrayOfNulls<Any?>(constructorParameters.size)

                    row.cellIterator().asSequence().forEachIndexed { col, cell ->
                        val property = sortedMemberProperties[col]

                        val propertyReturnType = property.returnType

                        val propertyType = propertyReturnType.jvmErasure

                        val constructorIndex = constructorParameters.indexOfFirst { param -> param.name == property.name }

                        when (cell.cellType) {
                            CellType.BLANK -> {
                                if (propertyReturnType.isMarkedNullable) {
                                    constructorParams[constructorIndex] = null
                                }
                            }

                            CellType.STRING -> {
                                constructorParams[constructorIndex] = cell.stringCellValue
                            }

                            CellType.BOOLEAN -> {
                                constructorParams[constructorIndex] = cell.booleanCellValue
                            }

                            CellType.NUMERIC -> {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    when (propertyType) {
                                        LocalDate::class -> {
                                            constructorParams[constructorIndex] = cell.localDateTimeCellValue.toLocalDate()
                                        }

                                        LocalDateTime::class -> {
                                            constructorParams[constructorIndex] = cell.localDateTimeCellValue
                                        }

                                        Calendar::class -> {
                                            val calendar = Calendar.getInstance()
                                            calendar.time = cell.dateCellValue
                                            constructorParams[constructorIndex] = calendar
                                        }

                                        Date::class -> {
                                            constructorParams[constructorIndex] = cell.dateCellValue
                                        }
                                    }

                                } else {
                                    when (propertyType) {
                                        Double::class -> constructorParams[constructorIndex] = cell.numericCellValue
                                        Float::class -> constructorParams[constructorIndex] = cell.numericCellValue.toFloat()
                                        Int::class -> constructorParams[constructorIndex] = cell.numericCellValue.toInt()
                                        Long::class -> constructorParams[constructorIndex] = cell.numericCellValue.toLong()
                                        BigDecimal::class -> constructorParams[constructorIndex] = cell.numericCellValue.toBigDecimal()
                                    }
                                }
                            }

                            else -> {}
                        }

                    }

                    @Suppress("UNCHECKED_CAST")
                    result += constructor.call(*constructorParams) as T
                }

                result
            }
        }
    }

}
