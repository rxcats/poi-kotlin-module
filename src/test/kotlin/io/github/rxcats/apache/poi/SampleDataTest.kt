package io.github.rxcats.apache.poi

import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.Calendar
import java.util.Date

class SampleDataTest {
    @ExcelSheet("Sheet1")
    data class SampleData(
        @ExcelCellOrder(10)
        @ExcelCellHeader("Column1")
        val col1: String,

        @ExcelCellOrder(20)
        @ExcelCellHeader("Column2")
        val col2: Boolean,

        @ExcelCellOrder(30)
        @ExcelCellHeader("Column3")
        val col3: Double,

        @ExcelCellOrder(40)
        @ExcelCellHeader("Column4")
        val col4: Int,

        @ExcelCellOrder(50)
        @ExcelCellHeader("Column5")
        val col5: BigDecimal,

        @ExcelCellOrder(60)
        @ExcelCellHeader(header = "Column6")
        val col6: LocalDateTime,

        @ExcelCellOrder(70)
        @ExcelCellHeader(header = "Column7")
        val col7: LocalDate,

        @ExcelCellOrder(80)
        @ExcelCellHeader(header = "Column8")
        val col8: Calendar,

        @ExcelCellOrder(90)
        @ExcelCellHeader(header = "Column9")
        val col9: Date,

        @ExcelCellOrder(100)
        @ExcelCellHeader(header = "Column10")
        val col10: Long,

        @ExcelCellOrder(110)
        @ExcelCellHeader(header = "Column11")
        val col11: Long? = null,
    ) {
        @ExcelCellIgnore
        val ignoreField: Int? = null
    }

    @Test
    fun sampleDataTest() {
        val example = SampleData(
            col1 = "string",
            col2 = true,
            col3 = 1.1,
            col4 = 1,
            col5 = 1.99.toBigDecimal(),
            col6 = LocalDateTime.now(),
            col7 = LocalDate.now(),
            col8 = Calendar.getInstance(),
            col9 = Date(),
            col10 = 22123123123123L,
            col11 = null
        )

        val bytes = excelExport(listOf(example))

        val result = excelImport<SampleData>(bytes)

        assertThat(result).hasSize(1)
        assertThat(result.first().col1).isEqualTo(example.col1)
        assertThat(result.first().col2).isTrue()
        assertThat(result.first().col3).isEqualTo(example.col3)
        assertThat(result.first().col4).isEqualTo(example.col4)
        assertThat(result.first().col5).isEqualTo(example.col5)

        assertThat(result.first().col6.year).isEqualTo(example.col6.year)
        assertThat(result.first().col6.monthValue).isEqualTo(example.col6.monthValue)
        assertThat(result.first().col6.dayOfMonth).isEqualTo(example.col6.dayOfMonth)
        assertThat(result.first().col6.hour).isEqualTo(example.col6.hour)
        assertThat(result.first().col6.minute).isEqualTo(example.col6.minute)
        assertThat(result.first().col6.second).isEqualTo(example.col6.second)

        assertThat(result.first().col7).isEqualTo(example.col7)
        assertThat(result.first().col7.year).isEqualTo(example.col7.year)
        assertThat(result.first().col7.monthValue).isEqualTo(example.col7.monthValue)
        assertThat(result.first().col7.dayOfMonth).isEqualTo(example.col7.dayOfMonth)

        val col8After = LocalDateTime.ofInstant(result.first().col8.toInstant(), ZoneId.systemDefault())
        val col8Before = LocalDateTime.ofInstant(example.col8.toInstant(), ZoneId.systemDefault())
        assertThat(col8After.year).isEqualTo(col8Before.year)
        assertThat(col8After.monthValue).isEqualTo(col8Before.monthValue)
        assertThat(col8After.dayOfMonth).isEqualTo(col8Before.dayOfMonth)
        assertThat(col8After.hour).isEqualTo(col8Before.hour)
        assertThat(col8After.minute).isEqualTo(col8Before.minute)
        assertThat(col8After.second).isEqualTo(col8Before.second)

        val col9After = LocalDateTime.ofInstant(result.first().col9.toInstant(), ZoneId.systemDefault())
        val col9Before = LocalDateTime.ofInstant(example.col9.toInstant(), ZoneId.systemDefault())
        assertThat(col9After.year).isEqualTo(col9Before.year)
        assertThat(col9After.monthValue).isEqualTo(col9Before.monthValue)
        assertThat(col9After.dayOfMonth).isEqualTo(col9Before.dayOfMonth)
        assertThat(col9After.hour).isEqualTo(col9Before.hour)
        assertThat(col9After.minute).isEqualTo(col9Before.minute)
        assertThat(col9After.second).isEqualTo(col9Before.second)

        assertThat(result.first().col10).isEqualTo(example.col10)
        assertThat(result.first().col11).isNull()
        assertThat(result.first().ignoreField).isNull()
    }

}
