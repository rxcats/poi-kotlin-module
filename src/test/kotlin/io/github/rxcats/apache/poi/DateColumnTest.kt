package io.github.rxcats.apache.poi

import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import java.time.LocalDate
import java.time.LocalDateTime

class DateColumnTest {

    @Test
    fun dateColumnTest() {
        @ExcelSheet("DateColumnData")
        data class DateColumnData(
            @ExcelCellType(CellStyleType.DATE)
            val date: LocalDate,

            @ExcelCellType(CellStyleType.DATETIME)
            val dateTime: LocalDateTime
        )

        val bytes = excelExport(DateColumnData(date = LocalDate.now(), dateTime = LocalDateTime.now()))
        assertThat(bytes).hasSizeGreaterThan(0)
    }

}
