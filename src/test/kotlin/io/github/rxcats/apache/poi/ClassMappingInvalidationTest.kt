package io.github.rxcats.apache.poi

import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.assertThrows

class ClassMappingInvalidationTest {

    @Test
    fun noDataClassErrorTest() {
        @ExcelSheet("NoDataSheet")
        class NoData(
            val id: String
        )

        val e = assertThrows<IllegalArgumentException> {
            excelExport(listOf(NoData(id = "no1")))
        }

        assertThat(e.message).contains("must be a data class")
    }

    @Test
    fun propertyDeclarationErrorTest() {
        @ExcelSheet("PropertyDeclarationErrorData")
        data class PropertyDeclarationErrorData(
            val id: String
        ) {
            val errorProperty: String = ""
        }

        val e = assertThrows<IllegalArgumentException> {
            excelExport(listOf(PropertyDeclarationErrorData(id = "no1")))
        }

        assertThat(e.message).contains("properties must all be declared in the constructor")
    }

}
