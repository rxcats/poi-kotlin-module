package io.github.rxcats.apache.poi

import org.apache.poi.ss.usermodel.IndexedColors

@Target(AnnotationTarget.CLASS)
annotation class ExcelSheet(
    val name: String = "Sheet1",
)

@Target(AnnotationTarget.PROPERTY)
annotation class ExcelCellHeader(
    val header: String = "",
    val cellColor: IndexedColors = IndexedColors.GREY_25_PERCENT
)

@Target(AnnotationTarget.PROPERTY)
annotation class ExcelCellType(
    val type: CellStyleType = CellStyleType.DEFAULT,
)

@Target(AnnotationTarget.PROPERTY)
annotation class ExcelCellOrder(
    val value: Int = 0
)

@Target(AnnotationTarget.PROPERTY)
annotation class ExcelCellIgnore
