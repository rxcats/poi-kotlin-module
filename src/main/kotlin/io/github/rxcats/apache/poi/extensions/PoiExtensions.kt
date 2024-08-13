package io.github.rxcats.apache.poi.extensions

import io.github.rxcats.apache.poi.CellStyleType
import org.apache.poi.xssf.usermodel.XSSFDataFormat

fun XSSFDataFormat.getFormat(type: CellStyleType): Short {
    return this.getFormat(type.format)
}
