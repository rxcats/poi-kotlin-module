package io.github.rxcats.apache.poi

enum class CellStyleType(val format: String) {
    GENERAL("General"),
    INT("0"),
    DECIMAL("0.00"),
    PRECISE("0.000000000"),
    CURRENCY("#,##0.00"),
    DATE("m/d/yyyy"),
    DATETIME("m/d/yyyy h:mm:ss AM/PM"),
    PERCENT("0.00%"),
    DEFAULT(""),
}
