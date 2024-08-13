[![codecov](https://codecov.io/gh/rxcats/poi-kotlin-module/branch/main/graph/badge.svg)](https://codecov.io/gh/rxcats/poi-kotlin-module)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Maven Central Version](https://img.shields.io/maven-central/v/io.github.rxcats/poi-kotlin-module)](https://central.sonatype.com/artifact/io.github.rxcats/poi-kotlin-module)

# Apache POI Kotlin Module

Kotlin module for Apache POI

- You can use the kotlin data classes to convert to excel.
- properties can be immutable; i.e. val is allowed

## Requirements

Java 17 and 21

## Quickstart

Mapping and convert to excel bytes

```kotlin
@ExcelSheet("Sheet1") // sheet name annotation
data class ItemTable(
    @ExcelCellOrder(10) // column order annotation
    @ExcelCellHeader("Id") // column name annotation
    val id: Int,

    @ExcelCellOrder(20)
    @ExcelCellHeader("Name")
    val name: String,

    @ExcelCellOrder(30)
    @ExcelCellHeader("ItemType")
    val itemType: String,

    @ExcelCellOrder(40)
    @ExcelCellHeader("Price")
    val price: Long,

    @ExcelCellOrder(50)
    @ExcelCellHeader("Description")
    val description: String,
)

// create data and export to bytes
val bytes = excelExport(listOf(ItemTable(
    id = TODO(),
    name = TODO(),
    itemType = TODO(),
    price = TODO(),
    description = TODO()
)))

// import file
val list = excelImport<ItemTable>(File("filename.xlsx"))

```
