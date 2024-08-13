package io.github.rxcats.apache.poi

import java.io.File
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.declaredMemberProperties
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.primaryConstructor

internal fun sheetNameOf(dataClass: KClass<*>): String = dataClass.findAnnotation<ExcelSheet>()?.name
    ?: dataClass.simpleName
    ?: "Sheet1"

internal fun sortedMemberPropertiesOf(dataClass: KClass<*>): List<KProperty1<*, *>> = dataClass.declaredMemberProperties
    .filterNot { it.hasAnnotation<ExcelCellIgnore>() }
    .sortedBy { it.findAnnotation<ExcelCellOrder>()?.value ?: 0 }

internal fun validate(dataClass: KClass<*>, sortedMemberProperties: List<KProperty1<*, *>>) {
    require(dataClass.isData) { "$dataClass must be a data class" }
    require(sortedMemberProperties.size == dataClass.primaryConstructor!!.parameters.size) {
        "${dataClass.simpleName} properties must all be declared in the constructor"
    }
}

inline fun <reified T> excelExport(vararg data: T): ByteArray {
    return ExcelWriter(T::class).toByteArray(data.toList())
}

inline fun <reified T> excelExport(data: List<T>): ByteArray {
    return ExcelWriter(T::class).toByteArray(data)
}

inline fun <reified T> excelImport(data: ByteArray): List<T> {
    return ExcelReader(T::class).parse(data)
}

inline fun <reified T> excelImport(file: File): List<T> {
    return ExcelReader(T::class).parse(file)
}
