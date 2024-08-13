package io.github.rxcats.apache.poi.extensions

import java.util.Calendar

fun Calendar.epoch(): Calendar {
    this.timeInMillis = 0
    return this
}
