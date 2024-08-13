package io.github.rxcats.apache.poi

import io.github.rxcats.apache.poi.extensions.epoch
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import java.time.Instant
import java.util.Calendar

class CalendarExtensionsTest {

    @Test
    fun epochTest() {
        val cal = Calendar.getInstance().epoch()
        assertThat(cal.toInstant()).isEqualTo(Instant.EPOCH)
    }

}
