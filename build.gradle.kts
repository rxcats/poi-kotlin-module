import com.vanniktech.maven.publish.KotlinJvm
import com.vanniktech.maven.publish.SonatypeHost

plugins {
    alias(libs.plugins.spring.boot)
    alias(libs.plugins.spring.dependency.management)
    alias(libs.plugins.kotlin.jvm)
    alias(libs.plugins.kotlin.spring)
    alias(libs.plugins.maven.publish)
    jacoco
}

repositories {
    mavenCentral()
}

dependencies {
    implementation(rootProject.libs.kotlin.reflect)
    api(rootProject.libs.bundles.apache.poi)

    testImplementation(rootProject.libs.spring.boot.starter.test)
    testImplementation(rootProject.libs.kotlin.test.junit5)
    testRuntimeOnly(rootProject.libs.junit.platform.launcher)
}

kotlin {
    compilerOptions {
        freeCompilerArgs.addAll("-Xjsr305=strict")
    }
}

tasks.withType<Test> {
    useJUnitPlatform()
    jvmArgs("-XX:+EnableDynamicAgentLoading", "-Xshare:off")
}

mavenPublishing {
    configure(KotlinJvm(sourcesJar = true))
    publishToMavenCentral(SonatypeHost.CENTRAL_PORTAL, automaticRelease = true)
    signAllPublications()
    coordinates("io.github.rxcats", "poi-kotlin-module", "0.0.1")

    pom {
        name.set("Apache POI Kotlin Module")
        description.set("Kotlin module for Apache POI")
        inceptionYear.set("2024")
        url.set("https://github.com/rxcats/poi-kotlin-module")
        licenses {
            license {
                name.set("The Apache License, Version 2.0")
                url.set("https://www.apache.org/licenses/LICENSE-2.0.txt")
                distribution.set("https://www.apache.org/licenses/LICENSE-2.0.txt")
            }
        }
        developers {
            developer {
                id.set("rxcats")
                name.set("Yong-Jun Park")
                url.set("https://github.com/rxcats")
            }
        }
        scm {
            url.set("https://github.com/rxcats/poi-kotlin-module")
        }
    }
}
