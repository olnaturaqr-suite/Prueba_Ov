
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
    exit 1
}

$wrapperJar = Join-Path $PSScriptRoot ".mvn\wrapper\maven-wrapper.jar"
java -classpath $wrapperJar "-Dmaven.multiModuleProjectDirectory=$PSScriptRoot" org.apache.maven.wrapper.MavenWrapperMain spring-boot:run
