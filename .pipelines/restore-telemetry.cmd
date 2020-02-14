cd /D "%~dp0"

nuget.exe restore -ConfigFile nuget.config packages.config

move /Y Microsoft.Telemetry.Inbox.Native.10.0.18362.1-190318-1202.19h1-release.amd64fre\build\native\inc\MicrosoftTelemetry.h ..\src\common\Telemetry\TraceLoggingDefines.h