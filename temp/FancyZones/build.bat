@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2017\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" x86    
msbuild.exe fancyzones\fancyzones.vcxproj