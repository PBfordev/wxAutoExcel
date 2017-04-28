REM This batch generates the list of the header and source files
REM to be used in CMakeLists.txt

@echo off

echo ###################################################################### > files.cmake
echo # Author:      PB >> files.cmake
echo # Purpose:     List of all .h and .cpp files for wxAutoExcel library >> files.cmake
echo # Copyright:   (c) 2017 PB ^<pbfordev@gmail.com^> >> files.cmake
echo # Licence:     wxWindows licence >> files.cmake
echo ###################################################################### >> files.cmake
echo: >> files.cmake 
echo set(INCLUDES >> files.cmake
dir ..\..\include\wx\*.h /B /oN >> files.cmake
echo 	) >> files.cmake
echo: >> files.cmake
echo set(SOURCES >> files.cmake
dir ..\..\src\*.cpp /B /oN >> files.cmake
echo 	) >> files.cmake
echo: >> files.cmake
echo foreach(Inc ${INCLUDES}) >> files.cmake
echo 	set(SRCS ${SRCS} include/wx/${Inc}) >> files.cmake
echo endforeach(Inc) >> files.cmake
echo: >> files.cmake 
echo foreach(Src ${SOURCES}) >> files.cmake
echo 	set(SRCS ${SRCS} src/${Src}) >> files.cmake
echo endforeach(Src) >> files.cmake