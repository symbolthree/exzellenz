@ECHO OFF
set JAVA_OPTS=-Xms64m -Xmx1024m -XX:+UseParallelGC
start javaw.exe -Dfile.encoding=UTF8 -cp ^
.\exzellenz-@build.version@.jar ^
-splash:splash.gif ^
symbolthree.oracle.excel.EXZ