#!/bin/sh

export JAVA_OPTS="-Xms64m -Xmx1024m -XX:+UseParallelGC"
java -Dfile.encoding=UTF8 -cp \
./exzellenz-@build.version@.jar \
symbolthree.oracle.excel.EXZELLENZ $1
