set PATH=%PATH%;C:\jdk1.5.0_08\bin;c:\horb2.0\bin
set CLASSPATH=%CLASSPATH%;.;c:\horb2.0\lib\horb20.jar;c:\horb2.0\classes
set CLASSPATH=%CLASSPATH%;c:\cz\Lib\classes12.zip;c:\cz\Lib\CZClass.jar;c:\cz\Lib\CZSystemLib.jar

cd c:\cz\
C:\jdk1.5.0_08\bin\java -Xms128M -Xmx256M cz.CZMain
exit
