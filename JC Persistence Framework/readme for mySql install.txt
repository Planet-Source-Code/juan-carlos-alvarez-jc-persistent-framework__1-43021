Download mySql, mySql ODBC and GUI Control Center (this last one not necesary for test the jcframework).

Install mySql and mySql ODBC.

Move or copy the folder "test2" from:
C:\JC Persistence Framework\JCFramework\Test2\Base\mysql\test2
to
C:\mysql\data\

Setup the mySql server as a W2000 or NT service:
In MS-DOS type (go to C:\mysql\bin first):
C:\mysql\bin> mysqld-max-nt --install
You get a message:
Service succesfull installed.

Now you can work with it like any service in W2000.
Run MySql service (Control Panel->Administrative Tools->Services->MySql).

Change the ini file for work with mysql (C:\JC Persistence Framework\JCFramework\IniFiles\xmlpath.ini), and jcframework is ready for run.