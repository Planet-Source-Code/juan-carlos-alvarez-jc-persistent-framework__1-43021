Decompress all files to "c:"

Check the option: "Use Folder Names"

If you unzip to a different place, edit the xml file to have 
the Access databases point to the full path of the Access.mdb (under test/xml/msAccess.xml in the .zip file structure) included with the package.

There is a Project Group (GrupoJCFramework.vbg) with all projects in it, ready for debug purposes. Open it with VB 6.

If you want to run the project "Test2" set a reference to the JCFramework.dll (first you need to generate the jcframework.dll).

If you want to test any of the tests (projects) change the line as explained in "xmlpath.ini" and set that project to initial by default.

Any comments to: jcarlos.alvarez@abitab.com.uy

Check for updates in:

http://sourceforge.net/projects/jcframework