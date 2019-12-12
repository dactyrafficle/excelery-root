# exceleryroot
1. tools for excel; 2. marshland plant

** .xlam **

When you save a file as .xlam, the default location at the time of writing this is:

c:\users\[username]\appdata\roaming\microsoft\addins\

In excel.exe, go to file>options>addins and at the bottom it says 'manage' next to a drop down which should say 'excel add-ins' and a button which reads 'go...'. Press that button! A dialogue box opens, and you can select the .xlam file.

Alternatively, you can go to the developper tab in the ribbon, and hit 'add-ins' which will produce the same dialogue box.

Once the add-in is activated, if you go to the vba editor, you should be able to see the .xlam file and its modules (unless it is protected) in the LHS column.

** tools > references **
Some libraries like FSO (file system object) or xmlhttp (to make get/post requests), or html parser - you can turn them on
also, to use the xlam functions in your current wbs vba code, you can include that from references too

** WORD FILES

activedocument is the main object, and that can be split up into sections

activedocument.sections.count

a section can have a header, range, footer; i think range is like body

because activedocument.sections(1).range is a thing that can have tables, but so is

activedocument.sections(1).headers(1).range.tables(1) etc etc.
