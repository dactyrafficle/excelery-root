# excelery-root
1. tools for excel; 2. marshland plant

When you save a file as .xlam, the default location at the time of writing this is:

c:\users\[username]\appdata\roaming\microsoft\addins\

In excel.exe, go to file>options>addins and at the bottom it says 'manage' next to a drop down which should say 'excel add-ins' and a button which reads 'go...'. Press that button! A dialogue box opens, and you can select the .xlam file.

Alternatively, you can go to the developper tab in the ribbon, and hit 'add-ins' which will produce the same dialogue box.

Once the add-in is activated, if you go to the vba editor, you should be able to see the .xlam file and its modules (unless it is protected) in the LHS column.
