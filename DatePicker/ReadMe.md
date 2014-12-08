# DatePicker Read Me!

I can't figure out  yet how to put all the Excel VBA code here (i.e., the actual form) so I'll just describe it 
   and the process for making this add-in
   
First, create a new Excel file, and open the VBA editor.

Add the ThisWorkbookModuleCode to the module for ThisWorkbook.

Insert a new module, DatePickerCode, and add the code to that.

Insert a new userform.  From the Tools, menu, select Additional Controls, and check the box for Microsoft 
  MonthView Control 6.0 (SP4).  Other versions would probably work.  Insert a MonthView control and edit 
  the form to look as you like.  I also renamed the control to MyMonthView.
  
Save the file as an add-in with macros, .xlam

In Excel Options, manage your add-ins and add the add-in you just created.

Enjoy!
