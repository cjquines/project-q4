# project-q4
A simple database, made for the sole purpose of my fourth quarter ICT project. 

![Form1](http://i.imgur.com/qkTCi12.png)

-	A database management system for handling students.
-	Supports storage of student IDs, names, sections, subject, grade, down payment, total payment, number of payments and contribution.
-	Automatically calculates grades based on DO s. 2015 no. 08. Choice of subject is used in computation. Grades automatically transmuted.
-	Automatically calculates payment contribution based on total payment, down payment and number of payments.
-	Can add, edit, and delete records from the student database.
-	Can navigate the records of the database.

### Programming notes:
-	Uses `navigate_records()` subroutine to navigate records, reducing code redundancy. `navigate_records()` fills text boxes and checks radio buttons automatically.
-	`btnCalculate_Click()` validates inputs. It checks whether all required fields are filled in (only required ones: for example, the user can fill in the computed grade and leave the raw scores and number of items blank and it will still work), whether all radio buttons have a choice, whether fields that contain numbers are numbers and whether the numbers are nonnegative. A limitation is that it does not check if numbers are integers, but only checks if they are numeric. Another limitation is that it does not sanitize inputs – [Robert ') DROP table_students ('](https://xkcd.com/327/) comes to mind here.
-	`btnCalculate_Click()` then calculates grades. Grades are calculated based on DepEd Order s. 2015 no. 08, based on raw scores and number of items for written work, performance task and quarterly assessment. Percentages are also based on DO s. 2015 no. 08. The grade is then transmuted according to the same order to produce the final computed grade.
-	`btnCalculate_Click()` then calculates payments, which is simply total minus down payment over number of payments, rounded up.
-	`transmutation()` is a helper function for `btnCalculate_Click()`. Given a double between 0 and 1, it transmutes it to a numerical rating according to DO s. 2015 no. 08. The function is my own, which I interpolated from the DepEd Order’s transmutation table.
-	`btnClear_Click()` both clears text boxes and unchecks radio buttons.
-	`rdoNumber_CheckedChanged()` activates when any of the radio buttons under payment_number are changed. If rdoNumberOther is checked, then it enables the textbox beside it, otherwise it disables it.
-	`btnAdd_Click()` adds a record to the database. It assigns local variables subject and payments to hold subject choice and number of payments, then adds it to the database through an SQL query.
-	`btnEdit_Click()` edits a given record. The ID is used to base the editing. It updates all the fields of a record given its ID. It may be buggy if the ID is the value being changed, because WHERE relies on the old value rather than the new one.
-	`btnDelete_Click()` deletes a given record, based on its ID. If the ID is changed, its behavior will be buggy, for the same reason as the sub `btnEdit_Click()` fails, because WHERE relies on the old value rather than the new one. The function asks for validation before deleting.
-	`btnFirst_Click()`, `btnPrevious_Click()`, `btnNext_Click()` and `btnLast_Click()` are used to navigate the records. They check whether or not the previous or next record does not exist.
-	`btnUpdate_Click()` disposes `Form1` and opens `Form2`. `Form2`’s `btnUpdate_Click()` then hides `Form2` and opens `Form1` again, making it reload the contents of the database. This is not the only way to reload the contents of the database, but this is easiest to implement.
-	All SQL keywords are capitalized, as is good programming practice. Data types are checked to match with each other without converting between doubles and integers, which may cause loss of data. Connections are opened and closed only when adding, editing, or deleting, preventing unwanted modifications to the database and unwanted memory usage.

### To-do:
-	Sanitize inputs, so we don’t have [Robert ') DROP TABLE students ('](https://xkcd.com/327/).
-	Validate to make sure inputs are integers and not doubles.
-	Change functionality of `btnEdit_Click()` and `btnDelete_Click()` to rely on index rather than ID, because the ID can be changed.
-	Make the database update without having to open another form. If this can be accomplished, make it update after every add, edit or delete.

### Screenshots:

Here's `Form1`, the heart of the program:

![Form1](http://i.imgur.com/qkTCi12.png)

`Form2`, the update form:

![Form2](http://i.imgur.com/jjFN8AB.png)

And `Quines.mdb`, the database used to store the student information:

![Quines.mdb](http://i.imgur.com/GKVQiht.png)

### Thanks:

To Sir Ruel Dogma, for teaching me how to use Visual Basic. He is an awesome teacher and this project won't be possible without him. Seriously, if he didn't assign this project for us I wouldn't make it.

### License:

This work is released under CC-BY 4.0. Look at [`LICENSE.md`](https://github.com/cjquines/project-q4/blob/master/LICENSE.md) for more information.
