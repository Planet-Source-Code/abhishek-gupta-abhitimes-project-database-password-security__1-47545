===============================
Maximize it for better viewing
===============================

This project gives you the full Database level security.
You cannot run your project untill you give the correct database password. Since the database is also password locked, nobody can open it also. If you make EXECUTABLE file of your application, then not to worry about someone else can run your application or open the database.
Good luck & please vote me if you really fount it useful.


This project contains following files:
---------------------------------------------------------------------
1.  Project1.vbp	//main Project file contains ADODB reference
2.  frmDBPassword.frm	//the form, where the password will be entered
3.  PrjCn.file		//Project connection file, including DBName & Project startup Form Name
4.  db1.mdb		//the Database file
5.  Form1.frm		//Project Startup Forms
---------------------------------------------------------------------

procedure for Beginners:-

Step 1: You mest set a password for the Database.
=======	To set the password - 
	(1) Create your database in MS Access 2000. (i prefer MS Access 2000)
	(2) now close database.
	(3) Start MS Access 2000 again. 
	(4) Open your database "Exclusively". To do this select Database name from the list and press the DOWN ARROW attached with OPEN Button. Select "Open Exclusive".
	(5) When your database is openned. Click Tools > Security > Set database password.
	(6) Enter your Password and close the database. 
	NOTE: To Unset the password, follow the steps 3-5 and select "Unset database Password"


Step 2:	Open Project in VB6.0 and make changes in the "PrjCn.file".
=======	(1) Set your Database Name P
	(2) Set your Project Startup File Name
	(3) Save the Changes

Step 3:	Run application. the default Password is "test"
=======

Step 4: 
=======	
If you want to merge this "project database security" into your project, Just include the files and set Project Startup Form Property to "frmDBConnection.frm"

---------------------------------------------------------------------
IMPORTANT: you are free to use and modify this project. Any new changes or ideas are welcomed.
---------------------------------------------------------------------

********************************************
*                  \\|//                   *
*                  (@ @)                   *
*   +---------oOO---(_)---OOo-----------+  *
*   |					|  *
*   | Mail your suggestions/comments at	|  *
*   |          Abhishek Gupta           |  *
*   |   Email: abhitimes@rediffmail.com |  *
*   |URL  : http://abhitimes.tripod.com |  *
*   |					|  *
*   +-----------------------------------+  *
*                 |__|__|                  *
*                  || ||                   *
*                 Ooo Ooo                  *
******************************************** 