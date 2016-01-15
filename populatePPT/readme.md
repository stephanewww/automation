#Overview

This pet project is about populating a PPTX file with data coming from an eXcel file. Writing this as a user story, the goal is

```UserStory

As a manager
I have to use standard ppt files AND to complete them with data coming from a standard excel sheet
To produce standardised presentation for my boss/clients/etc

```

#How To

In order to make it work you need to 

1. Download all the .bas and .cls files
2. Open your standard ppt file, open the macro part (and made them available)
3. Import the .bas and .cls files in your file (in the macro part)
4. Add the missing references (in the toolbar, menu Tools --> References):
  * Microsoft Excel component
  * Microsoft Script Control
  * Microsoft Scripting runtime
  * Microsoft Scriptlet librairy
  * Microsoft Shells control and Automation
  * Microsoft VBScript Regular Expression
  * Microsoft WMI Scripting Librairy
  * Microsoft JScript
  * Windows Script Host Object Model
  * WSHControllerLibrairy
5. Save your ppt file as a pptm or potm file to keep the macro working
6. Add with a mustache-like syntax the origin of your data
7. Launch the `populateMyPPT` macro

#Syntax

I choose a mustache-like syntax because it's easy to write and the pattern are not to hard to find without much of regular expression.

If you want to declare a variable in a shape you have to do it that way : `{{worksheetName!range}}`. For instance, if I want to use the value of the cell B3 in the worksheet named "S01", I have to write `{{S01!B3}}`.
By default you can format the mustache syntax directly in your PPT template to keep it after.

#Version

##0.1

@date : 20160114

@content :

* Mustache-like formating for one cell data
* Choose your Excel file
* Save the file automatically to `filename_populated.pptm`

