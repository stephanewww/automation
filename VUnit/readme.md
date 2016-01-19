#Purpose

This class is design for Unit Testing and TDD for VBA.
It's a really easy to use and short implementation of Unit Testing, far less extended that VBUnit or VBLiteUnit. I'm really impressed to see what you can do with those module / dll, but :

1. I don't like dlls
2. It's far too complex for most of the use in standard VBA dev (for me at least)
3. I don't like the part where you have to correct your Unit Test before you can run them

#How To use

Pretty simple use : 

* Copy/past the VUnit.cls as a new Module Class in your Office Application
* Add the requiered references
* Write your Unit test in a new module like the following example

```vba

Sub atestClassTest()

Dim cTest As New VUnit

cTest.init 'initializing the new VUnit Object

cTest.assertTrue "True : 1<2", 1<2
cTest.assertFalse "False 1>2", 1>2
cTest.assertEquals "Equals 1 = 1", 1,1
cTest.assertNotEquals "A<>B", A,B

End Sub

```

# Content

You have basically just the strict necessary for unit testing :

* AssertEquals when you expect an equality
* AssertNotEquals else
* AssertTrue and AssertFalse which are self explanatory
* addUTest when you need a more complex test

If you want to add new test or standard UnitTest, feel free

#Changelog

##V0.02 18/01/2016

* rewrite the whole AddUtest and basic Unit Tests
* fix for presenting the results with the number of the test, the value and the name

## V0.01 10/04/2014

first draft
