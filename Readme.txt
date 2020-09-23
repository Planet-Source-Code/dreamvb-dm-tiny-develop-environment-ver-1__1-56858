DM Tiny Develop Environment Scripting Version 1.
--------------------------------------------------------------------------------------------------------------

Written By Ben Jones


+ What is it
+ What’s new in version 1
+ How can I compile
+ Notes about this version
+ Updates Planned for next version

What is it:

DM Tiny Develop Environment Scripting is a small application that uses the VB Scripting language to build small robust 

Applications for window developers, or for small business that want to deploy applications, to there clients.

What new in this version:
--------------------------------------------------------------------------------------------------------------

New .NET type toolbar
Loads of bug fixes from Beta versions
Fixed project path bugs from the beta version
Added templates for both VBScript and JavaScript
Added codeinsight to the editor only supports VBScript at this time.
Added better coding support for controls. So no need for select case on controls.
Added config options
Added Loads of editor features:
Added Insert Menu
Added Convert text options
Added Plug-in support and plug-in template for you to use.
Added new about box
Added some new projects to help you get started.
Added ASCII insert for the editors
Added a list of all the built in functions also added to the code insight for the editor
Added new open project dialog. box.
Added Find Text to edit menu
Added Goto options to edit menu
Fixed but when opening a project while one was already open
Added project templates for both Java script and VBScript
Added support for Windows XP Themes
And more

How can I compile:
--------------------------------------------------------------------------------------------------------------

To Compile follow the steps below:

1 Open Visual Basic 6
2 Locate and open the project file DmDevEnv\DmDev Env.vbp
3 Form File menu select Make DmDevEnv.exe

4 Locate and open the project file DmDevEnv\exehead\exehead.vbp
5 Form File menu select Make exehead.exe
  Make sure the exe is compiled to DmDevEnv\exehead\exehead.exe

6 Close Visual Basic 6
7 Move to the folder DmDevEnv\exehead\
8 Run upx.bat
  This will then allow you to make the exehead smaller and also your compiled applications. This is however optional
  If you don't have upx.exe then you can get a copys from http://www.upx.org/
--------------------------------------------------------------------------------------------------------------

Notes about this version:

Please note that Java Script is not at the moment fully supported and some applications may not work.
I also found a small problem upon compiling the applications to exe with Java projects. I will how ever fix this soon. 

VBScript ones work fine.



Now I also noticed a but in my .NET type toolbar with the tools on. some times if you click the Ascii Char not all the items show, Well thay all load but after about 5 items all the chars are the same. also the over effect does not work. now this is strage as this does not happen all the time. so it makes me beleve that this is something to do with vb it's self.
anyway any ideas let me know or add them in. remmber I will creait anyone that helps me.
--------------------------------------------------------------------------------------------------------------

Updates Planned for next version:

# Adding more controls
# Adding more support for Java Script
# Fixing some more bugs
# Adding more MakeExe Options
# Some more fixes for the form designer and updates.
# Options to add modules to projects like in VB
# Adding console application support to new project options
# Adding resource file format
# Adding a framework
# VB Form Converter
# Automatic detection of new plug-ins
# And more.
--------------------------------------------------------------------------------------------------------------

Well that about all for this version hope you like it.
As always if you like to help in the development or have questions of things you like added let me know at

vbdream2k@yahoo.com

