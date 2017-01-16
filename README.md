# FollowUpSharp
Reimplementation of [qfuAuto](https://github.com/jmarkman/qfuAuto) in C#

### The problem:

[qfuAuto](https://github.com/jmarkman/qfuAuto) relies on GUI automation via the [PyAutoGui](https://github.com/asweigart/pyautogui) library by Al Sweigart. Implementation of that script went poorly, however, mainly because of how others at my job wanted it to be used aka having it act as a user-launched script instead of something that sat on a machine in a corner. That means:

*No one has the same monitor size (resolution implied)
*Computer speed varies from person to person
*The person is rendered useless as the GUI automation does its thing

### The solution:

Reimplement it in C#. While Python has access to Windows events via Win32Com, there isn't a whole lot of documentation on that kind of thing. C# has the benefit of the [EPPlus](http://epplus.codeplex.com/) library as well as the Microsoft.Office.Interop libraries for the various Office products. I can also call upon the SQL library SqlClient to access our company database and perform the query we need (though this ability obviously isn't unique to C#), and when all is said and done, compile it to an .exe without having to go through the madness of py2exe or cxFreeze just to get an executable that can be shared with those not technolgically inclined.
