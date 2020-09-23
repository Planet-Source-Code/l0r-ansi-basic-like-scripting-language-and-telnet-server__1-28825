<div align="center">

## ANSI Basic Like Scripting Language and Telnet server


</div>

### Description

This code will allow you to run a script like text file with simple commands for full control over a telnet client. This code makes ANSI easy. This code was made in less than 24 hours so it may need further work, if you culd add to it please tell me. I plan on developing it further. Simply run it, then connect to localhost port 23 in an ANSI able telnet client. an example program is included. Please before flaming me on this code you might want to understand and appriciate it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-11-12 16:45:52
**By**             |[l0r](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/l0r.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ANSI\_Basic3460511112001\.zip](https://github.com/Planet-Source-Code/l0r-ansi-basic-like-scripting-language-and-telnet-server__1-28825/archive/master.zip)

### API Declarations

```
EXAMPLE PROGRAM: (in my language)
CLS
Color 2, 5, 2
Box 2, 2, 70, 14, 4
	Wait 1000
Box 4, 4, 71, 15, 3
	Wait 1000
Box 6, 6, 72, 16, 2
	Wait 1000
color 6, 4, 3
Box 8, 8, 73, 17, 1
	Wait 1000
	Move 12, 10
Print "I don't like them... Let's erase."
CLS
Move 2, 2
Print "The current time is: $TIME"
Move 3, 2
Print "The current date is: $DATE"
Wait 2000
Box 10, 30, 18, 3, 2
Move 11, 31
Print "IP: $IP"
Wait 3000
CLS
Move 2, 2
	Print "Cool huh?"
Move 5, 30
	Print "This session will terminate in: 5"
Move 5, 62
	Wait 1000
Print "4"
Move 5, 62
	Wait 1000
Print "3"
Move 5, 62
	Wait 1000
Print "2"
Move 5, 62
	Wait 1000
Print "1"
TERMINATE
End
LANGUAGE EXPLANATION:
Notes:
Color Function: Change the screens color modes. (only some terminals support this)
Color <Forground>, <Background>, <Extended>
Extended Attributes Code:
0	Reset all attributes
1	Bright
2	Dim
3    NONE
4	Underscore
5	Blink
6    NONE
7	Reverse
8	Hidden
Background & Foreground Colors
0	Black
1	Red
2	Green
3	Yellow
4	Blue
5	Magenta
6	Cyan
7	White
Box Function: Easy way to create an ansi box.
Box <X Position>, <X Position>, <Width>, <Height>, <Style>*
Styles:
1    Single Line
2    Double Line
3    Solid Line
4    Large Line
*you are limited to screen resolution remember that!
Move Function: Move the cursor to the specified location.
Move <X Position>, <Y Position>
Print Function: This will print text to the screen.
Print <Text> (aliases may be inserted here)
Wait Function: Pause for the specified amount of time before processing the next line.
Wait <Milliseconds>
Scroll Function: Scroll the screen from the current section.
Scroll <number>
Go Functions:
GO_UP Set cursor position up one from its current location.
GO_DN Set cursor position down one from its current location.
GO_BK Set cursor position back one from its current location.
GO_FW Set cursor position forward one from its current location.
Line Wrapping Function: Disabled or enabled terminal line wrapping.
WRAP <ON|OFF>
Other Functions:
   TERMINATE Disconnect user.
     RESET Reset all terminal settings.
      CLS Clear the terminals screen.
      END Specify the end of your code.
  SAVE_CURSOR Save current cursor position.
 RESTORE_CURSOR Restore saved cursor position.
  SAVE_ATTRIB Save current cursor position and attributes.
 RESTORE_ATTRIB Restore saved cursor position and attributes.
ALIASES:
  $IP Users ip address.
 $HOST Users hostname.
 $DATE Current server date.
 $TIME Current server time.
```





