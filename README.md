# WinControl: Automation Tool

WinControl is a scriptable tool for automating Windows applications. It allows you to control windows, simulate user interactions, and log automation processes.

## Features

### Writing Comments
Use # to write comments in your script

### Typing Control
Set the typing speed (delay between keystrokes)
```
SetDelay 100    # Delay in milliseconds
```
### Process Control
Attach and control a specific process
```
AttachProcess "notepad.exe"   # Attach to the Notepad process
BringToFront                  # Bring the window to the front
Sleep 1000                    # Wait for 1000 milliseconds
```
### Basic Input
Simulate mouse clicks and keystrokes
```
Click 100 200                 # Click at specified coordinates
RightClick 100 200            # Right-click at specified coordinates
DoubleClick 300 400           # Double-click at specified coordinates
SendKeystroke "Hello"         # Type text
```
### Modifier Keys
Single modifier keys
```
SendModKey "CTRL" "C"         # Simulate Ctrl+C
SendModKey "ALT" "TAB"        # Simulate Alt+Tab
```
Multiple modifier keys
```
SendMultiModKey "CTRL" "ALT" "DELETE"   # Simulate Ctrl+Alt+Delete
```
Special keys supported: TAB, ENTER, ESC, DELETE.
### Element Interactions
Interact with elements by properties
```
ClickElementByProperties "id" "class" "type"        # Click
RightClickElementByProperties "id" "class" "type"   # Right-click
DoubleClickElementByProperties "id" "class" "type"  # Double-click
```
Use "null" for properties you don't need.
### Conditional Execution
Perform actions based on conditions
```
IF ElementExists "id" "class" "type"
   SendKeystroke "Element found"
ENDIF

IF ElementNotExists "id" "class" "type"
   SendKeystroke "Element not found"
ENDIF

IF ContainsElementText "textbox_id" "textbox_class" "50011" "World"
   SendKeystroke "Text contains 'World'!"
ENDIF
```
### Variables
Define and use variables in your script
```
SET mytext "Hello World"
SendKeystroke "$mytext"  # Types "Hello World"

IF ContainsElementText "textbox_id" "textbox_class" "50011" "$mytext"
   # Your logic here
ENDIF
```
### Logging
Start, customize, and end logging with detailed messages
```
StartLog "AutomationTest"
LogHeader "Starting test"
Log "Normal message"
LogWarning "Warning message"
LogError "Error message"
EndLog
```
## Future Enhancements
Test control: Implement pass/fail reporting</br >
Offset clicking: Add support for offset clicks relative to an element</br >
Process identifier support: Further implement process attachment by PID</br >
Automatic logging: Enable automatic function logging</br >
Exception handling: Enhance error handling in both scripts and code</br >
Code refactoring: Simplify and beautify the code

## Requirements
Windows 10 or 11 SDK is required for compiling
