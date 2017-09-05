# VBAHook

It has only 1 file.

Please follow the instructions to use it:

1. Download the hook.bas;
2. Open Excel application (**32 bit**);
3. Open the Visual Basic Editor;
4. Import the file into a new module;
5. Run the "Hook" function in the Immediate window;

You can then open any Excel spreadsheet which has VBA password without entering the password.

## How does it work?

- Excel will call a system function to create the password dialog box.
- If user enters the right password and click OK, this function returns 1. If user enters the wrong password or click Cancel, this function returns 0.
- After the dialog box is closed, Excel checks the returned value of the system function
- If this value is 1, Excel will "think" that the password is right, hence the locked VBA project will be opened.
- The code below swaps the memory of the original function used to display the password dialog with a user defined function that will always return 1 when being called.
