VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Pointers"
   ClientHeight    =   4335
   ClientLeft      =   10455
   ClientTop       =   3435
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3270
   Begin VB.CommandButton Command2 
      Caption         =   "Arrays"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Variables"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'32 BIT POINTERS IN VB USING vbpointers32.dll
'Made by: speedy
'Date: April 10, 2002
'
'With vbpointers32.dll, you can now use 32 bit pointers in vb.
'VB doesn't let you do this, but you can get around this.
'You can do it the VarPtr() ways combined with Copymemory
'(the hard way) or you can do it the easy way with this dll.
'
'This dll lets you reference a variable thats less than or
'equal to 32 bits, like a byte, integer, long, single and
'booleans (no objects or strings yet), and put it into a
'long variable! You can also do arrays, multi-dim arrays
'and types! The source code of the dll is in:
'vbpointers32.asm.


'this function puts a pointer to 'from' into 'pointer'
Private Declare Sub getPointer Lib "vbpointers32.dll" (ByRef pointer As Long, ByRef from As Any)

'this functions gets the value from the variable directed by 'pointer' and puts it into 'into'
Private Declare Sub valueFromPointer Lib "vbpointers32.dll" (ByRef pointer As Long, ByRef into As Any)

Private Sub Command1_Click()
    Dim value1 As Single        'original value
    Dim value2 As Single        'where to put a copy of value1 into, using the pointer
    Dim pointer As Long         'the pointer to value1
    
    Cls
    Print "Single"
    Print "Value1 is assigned a value, then a get"
    Print "it's pointer. Value2 then copies Value1"
    Print "by using the pointer:"
    Print
    
    'give value1 a random number
    value1 = Rnd * 10
    Print "Value1:"; value1
    Print "Value2:"; value2
    
    Print
    
    'get the pointer of value1
    getPointer pointer, value1
    Print "getPointer pointer, value1"
    Print "Pointer to Value1:"; pointer
    
    'change value1 to prove we're not cheating
    value1 = Rnd * 10
    Print "Now change Value1:"; value1
    
    
    Print
    
    'retrieve the value directed by our pointer
    valueFromPointer pointer, value2
    Print "valueFromPointer pointer, value2"
    Print "Value2:"; value2
End Sub

Private Sub Command2_Click()
    Dim values(3) As Long       'our array of numbers
    Dim value2 As Long          'where to put a copy of a value in the array into, using the pointer
    Dim pointer As Long         'the pointer to the array
    Dim arrnum As Long          'temp storage: the array element we want
    
    Cls
    Print "Long"
    Print "An array of 0-3 elements are given random"
    Print "numbers. A pointer is given to the first"
    Print "element. Enter the element number you wish"
    Print "to retrieve with the pointer:"
    Print
    Print "By the way, you can also do this with multi- "
    Print "dimensional arrays."
    Print
    
    'give the array random values
    For i = 0 To 3
        values(i) = Rnd * 10000
        Print "Array(" & i & "):"; values(i)
    Next i
    Print "Value2:"; value2
    
    Print
    
    'get the pointer to the base of the array
    getPointer pointer, values(0)
    Print "Pointer to Array:"; pointer
    
    'get our element number
    arrnum = InputBox("Element Number (0-3):")
    
    'redo our pointer by adding 4 bytes (32 bits) per every element
    'until we get want we want
    pointer = (pointer + CLng(arrnum * 4))
    Print "Element Number " & arrnum & " at " & pointer
    
    'get the array element using the pointer
    valueFromPointer pointer, value2
    Print "Value2:"; value2
End Sub

Private Sub Form_Load()
    FontBold = True
    Print "vbpointers32.dll"
    Print
    FontBold = False
    Print "Pointers with this dll only work with these"
    Print "data types:"
    Print
    FontBold = True
    Print "Byte"
    Print "Boolean"
    Print "Integer"
    Print "Long"
    Print "Single"
    Print
    Print "It will also work with arrays, multi-dim"
    Print "arrays of them, and types."
    FontBold = False
    Print
    Print "Click a Example:"
    
    
    'redo the random numbers
    Randomize
End Sub
