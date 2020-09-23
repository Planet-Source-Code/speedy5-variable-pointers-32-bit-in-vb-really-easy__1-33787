;vbpointers32.dll
;Made by: speedy
;Date: April 2002
;
;This is REALLY simple but it works.
;Even i couldn't believe how easy this was.
;I could have made it easier by having return
;values instead of changing the value in the
;method, like returning the pointer instead of
;passing the pointer's variable and changing it
;inside. The only problem was that valueFromPointer
;only let me return longs not singles or anything
;else. So i decided to be fair and have both
;methods change the variable inside. I hope you
;get it!
;
;By the way, I only have 3 days of asm experience.
;	:)



SEGMENT code USE32
GLOBAL _DllMain
_DllMain:   mov eax, 1
retn    12

;this function puts a pointer to 'from' into 'pointer'
;Declare Sub getPointer Lib "vbpointers32.dll" (ByRef pointer As Long, ByRef from As Any)
GLOBAL getPointer
getPointer:
enter 0, 0
push ebx

mov ebx, [ebp+12]
mov eax, [ebp+8]
mov [eax], ebx

pop ebx
leave
retn 8

;this functions gets the value from the variable directed by 'pointer' and puts it into 'into'
;Declare Sub valueFromPointer Lib "vbpointers32.dll" (ByRef pointer As Long, ByRef into As Any)
GLOBAL valueFromPointer
valueFromPointer:
enter 0, 0
push ebx

mov ebx, [ebp+8]
mov ebx, [ebx]
mov ebx, [ebx]

mov eax, [ebp+12]
mov [eax], ebx

pop ebx
leave
retn 8

ENDS