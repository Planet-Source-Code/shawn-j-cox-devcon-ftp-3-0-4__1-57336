Revamp from yzyFTP. 

UI Translated 100% to english
NOW 50% SPANISH FREE! (variables/controls still bear spanish names)
Made some memory improvements
Major GUI changes
added extra error handling

Please vote!

from Fatty:
 I fixed a couple UI glitches/cosmetic problems.
Added a couple DoEvents to the upload/download functions so the forms of
the program update.

 There is no possible way to fix the freeze problem during connect. This
is because we use an API call to WinINet.dll, and the program is forced
to wait for a result before continuing.
 