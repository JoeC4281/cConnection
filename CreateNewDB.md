**Usage**
```vbscript
Function CreateNewDB(
[ByVal FileName As String = ":memory:"],
[ByVal EncrKey As String],
[ByVal EnableVBFunctions As Boolean = True]
) As Boolean
```

**Example VBScript Code**

```vbscript
Dim rc6 'as Object
Dim dbf 'as Object

Set rc6 = CreateObject("rc6.cConnection")

With rc6
  .CreateNewDB
  .ExecCmd("CREATE TABLE S (str TEXT NOT NULL)")
  Set dbf=.CreateCommand("Insert Into S Values(@str)")
  dbf.SetText 1, "Test"
  dbf.Execute
  .CopyDatabase "R:\Test.db"
End With 

Set rc6 = nothing
Set dbf = nothing
```
**Example Harbour/HBrun Code**
```foxpro
PROC Main

   LOCAL rc6
   LOCAL dbf

   rc6 := win_oleCreateObject( "rc6.cConnection" )

   rc6:CreateNewDB()
   rc6:ExecCmd( "CREATE TABLE S (str TEXT NOT NULL)" )
   dbf := rc6:CreateCommand( "Insert Into S Values(@str)" )
   dbf:SetText(1, "Test")
*  dbf:SetInt32(2, 1984)
*  dbf:SetDouble(3, 3.14)
*  dbf:SetBlob(4, <What goes here?>)
   dbf:Execute()
   rc6:CopyDatabase("R:\Test.db")

   RETURN

```

