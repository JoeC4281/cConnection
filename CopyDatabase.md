**Usage**
```vbscript
Function CopyDatabase(
[ByVal DstDBFileName As String = ":memory:"],
[ByVal EncrKey As String],[ByVal CompactDstDB As Boolean]
) As cConnection
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