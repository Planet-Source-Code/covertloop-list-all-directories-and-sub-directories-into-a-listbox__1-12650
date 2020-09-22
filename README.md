<div align="center">

## List All Directories And Sub\-Directories Into A Listbox


</div>

### Description

This will list all the directories and subdirectories on your hard drive into a listbox.

Vote if ya wanna.
 
### More Info
 
Don't Forget The Declarations


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CovertLoop](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/covertloop.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/covertloop-list-all-directories-and-sub-directories-into-a-listbox__1-12650/archive/master.zip)

### API Declarations

```
Sub Listdir(path)
Dim d(1000)
Dir1.path = path
For lop = 0 To Dir1.ListCount - 1
d(cnt) = Dir1.List(lop)
cnt = cnt + 1
Next lop
For lop = 0 To cnt - 1
List1.AddItem d(lop)
cur_depth = cur_depth + 1
listdir d(lop)
Next lop
cur_depth = curr_depth - 1
End Sub
```


### Source Code

```
Add the declarations mentioned above, then put this in when you want it to run:
Listdir ("C:\")
```

