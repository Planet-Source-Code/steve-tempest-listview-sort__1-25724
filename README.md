<div align="center">

## ListView Sort


</div>

### Description

This Is a handy code for sorting a ListView Box by whichever column header is clicked.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Tempest](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-tempest.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-tempest-listview-sort__1-25724/archive/master.zip)





### Source Code

```
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView1.Sorted = True
  If ListView1.SortKey = ColumnHeader.Index - 1 Then
    If ListView1.SortOrder = lvwAscending Then
      ListView1.SortOrder = lvwDescending
    Else
      ListView1.SortOrder = lvwAscending
    End If
  Else
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = ColumnHeader.Index - 1
  End If
End Sub
```

