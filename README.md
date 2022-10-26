# visualbasic

## Function To check if the Value is in Array
```
Function IsInArray(Value,MyArray)
on error resume next
 For i = 0 to UBound(MyArray)
  if Value = MyArray(i) then
  IsInArray = True
  exit function
  end if
 Next
 IsInArray = False
on error goto 0
end function
```
