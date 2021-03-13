Option Explicit

'ARR_ CAN BE LBOUND 0 OR 1, IS NO MATTER
'isExact is only there to keep the same arg sequence as vlookup
'return as array is preferable over collection bc sum/count work on an array
Function ALOOKUP3(lookup_value As Variant, arr_ As Variant, lookup_column_reference As Variant, Optional isApprox As Boolean = False, Optional returnAsArray As Boolean = True) As Variant

  'IF arr_ IS RANGE, SHRINK AND CONVERT TO ARR
  Dim arr As Variant
  If TypeName(arr_) = "Range" Then
    Dim r As Range
    Set r = Intersect(arr_.Parent.UsedRange, arr_)
    arr = r.Value
  Else
    arr = arr_
  End If
  'Debug.Print UBound(arr, 1) - LBound(arr, 1) + 1
  
  Dim n_col As Long
  n_col = 0
  If IsNumeric(lookup_column_reference) Then 'IS IT A NUMBER?
   If lookup_column_reference / Int(lookup_column_reference) = 1 Then 'IS IT AN INTEGER?
     n_col = lookup_column_reference
   End If
  Else 'GUESS ITS NOT A NUMBER
    Dim x As Long
    For x = 1 To UBound(arr, 2)
     If (UCase(lookup_column_reference) = UCase(arr(1, x))) Then 'TEST IF HEADER
       n_col = x
     End If
    Next x

    If n_col = 0 Then
     ALOOKUP3 = "#COL"
     Exit Function
    End If
    
  End If
  
  'MAKE AN ARRAY THE SAME SIZE AS SEARCH_ARR, arr
  'CREATING AND RESIZING AN ARRAY IS EXPENSIVE, SO LETS DO IT ONCE
  'AND MAKE IT THE MAX SIZE WE COULD POSSIBLY NEED, IN CASE EVERY ROW IS A HIT
  Dim indexOfHits() As Variant
  ReDim indexOfHits(LBound(arr, 1) To UBound(arr, 1))
  
  'LET US COUNT HOW MANY ROWS ARE HITS
  'ITS PROBABLY MORE ACCURATE TO THINK OF THIS AS WHAT OUR UBOUND WILL BE
  'BECAUSE IF LBOUND(arr,1) IS A WEIRD NUMBER, THINKING OF THIS AS THE NUMBER OF HITS WILL BE INACCURATE
  Dim uboundOfIndexOfHits As Long
  uboundOfIndexOfHits = LBound(arr, 1) - 1 'THIS WAS NOT MY IDEA THO: FROM USER TinMan IN A POST ON SO
  'Debug.Print "uboundOfIndexOfHits_before: " & uboundOfIndexOfHits
  
  'LOOP OVER arr TO STORE THE INDICES AND GET A UBOUND
  Dim y As Long
  For y = LBound(arr, 1) To UBound(arr, 1)
    If arr(y, LBound(arr, 2)) = lookup_value Then
     uboundOfIndexOfHits = uboundOfIndexOfHits + 1
     indexOfHits(uboundOfIndexOfHits) = y
    End If
  Next y
  'AFTER THIS LOOP, WE KNOW THE INDEX OF EACH ROW THAT IS A HIT
  'AND THE UBOUND OF OUR FINAL ARRAY, results
  'Debug.Print "uboundOfIndexOfHits_after: " & uboundOfIndexOfHits
  
  Dim results As Variant, index As Long
  If uboundOfIndexOfHits >= LBound(arr, 1) Then 'THERE ARE RESULTS
   'MAKE SURE results HAS ENOUGH SPACE TO STORE ALL THE DATA WE NEED
   'LBOUND TO LBOUND JUST ENSURES THERE IS 1 SPOT
   ReDim results(LBound(arr, 1) To uboundOfIndexOfHits, LBound(arr, 2) To LBound(arr, 2)) 'HERE, LBOUND TO LBOUND, BC I WANT ONLY 1 COLUMN
   
   'JUST CHECKING THAT IT IS AN mx1 ARRAY
   Dim m As Long, n As Long
   m = UBound(results, 1) - LBound(results, 1) + 1
   n = UBound(results, 2) - LBound(results, 2) + 1
   'Debug.Print "results, mxn: " & m & "x" & n
   
   'LOOP OVER HITS
   'REMEMBER, UBound(indexOfHits) IS PROBABLY A LOT BIGGER THAN uboundOfIndexOfHits
   For y = LBound(indexOfHits, 1) To uboundOfIndexOfHits
    index = indexOfHits(y) '1D-ARRAY
    results(y, LBound(arr, 2)) = arr(index, n_col) 'BUT results IS A 2D-ARRAY
    'Debug.Print "index = " & index & ": " & results(y, LBound(arr, 2))
   Next y
   'NOW, WE HAVE AN ARRAY, results(), WHICH CONTAINS ALL THE HITS
   
   If returnAsArray Then
     ALOOKUP3 = results
     Exit Function
   Else
     ALOOKUP3 = results(LBound(results, 1), LBound(arr, 2))
     Exit Function
   End If

  Else
   'THERE ARE NO RESULTS
   ALOOKUP3 = Array()
   Exit Function
  End If
  
End Function
