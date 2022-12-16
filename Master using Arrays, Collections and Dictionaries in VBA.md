# master using Arrays, Collections and Dictionaries in VBA

[TOC]

#### How to use Arrays with Ranges

Range

> ​    Dim rg As Range
> ​    Set rg = shData.Range("A1").CurrentRegion
>
> ​	Dim i As Long
> ​	For i = 1 To rg.Rows.Count
> ​    Debug.Print rg.Cells(i, 1), rg.Cells(i, 2)
> ​	Next i  

Read from Range to Array

> Dim arr As Variant
>     arr = shData.Range("A1").CurrentRegion.Value2
>
> Dim i As Long
> For i = LBound(arr, 1) To UBound(arr, 1)
>     Debug.Print arr(i, 1), arr(i, 2)
> Next i

From Array back to Range:

> shData.Range("F10").Resize(UBound(arr, 1), UBound(arr, 2)) = arr		

Example 1:  *Ex1_ArrayFilter*

Private Sub *ArrayCopyFilterExample*()
    
    ' 1. Get the range and place in array
    Dim arr As Variant
    arr = shData.Range("a1").CurrentRegion.Value2
    
    ' 2. Store the row and column size
    Dim rowCount As Long, columnCount As Long
    rowCount = UBound(arr, 1)
    columnCount = UBound(arr, 2)
    
    ' 3. Create the output array
    Dim outputArray As Variant
    ReDim outputArray(1 To rowCount, 1 To columnCount)
    
    Dim i As Long, j As Long
    Dim salesPerson As String, currentRow As Long
    currentRow = 0
    ' 4. Read through the data and filter
    For i = 1 To rowCount
        salesPerson = arr(i, 2)
        ' 4a. Check if the sales person is Jenny
        If salesPerson = "Jenny" Then
            currentRow = currentRow + 1
            ' 4b. Copy the current row to the new array
            For j = 1 To columnCount
                outputArray(currentRow, j) = arr(i, j)
            Next j
        End If
    
    Next i
    ' 5. Write out array
    shData.Range("F1").CurrentRegion.Offset(1).ClearContents
    shData.Range("F2").Resize(currentRow, UBound(outputArray, 2)).Value2 = outputArray

End Sub



#### Arrays vs Dictionaries/Collections

Example 2:  Ex2_Unique_Array2D

Example 2B: Using Transpose

- Collection basics

  ​        

> Dim coll As New Collection
> coll.Add "Apple"
> coll.Add "Orange", "Prod001"
> coll.Add item:="Pear", key:="Prod002"



> Debug.Print "Item at key position 1 is: " & coll(1) ' retrieve Apple
> Debug.Print "Item at key Prod001 is: " & coll("Prod001") ' retrieve Orange
> Debug.Print "Item at key Prod002 is: " & coll("Prod002") ' retrieve Pear

> Dim item As Variant
> For Each item In coll
>     Debug.Print item
> Next item



- Example 2:  Unique values - Collection	

Sub ReadMarksCollUnique()

    ' 1. Get the range
    Dim arr As Variant
    arr = shEx2UniqueIn.Range("A1").CurrentRegion.Value2
    
    ' 2. Read through the data
    Dim i As Long, lastName As String
    Dim coll As New Collection
    
    ' For this to skip the error you must have the default error trapping settings i.e:
    ' Tools->Options, General tab, error trapping section, select "Break on unhandled errors"
    On Error Resume Next
    For i = 2 To UBound(arr, 1)
        lastName = arr(i, 1)
        ' 2.A Add item to collection - if already exists then ignore error
        coll.Add item:=lastName, key:=lastName
    Next i
    On Error GoTo 0
    
    ' 3. Create the output array
    Dim outputArray As Variant
    ReDim outputArray(1 To coll.Count, 1 To 1)
    
    ' 4. Write from collection to output array
    For i = 1 To coll.Count
        outputArray(i, 1) = coll(i)
    Next i
    
    ' 5. Write to the worksheet
    shEx2UniqueOut.Range("A1").CurrentRegion.Offset(1).ClearContents
    shEx2UniqueOut.Range("A2").Resize(UBound(outputArray, 1), 1) = outputArray

End Sub

- Dictionary Basics

> Dim dict As New Dictionary
> dict.Add "Prod001", "Apple"
> dict.Add "Prod002", "Orange"
> dict.Add "Prod003", "Pear"
>
> Debug.Print "Prod001 exists is " & dict.Exists("Prod001")
> Debug.Print "Prod999 exists is " & dict.Exists("Prod999")

>
> Dim key As Variant
> For Each key In dict.Keys
>     Debug.Print key, dict(key)
> Next key

- Example 2: Unique - Dictionary

Sub ReadMarksDictUnique()

    ' 1. Get the range
    Dim arr As Variant
    arr = shEx2UniqueIn.Range("A1").CurrentRegion.Value2
    
    ' 2. Declare and create the dictionary
    Dim dict As New Dictionary
    
    Dim i As Long, lastName As String
    ' 3. Read through the data
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        lastName = arr(i, 1)
        ' 3.A Add to the dictionary
        If dict.Exists(lastName) = False Then
            dict.Add lastName, 0
        End If
    Next i
    
    ' 4. Write out the names
    shEx2UniqueOut.Range("A1").CurrentRegion.Offset(1).ClearContents
    shEx2UniqueOut.Range("A2").Resize(dict.Count, 1) = WorksheetFunction.Transpose(dict.Keys)

End Sub

- Example 3: Sum Data Dictionary

Sub *ReadMarksDictSumValues*()

    ' 1. Get the range
    Dim arr As Variant
    arr = shEx3SalesSum.Range("A1").CurrentRegion.Value2
    
    ' Add reference using Tools->Reference and check "Microsoft Scripting Runtime"
    Dim dict As New Dictionary
    
    Dim i As Long, salesPerson As String, salesAmount As Currency
    ' 2. Read through the data
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        salesPerson = arr(i, 2)
        salesAmount = arr(i, 4)
        ' 2.A Add to the dictionary
        dict(salesPerson) = dict(salesPerson) + salesAmount
    Next i
    
    ' 3. Write out the results
    With shEx3SalesSum
        .Range("F1").CurrentRegion.Offset(1).ClearContents
        .Range("F2").Resize(dict.Count, 1) = WorksheetFunction.Transpose(dict.Keys)
        .Range("G2").Resize(dict.Count, 1) = WorksheetFunction.Transpose(dict.Items)
    End With

End Sub

- Example 4: Sum Multiple fields

Private Sub ReadMarksDictSumMulti()

    ' 1. Get the range
    Dim arr As Variant
    arr = shEx4SalesMulti.Range("A1").CurrentRegion.Value2
    
    Dim outputArray As Variant
    ReDim outputArray(1 To UBound(arr, 1), 1 To 3)
    
    ' 2. Read through the data
    
    Dim dict As New Dictionary, salesPerson As String
    Dim i As Long, newRowID As Long, outputRowID As Long
    newRowID = 0
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        salesPerson = arr(i, 2)
        
        ' 2.A If the salesPerson doesn't exist then add a new record for them
        If dict.Exists(salesPerson) = False Then
            ' Get new row id
            newRowID = newRowID + 1
            ' Add sales person as key and new row id as value
            dict.Add salesPerson, newRowID
            ' Add the sales person's data to the row at newrowID in the output array
            outputArray(newRowID, 1) = salesPerson
        End If
        
        outputRowID = dict(salesPerson)
        ' 2.B Update the salesperson record by adding the new values to the totals
        outputArray(outputRowID, 2) _
                    = outputArray(outputRowID, 2) + arr(i, 3)
        outputArray(outputRowID, 3) _
                    = outputArray(outputRowID, 3) + arr(i, 4)
       
    Next i
    
    ' 3. Write out the results
    shEx4SalesMulti.Range("F1").CurrentRegion.Offset(1).ClearContents
    shEx4SalesMulti.Range("F2").Resize(UBound(outputArray, 1), UBound(outputArray, 2)) _
        = outputArray

End Sub



### Explained in Simple English

- The Basic Example

  Sub *UseClass*()

      Dim oCust1 As New clsCustomer
      
      oCust1.firstName = "John"
      oCust1.lastName = "Smith"
      oCust1.Unite

  End Sub

  '***clsCustomer***
  `Public firstName$,lastName$`

  *Unite*() => `Debug.Print firstName & " " & lastName`

  

- Practical Use of Class

  1. Extending the Collection (Sort & Reverse)

     Sub Collection()

         Dim coll As New Collection
         
         coll.Add "D": coll.Add "B"
         coll.Add "A": coll.Add "E"
         coll.Add "C"
         
         DebugPrint coll, "Before sort:"
         QuickSort coll, 1, coll.Count
         DebugPrint coll, "After sort:"

     End Sub

     Sub *Collection*()

     ```
     Dim coll As New EMMCollection
     
     coll.Add "D"
     coll.Add "B"
     coll.Add "A"
     coll.Add "E"
     coll.Add "C"
     
     coll.DebugPrint "Before sort:"
     coll.Sort: coll.DebugPrint "After sort:"
     coll.Reverse: coll.DebugPrint "After sort:"
     ```

     End Sub

     [EMMCollection cls]()

  2. Filtering data on a UserForm



### How Objects Really Work in Memory

- Let and Set
  Sub *BasicDim*()

      Dim total As Long, name As String
      Let total = 67: Let name = "Bill"
      
      Dim coll As Collection
      Set coll = New Collection
      
      MsgBox total

  End Sub

- An Example of Set and New
  Sub *ExampleDimSet*() 
  
      Dim customers As New Collection
      
      Dim rg As Range
      Set rg = shData.Range("a1").CurrentRegion
      
      Dim i As Long, customer As clsCustomer
      For i = 2 To rg.Rows.Count
          Set customer = New clsCustomer
          customer.firstName = rg.Cells(i, 1).Value
          customer.lastName = rg.Cells(i, 2).Value
          customer.country = rg.Cells(i, 3).Value
          customers.Add customer
      Next i
  End Sub
  
- Setting one object to another

  Sub *Copy*()

  ```
  Dim customer1 As New clsCustomer
  
  customer1.firstName = "Jane"
  customer1.lastName = "Murphy"
  
  'Dim customer2 As clsCustomer: Set customer2 = customer1
   Dim customer2 As New clsCustomer
   customer2.firstName = customer1.firstName
  
   customer2.firstName = "Bob"
   customer1.firstName = "Tim"
  ```
  
  End Sub
  
### Class Interfaces

  - The *shData* Sheet
  
    | Amount | Interest Type |
    | ------ | ------------- |
    | 1000   | A             |
    | 2000   | B             |
    | 1300   | A             |
    | 2400   | B             |
  
  - The Basic Modules:
  
    - Section 1
    
      Sub Main()
    
          Dim rg As Range
          Set rg = shData.Range("a1").CurrentRegion
          
          Dim oInterestA As clsInterestA, oInterestB As clsInterestB
          
          Dim amount As Double, interestType As String
          
          Dim i As Long, result As Double
          For i = 2 To rg.Rows.Count
              amount = rg.Cells(i, 1).Value
              interestType = rg.Cells(i, 2).Value
              
              If interestType = "A" Then
                  Set oInterestA = New clsInterestA
                  result = oInterestA.Calculate(amount)
              ElseIf interestType = "B" Then
                  Set oInterestB = New clsInterestB
                  result = oInterestB.Calculate(amount)
              Else
                  MsgBox "Invalid type " & interestType
              End If
              
              Debug.Print result
          Next i
    
      End Sub
      
      ```
      clsInterestA (B):
      Function Calculate(ByVal amount As Double) As Double
          Calculate = amount * 1.1 (1.5)
      End Function
      *OutPut : = >  1100 3000 1430 3600* 
      ```

  - Section 2

    ```
    clsInterestA (B):
    Private m_Amount As Double
    
    Sub Calculate(ByVal amount As Double)
        m_Amount = amount * 1.1 (1.5)
    End Sub
    Sub PrintResult()
        Debug.Print TypeName(Me) & ": " & m_Amount
    End Sub
    ```

    Sub Main()

        Dim rg As Range
        Set rg = shData.Range("a1").CurrentRegion
        
        Dim oInterestA As clsInterestA, oInterestB As clsInterestB
        
        Dim amount As Double, interestType As String
        
        'Read through the data
        Dim i As Long, result As Double
        For i = 2 To rg.Rows.Count
            amount = rg.Cells(i, 1).Value
            interestType = rg.Cells(i, 2).Value
            
            If interestType = "A" Then
                Set oInterestA = New clsInterestA
                oInterestA.Calculate amount
                'oInterestA.PrintResult
            ElseIf interestType = "B" Then
                Set oInterestB = New clsInterestB
                oInterestB.Calculate amount
                'oInterestB.PrintResult
            Else
                MsgBox "Invalid type " & interestType
            End If
            
            'some code
            
            'print
            If interestType = "A" Then
                oInterestA.PrintResult
            ElseIf interestType = "B" Then
                oInterestB.PrintResult
            Else
                MsgBox "Invalid type " & interestType
            End If
            
        Next i

    End Sub

  - Use Interface

    Sub *Main*()

        Dim rg As Range
        Set rg = shData.Range("a1").CurrentRegion
        
        Dim oInterest As iInterest
        Dim amount As Double, interestType As String
        
        'Read through the data
        Dim i As Long, result As Double
        For i = 2 To rg.Rows.Count
            amount = rg.Cells(i, 1).Value
            interestType = rg.Cells(i, 2).Value
            
            Set oInterest = ClassFactory(interestType)
            oInterest.Calculate amount
            'some code
            
            'print
            oInterest.PrintResult
            
        Next i

    End Sub

    Function *ClassFactory*(ByVal *interestType* As String) As *iInterest*

        Dim oInterest As iInterest
        
        If interestType = "A" Then
            Set oInterest = New clsInterestA
        ElseIf interestType = "B" Then
            Set oInterest = New clsInterestB
        Else
            MsgBox "Invalid type " & interestType
        End If
        
        Set ClassFactory = oInterest

    End Function

    


### Build a Trading Simulator

| List A | List B |      | list2Only | both | list1Only |
| ------ | ------ | ---- | --------- | ---- | --------- |
| A1     | B3     |      | B3        | CC1  | A1        |
| A2     | B1     |      | B1        | CC2  | A2        |
| CC1    | CC1    |      | B5        |      | A3        |
| A3     | B5     |      | B4        |      | A4        |
| A4     | CC2    |      |           |      |           |
| CC2    | B4     |      |           |      |           |
| A2     | B5     |      |           |      |           |
| CC2    | B3     |      |           |      |           |

- modCompare :

  - Public Enum *eResultType*
        `list1Only = 1`
        `both = 2`
        `list2Only = 3`

  - *MainCompare*

        Dim resultType As eResultType
        resultType = eResultType.list2Only
        
        Dim rgBase As Range, rgCompare As Range, rgResult As Range
        Set rgBase = shData.Range("A2:A9")
        Set rgCompare = shData2.Range("A2:A9")
        Set rgResult = shResult.Range("A1")
          
        Dim dictBase As Dictionary
        Set dictBase = ReadList(rgBase.Value)
        
        Dim dictResult As Dictionary
        Set dictResult = CompareLists(dictBase, rgCompare.Value, resultType)
        
        PrintDictionary dictResult
        
        Call WriteResult(rgResult, dictResult, resultType)

  - ReadList

    Public Function *ReadList*(arr As Variant) As *Dictionary*

        Dim dict As New Dictionary
        
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            dict(arr(i, 1)) = 0
        Next i
        
        Set ReadList = dict

    End Function

  - CompareLists

    Function *CompareLists*(dict As *Dictionary*, arr As *Variant*, resultType As *eResultType*) As *Dictionary*
        
        Dim dictResult As New Dictionary, dict2Only As New Dictionary
        
        Dim i As Long, item As Variant
        For i = LBound(arr) To UBound(arr)
            item = arr(i, 1)
            If dict.Exists(item) = True Then
                dictResult(item) = 0
                dict.Remove item
            Else
                dict2Only(item) = 0
            End If
        Next i
        
        If resultType = both Then
            Set CompareLists = dictResult
        ElseIf resultType = list1Only Then
            Set CompareLists = dict
        Else
            Set CompareLists = dict2Only
        End If

    End Function

  - PrintDictionary

    Sub *PrintDictionary*(dict As *Dictionary*)

        Dim key As Variant
        For Each key In dict
            Debug.Print key, dict(key)
        Next key

    End Sub

  - WriteResult

    Sub *WriteResult*(rg As *Range*, dict As *Dictionary*, resultType As *eResultType*)

        rg.CurrentRegion.ClearContents
        
        rg.Value = "Result"
        
        rg.Offset(1).Resize(dict.Count, 1).Value = WorksheetFunction.Transpose(dict.Keys)

    End Sub

