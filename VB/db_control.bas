Attribute VB_Name = "Module1"
Public Conn As New adodb.Connection
Public rs As New adodb.Recordset
'定义设备信息结构体
Public DS1 As DeviceState
Public Adodc1 As Adodc

Public Function adodbjet(Optional DBfile As String, Optional pwd As String) As adodb.Connection
On Error GoTo err
Dbpath = App.Path & "\LQD_database.mdb"
cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Dbpath
Conn.Open cn
err:
If err.Number Then
    MsgBox "数据库出错:" & err.Number
End
End If
End Function
Public Function databaseInit()
    Set cn = adodbjet
End Function

Public Function addRecord(Index As Byte)
SQL = "select * from 设备数据 ORDER BY 编号"
rs.Open SQL, Conn, 3, 3
If Not rs.EOF Or Not BOF Then
    rs.AddNew
    rs("设备ID") = DS1.DR(Index).id
    rs("设备名称") = DS1.DR(Index).name
    rs("日期") = Format(DS1.DR(Index).Date, "0000/00/00")
    rs("时间") = DS1.DR(Index).Time
    rs("温度") = Format(DS1.DR(Index).Temperature, "0.0")
    rs("风速") = Format(DS1.DR(Index).WindSpeed, "0.0")
    rs("风冷指数") = Format(DS1.DR(Index).WCI, "0.0")
    rs("等价制冷温度") = Format(DS1.DR(Index).ECT, "0.0")
    rs("相当温度") = Format(DS1.DR(Index).TEQ, "0.0")
    rs("冻伤危害性") = DS1.DR(Index).WeiHai
    rs("高强度作业") = DS1.DR(Index).HighLabor
    rs("中等强度作业") = DS1.DR(Index).MidLabor
    rs("安静作业") = DS1.DR(Index).LowLabor
    rs.Update
End If
rs.Close
End Function


Public Function 读取()
SQL = "select * from 设备数据 ORDER BY 编号"
rs.Open SQL, Conn, 1, 1
If rs.RecordCount <> 0 Then
        rs.MoveFirst
        List1.Clear
    Do While rs.EOF = False
        With List1
        .AddItem rs("name")
        End With
        rs.MoveNext
    Loop
  List1.ListIndex = List1.ListCount - 1
 End If
 rs.Close
End Function
Public Function 修改(id As Long, 数据 As String)
    SQL = "select * from 设备数据 where id=" & id & " ORDER BY 编号"
    rs.Open SQL, Conn, 1, 3
    If Not rs.BOF Or Not rs.EOF Then
        rs("name") = 数据
        rs.Update
    Else
                If (MsgBox("无id为" + Text2.Text + "的数据，是否创建新数据？", vbOKCancel + vbExclamation, "警告") = vbOK) Then
                            rs.AddNew
                            rs("name") = Text2.Text
                            rs.Update
                 End If
     End If
     rs.Close
 
    
End Function
Public Function 删除(id As Long)
    SQL = "Delete from 设备数据 where id=" & id & " ORDER BY 编号"
    Conn.Execute SQL
  
End Function
