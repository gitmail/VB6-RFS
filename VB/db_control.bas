Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public rs As New ADODB.Recordset
'�����豸��Ϣ�ṹ��
Public DS1 As DeviceState
Public Function adodbjet(Optional DBfile As String, Optional pwd As String) As ADODB.Connection
On Error GoTo err
Dbpath = App.Path & "\LQD_database.mdb"
cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Dbpath
Conn.Open cn
err:
If err.Number Then
    MsgBox "���ݿ����:" & err.Number
End
End If
End Function
Public Function databaseInit()
    Set cn = adodbjet
End Function

Public Function addRecord(Index As Byte)
SQL = "select * from �豸����"
rs.Open SQL, Conn, 3, 3
If Not rs.EOF Or Not BOF Then
    rs.AddNew
    rs("�豸ID") = DS1.DR(Index).id
    rs("�豸����") = DS1.DR(Index).name
    rs("����") = Format(DS1.DR(Index).Date, "0000/00/00")
    rs("ʱ��") = DS1.DR(Index).Time
    rs("�¶�") = Format(DS1.DR(Index).Temperature, "0.0")
    rs("����") = Format(DS1.DR(Index).WindSpeed, "0.0")
    rs("����ָ��") = Format(DS1.DR(Index).WCI, "0.0")
    rs("�ȼ������¶�") = Format(DS1.DR(Index).ECT, "0.0")
    rs("�൱�¶�") = Format(DS1.DR(Index).TEQ, "0.0")
    rs("����Σ����") = DS1.DR(Index).WeiHai
    rs("��ǿ����ҵ") = DS1.DR(Index).HighLabor
    rs("�е�ǿ����ҵ") = DS1.DR(Index).MidLabor
    rs("������ҵ") = DS1.DR(Index).LowLabor
    rs.Update
End If
rs.Close
End Function


Public Function ��ȡ()
SQL = "select * from �豸����"
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
Public Function �޸�(id As Long, ���� As String)
    SQL = "select * from �豸���� where id=" & id
    rs.Open SQL, Conn, 1, 3
    If Not rs.BOF Or Not rs.EOF Then
        rs("name") = ����
        rs.Update
    Else
                If (MsgBox("��idΪ" + Text2.Text + "�����ݣ��Ƿ񴴽������ݣ�", vbOKCancel + vbExclamation, "����") = vbOK) Then
                            rs.AddNew
                            rs("name") = Text2.Text
                            rs.Update
                 End If
     End If
     rs.Close
 
    
End Function
Public Function ɾ��(id As Long)
    SQL = "Delete from �豸���� where id=" & id
    Conn.Execute SQL
  
End Function



