Attribute VB_Name = "Module2"
'************�豸��Ϣ��ؽṹ��************
Public Type DEVICEDRIVER
id As Integer
name As String
Date As String
Time As String
Temperature As Single
WindSpeed As Single
WCI As Single
ECT As Single
TEQ As Single
WeiHai As String
LowLabor As String
MidLabor As String
HighLabor As String
End Type

Public Type DeviceState
DR(6) As DEVICEDRIVER '�豸��Ϣ
DeviceCount As Integer '�豸����
End Type

Public db_auto_refresh As Boolean
'���ݿ�洢����
Public usedatabase As Boolean
