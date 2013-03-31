Attribute VB_Name = "Module2"
'************设备信息相关结构体************
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
DR(6) As DEVICEDRIVER '设备信息
DeviceCount As Integer '设备个数
End Type

Public db_auto_refresh As Boolean
'数据库存储开关
Public usedatabase As Boolean
