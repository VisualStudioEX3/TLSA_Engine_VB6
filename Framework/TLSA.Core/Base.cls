VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CORE_EXCEPTIONS
    FILE_NOT_FOUND = SYS_ErrorCodes.SYS_FILENOTFOUND
    EMPTY_LIST = SYS_ErrorCodes.SYS_EMPTYLIST
    INVALID_PATH = SYS_ErrorCodes.SYS_INVALIDPATH
    FILE_NOT_LOAD = SYS_ErrorCodes.SYS_NOTCREATED
    HIGH_RESOLUTION_TIMERS_NOT_SUPPORT = SYS_ErrorCodes.SYS_HITIMERNOTSUPPORT
    ARGUMENT_EXCEPTION = 5
    UNKNOWN_EXCEPTION = SYS_ErrorCodes.SYS_UNKNOWNERROR
End Enum

Private varGenerics As New Core.Generics
Private varIO As New Core.IO
Private varMath As New Core.Math
Private varInfo As New Core.Information
Private varText As New Core.Text
Private varGUID As New Core.GUID

Public Property Get Generics() As Core.Generics
    Set Generics = varGenerics
End Property

Public Property Get IO() As Core.IO
    Set IO = varIO
End Property

Public Property Get Math() As Core.Math
    Set Math = varMath
End Property

Public Property Get Information() As Core.Information
    Set Information = varInfo
End Property

Public Property Get Text() As Core.Text
    Set Text = varText
End Property

Public Property Get GUID() As Core.GUID
    Set GUID = varGUID
End Property
