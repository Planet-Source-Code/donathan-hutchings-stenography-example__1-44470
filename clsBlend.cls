VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'*  Author: Donathan Hutchings                                              *
'*  Date:   04/03/2003                                                      *
'*                                                                          *
'*  Purpose:    To embed a message into a graphic file with little or no    *
'*              degradation to the image view and to distribute the message *
'*              through out the file.                                       *
'****************************************************************************
Option Explicit

Event Progress(Value As Integer)

'This method embeds the text inside a graphic file and creates a key file
'to be used to read the message later.
Public Function Blend(SourceFileName As String, TargetFileName As String, _
                      KeyFileName As String, Text As String, _
                      Optional Errcode As Single = 0, _
                      Optional ErrMsg As String = "") As Boolean
                      
    Dim FF As Integer
    Dim FileLength As Single
    Dim TextLength As Single
    Dim ByteArray() As Byte
    Dim TextArray() As Byte
    Dim KeyArray() As Single
    Dim x As Single
    Dim Iterator As Single
    Dim Offset As Single
    
    'This value is used to preserve the image header.  This may change
    'depending on the format of the image.
    Const HeaderSize As Long = 200
    
    On Local Error GoTo ReportError
    
    'Determine the size of the image file.
    FileLength = FileLen(SourceFileName)
    
    'Determine the length of the message.   If the message
    'is larger than the image file, the image view maybe
    'distorted or unviewable.
    TextLength = Len(Text)
    
    'Convert to a byte array.
    TextArray() = StrConv(Text, vbFromUnicode)
    
    'Determine the offset to use when embeding the message.
    'This gives a a reliable starting position based on the
    'sizes of the image and message.
    Offset = (FileLength - HeaderSize) \ TextLength
    
    'Open the image and grab the binary data.
    FF = FreeFile
    ReDim ByteArray(FileLength)
    Open SourceFileName For Binary Access Read As #FF
        Get #FF, , ByteArray()
    Close #FF
    
    'Loop through the image data and begin placing one message byte at
    'a time.
    For x = Offset + HeaderSize To FileLength - 1 Step Offset
        DoEvents
        
        'Make sure we are not at the end of the message.
        If Iterator < TextLength Then
            'Swap the image byte with the message byte.
            ByteArray(x) = TextArray(Iterator)
            
            'Expand our key data array.  This array stores the locations
            'of each message byte.
            ReDim Preserve KeyArray(Iterator)
            
            'Store the location.
            KeyArray(Iterator) = x
            Iterator = Iterator + 1
            RaiseEvent Progress((Iterator / TextLength) * 100)
        Else
            'since the message is now in the image data, no sense in
            'iteration through the rest of the image array.
            Exit For
        End If
    Next
    
    'Create a new image file with the message inside.
    FF = FreeFile
    Open TargetFileName For Binary Access Write As #FF
        Put #FF, , ByteArray()
    Close #FF
    
    'Create the key file.
    FF = FreeFile
    Open KeyFileName For Binary Access Write As #FF
        Put #FF, , KeyArray()
    Close #FF
    
    Blend = True
    Exit Function
    
ReportError:
    Blend = False
    Errcode = Err.Number
    ErrMsg = Err.Description
    Exit Function
End Function

'This method reads the message stored in a graphics file.
Public Function Read(SourceFileName As String, KeyFileName As String, _
                     Text As String, Optional Errcode As Single = 0, _
                     Optional ErrMsg As String = "") As Boolean
                      
    Dim FF As Integer
    Dim FileLength As Single
    Dim TextLength As Single
    Dim ByteArray() As Byte
    Dim TextArray() As Byte
    Dim KeyArray() As Single
    Dim Buffer As String
    Dim x As Single
    Dim Iterator As Single
    
    On Local Error GoTo ReportError:
    
    'Determine our lengths.
    FileLength = FileLen(SourceFileName)
    TextLength = FileLen(KeyFileName)
    
    'Read the image byte data into a byte array.
    FF = FreeFile
    ReDim ByteArray(FileLength)
    Open SourceFileName For Binary Access Read As #FF
        Get #FF, , ByteArray()
    Close #FF
    
    'Read our message location data from the key file.
    FF = FreeFile
    ReDim KeyArray(TextLength)
    Open KeyFileName For Binary Access Read As #FF
        Get #FF, , KeyArray
    Close #FF
        
    'This loop determines where our actual location end is in
    'the array.  It's sort of a null trim.
    Iterator = 0
    For x = 0 To UBound(KeyArray)
        If KeyArray(x) = 0 Then
            Exit For
        End If
        Iterator = Iterator + 1
    Next
    
    'Loop through the image data to find our message and construct
    'our actual message.
    For x = 0 To Iterator - 1
        DoEvents
        
        'Resize our message byte array.
        ReDim Preserve TextArray(x)
        
        'Get the message byte from the image data.
        TextArray(x) = ByteArray(KeyArray(x))
        RaiseEvent Progress((x / Iterator) * 100)
    Next
    
    'Convert the byte data back to our original message.
    Text = StrConv(TextArray(), vbUnicode)
    
    Read = True
    Exit Function
    
ReportError:
    Read = False
    Errcode = Err.Number
    ErrMsg = Err.Description
    Exit Function
End Function


