Attribute VB_Name = "modGlobals"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' online player variables
Public PlayersOnline() As Integer
Public High_Index As Integer
Public TotalPlayersOnline As Integer

' Message of the Day
Public MOTD As String

' Maximum classes
Public Max_Classes As Byte

' Scripting Globals
Global MyScript As clsSadScript
Public clsScriptCommands As clsCommands
Public DebugScripting As Boolean
