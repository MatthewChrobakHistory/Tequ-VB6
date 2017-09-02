Attribute VB_Name = "modGlobals"
Option Explicit

Public Running As Boolean
Public Connecting As Boolean

Public OnlineMode As Boolean

Public MyIndex As Long
Public LastMap As Long 'this variable is used to avoid rendering the picMap over and over again.

'MAP EDITOR
Public InEditor As Boolean
Public LabelIndex As Byte
Public TransparentLabels As Boolean
Public CurMap As Long
