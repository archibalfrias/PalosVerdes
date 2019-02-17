Attribute VB_Name = "modDeclaration"
Public TournamentKey        As Double
Public WithTeamPlay         As Integer
Public TeamPlayer2Cnt       As Integer
Public NoofPlayerPerTeam    As Integer
Public AllowedTeam          As Integer
Public WithIndividualPlay   As Integer
Public HandicapDivisor      As Integer
Public DaysPlayerToPlay     As Integer
Public TournamentName       As String
Public TournamentRange      As String
Public ScoringType          As Long
Public PointsToCnt          As Long
Public TopHandicap          As Double
Public PointsToCntIndi      As Long
Public TeamAverage          As Long
Public TopIndex             As Double
Public dParGrossPoints      As Double
Public LocationCnt          As Double
Public LocationKey          As Long
Public TourNoOfPlays        As Long

Public PlayerKey            As Long

Public xlsApp               As Object

'------- UPPER CASE
Public Const ES_UPPERCASE = &H8&
Public Const GWL_STYLE = (-16)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
