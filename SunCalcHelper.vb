Namespace SunCalcForDotNet
    ''' <summary>
    ''' Simulates the behaviour of the javascript Date type
    ''' Constructors:
    '''     New                 - Initializes an instance with the current date
    '''     New(ms As Double)   - Initializes an instance with the number of milliseconds passed as parameter
    '''     New(fecha As Date)  - Initializes an instance with the date (as .NET date type) pased as parameter
    '''     
    ''' To retrieve the date as .NET date tupe, use the GetNETDate method.
    ''' </summary>
    Public Class clsJSDate

        Private ReadOnly Jan11970 As New Date(1970, 1, 1, 0, 0, 0, 0)

        Public Property miliseconds As Double


        Public Sub New()

            Dim startTime As Date = Jan11970
            Dim endTime As Date = Now.ToUniversalTime
            Dim lapso As TimeSpan = endTime.Subtract(startTime)

            miliseconds = lapso.TotalMilliseconds
        End Sub

        Public Sub New(ms As Double)
            miliseconds = ms
        End Sub

        Public Sub New(fecha As Date)

            Dim startTime As Date = Jan11970
            Dim endTime As Date = fecha
            Dim lapso As TimeSpan = endTime.Subtract(startTime)

            miliseconds = lapso.TotalMilliseconds

        End Sub

        Public Function ValueOf() As Double
            Return miliseconds
        End Function

        Public Sub setUTCHours(h As Integer, Optional m As Integer = 0, Optional s As Integer = 0, Optional ms As Integer = 0)

            Dim NETDateTime As Date = Jan11970
            NETDateTime = NETDateTime.AddMilliseconds(miliseconds)

            Dim NewNETDateTime As Date = NETDateTime.Date
            NewNETDateTime.AddHours(h)
            NewNETDateTime.AddMinutes(m)
            NewNETDateTime.AddSeconds(s)
            NewNETDateTime.AddMilliseconds(ms)

            Dim startTime As Date = Jan11970
            Dim endTime As Date = NewNETDateTime
            Dim lapso As TimeSpan = endTime.Subtract(startTime)

            miliseconds = lapso.TotalMilliseconds

        End Sub

        Public Sub setHours(h As Integer, Optional m As Integer = 0, Optional s As Integer = 0, Optional ms As Integer = 0)

            Dim NETDateTime As Date = Jan11970
            NETDateTime = NETDateTime.AddMilliseconds(miliseconds)

            Dim NewNETDateTime As Date = NETDateTime.Date
            NewNETDateTime.AddHours(h)
            NewNETDateTime.AddMinutes(m)
            NewNETDateTime.AddSeconds(s)
            NewNETDateTime.AddMilliseconds(ms)

            Dim startTime As Date = Jan11970
            Dim endTime As Date = NewNETDateTime
            Dim lapso As TimeSpan = endTime.Subtract(startTime)

            miliseconds = lapso.TotalMilliseconds

        End Sub

        Public Function GetNETDate() As Date

            Dim NETDateTime As Date = Jan11970
            NETDateTime = NETDateTime.AddMilliseconds(miliseconds)

            Return NETDateTime
        End Function

    End Class


    ' Clases to store the results of SunCalcforDotNet functions
    ' -----------------------------------------------------------

    Public Class clsSunCoords
        Public Property dec As Double
        Public Property ra As Double
    End Class

    Public Class clsSunPosition
        Public Property azimuth As Double
        Public Property altitude As Double
    End Class

    Public Class clsTimes
        Public Property val As Double
        Public Property enumVal1 As Integer
        Public Property enumVal2 As Integer
        Public Property str1 As String
        Public Property str2 As String
    End Class

    Public Class clsSunTimes
        Public Property evento As String
        Public Property fecha As clsJSDate
    End Class

    Public Class clsMoonCoords
        Public Property ra As Double
        Public Property dec As Double
        Public Property dist As Double
    End Class

    Public Class clsMoonPosition
        Public Property azimuth As Double
        Public Property altitude As Double
        Public Property distance As Double
        Public Property parallacticAngle As Double
    End Class

    Public Class clsMoonIllumination
        Public Property fraction As Double
        Public Property phase As Double
        Public Property angle As Double
    End Class

    Public Class clsMoonTimes
        Public Property rise As Date
        Public Property sett As Date
        Public Property alwaysUp As Boolean
        Public Property alwaysDown As Boolean

        'Public Sub New()
        '    alwaysUp = False
        '    alwaysDown = False
        'End Sub
    End Class

    Public Class clsHelper

        Public Shared Function RadiansToDegrees(rads As Double) As Double

            Return rads * 180 / Math.PI
        End Function

    End Class
End Namespace