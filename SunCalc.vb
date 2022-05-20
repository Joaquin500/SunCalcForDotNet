Namespace SunCalcForDotNet

#Region "Info"
    ' --------------------------------------------------------------------------------------------------
    ' Translation into VB.NET of Vladimir Agafonkin's SunCalc for Javascript
    ' (c) 2022, Joaquin Suarez
    ' https://github.com/Joaquin500/SunCalcForDotNet
    '
    '
    ' (c) 2011-2015, Vladimir Agafonkin
    ' SunCalc Is a JavaScript library for calculating sun/moon position And light phases.
    ' https://github.com/mourner/suncalc
    ' --------------------------------------------------------------------------------------------------

    ' I have chosen to keep the library structure as similar as possible to the original SunCalc.
    ' Thus, those who have already used the original SunCalc will be familiar with this version.
    ' I have maintained original Vladimir code remarks. Any remark begining with @Vladimir: is from him.
    ' Any regular remark is mine.

    ' The library is made of two (2) files. Make sure you have both when you try to use it. These files are:
    '   * SunCalc.vb            - Main file. Contains the SunCalc code
    '   * SunCalcHelper.vb      - Helper classes for the main file

    ' --------------------------------------------------------------------------------------------------

    ' Basic Usage:
    '
    ' Imports SunCalcForDotNet
    '
    ' ' ...surrounding class/method...
    '
    ' Dim sc as New SunCalc
    ' Dim fecha As Date = Now     ' The date you want Sun information for
    ' Dim timezone As New TimeSpan(2, 0, 0)   ' Your time zone
    '
    ' ' Call SunCalc to get the time for Sun's main events on the date passed
    ' ' Create a dictionary variable to store the results
    ' Dim suntimes As Dictionary(Of Integer, Date) = sc.getSunTimes(fecha, myLat, myLng)
    '
    ' ' Extract the desired sun event dates from the dictionary
    ' Dim sunrise As Date = suntimes(enumSunTimes.sunrise).Add(timezone)
    ' Dim sunset As Date = suntimes(enumSunTimes.sunset).Add(timezone)
    '
    ' TextBox1.Text = String.Empty
    ' TextBox1.Text = TextBox1.Text & "sunrise: " & sunrise.ToString & vbCrLf
    ' TextBox1.Text = TextBox1.Text & "sunset: " & sunset.ToString & vbCrLf
    '
    ' ' ...surrounding class/method...

    ' --------------------------------------------------------------------------------------------------

    ' Public Methods:
    '
    ' Public Function getSunPosition(jsDate As clsJSDate, lat As Double, lng As Double) As clsSunPosition
    ' Public Sub addSunTime(angle As Double, riseName As String, setName As String)
    ' Public Function getSunTimes(jsdate As clsJSDate, lat As Double, lng As Double, Optional height As Double = 0.0R) As Dictionary(Of String, clsJSDate)

    ' Public Function getMoonPosition(jsDate As clsJSDate, lat As Double, lng As Double) As clsMoonPosition
    ' Public Function getMoonIllumination(jsDate As clsJSDate) As clsGetMoonIllumination
    ' Public Function getMoonTimes(jsDate As clsJSDate, lat As Double, lng As Double, inUTC As Boolean) As clsGetMoonTimes

    ' ========================================================================================================================================================
#End Region


    ' to ease the data retrieval from the dictionary of results of getSunTimes()
    Public Enum enumSunTimes As Integer

        solarNoon = 1
        nadir = 2
        sunrise = 3
        sunset = 4
        sunsetStart = 5
        sunriseEnd = 6
        dawn = 7
        dusk = 8
        nauticalDawn = 9
        nauticalDusk = 10
        night = 11
        nightEnd = 12
        goldenHour = 13
        goldenHourEnd = 14
        blueHour = 15
        blueHourEnd = 16

    End Enum

    Public Class SunCalc

        Private PI As Double = Math.PI
        Private rad As Double = PI / 180
        Private Function sin(x As Double) As Double
            Return Math.Sin(x)
        End Function
        Private Function cos(x As Double) As Double
            Return Math.Cos(x)
        End Function
        Private Function tan(x As Double) As Double
            Return Math.Tan(x)
        End Function
        Private Function asin(x As Double) As Double
            Return Math.Asin(x)
        End Function
        Private Function atan(x As Double, y As Double) As Double
            Return Math.Atan2(x, y)
        End Function
        Private Function acos(x As Double) As Double
            Return Math.Acos(x)
        End Function

        ' @Vladimir: sun calculations are based on http://aa.quae.nl/en/reken/zonpositie.html formulas


        ' @Vladimir: date/time constants and conversions

        Private ReadOnly dayMs As Double = 1000 * 60 * 60 * 24
        Private ReadOnly J1970 As Double = 2440588
        Private ReadOnly J2000 As Double = 2451545

        Private Function toJulian(jsdate As clsJSDate) As Double
            Return jsdate.ValueOf() / dayMs - 0.5 + J1970
        End Function
        Private Function fromJulian(j As Double) As clsJSDate
            Return New clsJSDate((j + 0.5 - J1970) * dayMs)
        End Function
        Private Function toDays(jsdate As clsJSDate) As Double
            Return toJulian(jsdate) - J2000
        End Function


        ' @Vladimir: general calculations for position

        Private ReadOnly e As Double = rad * 23.4397     ' @Vladimir: obliquity of the Earth

        Private Function rightAscension(l As Double, b As Double) As Double
            Return atan(sin(l) * cos(e) - tan(b) * sin(e), cos(l))
        End Function
        Private Function declination(l As Double, b As Double) As Double
            Return asin(sin(b) * cos(e) + cos(b) * sin(e) * sin(l))
        End Function

        Private Function azimuth(H As Double, phi As Double, dec As Double) As Double
            Return atan(sin(H), cos(H) * sin(phi) - tan(dec) * cos(phi))
        End Function
        Private Function altitude(H As Double, phi As Double, dec As Double) As Double
            Return asin(sin(phi) * sin(dec) + cos(phi) * cos(dec) * cos(H))
        End Function

        Private Function siderealTime(d As Double, lw As Double) As Double
            Return rad * (280.16 + 360.9856235 * d) - lw
        End Function

        Private Function astroRefraction(h As Double) As Double
            If (h < 0) Then     ' @Vladimir: the Then following formula works For positive altitudes only.
                h = 0           ' @Vladimir: If h = -0.08901179 a div/0 would occur.
            End If

            ' @Vladimir: formula 16.4 of "Astronomical Algorithms" 2nd edition by Jean Meeus (Willmann-Bell, Richmond) 1998.
            ' @Vladimir: 1.02 / tan(h + 10.26 / (h + 5.10)) h in degrees, result in arc minutes -> converted to rad
            Return 0.0002967 / Math.Tan(h + 0.00312536 / (h + 0.08901179))
        End Function


        ' @Vladimir: general sun calculations

        Private Function solarMeanAnomaly(d As Double) As Double
            Return rad * (357.5291 + 0.98560028 * d)
        End Function

        Private Function eclipticLongitude(M As Double) As Double

            Dim C As Double = rad * (1.9148 * sin(M) + 0.02 * sin(2 * M) + 0.0003 * sin(3 * M)) ' @Vladimir: equation of center
            Dim P As Double = rad * 102.9372 ' @Vladimir: perihelion Of the Earth

            Return M + C + P + PI
        End Function

        Private Function sunCoords(d As Double) As clsSunCoords

            Dim M = solarMeanAnomaly(d)
            Dim L = eclipticLongitude(M)

            Return New clsSunCoords With {
            .dec = declination(L, 0),
            .ra = rightAscension(L, 0)
            }
        End Function


        ' =========================================================================================
        ' GET SUN POSITION
        '
        ' @Vladimir: calculates sun position for a given date and latitude/longitude
        ' =========================================================================================
        Public Function getSunPosition(eventDate As Date, lat As Double, lng As Double) As clsSunPosition

            Dim jsDate As New clsJSDate(eventDate)

            Dim lw = rad * -lng
            Dim phi = rad * lat
            Dim d = toDays(jsDate)

            Dim c = sunCoords(d)
            Dim H As Double = siderealTime(d, lw) - c.ra

            Return New clsSunPosition With {
            .azimuth = azimuth(H, phi, c.dec),
            .altitude = altitude(H, phi, c.dec)
            }
        End Function


        ' @Vladimir: sun times configuration (angle, morning name, evening name)
        ' enumVal1 and enumVal2 are to ease the data retrieval from the dictionary of results of getSunTimes()

        Public times As New Stack(Of clsTimes)({
        New clsTimes With {.val = -0.833, .str1 = "sunrise", .enumVal1 = enumSunTimes.sunrise, .str2 = "sunset", .enumVal2 = enumSunTimes.sunset},
        New clsTimes With {.val = -0.3, .str1 = "sunriseEnd", .enumVal1 = enumSunTimes.sunriseEnd, .str2 = "sunsetStart", .enumVal2 = enumSunTimes.sunsetStart},
        New clsTimes With {.val = -6, .str1 = "dawn", .enumVal1 = enumSunTimes.dawn, .str2 = "dusk", .enumVal2 = enumSunTimes.dusk},
        New clsTimes With {.val = -12, .str1 = "nauticalDawn", .enumVal1 = enumSunTimes.nauticalDawn, .str2 = "nauticalDusk", .enumVal2 = enumSunTimes.nauticalDusk},
        New clsTimes With {.val = -18, .str1 = "nightEnd", .enumVal1 = enumSunTimes.nightEnd, .str2 = "night", .enumVal2 = enumSunTimes.night},
        New clsTimes With {.val = 6, .str1 = "goldenHourEnd", .enumVal1 = enumSunTimes.goldenHourEnd, .str2 = "goldenHour", .enumVal2 = enumSunTimes.goldenHour,
        New clsTimes With {.val = 4, .str1 = "blueHourEnd", .enumVal1 = enumSunTimes.blueHourEnd, .str2 = "blueHour", .enumVal2 = enumSunTimes.blueHour}
        })

        ' =========================================================================================
        ' ADD SUN TIME
        '
        ' @Vladimir: adds a custom time to the times config
        ' The retrieval of custom times data from the dictionary of results of getSunTimes() must
        ' be done using their keys, unless you add them to enumSunTimes
        ' =========================================================================================
        Public Sub addSunTime(angle As Double, riseName As String, setName As String)
            times.Push(New clsTimes With {.val = angle, .str1 = riseName, .str2 = setName})
        End Sub


        ' @Vladimir: calculations for sun times

        Private Const J0 As Double = 0.0009

        Private Function julianCycle(d As Double, lw As Double) As Double
            Return Math.Round(d - J0 - lw / (2 * PI))
        End Function

        Private Function approxTransit(Ht As Double, lw As Double, n As Double) As Double
            Return J0 + (Ht + lw) / (2 * PI) + n
        End Function
        Private Function solarTransitJ(ds As Double, M As Double, L As Double) As Double
            Return J2000 + ds + 0.0053 * sin(M) - 0.0069 * sin(2 * L)
        End Function

        Private Function hourAngle(h As Double, phi As Double, d As Double) As Double
            Return acos((sin(h) - sin(phi) * sin(d)) / (cos(phi) * cos(d)))
        End Function
        Private Function observerAngle(height As Double) As Double
            Return -2.076 * Math.Sqrt(height) / 60
        End Function

        ' @Vladimir: returns set time for the given sun altitude
        Private Function getSetJ(h As Double, lw As Double, phi As Double, dec As Double, n As Double, M As Double, L As Double) As Double

            Dim w As Double = hourAngle(h, phi, dec)
            Dim a As Double = approxTransit(w, lw, n)

            Return solarTransitJ(a, M, L)
        End Function


        ' =========================================================================================
        ' GET SUN TIMES
        '
        ' @Vladimir: calculates sun times for a given date, latitude/longitude, and, optionally,
        ' @Vladimir: the observer height (in meters) relative to the horizon
        ' =========================================================================================
        Public Function getSunTimes(eventDate As Date, lat As Double, lng As Double, Optional height As Double = 0.0R) As Dictionary(Of Integer, Date)

            Dim jsdate As New clsJSDate(eventDate)

            Dim lw = rad * -lng
            Dim phi = rad * lat

            Dim dh = observerAngle(height)

            Dim d = toDays(jsdate)
            Dim n = julianCycle(d, lw)
            Dim ds = approxTransit(0, lw, n)

            Dim M = solarMeanAnomaly(ds)
            Dim L = eclipticLongitude(M)
            Dim dec = declination(L, 0)

            Dim Jnoon = solarTransitJ(ds, M, L)

            Dim i As Integer
            Dim Len As Integer
            Dim time As clsTimes
            Dim h0 As Double
            Dim Jset As Double
            Dim Jrise As Double


            Dim result As New Dictionary(Of Integer, Date) From {
            {enumSunTimes.solarNoon, fromJulian(Jnoon).GetNETDate},
            {enumSunTimes.nadir, fromJulian(Jnoon - 0.5).GetNETDate}
        }

            Len = times.Count
            i = 0

            While i < Len

                time = times(i)
                h0 = (time.val + dh) * rad

                Jset = getSetJ(h0, lw, phi, dec, n, M, L)
                Jrise = Jnoon - (Jset - Jnoon)

                result.Add(time.enumVal1, fromJulian(Jrise).GetNETDate)
                result.Add(time.enumVal2, fromJulian(Jset).GetNETDate)

                i = i + 1

            End While

            Return result

        End Function


        ' @Vladimir: moon calculations, based on http://aa.quae.nl/en/reken/hemelpositie.html formulas

        Private Function moonCoords(d As Double) As clsMoonCoords ' @Vladimir: geocentric ecliptic coordinates Of the moon
            ' d is the date in days since 1 January 2000, 12:00 UTC 

            Dim L = rad * (218.316 + 13.176396 * d) ' @Vladimir: ecliptic longitude
            Dim M = rad * (134.963 + 13.064993 * d) ' @Vladimir: mean anomaly
            Dim F = rad * (93.272 + 13.22935 * d)   ' @Vladimir: mean distance

            Dim ll = L + rad * 6.289 * sin(M)       ' @Vladimir: longitude
            Dim b = rad * 5.128 * sin(F)            ' @Vladimir: latitude
            Dim dt = 385001 - 20905 * cos(M)        ' @Vladimir: distance To the moon In km

            Return New clsMoonCoords With {
            .ra = rightAscension(ll, b),
            .dec = declination(ll, b),
            .dist = dt
        }
        End Function


        ' =========================================================================================
        ' GET MOON POSITION
        ' =========================================================================================
        Private Overloads Function getMoonPosition(jsDate As clsJSDate, lat As Double, lng As Double) As clsMoonPosition
            Return _getMoonPosition(jsDate, lat, lng)
        End Function
        Public Overloads Function getMoonPosition(eventDate As Date, lat As Double, lng As Double) As clsMoonPosition
            Dim jsDate As New clsJSDate(eventDate)
            Return _getMoonPosition(jsDate, lat, lng)
        End Function
        Private Function _getMoonPosition(jsDate As clsJSDate, lat As Double, lng As Double) As clsMoonPosition
            Dim lw = rad * -lng
            Dim phi = rad * lat
            Dim d = toDays(jsDate)

            Dim c = moonCoords(d)
            Dim H = siderealTime(d, lw) - c.ra
            Dim hh = altitude(H, phi, c.dec)
            ' @Vladimir: formula 14.1 of "Astronomical Algorithms" 2nd edition by Jean Meeus (Willmann-Bell, Richmond) 1998.
            Dim pa = atan(sin(H), tan(phi) * cos(c.dec) - sin(c.dec) * cos(H))

            hh = hh + astroRefraction(hh) ' @Vladimir: altitude correction For refraction

            Return New clsMoonPosition With {
            .azimuth = azimuth(H, phi, c.dec),
            .altitude = hh,                     ' radians
            .distance = c.dist,
            .parallacticAngle = pa
        }
        End Function


        ' =========================================================================================
        ' GET MOON ILLUMINATION
        '
        ' @Vladimir: calculations for illumination parameters of the moon,
        ' @Vladimir: based on http:'idlastro.gsfc.nasa.gov/ftp/pro/astro/mphase.pro formulas And
        ' @Vladimir: Chapter 48 of "Astronomical Algorithms" 2nd edition by Jean Meeus (Willmann-Bell, Richmond) 1998.
        ' =========================================================================================
        Public Function getMoonIllumination(eventDate As Date) As clsMoonIllumination

            If IsNothing(eventDate) Then eventDate = Now
            Dim jsDate As New clsJSDate(eventDate)

            Dim d = toDays(jsDate)
            Dim s = sunCoords(d)
            Dim m = moonCoords(d)

            Dim sdist As Double = 149598000 ' distance from Earth To Sun In km

            Dim phi As Double = acos(sin(s.dec) * sin(m.dec) + cos(s.dec) * cos(m.dec) * cos(s.ra - m.ra))
            Dim inc As Double = atan(sdist * sin(phi), m.dist - sdist * cos(phi))
            Dim angle As Double = atan(cos(s.dec) * sin(s.ra - m.ra), sin(s.dec) * cos(m.dec) - cos(s.dec) * sin(m.dec) * cos(s.ra - m.ra))

            Return New clsMoonIllumination With {
            .fraction = (1 + cos(inc)) / 2,
            .phase = 0.5 + 0.5 * inc * (If(angle < 0, -1, 1)) / Math.PI,
            .angle = angle
        }
        End Function


        Private Function hoursLater(jsDate As clsJSDate, h As Double) As clsJSDate
            Return New clsJSDate(jsDate.ValueOf() + h * dayMs / 24)
        End Function


        ' =========================================================================================
        ' GET MOON TIMES
        '
        ' @Vladimir: calculations for moon rise/set times are based on http://www.stargazing.net/kepler/moonrise.html article
        ' =========================================================================================
        Public Function getMoonTimes(eventDate As Date, lat As Double, lng As Double) As clsMoonTimes

            Dim tempDate As Date = New Date(eventDate.Year, eventDate.Month, eventDate.Day, 0, 0, 0)

            Dim jsDate As New clsJSDate(tempDate)

            Dim t = New clsJSDate(jsDate.GetNETDate.Date)

            Dim hc As Double = 0.133 * rad
            Dim h0 As Double = getMoonPosition(t, lat, lng).altitude - hc
            Dim h1 As Double, h2 As Double, rise As Double, sett As Double, a As Double, b As Double, xe As Double, ye As Double, d As Double, roots As Double, x1 As Double, x2 As Double, dx As Double

            ' @Vladimir: go in 2-hour chunks, each time seeing if a 3-point quadratic curve crosses zero (which means rise Or set)
            For i = 1 To 24 Step 2
                h1 = getMoonPosition(hoursLater(t, i), lat, lng).altitude - hc
                h2 = getMoonPosition(hoursLater(t, i + 1), lat, lng).altitude - hc

                a = (h0 + h2) / 2 - h1
                b = (h2 - h0) / 2
                xe = -b / (2 * a)
                ye = (a * xe + b) * xe + h1
                d = b * b - 4 * a * h1
                roots = 0

                If (d >= 0) Then
                    dx = Math.Sqrt(d) / (Math.Abs(a) * 2)
                    x1 = xe - dx
                    x2 = xe + dx
                    If (Math.Abs(x1) <= 1) Then roots = roots + 1
                    If (Math.Abs(x2) <= 1) Then roots = roots + 1
                    If (x1 < -1) Then x1 = x2
                End If

                If (roots = 1) Then
                    If (h0 < 0) Then
                        rise = i + x1
                    Else
                        sett = i + x1
                    End If
                ElseIf (roots = 2) Then
                    rise = i + (If(ye < 0, x2, x1))
                    sett = i + (If(ye < 0, x1, x2))
                End If

                If (rise <> 0) And (sett <> 0) Then Exit For

                h0 = h2

            Next

            Dim result As New clsMoonTimes

            If (rise <> 0) Then result.rise = hoursLater(t, rise).GetNETDate

            If (sett <> 0) Then result.sett = hoursLater(t, sett).GetNETDate

            If (rise = 0 And sett = 0) Then
                If ye > 0 Then
                    result.alwaysUp = True
                Else
                    result.alwaysDown = True
                End If
            End If

            Return result
        End Function

    End Class
End Namespace