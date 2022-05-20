# SunCalcForDotNet
SunCalcForDotNet (or SunCalc for .NET) is a translation into VB.NET of well known javascript Vladimir Agafonkin's SunCalc.
SunCalcForDotNet is made of 2 files (both required for the library to work):
  * SunCalc.vb            - Main file. Contains the SunCalc code
  * SunCalcHelper.vb      - Helper classes for the main file
# Basic usage
Add both files to your project.

1. Add
```vb.net
Imports SunCalcForDotNet
```
  on top of class or module you intend to use to work with SunCalc

2. Inside your class, instantiate SunCalc:

```vb.net
    Dim sc as New SunCalc
```

3. Define the vars needed for the call to the method you are going to use:

```vb.net
    Dim myDate As Date = Now
```

4. Then call the method of your choice, assigning its result to the appropiate var. For example:

```vb.net
    Dim suntimes As Dictionary(Of Integer, Date) = sc.getSunTimes(fecha, myLat, myLng)
```

5. Extract the desired sun event dates from the dictionary

```vb.net
    Dim sunrise As Date = suntimes(enumSunTimes.sunrise).Add(timezone)
    Dim sunset As Date = suntimes(enumSunTimes.sunset).Add(timezone)
```

**All dates returned by SunCalc are UTC**, so you might want to convert them into your time zone:

```vb.net
    Dim timezone As New TimeSpan(2, 0, 0)   ' Your time zone
    sunrise = sunrise.Add(timezone)
    sunset = sunset.Add(timezone)
```
