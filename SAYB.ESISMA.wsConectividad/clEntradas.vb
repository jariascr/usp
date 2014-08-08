Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.Threading

Public Class clEntradas

    Public Shared Sub Main()

        ' Create the source, if it does not already exist.
        If Not EventLog.SourceExists("MySource") Then
            EventLog.CreateEventSource("MySource", "MyNewLog")
            Console.WriteLine("CreatingEventSource")
        End If

        ' Create an EventLog instance and assign its source.
        Dim myLog As New EventLog()
        myLog.Source = "MySource"

        ' Write an informational entry to the event log.    
        myLog.WriteEntry("Writing to event log.")
    End Sub 'Main 

End Class



