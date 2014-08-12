'+--------------------------------------------------+
'| Author: Steven Gatto - IT Intern                 |
'| Last Updated: 7/25/14                            |
'|                                                  |
'| This program is aimed to help the receiving      |
'| dept automate label printing and increase        |
'| productivity.                                    |
'+--------------------------------------------------+

Module AllowedPrinters

    ' The printer names that will appear in the printer selection drop down list
    Public ReadOnly DanPrinterDisplayName As String = "Dan's LabelWriter"
    Public ReadOnly RonPrinterDisplayName As String = "Ron's LabelWriter"
    Public ReadOnly ErikPrinterDisplayName As String = "Erik's LabelWriter"
    Public ReadOnly PrestonPrinterDisplayName As String = "Preston's LabelWriter"

    ' (currently unused)
    ' Computer names for the receiving dept
    Public ReadOnly DanComputerName As String = "IPG290-16B-01"
    Public ReadOnly RonComputerName As String = "IPG290-WH16B-01"
    Public ReadOnly ErikComputerName As String = "IPG290-3102-10"
    Public ReadOnly PrestonComputerName As String = "IPG290-3102-12"

    ' (currently unused)
    ' Name of each shared printer from the receiving dept
    Public ReadOnly DanPrinterName As String = "DYMO LabelWriter 450 Twin Turbo"
    Public ReadOnly RonPrinterName As String = "DYMO LabelWriter 450 Twin Turbo (Copy 1)"
    Public ReadOnly ErikPrinterName As String = "DYMO LabelWriter 450 Twin Turbo"
    Public ReadOnly PrestonPrinterName As String = "DYMO LabelWriter 450 Twin Turbo"

    ' The full network paths of the shared printers
    Public ReadOnly DanPrinterNetworkPath As String = "\\" & DanComputerName & "\" & DanPrinterName
    Public ReadOnly RonPrinterNetworkPath As String = "\\" & RonComputerName & "\" & RonPrinterName
    Public ReadOnly ErikPrinterNetworkPath As String = "\\" & ErikComputerName & "\" & ErikPrinterName
    Public ReadOnly PrestonPrinterNetworkPath As String = "\\" & PrestonComputerName & "\" & PrestonPrinterName

End Module
