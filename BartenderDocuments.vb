'+--------------------------------------------------+
'| Author: Steven Gatto - IT Intern                 |
'| Last Updated: 7/25/14                            |
'|                                                  |
'| This program is aimed to help the receiving      |
'| dept automate label printing and increase        |
'| productivity.                                    |
'+--------------------------------------------------+

Module BartenderDocuments

    ' used for running locally
    'Public ReadOnly CustomLargeLabel As String = "C:\Users\sgatto\Documents\Visual Studio 2010\Projects\1-ClickPublish\1-ClickPublish\Resources\Label Templates\LayeredCustomLargeLabel.btw"
    'Public ReadOnly CustomSmallLabel As String = "C:\Users\sgatto\Documents\Visual Studio 2010\Projects\1-ClickPublish\1-ClickPublish\Resources\Label Templates\LayeredCustomSmallLabel.btw"
    'Public ReadOnly LargeStaticLabel As String = "C:\Users\sgatto\Documents\Visual Studio 2010\Projects\1-ClickPublish\1-ClickPublish\Resources\Label Templates\LargeStaticRecvLabel.btw"
    'Public ReadOnly SmallStaticLabel As String = "C:\Users\sgatto\Documents\Visual Studio 2010\Projects\1-ClickPublish\1-ClickPublish\Resources\Label Templates\SmallStaticRecvLabel.btw"

    ' used for running on iis server
    Public ReadOnly CustomLargeLabel As String = "C:\RecvLabelGenerator\Resources\Label Templates\LayeredCustomLargeLabel.btw"
    Public ReadOnly CustomSmallLabel As String = "C:\RecvLabelGenerator\Resources\Label Templates\LayeredCustomSmallLabel.btw"
    Public ReadOnly LargeStaticLabel As String = "C:\RecvLabelGenerator\Resources\Label Templates\LargeStaticRecvLabel.btw"
    Public ReadOnly SmallStaticLabel As String = "C:\RecvLabelGenerator\Resources\Label Templates\SmallStaticRecvLabel.btw"

End Module
