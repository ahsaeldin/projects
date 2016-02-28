Public Module PublicEnums
    Public Enum ShutdownTypeEnum
        LogOff = 0
        Shutdown = 1
        Restart
        ForcedLogOff = 4
        ForcedShutdown = 5
        ForcedRestart = 6
        Sleep
        Hibernate
        ForcedHibernate
        ForcedSleep
    End Enum

    Public Enum TimeUnit
        Seconds
        Minutes
        Hours
    End Enum

    Public Enum WeekDays
        SUN = 1
        MON
        TUE
        WED
        THU
        FRI
        SAT
        ALL
        Ordinal 'like 3#2 the second Tue.
    End Enum

    Public Enum Months
        JAN = 1
        FEB
        MAR
        APR
        MAY
        JUN
        JUL
        AUG
        SEP
        OCT
        NOV
        DEC
    End Enum

    'http://en.wikipedia.org/wiki/English_numerals#Ordinal_numbers
    Public Enum OrdinalNumbers
        First = 1
        Second
        Third
        Fourth
        Fifth
        Sixth
    End Enum

    ''' <summary>
    ''' Database Type Enumeration.
    ''' </summary>
    ''' <remarks>Used as a value for both DatabaseType and DefaultDatabase properties.</remarks>
    Enum DbType
        SQLite = 0
        MySQL
        SQLServer
        Oracle
        PostgreSQL
        Firebird
    End Enum

    Enum Dbprovider
        DevartSQlite = 0
        System_Data_SQLite
    End Enum
End Module

