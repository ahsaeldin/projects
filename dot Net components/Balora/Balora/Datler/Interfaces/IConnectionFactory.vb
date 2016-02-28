'Implemented mainly by ConnectionFactory class.
Imports System.Data.Common

Namespace Datler
    Friend Interface IConnectionfactory

        Function createconnection() As DbConnection

    End Interface
End Namespace
