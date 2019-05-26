Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports System.ServiceProcess
Imports System.Text
Imports System.Threading.Tasks
Imports System.Timers
Imports System.IO
Imports Npgsql

Public Class ffiscal
    Private EventLog1 As EventLog
    Private eventId As Integer = 1
    Dim k As Integer, prf As Decimal, cnt As Decimal
    Dim impf As ImpresoraFiscal.Impresoras
    Dim resp As Boolean
    Dim carticulo As String, darticulo As String, preciof As String, totalf As String, cantidadf As String
    Dim impuestof As String, tporce(99) As String, tnovta As String
    Public t As New Timers.Timer

    ' To access the constructor in Visual Basic, select New from the
    ' method name drop-down list. 
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        Me.EventLog1 = New System.Diagnostics.EventLog
        If Not System.Diagnostics.EventLog.SourceExists("fiscal_service") Then
            System.Diagnostics.EventLog.CreateEventSource("fiscal_service",
        "Fiscallog")
        End If
        EventLog1.Source = "fiscal_service"
        EventLog1.Log = "Fiscallog"
    End Sub

    Protected Overrides Sub OnStart(ByVal args As String())
        EventLog1.WriteEntry("In OnStart.")
        ' Set up a timer that triggers every minute.
        'Public t As Timer = New Timer()
        t.Interval = 30000 ' 60 seconds
        AddHandler t.Elapsed, AddressOf ImprimeFac
        t.Start()
    End Sub
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)

        'EventLog1.WriteEntry("Verificando si hay factura a imprimir", EventLogEntryType.Information, eventId)
        'eventId = eventId + 1

        EventLog1.WriteEntry("Verifica si hay factura Fiscal para imprimir", EventLogEntryType.Information, eventId)
        eventId = eventId + 1
        'ImprimeFac()

        'connString = "Host=myserver;Username=mylogin;Password=mypass;Database=mydatabase"
        'Dim builder As New Npgsql.NpgsqlConnectionStringBuilder
        'With builder
        '    .Host = "186.72.149.211"
        '    .Database = "conta"
        '    .Username = "odoo"
        '    .Password = "odoo"
        '    .Port = 5433
        '    .Timeout = 20
        'End With

        'Dim conn As NpgsqlConnection = New NpgsqlConnection(builder.ConnectionString)
        'conn.Open()

        'Dim command As NpgsqlCommand = New NpgsqlCommand("select * from fiscal_header", conn)
        'Dim dr As NpgsqlDataReader
        'Dim dr1 As NpgsqlDataReader
        'Dim fac As Integer
        'Dim strcliente As String, strruc As String
        'dr = command.ExecuteReader()

        'If dr.HasRows Then
        '    'MsgBox("hasrows is true")
        '    If dr.GetString(18) = "invoice" And Not dr(14) Then

        '        EventLog1.WriteEntry("Imprimiendo factura Fiscal", EventLogEntryType.Information, eventId)
        '        eventId = eventId + 1

        '        fac = Trim(dr(0))
        '        strcliente = dr.GetString(5)
        '        tnovta = dr.GetString(16)
        '        strruc = dr.GetString(9)

        '        impf = New ImpresoraFiscal.Impresoras
        '        resp = impf.Iniciar()
        '        resp = impf.Encabezado(strruc, strcliente)

        '        dr.Close()
        '        Dim cmdupd As NpgsqlCommand = New NpgsqlCommand("update fiscal_header set printed=true where transaction_type = 'invoice' And Not printed and id='" & fac & "'", conn)
        '        cmdupd.ExecuteNonQuery()

        '        Dim cmd1 As NpgsqlCommand = New NpgsqlCommand("select * from fiscal_details where id='" & fac & "'", conn)
        '        dr1 = cmd1.ExecuteReader()

        '        While (dr1.Read())
        '            If Trim(dr1(4)).Length > 13 Then
        '                carticulo = Trim(dr1(4)).Substring(1, 13)
        '            Else
        '                carticulo = Trim(dr1(4))
        '            End If
        '            If Trim(dr1(7)).Length > 29 Then
        '                darticulo = Trim(dr1(7)).Substring(0, 28)
        '            Else
        '                darticulo = Trim(dr1(7))
        '            End If

        '            ' Se ajusta el precio
        '            prf = dr1(11)
        '            preciof = prf.ToString("0.00")

        '            ' Se ajusta la cantidad con 3 ceros
        '            cnt = dr1(12)
        '            cantidadf = cnt.ToString("0.000")

        '            resp = impf.Detalle(carticulo, darticulo, Trim(dr1(5)), cantidadf, preciof, Trim(dr1(2)))

        '        End While
        '        dr1.NextResult()
        '        resp = impf.Cierre(tnovta, "Gracias por si visita")
        '        dr1.Close()
        '    End If
        'End If

        'conn.Close()
        'EventLog1.WriteEntry("Continuo", EventLogEntryType.Information, eventId)
        'eventId = eventId + 1

    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        EventLog1.WriteEntry("In OnStop.")
        t.stop()
    End Sub
    Protected Overrides Sub Oncontinue()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        EventLog1.WriteEntry("In Oncontinue.")
        t.Start()
    End Sub

    Public Function ImprimeFac()

        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        EventLog1.WriteEntry("Verifica si hay factura Fiscal para imprimir", EventLogEntryType.Information, eventId)
        eventId = eventId + 1

        'connString = "Host=myserver;Username=mylogin;Password=mypass;Database=mydatabase"
        Dim builder As New Npgsql.NpgsqlConnectionStringBuilder
        With builder
            .Host = "186.72.149.211"
            .Database = "conta"
            .Username = "odoo"
            .Password = "odoo"
            .Port = 5433
            .Timeout = 20
        End With

        Dim conn As NpgsqlConnection = New NpgsqlConnection(builder.ConnectionString)
        conn.Open()

        Dim command As NpgsqlCommand = New NpgsqlCommand("select * from fiscal_header", conn)
        Dim dr As NpgsqlDataReader
        Dim dr1 As NpgsqlDataReader
        Dim fac As Integer
        Dim strcliente As String, strruc As String
        dr = command.ExecuteReader()

        If dr.HasRows Then
            'MsgBox("hasrows is true")
            If dr.GetString(18) = "invoice" And Not dr(14) Then

                EventLog1.WriteEntry("Imprimiendo factura Fiscal", EventLogEntryType.Information, eventId)
                eventId = eventId + 1

                fac = Trim(dr(0))
                strcliente = dr.GetString(5)
                tnovta = dr.GetString(16)
                strruc = dr.GetString(9)

                impf = New ImpresoraFiscal.Impresoras
                resp = impf.Iniciar()
                resp = impf.Encabezado(strruc, strcliente)

                dr.Close()
                Dim cmdupd As NpgsqlCommand = New NpgsqlCommand("update fiscal_header set printed=true where transaction_type = 'invoice' And Not printed and id='" & fac & "'", conn)
                cmdupd.ExecuteNonQuery()

                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand("select * from fiscal_details where id='" & fac & "'", conn)
                dr1 = cmd1.ExecuteReader()

                While (dr1.Read())
                    If Trim(dr1(4)).Length > 13 Then
                        carticulo = Trim(dr1(4)).Substring(1, 13)
                    Else
                        carticulo = Trim(dr1(4))
                    End If
                    If Trim(dr1(7)).Length > 29 Then
                        darticulo = Trim(dr1(7)).Substring(0, 28)
                    Else
                        darticulo = Trim(dr1(7))
                    End If

                    ' Se ajusta el precio
                    prf = dr1(11)
                    preciof = prf.ToString("0.00")

                    ' Se ajusta la cantidad con 3 ceros
                    cnt = dr1(12)
                    cantidadf = cnt.ToString("0.000")

                    resp = impf.Detalle(carticulo, darticulo, Trim(dr1(5)), cantidadf, preciof, Trim(dr1(2)))

                End While
                dr1.NextResult()
                resp = impf.Cierre(tnovta, "Gracias por si visita")
                dr1.Close()
            End If
        End If

        'While (dr.Read())
        '    MsgBox(dr(0))
        'End While

        'Dim command As NpgsqlCommand = New NpgsqlCommand("insert into tablename(name) values('memo')", conn)

        'Dim retcnt = command.ExecuteNonQuery()

        conn.Close()

    End Function

End Class
