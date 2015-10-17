Imports System.IO

Module Module1
    Private serveurSMTP As String = ""
    Private ProcessPath As String = "C:\Program Files\VMware\VMware VI Remote CLI\bin\vicfg-cfgbackup.pl"

    'Format de la date utilisé dans le journal et pour la création des dossiers
    Private dateToday = Format(Date.Now, "yyyyMMdd,HH") & "h"
    Private dateLastWeek = Format(Now.AddDays("-7"), "yyyyMMdd")

    Private alerteMsg As String = ""

    Sub Main()
        Dim limiteJour As Integer = 21
        Dim cheminBackup As String = "D:\Sauvegarde\ESXi\BackupKernel\"

        'Creation du dossier par date
        If (Directory.Exists(cheminBackup + dateToday)) Then
            Directory.Delete(cheminBackup + dateToday, True)
        End If
        Directory.CreateDirectory(cheminBackup + dateToday)

        'purge Kernel
        purgeKernel(cheminBackup, limiteJour)

        'sauvegarde esx
        sauvegardeESXi("", "backup", "", cheminBackup & dateToday & "\VMT01006")
        sauvegardeESXi("", "root", "", cheminBackup & dateToday & "\VMT01007")
        sauvegardeESXi("", "backup", "", cheminBackup & dateToday & "\VMT01008")
        sauvegardeESXi("", "backup", "", cheminBackup & dateToday & "\VMT01009")
        sauvegardeESXi("", "backup", "", cheminBackup & dateToday & "\VMT01010")

        'alerte
        If alerteMsg <> "" Then
            alerteMail(alerteMsg)
        End If

    End Sub

    Private Sub purgeKernel(ByVal cheminLogs As String, ByVal limiteJour As Integer)
        'Suppression des dossiers inférieurs à une semaine
        Dim dirs() As String = Directory.GetDirectories(cheminLogs)
        For Each dir As String In dirs
            Dim olddirs As String = dir.Remove(0, cheminLogs.Length)
            olddirs = olddirs.Remove(8, 4)

            If (olddirs.CompareTo(dateLastWeek) < 0) Then
                If (Directory.Exists(dir)) Then
                    Directory.Delete(dir, True)
                End If
            End If
        Next
    End Sub

    Private Function sauvegardeESXi(ByVal adresseIP As String, ByVal login As String, ByVal motDePasse As String, ByVal fichierSauvegarde As String) As Boolean
        Dim objProcess As System.Diagnostics.Process

        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.Arguments = " --server " & adresseIP & " --username " & login & " --password " & motDePasse & " -s " & fichierSauvegarde
            'Console.WriteLine(objProcess.StartInfo.FileName & objProcess.StartInfo.Arguments)
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()

            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()

            'Free resources associated with this process
            objProcess.Close()
        Catch
            alerteMsg = alerteMsg & "Sauvegarde du serveur " & adresseIP & " : " & vbTab & " [ERREUR]" & vbCrLf
            sauvegardeESXi = False
        End Try
        sauvegardeESXi = True
    End Function

    Private Sub alerteMail(ByVal message As String)
        Dim objet As String = "Reporting de la sauvegarde des serveurs ESXi"

        Dim eMail As New System.Net.Mail.MailMessage()
        Dim mailSender As New System.Net.Mail.SmtpClient(serveurSMTP)

        eMail.From = New System.Net.Mail.MailAddress("fmartin@global-sp.net")
        eMail.To.Add(New System.Net.Mail.MailAddress("fmartin@global-sp.net"))
        eMail.Subject = objet
        eMail.Body = message
        mailSender.Send(eMail)
    End Sub

End Module
