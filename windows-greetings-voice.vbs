' windows-greetings-voice makes Windows greet you, tell the time and optionally check Outlook.
'
' Copyright (c) 2018, Mnheia <mnheia@gmail.com>
'
' This module is free software; you can redistribute it and/or modify it
' under the terms of GNU general public license (gpl) version 3.
' See the LICENSE file for details.

' Configuration
Language = "en"       ' en, de, ru, es, fr, zh
Use24Hour = False
CheckOutlook = True

Set Sapi = WScript.CreateObject("SAPI.SpVoice")
SetVoice Sapi, Language

Select Case LCase(Language)
    Case "de"
        morningText = "Guten Morgen"
        afternoonText = "Guten Tag"
        eveningText = "Guten Abend"
        dayText = "Heute ist"
        timeText = "Es ist"
        checkingText = U("Ich pr\u00FCfe auf neue Nachrichten. Bitte warten.")
        unreadText = "Ungelesene Nachrichten"
    Case "ru"
        morningText = U("\u0414\u043E\u0431\u0440\u043E\u0435 \u0443\u0442\u0440\u043E")
        afternoonText = U("\u0414\u043E\u0431\u0440\u044B\u0439 \u0434\u0435\u043D\u044C")
        eveningText = U("\u0414\u043E\u0431\u0440\u044B\u0439 \u0432\u0435\u0447\u0435\u0440")
        dayText = U("\u0421\u0435\u0433\u043E\u0434\u043D\u044F")
        timeText = U("\u0421\u0435\u0439\u0447\u0430\u0441")
        checkingText = U("\u041F\u0440\u043E\u0432\u0435\u0440\u044F\u044E \u043D\u043E\u0432\u044B\u0435 \u0441\u043E\u043E\u0431\u0449\u0435\u043D\u0438\u044F. \u041F\u043E\u0436\u0430\u043B\u0443\u0439\u0441\u0442\u0430, \u043F\u043E\u0434\u043E\u0436\u0434\u0438\u0442\u0435.")
        unreadText = U("\u041D\u0435\u043F\u0440\u043E\u0447\u0438\u0442\u0430\u043D\u043D\u044B\u0435 \u0441\u043E\u043E\u0431\u0449\u0435\u043D\u0438\u044F")
    Case "es"
        morningText = U("Buenos d\u00EDas")
        afternoonText = "Buenas tardes"
        eveningText = "Buenas noches"
        dayText = "Hoy es"
        timeText = "La hora actual es"
        checkingText = "Comprobando mensajes nuevos. Por favor espere."
        unreadText = U("Mensajes no le\u00EDdos")
    Case "fr"
        morningText = "Bonjour"
        afternoonText = U("Bon apr\u00E8s-midi")
        eveningText = "Bonsoir"
        dayText = "Nous sommes le"
        timeText = "Il est"
        checkingText = "Recherche de nouveaux messages. Veuillez patienter."
        unreadText = "Messages non lus"
    Case "zh"
        morningText = U("\u65E9\u4E0A\u597D")
        afternoonText = U("\u4E0B\u5348\u597D")
        eveningText = U("\u665A\u4E0A\u597D")
        dayText = U("\u4ECA\u5929\u662F")
        timeText = U("\u73B0\u5728\u65F6\u95F4\u662F")
        checkingText = U("\u6B63\u5728\u68C0\u67E5\u65B0\u90AE\u4EF6\uFF0C\u8BF7\u7A0D\u5019\u3002")
        unreadText = U("\u672A\u8BFB\u90AE\u4EF6")
    Case Else
        morningText = "Good morning"
        afternoonText = "Good afternoon"
        eveningText = "Good evening"
        dayText = "The current day is"
        timeText = "The current time is"
        checkingText = "Checking for new messages. Please standby."
        unreadText = "Unread messages"
End Select

If Hour(Time) < 12 Then
    Sapi.Speak morningText
ElseIf Hour(Time) < 17 Then
    Sapi.Speak afternoonText
Else
    Sapi.Speak eveningText
End If

Sapi.Speak dayText
Sapi.Speak Date

Sapi.Speak timeText
SpeakCurrentTime Sapi, Use24Hour, Language

If CheckOutlook Then
    Sapi.Speak checkingText
    WScript.Sleep 2000

    Set otl = CreateObject("Outlook.Application")
    Set session = otl.GetNamespace("MAPI")

    session.Logon
    Set inbox = session.GetDefaultFolder(6)

    c = 0
    For Each m In inbox.Items
        If m.Unread Then c = c + 1
    Next

    session.Logoff

    Sapi.Speak unreadText
    Sapi.Speak CStr(c)
End If

Sub SpeakCurrentTime(Sapi, Use24Hour, Language)
    h = Hour(Time)
    m = Minute(Time)

    If Use24Hour Then
        Sapi.Speak CStr(h)
        If m < 10 Then Sapi.Speak "0"
        Sapi.Speak CStr(m)
    Else
        displayHour = h Mod 12
        If displayHour = 0 Then displayHour = 12

        Sapi.Speak CStr(displayHour)

        If m > 0 Then
            If m < 10 Then Sapi.Speak "0"
            Sapi.Speak CStr(m)
        End If

        Select Case LCase(Language)
            Case "de"
                If h < 12 Then marker = "vormittags" Else marker = "nachmittags"
            Case "ru"
                If h < 12 Then marker = U("\u0443\u0442\u0440\u0430") Else marker = U("\u0434\u043D\u044F")
            Case "es"
                If h < 12 Then marker = U("de la ma\u00F1ana") Else marker = "de la tarde"
            Case "fr"
                If h < 12 Then marker = "du matin" Else marker = U("de l'apr\u00E8s-midi")
            Case "zh"
                If h < 12 Then marker = U("\u4E0A\u5348") Else marker = U("\u4E0B\u5348")
            Case Else
                If h < 12 Then marker = "A.M." Else marker = "P.M."
        End Select

        Sapi.Speak marker
    End If
End Sub

Sub SetVoice(Sapi, ByRef Language)
    Select Case LCase(Language)
        Case "de"
            languageId = "407"
        Case "ru"
            languageId = "419"
        Case "es"
            languageId = "40A"
        Case "fr"
            languageId = "40C"
        Case "zh"
            languageId = "804"
        Case Else
            Language = "en"
            languageId = "409"
    End Select

    Set voices = Sapi.GetVoices("Language=" & languageId)

    If voices.Count > 0 Then
        Set Sapi.Voice = voices.Item(0)
    Else
        Language = "en"
        Set voices = Sapi.GetVoices("Language=409")
        If voices.Count > 0 Then Set Sapi.Voice = voices.Item(0)
    End If
End Sub

Function U(text)
    result = ""

    Do While Len(text) > 0
        pos = InStr(text, "\u")

        If pos = 0 Then
            result = result & text
            Exit Do
        End If

        result = result & Left(text, pos - 1)
        result = result & ChrW(CLng("&H" & Mid(text, pos + 2, 4)))
        text = Mid(text, pos + 6)
    Loop

    U = result
End Function
