Copyright (c) 2018, Mnheia <mnheia@gmail.com>

# windows-greetings-voice
windows-greetings-voice is a VBS script that makes Windows greet you, tells you the date and time, and can optionally check for unread messages in Microsoft Outlook.

The script supports English, German, Russian, Spanish, French and Chinese voices. It can use either 12-hour or 24-hour time and Outlook integration can be enabled or disabled.

# Example
Edit the configuration at the top of the script:

```
Language = "en"       ' en, de, ru, es, fr, zh
Use24Hour = False
CheckOutlook = True
```

The script can be run through Task Scheduler.

If the requested language voice is not installed, the script falls back to an English/default SAPI voice.

# Requirements
- Microsoft Windows
- Microsoft Outlook if `CheckOutlook = True`
- A compatible SAPI voice for the selected language

# Bugs
Please report any bugs or feature requests through the web interface at https://github.com/mnheia/windows-greetings-voice/issues
