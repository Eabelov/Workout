# Parse IIS Exchange Log

Простой скрипт для разбора IIS-лога Exchange (W3C Extended Log Format).

## Запуск

```powershell
.\Parse-IisExchangeLog.ps1
```

По умолчанию читает:

`C:\inetpub\logs\LogFiles\W3SVC1\u_ex260725.log`

## Примеры

```powershell
# Другой файл
.\Parse-IisExchangeLog.ps1 -Path "C:\inetpub\logs\LogFiles\W3SVC1\u_ex260725.log"

# Только ошибки 401
.\Parse-IisExchangeLog.ps1 -Status 401

# Первые 20 строк
.\Parse-IisExchangeLog.ps1 -Top 20

# Топ URI по количеству запросов
.\Parse-IisExchangeLog.ps1 | Group-Object cs-uri-stem | Sort-Object Count -Descending | Select-Object -First 10 Count, Name

# Выгрузка в CSV
.\Parse-IisExchangeLog.ps1 | Export-Csv .\exchange-iis.csv -NoTypeInformation -Encoding UTF8
```

Пример лога для проверки лежит в `samples/u_ex260725.log`.
