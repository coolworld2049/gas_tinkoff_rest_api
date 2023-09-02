# GAS Tinkoff Trades

![GAS Tinkoff Trades main image](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/main-image.jpg "GAS Tinkoff Trades main image")

Данный [Google Apps Script](https://developers.google.com/apps-script) предназначен для импорта сделок из Тинькофф
Инвестиций прямо в Google таблицы, для последующего анализа.

Я сделал этот скрипт для автоматизации ручного вбивания данных из приложения тинькофф, и надеюсь он окажется полезен
кому-нибудь ещё :)

## Установка

* Создать или открыть документ Google Spreadsheets https://docs.google.com/spreadsheets/
* В меню "Tools" выбрать "Script Editor"
* Дать проекту имя, например `TinkoffTrades`
* Скопировать код из [TinkoffGAS.gs](https://raw.githubusercontent.com/ErhoSen/gas-tinkoff-trades/master/TinkoffGAS.gs) и заменить
  им дефолтный текст скрипта.
* Получить [OpenApi-токен Тинькофф](https://www.tinkoff.ru/invest/settings/api/)
* Добавить свойство `OPENAPI_TOKEN` в разделе `Project properties -> Script properties` со значением токена, полученным
  выше.
* Сохранить скрипт 💾
* В документе Google Spreadsheets выбрать любую ненужную ячейку и присвоить ей имя `UPDATE_DATE` с помощью
  меню `Data`->`Named ranges`->`Add named range`. В эту ячейку по команде меню TI->Обновить вставляется текущая дата.
  Данная ячейка может использоваться в качестве необязательного параметра `UPDATE_DATE` любой функции
  для [принудительного обновления формул](https://stackoverflow.com/a/27656313). Например, при использовании
  функции `TI_GetLastPriceByFigi("BBG004730RP0",UPDATE_DATE)` можно реализовать обновление цены акции "Газпром" по
  команде меню TI->Обновить.
* В документе Google Spreadsheets выбрать любую ненужную ячейку и присвоить ей имя `IS_UPDATED` с помощью
  меню `Data`->`Named ranges`->`Add named range` и добавить чекбокс (TRUE/FALSE). Данная ячейка необходима для включения
  и выключения циклического обновления таблицы через определенный промежуток времени в секундах `REFRESH_EVERY_SEC` (
  указывается в скрипте).

На этом всё. Теперь при работе с этим документом на всех листах будут доступны функции:

* API
  v2: `TI_GetAccounts()`, `TI_GetAccountID()`, `TI_GetInstrumentsID()`, `TI_GetLastPriceByFigi()`, `TI_GetPortfolio()`, `TI_GetOperations()`, `TI_GetBidAskSpread()`
  и `TI_GetLastPrice()` (использует API v1 и может перестать работать)
* API v1: `getPriceByTicker()`, `getTrades()`, `getPortfolio()`, `getMaxBidByTicker()`, `getMinAskByTicker()`
  и `getBidAskSpread()`

## Пример

![example.gif](gif/example.gif)

## Функции API v2

* `TI_GetAccounts()` - получение информации по счетам (идентификатор, тип, название, статус, дата открытия, права
  доступа).

* `TI_GetAccountID(accountNum)` - возвращает идентификатор счета по его порядковому номеру (начиная с 0).

* `=TI_GetLastPriceByFigi(FIGI, UPDATE_DATE)` - требует на вход FIGI инструмента (может быть получен через `=VLOOKUP()`
  по отдельной вкладке с информацией по всем инструментам, полученной вызовом `TI_GetInstrumentsID()`) и опциональный
  параметр с именем ячейки `UPDATE_DATE`, которая будет обновлятся через меню TI->Обновить).

* `=TI_GetInstrumentsID()` - выводит информацию по всем инструментам, включая их тикер, FIGI, класс, биржу, валюту и
  ISIN.

* `=TI_GetOperations(accountId, from, to)` - выводит операции по заданному идентификатору счета (можно получить через
  функцию `TI_GetAccounts()`) и, опционально, фильтрацию по времени. Параметры `from` и `to` являются ссылками на ячейки
  с типом Дата.

* `=TI_GetPortfolio(accountId)` - выводит портфель по заданному идентификатору счета (можно получить через
  функцию `TI_GetAccounts()`.

## Функции API v1 (перестанут работать в 2023 году)

* `=getPriceByTicker(ticker, UPDATE_DATE)` - требует на
  вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опциональный параметр с именем
  ячейки `UPDATE_DATE`, которая будет обновлятся через меню TI->Обновить).

* `=getTrades(ticker, from, to)` - требует на
  вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опционально фильтрацию по времени.
  Параметры `from` и `to` являются строками и должны быть в [ISO 8601 формате](https://ru.wikipedia.org/wiki/ISO_8601)

* `=getPortfolio()` - выводит портфель.

## Особенности

* Среди настроек скрипта есть `TRADING_START_AT` - дефолтная дата, начиная с которой фильтруются операции `getTrades`.
  По умолчанию это `Apr 01, 2020 10:00:00`, но данную константу можно в любой момент поменять в исходном коде.

## Пример использования работы функций API v1

```
=getPriceByTicker("V", UPDATE_DATE)  # Возвращает текущую цену акции Visa
=getPriceByTicker("FXMM", UPDATE_DATE)  # Возвращает текущую цену фонда казначейских облигаций США

=getTrades("V") 
# Вернёт все операции с акцией Visa, которые произошли начиная с TRADING_START_AT и по текущий момент.
=getTrades("V", "2020-05-01T00:00:00.000Z") 
# Вернёт все операции с акцией Visa, которые произошли начиная с 1 мая и по текущий моментs.
=getTrades("V", "2020-05-01T00:00:00.000Z", "2020-05-05T23:59:59.999Z") 
# Вернёт все операции с акцией Visa, которые произошли в период с 1 и по 5 мая.
```

## Пример работы функций API v1

#### `=getTrades()`

![getTrades in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-trades-in-action.gif "getTrades in Action")

#### `=getPriceByTicker()`

![Get price by ticker in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-price-by-ticker.gif)
