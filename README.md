# GAS Tinkoff Trades

![GAS Tinkoff Trades main image](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/main-image.jpg "GAS Tinkoff Trades main image")

## Установка

https://github.com/coolworld2049/gas_tinkoff_rest_api/assets/82733942/a0b2b17f-2bf8-4a1b-9829-31217d40c928


Вы можете использовать готовый шаблон [template.xlsx](template.xlsx) который необходимо импортировать в Google
Spreadsheets

### Именованные диапазоны

![named-ranges.png](assets%2Fnamed-ranges.png)

- Обязательные именованные диапазоны
    * IS_UPDATED
    * UPDATE_DATE
    * IS_SANDBOX
- Необязательные именованные диапазоны
    * instrumentKind
    * orderType
    * direction
    * operationState
    * CurrencyIsoCode

1. В меню "Tools" выбрать "Script Editor".Дать проекту имя, например `TinkoffAPI`

2. Скопировать код
   из [TinkoffRestApiClient.gs](https://raw.githubusercontent.com/ErhoSen/gas-tinkoff-trades/master/TinkoffRestApiClient.gs)
   и заменить им дефолтный текст скрипта.

3. Выполнить функцию onOpen и предоставить проекту необходимые разрешения
   ![execute-onOpen.png](assets%2Fexecute-onOpen.png)

4. Получить [OpenApi-токен Тинькофф](https://www.tinkoff.ru/invest/settings/api/) и установить в
   переменную `OPENAPI_TOKEN`
   значение токена, полученного выше.

5. Сохраните скрипт 💾

Среди настроек скрипта есть `TRADING_START_AT` - дефолтная дата (`Apr 01, 2020 10:00:00`), начиная с которой фильтруются
операции `getTrades`.

На этом всё. Теперь при работе с этим документом на всех листах будут доступны функции API V2 (не все)
