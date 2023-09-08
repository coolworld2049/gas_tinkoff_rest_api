
/**@NotOnlyCurrentDoc */
// https://github.com/ErhoSen/gas-tinkoff-trades

const OPENAPI_TOKEN = ""
const CACHE = CacheService.getScriptCache()
const CACHE_MAX_AGE = 60 * 60 * 6 // 6 Hours
const TRADING_START_AT = new Date('Apr 01, 2020 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24
const REFRESH_EVERY_SEC = 2;


/**
 * Добавляет меню с командой вызова функции обновления значений служебной ячейки (для обновления вычислнений функций, ссылающихся на эту ячейку)
 *
 **/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var entries = [{
    name : "Обновить",
    functionName : "refresh"
  },
  {
    name : "Запуск обновления",
    functionName : "runRefresh"
  },
  {
    name : "Остановка обновления",
    functionName : "stopRefresh"
  }
  ]
  sheet.addMenu("TI", entries)
};


function _convertRangeToOneCell(range){
  if (range == null) return;
  if (range.getA1Notation == undefined) return;   // range should be the instance of Range
  return range.getCell(1, 1);
}

function refresh() {
  const updateRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('UPDATE_DATE');
  const updateDateRange = _convertRangeToOneCell(updateRange);
  Logger.log(`refresh range ${updateDateRange.getA1Notation()}`)
  if ((updateDateRange != null) && (checkIsUpdatedCell() == true)) {
    updateDateRange.setValue(Utilities.formatDate(new Date(), "GMT+3", "MM-dd-yyyy HH:mm:ss"));
  } else {
    SpreadsheetApp.getUi().ui.alert('You should specify the named range "UPDATE_DATE" for using this function.');
  }
}

/**
 * Управление циклическим обновлением значений акций
 */

function runRefresh() {
  const updateRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('UPDATE_DATE');
  for (i=0; i<43200; i++){
    var flag = checkIsUpdatedCell();
    if (flag == false) {
      counter = 0;
      return null;
    }
    updateRange.setBackground(null);
    Utilities.sleep(1000 * REFRESH_EVERY_SEC);
    refresh();
    updateRange.setBackgroundRGB(123, 229, 152);
  }
}

function stopRefresh() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var isUpdatedRange = _convertRangeToOneCell(sheet.getRangeByName("IS_UPDATED"));
  var flag = checkIsUpdatedCell();
  if (flag == true) {
    isUpdatedRange.setValue(false);
    Logger.log(`IS_UPDATED forced to ${isUpdatedRange.getValue()}`)
  }
}

function checkIsUpdatedCell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var isUpdatedRange = _convertRangeToOneCell(sheet.getRangeByName("IS_UPDATED"));
  if (isUpdatedRange != null) {
    var value = isUpdatedRange.getValue();
    Logger.log(`IS_UPDATED == ${value}`)
    return value;
  } else {
    SpreadsheetApp.getUi().ui.alert('You should specify the named range "IS_UPDATED" for using this function.');
  }
}

function isoToDate(dateStr){
  // How to format date string so that google scripts recognizes it?
  // https://stackoverflow.com/a/17253060
  const str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
  return new Date(str)
}

class TinkoffClientV1 {
  // Doc: https://tinkoffcreditsystems.github.io/invest-openapi/swagger-ui/
  // How to create a token: https://tinkoffcreditsystems.github.io/invest-openapi/auth/
  constructor(token) {
    this.token = token
    this.baseUrl = 'https://api-invest.tinkoff.ru/openapi/'
  }

  _makeApiCall(methodUrl) {
    const url = this.baseUrl + methodUrl
    Logger.log(`[API Call] ${url}`)
    const params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
    const response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText())
  }

  getInstrumentByTicker(ticker) {
    const url = `market/search/by-ticker?ticker=${ticker}`
    const data = this._makeApiCall(url, {})
    return data.payload
  }

  getOrderbookByFigi(figi, depth) {
    const url = `market/orderbook?depth=${depth}&figi=${figi}`
    const data = this._makeApiCall(url, {})
    return data.payload
  }

  getOperations(from, to, figi) {
    // Arguments `from` && `to` should be in ISO 8601 format
    const url = `operations?from=${from}&to=${to}&figi=${figi}`
    const data = this._makeApiCall(url, {})
    return data.payload.operations
  }

  getPortfolio(){
    const url = `portfolio`
    const data = this._makeApiCall(url, {})
    return data.payload.positions
  }
}

/**  ============================== Tinkoff V1 ==============================
*
*
**/
const tinkoffClientV1 = new TinkoffClientV1(OPENAPI_TOKEN)

function _getFigiByTicker(ticker) {
  const CACHE_KEY_PREFIX = 'figi_'
  const ticker_cache_key = CACHE_KEY_PREFIX + ticker

  const cached = CACHE.get(ticker)
  if (cached != null)
    return cached
  const {instruments,total} = tinkoffClientV1.getInstrumentByTicker(ticker)
  if (total > 0) {
    const figi = instruments[0].figi
    CACHE.put(ticker_cache_key, figi, CACHE_MAX_AGE)
    return figi
  } else {
    return null
  }
}

/**
 * Получение последней цены инструмента по тикеру
 * @param {"GAZP"} ticker Тикер инструмента
 * @return {}             Last price
 * @customfunction
 */
function getPriceByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {lastPrice} = tinkoffClientV1.getOrderbookByFigi(figi, 1)
  return lastPrice
}

/**
 * Получение Bid/Ask спреда инструмента по тикеру
 * @param {"GAZP"} ticker Тикер инструмента
 * @return {0.03}         Спред в %
 * @customfunction
 */
function getBidAskSpreadByTicker(ticker) { // dummy parameter is optional
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,bids,asks} = tinkoffClientV1.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return (asks[0].price-bids[0].price) / asks[0].price
  else
    return null
}

function getMaxBidByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,bids} = tinkoffClientV1.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return [
    ["Max bid", "Quantity"],
    [bids[0].price, bids[0].quantity]
    ]
  else
    return null
}

function getMinAskByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,asks} = tinkoffClientV1.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return [
      ["Min ask", "Quantity"],
      [asks[0].price, asks[0].quantity]
    ]
  else
    return null
}

function _calculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    const {quantity, price} = trades[j]
    totalQuantity += quantity
    totalSum += quantity * price
  }
  const weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}

/**
 * Получение списка операций по тикеру инструмента
 * @param {String} ticker Тикер инструмента для фильтрации
 * @param {String} from Начальная дата
 * @param {String} to Конечная дата
 * @return {Array} Массив результата
 * @customfunction
 */
function getTrades(ticker, from, to) {
  const figi = _getFigiByTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    const now = new Date()
    to = new Date(now + MILLIS_PER_DAY)
    to = to.toISOString()
  }
  const operations = tinkoffClientV1.getOperations(from, to, figi)

  const values = [
    ["ID", "Date", "Operation", "Ticker", "Quantity", "Price", "Currency", "SUM", "Commission"],
  ]
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, id, date, currency, commission} = operations[i]
    if (operationType == "BrokerCommission" || status == "Decline" || operationType == "Dividend")
      continue
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades) // calculate weighted values
    if (operationType == "Buy") {  // inverse values in a way, that it will be easier to work with
      totalQuantity = -totalQuantity
      totalSum = -totalSum
    }
    let com_val = 0
    if (commission){
      com_val = commission.value
    }else{
      com_val = null
    }
    values.push([
      id, isoToDate(date), operationType, ticker, totalQuantity, weigthedPrice, currency, totalSum, com_val
    ])
  }
  return values
}

/**
 * Получение портфеля
 * @return {Array}     Массив с результатами
 * @customfunction
 */
function getPortfolio() {
  const portfolio = tinkoffClientV1.getPortfolio()
  const values = []
  values.push(["Тикер","Название","Тип","Кол-во","Ср.цена покупки","Ст-ть покупки","Валюта","Доход","Тек.ст-ть","Валюта","НКД","Валюта"])
  for (let i=0; i<portfolio.length; i++) {
    let {ticker, name, instrumentType, balance, averagePositionPrice, averagePositionPriceNoNkd, expectedYield} = portfolio [i]
    let NKD=null
    let NKD_Curr=null
    if (averagePositionPriceNoNkd){
      NKD = averagePositionPrice.value - averagePositionPriceNoNkd.value
      averagePositionPrice.value = averagePositionPriceNoNkd.value
      NKD_Curr = averagePositionPriceNoNkd.currency
    }

    values.push([
      ticker, name, instrumentType, balance, averagePositionPrice.value, averagePositionPrice.value * balance, averagePositionPrice.currency, expectedYield.value, averagePositionPrice.value * balance + expectedYield.value, averagePositionPrice.currency, NKD, NKD_Curr
    ])
  }
  return values
}

/**  ============================== Tinkoff V2 ==============================
*
* https://tinkoff.github.io/investAPI/
* https://tinkoff.github.io/investAPI/swagger-ui/
*
**/
class _TinkoffClientV2 {
  constructor(token){
    this.token = token
    this.baseUrl = "https://invest-public-api.tinkoff.ru/rest/";
    //Logger.log(`[_TinkoffClientV2.constructor]`)
  }
  _makeApiCall(methodUrl,data){
    const url = this.baseUrl + methodUrl
    Logger.log(`[Tinkoff OpenAPI V2 Call] ${url}`)
    const params = {
      'method': 'post',
      'headers': {'accept': 'application/json', 'Authorization': `Bearer ${this.token}`},
      'contentType': 'application/json',
      'payload' : JSON.stringify(data),
      'muteHttpExceptions': true
    }

    let respCode, respText, rateLimited;
    do {
      const response = UrlFetchApp.fetch(url, params);
      const respHeaders = response.getAllHeaders();
      respCode = response.getResponseCode();
      respText = response.getContentText("UTF-8");

      rateLimited = Boolean(respHeaders['x-envoy-ratelimited']); // {x-ratelimit-reset, x-envoy-ratelimited, x-ratelimit-remaining}
      if (rateLimited) { // Выжидаем конец периода квоты запросов
        const timeToWait = 500+1000*Number(respHeaders ['x-ratelimit-reset']);
        Logger.log(`[_makeApiCall] Ожидаем конца квоты запросов ${timeToWait} милисек`);
        Utilities.sleep(timeToWait);
        Logger.log(`[_makeApiCall] Конец ожидания`);
      }
    } while (rateLimited)

    const respObj = JSON.parse(respText);
    if ( respCode == 200)
      return respObj
    else
      throw new Error(`Ошибка ${respCode} - ${respObj.message} - ${respObj.description}`);
  }
  _UnitsNanoToPrice(units, nano) {
    return Number(units) + nano/1000000000;
  }
  // ----------------------------- InstrumentsService -----------------------------
  _Bonds(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Bonds`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Shares(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Shares`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Futures(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Futures`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Etfs(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Etfs`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _FindInstrument(query, instrumentKind, apiTradeAvailableFlag) {
    var requestBody = {
    "query": query,
    "instrumentKind": instrumentKind,
    "apiTradeAvailableFlag": apiTradeAvailableFlag ? apiTradeAvailableFlag : true
  }
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/FindInstrument`
    const data = this._makeApiCall(url, requestBody)
    return data
  }
  _Currencies(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Currencies`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _GetInstrumentBy(idType,classCode,id) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/GetInstrumentBy`
    const data = this._makeApiCall(url, {'idType': idType, 'classCode': classCode, 'id': id})
    return data
  }
  // ----------------------------- MarketDataService -----------------------------
  _GetLastPrices(figi_arr) {
    const url = 'tinkoff.public.invest.api.contract.v1.MarketDataService/GetLastPrices'
    const data = this._makeApiCall(url,{'figi': figi_arr})
    return data
  }
  _GetOrderBookByFigi(figi,depth) {
    const url = `tinkoff.public.invest.api.contract.v1.MarketDataService/GetOrderBook`
    const data = this._makeApiCall(url,{'figi': figi, 'depth': depth})
    return data
  }
  // ----------------------------- OperationsService -----------------------------
  _GetOperations(accountId,from,to,state,figi) {
    const url = 'tinkoff.public.invest.api.contract.v1.OperationsService/GetOperations'
    const data = this._makeApiCall(url,{'accountId': accountId,'from': from,'to': to,'state': state,'figi': figi})
    return data
  }
  _GetOperationsByCursor(accountId,instrument_id,from,to,cursor,limit,operation_types,state,without_commissions,without_trades,without_overnights) {
    const url = 'tinkoff.public.invest.api.contract.v1.OperationsService/GetOperationsByCursor'
    const data = this._makeApiCall(url,{'accountId':accountId,'instrument_id':instrument_id,'from':from,'to':to,'cursor':cursor,'limit':limit,'operation_types':operation_types,'state':state,'without_commissions':without_commissions,'without_trades':without_trades,'without_overnights':without_overnights})
    return data
  }
  _GetPortfolio(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio'
    const data = this._makeApiCall(url,{'accountId': accountId})
    return data
  }
  // ----------------------------- OrdersService ----------------------------
  _CancelOrder(accountId, orderId) {
    const url = 'tinkoff.public.invest.api.contract.v1.OrdersService/CancelOrder';
    const data = this._makeApiCall(url, {'accountId': accountId, 'orderId': orderId});
    return data;
  }
  _GetOrderState(accountId, orderId) {
    const url = 'tinkoff.public.invest.api.contract.v1.OrdersService/GetOrderState';
    const data = this._makeApiCall(url, {'accountId': accountId, 'orderId': orderId});
    return data;
  }
  _GetOrders(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.OrdersService/GetOrders';
    const data = this._makeApiCall(url, {'accountId': accountId});
    return data;
  }
  _PostOrder(figi,quantity,priceNano,priceUnits,direction,accountId,orderType) {
    var requestBody = {
      "figi": figi,
      "quantity": quantity,
      "price": {
        "nano": priceNano, //целая часть суммы, может быть отрицательным числом
        "units": priceUnits //дробная часть суммы, может быть отрицательным числом
      },
      "direction": direction,
      "accountId": accountId,
      "orderType": orderType,
    }
    const url = 'tinkoff.public.invest.api.contract.v1.OrdersService/PostOrder';
    var data = null;
    try {
        data = this._makeApiCall(url, requestBody);
        SpreadsheetApp.getUi().alert(JSON.stringify(data));
      }
      catch (e) {
        SpreadsheetApp.getUi().alert(e.toString());
      }
    return data;
  }
  _ReplaceOrder(accountId,orderId,idempotencyKey,quantity,priceNano,priceUnits,price_type) {
    requestBody = {
      "accountId": accountId,
      "orderId": orderId,
      "idempotencyKey": idempotencyKey,
      "quantity": quantity,
      "price": {
        "nano": priceNano,
        "units": priceUnits
      },
      "priceType": price_type
    }
    const url = 'tinkoff.public.invest.api.contract.v1.OrdersService/ReplaceOrder';
    const data = this._makeApiCall(url, requestBody);
    return data;
  }
  // ----------------------------- SandboxService ---------------------------
  _CancelSandboxOrder(accountId,orderId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/CancelSandboxOrder';
    const data = this._makeApiCall(url, {'accountId': accountId, 'orderId': orderId});
    return data;
  }

  _CloseSandboxAccount(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/CloseSandboxAccount';
    const data = this._makeApiCall(url, {'accountId': accountId});
    return data;
  }

  _GetSandboxAccounts() {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxAccounts';
    const data = this._makeApiCall(url, {});
    return data;
  }

  _GetSandboxOperations(accountId,fromDate,toDate,state) {
    var requestBody = {
    "accountId": accountId,
    "from": fromDate,
    "to": toDate,
    "state": state,
    }
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxOperations';
    const data = this._makeApiCall(url, requestBody);
    return data;
  }

  _GetSandboxOperationsByCursor(accountId,instrumentId,fromDate,toDate,cursor,limit,operationTypesList,state, withoutCommissions, withoutTrades, withoutOvernights) {
    var requestBody = {
    "accountId": accountId,
    "instrumentId": instrumentId,
    "from": fromDate,
    "to": toDate,
    "cursor": cursor,
    "limit": limit,
    "operationTypes": operationTypesList,
    "state": state,
    "withoutCommissions": withoutCommissions,
    "withoutTrades": withoutTrades,
    "withoutOvernights": withoutOvernights
    }
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxOperationsByCursor';
    const data = this._makeApiCall(url, requestBody);
    return data;
  }

  _GetSandboxOrderState(accountId,orderId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxOrderState';
    const data = this._makeApiCall(url, {'accountId': accountId, 'orderId': orderId});
    return data;
  }

  _GetSandboxOrders(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxOrders';
    const data = this._makeApiCall(url, {"accountId": accountId});
    return data;
  }

  _GetSandboxPortfolio(accountId, currency) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxPortfolio';
    const data = this._makeApiCall(url, {"accountId": accountId, "currency": "RUB"});
    return data;
  }

  _GetSandboxPositions(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxPositions';
    const data = this._makeApiCall(url, {"accountId": accountId});
    return data;
  }

  _GetSandboxWithdrawLimits(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/GetSandboxWithdrawLimits';
    const data = this._makeApiCall(url, {"accountId": accountId});
    return data;
  }

  _OpenSandboxAccount() {
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/OpenSandboxAccount';
    const data = this._makeApiCall(url, {});
    return data;
  }

  _PostSandboxOrder(figi,
    quantity,
    priceNano,
    priceUnits,
    direction,
    accountId,
    orderType) {
    var requestBody = {
      "figi": figi,
      "quantity": quantity,
      "price": {
        "nano": priceNano,
        "units": priceUnits
      },
      "direction": direction,
      "accountId": accountId,
      "orderType": orderType
    }
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/PostSandboxOrder';
    var data = null;
    try {
        data = this._makeApiCall(url, requestBody);
        SpreadsheetApp.getUi().alert(JSON.stringify(data));
      }
      catch (e) {
        SpreadsheetApp.getUi().alert(e.toString());
      }
    return data;
  }

  _ReplaceSandboxOrder(
    accountId,
    orderId,
    idempotencyKey,
    quantity,
    priceNano,
    priceUnits,
    price_type) {
    var requestBody = {
      "accountId": accountId,
      "orderId": orderId,
      "idempotencyKey": idempotencyKey,
      "quantity": quantity,
      "price": {
        "nano": priceNano,
        "units": priceUnits,
        "type": price_type
      }
    };
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/ReplaceSandboxOrder';
    const data = this._makeApiCall(url, requestBody);
    return data;
  }

  _SandboxPayIn(accountId,amount_nano,amount_currency,amount_units) {
    var requestBody = {
      "accountId": accountId,
      "amount": {
        "nano": amount_nano,
        "currency": amount_currency,
        "units": amount_units
      }
    }
    const url = 'tinkoff.public.invest.api.contract.v1.SandboxService/SandboxPayIn';
    const data = this._makeApiCall(url, requestBody);
    return data;
  }
  // ----------------------------- UsersService -----------------------------
  _GetAccounts() {
    const url = 'tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts'
    const data = this._makeApiCall(url,{})
    return data
  }
  _GetUserTariff() {
    const url = 'tinkoff.public.invest.api.contract.v1.UsersService/GetUserTariff'
    const data = this._makeApiCall(url,{})
    return data
  }
  _GetInfo() {
    const url = 'tinkoff.public.invest.api.contract.v1.UsersService/GetInfo'
    const data = this._makeApiCall(url,{})
    return data
  }
}

const tinkoffClientV2 = new _TinkoffClientV2(OPENAPI_TOKEN)

function _GetTickerNameByFIGI(figi) {
  //Logger.log(`[TI_GetTickerByFIGI] figi=${figi}`)   // DEBUG
  const {ticker,name} = tinkoffClientV2._GetInstrumentBy('INSTRUMENT_ID_TYPE_FIGI',null,figi).instrument
  return [ticker,name]
}

function TI_GetInstrumentsID() {
  const AllInstruments = tinkoffClientV2._Bonds('INSTRUMENT_STATUS_BASE').instruments.concat(tinkoffClientV2._Shares('INSTRUMENT_STATUS_BASE').instruments, tinkoffClientV2._Futures('INSTRUMENT_STATUS_BASE').instruments, tinkoffClientV2._Etfs('INSTRUMENT_STATUS_BASE').instruments)
  Logger.log(`[TI_GetInstrumentsID()] Number of instruments: ${AllInstruments.length}`)

  const values = []
  values.push(["Тикер","FIGI","Название","Класс","Биржа","Валюта","Лот","ISIN","UID"])

  for (let i=0; i < AllInstruments.length; i++) {
    values.push([
      AllInstruments[i].ticker,AllInstruments[i].figi,AllInstruments[i].name,AllInstruments[i].classCode,AllInstruments[i].exchange,AllInstruments[i].currency,AllInstruments[i].lot,AllInstruments[i].isin,AllInstruments[i].uid
    ])
  }
  return values
}

function TI_FindInstrument(query, instrumentKind) {
  const resp = tinkoffClientV2._FindInstrument(query, instrumentKind);
  if (resp.instruments.length == 1){
    var instrument = JSON.stringify(resp.instruments[0]);
    Logger.log(`instrument ${instrument}`)
    return resp.instruments[0];
  } else {
    var msg = `Refine your query`
    Logger.log(msg);
    return msg;
  }
}

function TI_GetLastPriceByFigi(figi) {
  if (figi) {
    const data = tinkoffClientV2._GetLastPrices([figi])
    if (data.lastPrices[0].price)
      return tinkoffClientV2._UnitsNanoToPrice(data.lastPrices[0].price.units, data.lastPrices[0].price.nano)
  }
  return null
}

function TI_GetLastPrice(ticker) {
  const figi = _getFigiByTicker(ticker)    // Tinkoff API v1 function !!!
  if (figi) {
    return TI_GetLastPriceByFigi(figi)
  }
}

function TI_GetAccounts() {
  const data = tinkoffClientV2._GetAccounts()

  const values = []
  values.push(["ID","Тип","Название","Статус","Открыт","Права доступа"])
  for (let i=0; i<data.accounts.length; i++) {
    values.push([data.accounts[i].id, data.accounts[i].type.replace('ACCOUNT_TYPE_',''), data.accounts[i].name, data.accounts[i].status.replace('ACCOUNT_STATUS_',''), isoToDate(data.accounts[i].openedDate), data.accounts[i].accessLevel.replace('ACCOUNT_ACCESS_LEVEL_','')])
  }

  return values
}

function TI_GetAccountID(accountNum) {
  if(accountNum >= 0) {
    const data = tinkoffClientV2._GetAccounts()

    return data.accounts[accountNum].id
  }
}


/**
 * Получение Bid/Ask спреда инструмента по тикеру
 * @param {"GAZP"} ticker Тикер инструмента
 * @return {0.03}         Спред в %
 * @customfunction
 */
function TI_GetBidAskSpread(ticker) {
  const figi = _getFigiByTicker(ticker)
  if(figi) {
    const {depth,bids,asks} = tinkoffClientV2._GetOrderBookByFigi(figi, 1)
    if ((bids.length > 0) && (asks.length > 0))
      return ((Number(asks[0].price.units)+asks[0].price.nano/1000000000) - (Number(bids[0].price.units)+bids[0].price.nano/1000000000)) / (Number(asks[0].price.units) + asks[0].price.nano/1000000000)
    else
      return null
  }
}

/**
 * Получение портфеля
 * @param {"12345678"} accountId  Номер брокерского счета
 * @return {Array}                Массив с результатами
 * @customfunction
 **/
function TI_GetPortfolio(accountId) {
  const portfolio = tinkoffClientV2._GetPortfolio(accountId)
  const values = []
  values.push(["Тикер","Название","Тип","Кол-во","Ср.цена покупки","Ст-ть покупки","Валюта","Доход","Тек.ст-ть","Валюта","НКД","Валюта"])
  for (let i=0; i<portfolio.positions.length; i++) {
    const [ticker,name] = _GetTickerNameByFIGI(portfolio.positions[i].figi)
    let quantity = Number(portfolio.positions[i].quantity.units) + portfolio.positions[i].quantity.nano/1000000000
    let averagePositionPrice = Number(portfolio.positions[i].averagePositionPrice.units) + portfolio.positions[i].averagePositionPrice.nano/1000000000
    let currentNkd=null
    let currentNkd_currency=null
    if(portfolio.positions[i].currentNkd) {
      currentNkd = Number(portfolio.positions[i].currentNkd.units) + portfolio.positions[i].currentNkd.nano/1000000000
      currentNkd_currency = portfolio.positions[i].currentNkd.currency
    }
    values.push([
      ticker,
      name,
      portfolio.positions[i].instrumentType,
      quantity,
      averagePositionPrice,
      quantity * averagePositionPrice,
      portfolio.positions[i].averagePositionPrice.currency,
      Number(portfolio.positions[i].expectedYield.units) + portfolio.positions[i].expectedYield.nano/1000000000,
      (Number(portfolio.positions[i].currentPrice.units) + portfolio.positions[i].currentPrice.nano/1000000000) * quantity,
      portfolio.positions[i].currentPrice.currency,
      currentNkd,
      currentNkd_currency
    ])
  }
  return values
}

function _TI_CalculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    const {quantity, price} = trades[j]
    let price_val = Number(price.units) + price.nano/1000000000
    totalQuantity += Number(quantity)
    totalSum += Number(quantity) * price_val
  }
  const weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}

/**
 * Получение операций по счету
 * @param {"12345678"} accountId  Номер брокерского счета
 * @param {2020-02-20} from_param [From date] - Optional
 * @param {2020-12-31} to_param   [To date] - Optional
 * @return {Array}                Массив с результатами
 * @customfunction
 **/
function TI_GetOperations(accountId,from_param, to_param) {
  const limit = 50

  const values = []
  if (!from_param){
    from = TRADING_START_AT.toISOString()
  } else {
    from = from_param.toISOString()
  }
  if (!to_param){
    to = new Date(new Date() + MILLIS_PER_DAY).toISOString()
  } else {
    to = to_param.toISOString()
  }

  let hasNext = false, nextCursor = null

  do {
    // _GetOperationsByCursor(accountId,instrument_id,from,to,cursor,limit,operation_types,state,without_commissions,without_trades,without_overnights)
    const data = tinkoffClientV2._GetOperationsByCursor(accountId, null, from, to, nextCursor, limit,['OPERATION_TYPE_UNSPECIFIED'],'OPERATION_STATE_EXECUTED',true,true,true)
    if (!data.code) {
      // Logger.log(`[TI_GetOperations] items.length=${data.items.length}`) // DEBUG!!!
      hasNext = data.hasNext
      nextCursor = data.nextCursor

      for (let i=0; i<data.items.length; i++) {
        const {date, type, tradesInfo, figi, payment, price, quantity, commission, accruedInt} = data.items[i]

        let payment_val = Number(payment.units)+payment.nano/1000000000
        let accruedInt_val = Number(accruedInt.units)+accruedInt.nano/1000000000

        let com_val = Number(commission.units)+commission.nano/1000000000
        let full_payment_val = payment_val + com_val

        if(tradesInfo) {
          [totalQuantity, totalSum, weigthedPrice] = _TI_CalculateTrades(tradesInfo.trades)
        } else {
          totalQuantity = Number(quantity)
          weigthedPrice = Number(price.units)+price.nano/1000000000
          totalSum = totalQuantity * weigthedPrice
        }

        if ((type == "OPERATION_TYPE_SELL") || (type == "OPERATION_TYPE_SELL_CARD") || (type == "OPERATION_TYPE_SELL_MARGIN")) {
          totalQuantity = -totalQuantity
          totalSum = -totalSum
        }

        if (!totalQuantity) {
          totalQuantity=null
          totalSum=null
          weigthedPrice=null
        }

        if (!accruedInt_val) {
          accruedInt_val = null
        }

        if (!com_val) {
          com_val = null
        }

        var ticker, name
        if (!figi) {
          ticker = null
        } else {
          [ticker,name] = _GetTickerNameByFIGI(figi)
        }

        values.unshift([isoToDate(date), ticker, type.replace('OPERATION_TYPE_',''), totalQuantity, weigthedPrice, totalSum, accruedInt_val, payment_val, com_val, full_payment_val, payment.currency])
      }
    } else {
      throw('error: ',data.message)
    }
  } while (hasNext)

  if (!from_param){
    values.unshift(["Дата","Тикер","Операция","Кол-во","Цена (средн)","Стоимость","НКД (T+)","Итого","Комиссия","Итого (с комисс.)","Валюта"])
  }
  return values
}

function TI_GetSandboxAccounts() {
  const data = tinkoffClientV2._GetSandboxAccounts();

  const values = []
  // values.push(["ID","Тип","Название","Статус","Открыт","Права доступа"])
  for (let i=0; i<data.accounts.length; i++) {
    values.push([data.accounts[i].id, data.accounts[i].type.replace('ACCOUNT_TYPE_',''), data.accounts[i].name, data.accounts[i].status.replace('ACCOUNT_STATUS_',''), isoToDate(data.accounts[i].openedDate), data.accounts[i].accessLevel.replace('ACCOUNT_ACCESS_LEVEL_','')])
  }

  return values
}

function TI_GetSandboxOrders(accountId) { // TODO: prices columns
  const data = tinkoffClientV2._GetSandboxOrders(accountId);
  const dataOrders = data.orders.map(item => item);
  const values = [];
  values.push(_JsonToKV(dataOrders[0]).keys);
  var dataOrdersValues = dataOrders.map(v => _JsonToKV(v).values)
  for (let i=0; i<dataOrders.length; i++) {
    var item = _JsonToKV(dataOrdersValues.pop(i));
    values.push(item.values)
  }
  return values
}

// ------------------------------------------- Installable Triggers -------------------------------------------

/**
 * set Figi by value from IntrumentName in a row on the "Orders" sheet
 */

function TI_Utils_UnitsNanoToPrice(units, nano) {
  var concatenatedValue = units.toString() + nano.toString();
  var combinedNumber = parseFloat(concatenatedValue);
  return combinedNumber;
}

function TI_Utils_UnitsNanoToPrice(jsonData) {
  var concatenatedValue = jsonData.units.toString() + jsonData.nano.toString();
  var combinedNumber = parseFloat(concatenatedValue);
  return combinedNumber;
}

function TI_Utils_PriceToUnitsNano(combinedValue) {
  var stringValue = combinedValue.toString().split(".");
  return { "units": Number(stringValue[0]), "nano": Number(stringValue[1]) };
}

function _JsonToKV(data) {
  var keys = [];
  var values = [];
  for (var key in data) {
    keys.push(key);
    values.push(data[key]);
  }
  return {"keys": keys, "values": values}
}

function _createTransaction(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeRow = ss.getActiveSheet().getRange("A10:J10").getValues();
  var orderTypeRange = ss.getActiveSheet().getRange(row,7);
  var orderIdRange = ss.getActiveSheet().getRange(row,9);

  // Validation
  var errors = []
  if (activeRow.includes(null) || activeRow.includes("")) {
    errors.push(`Error: Fill in all cells in a row ${row}`)
  }
  if (orderTypeRange.getValue() == "" || orderTypeRange.getValue() == "ORDER_TYPE_UNSPECIFIED") {
    errors.push(`Error: orderType in ${orderTypeRange.getA1Notation()} cant be empty or "ORDER_TYPE_UNSPECIFIED"`);
  }
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(errors.join(",\n"));
    return;
  }

  var activeSs = ss.getActiveSheet();
  var price = TI_Utils_PriceToUnitsNano(activeSs.getRange(row, 6).getValue());
  var direction = activeSs.getRange(row, 8);
  var directionValue = direction.getValue();
  var respData = null;

  // Creation
  if (orderIdRange.getValue() == "") {
    if (ss.getRangeByName("IS_SANDBOX").getValue() == true){
      SpreadsheetApp.getUi().alert("YOU'RE IN SANDBOX MODE");
      respData = tinkoffClientV2._PostSandboxOrder(
        activeSs.getRange(row, 4).getValue(),
        activeSs.getRange(row, 5).getValue(),
        price.nano,
        price.units,
        directionValue,
        activeSs.getRange(row, 1).getValue(),
        activeSs.getRange(row, 7).getValue()
      );
    } else {
      respData = tinkoffClientV2._PostOrder(
        activeSs.getRange(row, 4).getValue(),
        activeSs.getRange(row, 5).getValue(),
        price.nano,
        price.units,
        directionValue,
        activeSs.getRange(row, 1).getValue(),
        activeSs.getRange(row, 7).getValue()
      );
    }
    if (respData) {
      var data = _JsonToKV(respData);
      activeSs.getRange(10, 9,1, data.keys.length).setValues([data.keys]);
      activeSs.getRange(row, 9,1, data.keys.length).setValues([data.values]);
      activeSs.getRange(row, 10).setValue(respData.executionReportStatus.replace("EXECUTION_REPORT_STATUS_", ""));
      activeSs.getRange(row, 13).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.initialOrderPrice.units, respData.initialOrderPrice.nano)} ${respData.initialOrderPrice.currency}`);
      activeSs.getRange(row, 14).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.executedOrderPrice.units, respData.executedOrderPrice.nano)} ${respData.executedOrderPrice.currency}`);
      activeSs.getRange(row, 15).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.totalOrderAmount.units, respData.totalOrderAmount.nano)} ${respData.totalOrderAmount.currency}`);
      activeSs.getRange(row, 16).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.initialCommission.units, respData.initialCommission.nano)} ${respData.initialCommission.currency}`);
      activeSs.getRange(row, 17).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.executedCommission.units, respData.executedCommission.nano)} ${respData.executedCommission.currency}`);
      activeSs.getRange(row, 20).setValue(`${tinkoffClientV2._UnitsNanoToPrice(respData.initialSecurityPrice.units, respData.initialSecurityPrice.nano)} ${respData.initialSecurityPrice.currency}`);
      if (directionValue == "ORDER_DIRECTION_SELL"){
        orderIdRange.setBackground("red");
      } else if (directionValue == "ORDER_DIRECTION_BUY"){
        orderIdRange.setBackground("green");
      }
    }
  } else {
    SpreadsheetApp.getUi().alert(`The deal has already been assigned an orderId ${orderIdRange.getValue()}`)
  }
}

function _testCreateTransaction(){
  _createTransaction(11);
}

function onEditOrdersSheet(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getActiveSheet().getName() != "Orders") {
      return;
    }
    var activeCell = ss.getActiveCell();
    var col = activeCell.getColumn();
    var row = activeCell.getRow();
    if ((col == 2 || col == 3) && row > 10) { //InstrumentName, InstrumentKind
      var instrumentNameRange = ss.getActiveSheet().getRange(row,2);
      var instrumentKindRange = ss.getActiveSheet().getRange(row,3);
      var figiRange = ss.getActiveSheet().getRange(row,4);
      var priceRange = ss.getActiveSheet().getRange(row,6);

      figiRange.setValue(null).setFontColor(null);
      var foundInstrument = TI_FindInstrument(instrumentNameRange.getValue(), instrumentKindRange.getValue());
      Logger.log(`figiRange ${figiRange.getA1Notation()}, foundInstrument ${foundInstrument.figi}, type ${instrumentKindRange.getValue()}`);
      if (typeof(foundInstrument) != "string"){
        figiRange.setValue(foundInstrument.figi);
        priceRange.setValue(TI_GetLastPriceByFigi(foundInstrument.figi));
      } else {
        figiRange.setValue(foundInstrument).setFontColor("red");
        priceRange.setValue(null);
      }
    } else if (col == 8 && row > 10) {
      if (ss.getActiveSheet().getRange(row, col).getValue() != "ORDER_DIRECTION_UNSPECIFIED") {
        _createTransaction(row);
      }
    } else {
      Logger.log("conditions do not match")
      return;
    }
}
