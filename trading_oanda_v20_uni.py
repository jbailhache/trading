
import uno
import urllib
import datetime
import time
from threading import Thread
from com.sun.star.sheet.CellInsertMode import DOWN

# import oandapy
# from oandapy import API 

import json
# import requests

import sys

if sys.version_info.major < 3 :
	import httplib
else :
	import http.client
import math

hedging = False 

oDoc = XSCRIPTCONTEXT.getDocument()

createUNOService = (
	XSCRIPTCONTEXT
	.getComponentContext()
	.getServiceManager()
	.createInstance
		    )

feuille = oDoc.CurrentController.ActiveSheet

# document = oDoc.CurrentController.Frame
desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()
document = model.getCurrentController()
# document = oDoc.CurrentController
dispatcher = createUNOService("com.sun.star.frame.DispatchHelper")

key = feuille.getCellRangeByName("O5").String
# account_id = int(feuille.getCellRangeByName("O4").Value)

if hedging :
 long_account_id = feuille.getCellRangeByName("O3").String
 short_account_id = feuille.getCellRangeByName("O4").String
 main_account_id = long_account_id
else :
 account_id = feuille.getCellRangeByName("O4").String
 main_account_id = account_id

def semblable(s1,s2) :
 if s1 == s2 :
  return True
 if type(s1) != type('') :
  return False
 if type(s2) != type(''):
  return False
 return s1.strip().upper() == s2.strip().upper()


##########################################################################################

# account_id = "101-004-3943284-001"
# key = "b2dace0ec807b2d60e4fcf7cf72f72ed-e364407ed760c25e5fb247d324a09d2e"

def oanda (method, url, params) :
 conn = httplib.HTTPSConnection("api-fxpractice.oanda.com")
 # eparams = urllib.urlencode(params)
 eparams = json.dumps(params)
 headers = {"Content-Type" : "application/json", 
            "Authorization" : "Bearer " + key}
 conn.request(method, url, eparams, headers)
 response = conn.getresponse()
 # print "status = ", response.status
 content = response.read()
 # print "content = ", content
 # print "msg = ", response.msg
 # print "reason = ", response.reason
 content = content.decode('utf-8')	
 try:
  content = json.loads(content)
  return content
 except:
  return ""

def get_summary_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/summary", {})
 return response.get("account")

def get_summary () :
 return get_summary_account(main_account_id)

def get_balance_account (account_id) :
 summary = get_summary_account(account_id)
 balance = summary["balance"]
 return float(balance)

def get_balance () :
 if hedging :
  return get_balance_account(long_account_id) + get_balance_account(short_account_id)
 else :
  return get_balance_account(main_account_id)

def get_nav_account (account_id) :
 summary = get_summary_account(account_id)
 nav = summary['NAV']
 return float(nav)

def get_nav () :
 if hedging :
  return get_nav_account(long_account_id) + get_nav_account(short_account_id)
 else :
  return get_nav_account(main_account_id)

def get_details_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id, {})
 return response.get("account")

def get_details () :
 return get_details_account(main_account_id)

def get_instruments_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments", {})
 instruments = response["instruments"]
 return instruments

def get_instruments () :
 return get_instruments_account (main_account_id)

def get_instruments_names_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments", {})
 instruments = response["instruments"]
 instruments_names = []
 for instrument in instruments :
  instruments_names.append(instrument["name"])
 return instruments_names

def get_instruments_names () :
 return get_instruments_names_account (main_account_id)

def get_account_instrument (account_id, iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments?instruments=" + iname, {})
 instrument = response["instruments"][0]
 return instrument

def get_instrument (iname) :
 return get_account_instrument (main_account_id, iname)

def get_pip_location (iname) :
 global instruments
 for instrument in instruments :
  if instrument["name"] == iname :
   return instrument["pipLocation"]-1
 return None

def get_trades_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/trades", {})
 trades = response["trades"]
 return trades

def get_trades () :
 if hedging :
  return get_trades_account(long_account_id) + get_trades_account(short_account_id)
 else :
  return get_trades_account(main_account_id)

def get_orders_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/orders", {})
 orders = response["orders"]
 return orders 

def get_orders () :
 if hedging :
  return get_orders_account(long_account_id) + get_orders_account(short_account_id)
 else :
  return get_orders_account(main_account_id)

def get_prices_account (account_id, inames) :
 if inames == "" :
  return []
 else :
  response = oanda("GET", "/v3/accounts/" + account_id + "/pricing?instruments=" + inames, {})
  return response["prices"]

def get_prices (iname) :
 return get_prices_account(main_account_id,iname)

def get_ask_account (account_id, iname) :
 prices = get_prices_account (account_id, iname)[0]
 if len(prices["asks"]) > 0:
  ask = prices["asks"][0]["price"]
  return float(ask)
 else :
  return 0.0

def get_ask (iname) :
 return get_ask_account (main_account_id, iname)

def get_bid_account (account_id, iname) :
 prices = get_prices_account (account_id, iname)[0]
 if len(prices["bids"]) > 0 :
  bid = prices["bids"][0]["price"]
  return float(bid)
 else :
  return 0.0

def get_bid (iname) :
 return get_bid_account (main_account_id, iname)

def get_price_account (account_id, iname) :
 prices = get_prices_account (account_id, iname)[0]
 ask = prices["asks"][0]["price"]
 bid = prices["bids"][0]["price"]
 price = (float(ask)+float(bid))/2.0
 return price

def get_price (iname) :
 return get_price_account (main_account_id, iname)

def get_positions_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/positions", {})
 return response["positions"]

def get_positions () :
 if hedging :
  return get_positions_account(long_account_id) + get_positions_account(short_account_id)
 else :
  return get_positions_account(main_account_id)

def get_position_account_instr (account_id, iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/positions/" + iname, {})
 return response

def get_position_instr (iname) :
 return get_position_account_instr (main_account_id, iname)

def get_units_account_instr (account_id, iname) :
 position = get_position_account_instr (account_id, iname)
 if 'position' in position.keys() :
  s = int(position['position']['short']['units'])
  l = int(position['position']['long']['units'])
  u = s + l
  return u
 else :
  return 0

def get_units_instr (iname) :
 if hedging :
  return get_units_account_instr(long_account_id,iname) + get_units_account_instr(short_account_id,iname)
 else :
  return get_units_account_instr(main_account_id,iname)

def get_long_units_instr (iname) :
 return get_units_account_instr(long_account_id,iname) 

def get_short_units_instr (iname) :
 return get_units_account_instr(short_account_id,iname)

def set_units_account_instr (account_id, iname, np) :
 ap = get_units_account_instr (account_id, iname)
 dp = np - ap
 if dp != 0 :
  response = market_order_account (account_id, iname, dp)
  return response
 else :
  return None

def set_units_instr (iname, np) :
 return set_units_account_instr (main_account_id, iname, np)

def get_transactions_account (account_id, fromtime, totime) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/transactions?from=" + fromtime + "&to=" + totime, {})
 return response

def get_transactions (fromtime, totime) :
 return get_transactions_account (main_account_id, fromtime, totime)

def get_orders_account_instr (account_id, iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/orders?instrument=" + iname + "&count=500", {})
 return response["orders"]

def get_orders_instr (iname) :
 if hedging :
  return get_orders_account_instr(long_account_id,iname) + get_orders_account_instr(short_account_id,iname)
 else :
  return get_orders_account_instr(main_account_id,iname)

def get_pending_orders_account(account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/pendingOrders", {})
 return response["orders"]

def get_pending_orders () :
 if hedging :
  return get_pending_orders_account(long_account_id) + get_pending_orders_account(short_account_id)
 else :
  return get_pending_orders_account(main_account_id)

def market_order_account (account_id, iname, n) :
 params = {"order" : 
           {"instrument" : iname,
            "units" : str(n),
            "type" : "MARKET"
           }};
 # print "params = ", params
 response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
 return response

def market_order (iname, n) :
 if hedging :
  if n > 0 :
   return market_order_account (long_account_id, iname, n)
  elif n < 0 :
   return market_order_account (short_account_id, iname, n)
  else :
   return None 
 else :
  return market_order (main_account_id, iname, n)

def order_account (account_id, iname, n, typ, price) :
 global feuille
 # global account_id
 # feuille.getCellByPosition(1,1).String = "order"
 # price1 = ("%10.5f" % price).strip()
 pip_location = get_pip_location(iname)
 price1 = str(round(price, -pip_location))
 params = {"order" : 
           {"instrument" : iname,
            "units" : str(n),
            "type" : typ,
            "price" : price1
           }};
 # print "params = ", params
 # feuille.getCellByPosition(1,1).String = "params=" + str(params)
 response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
 # feuille.getCellByPosition(1,1).String = "response=" + str(response)
 # print "order(", typ, ",", price, ") -> ", response
 return response

def order (iname, n, typ, price) :
 if hedging :
  if n > 0 :
   return order_account (long_account_id, iname, n, typ, price)
  elif n < 0 :
   return order_account (short_account_id, iname, n, typ, price)
  else :
   return None
 else :
  return order_account (main_account_id, iname, n, typ, price)

def close_long_account_instr (account_id, iname) :
 params = {"longUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_long_instr (iname) :
 return close_long_account_instr (main_account_id, iname)

def close_short_account_instr (account_id, iname) :
 params = {"shortUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_short_instr (iname) :
 return close_short_account_instr (main_account_id, iname)

def close_account_instr (account_id, iname) :
 u = get_units_account_instr(account_id,iname)
 if u > 0 :
  return close_long_account_instr(account_id,iname)
 elif u < 0 :
  return close_short_account_instr(account_id,iname)

def close_instr (iname) :
 if hedging :
  return [close_long_account_instr(long_account_id,iname), close_short_account_instr(short_account_id,iname)] 
 else :
  return close_account_instr (main_account_id, iname)

def close_all_account (account_id) :
 inames = get_instruments_names_account(account_id)
 for iname in inames :
  # print iname
  close_account_instr(account_id,iname)

def close_all () :
 if hedging :
  # print "LONG:"
  close_all_account(long_account_id)
  # print ""
  # print "SHORT:"
  close_all_account(short_account_id)
 else :
  close_all_account(main_account_id)

def cancel_order_account (account_id, id) :
 response = oanda("PUT", "/v3/accounts/" + account_id + "/orders/" + str(id) + "/cancel", {})
 # print response
 return response

def cancel_order (id) :
 cancel_order_account (main_account_id, id)

def cancel_orders_account_instr (account_id, instr) :
 orders = get_orders_account_instr (account_id, instr)
 for o in orders :
  # print "Cancel order ", o
  cancel_order_account (account_id, o["id"])

def cancel_orders_instr (instr) :
 if hedging :
  cancel_orders_account_instr(long_account_id)
  cancel_orders_account_instr(short_account_id)
 else :
  cancel_orders_account_instr(main_account_id)

def cancel_all_account (account_id) :
 inames = get_instruments_names_account(account_id)
 for iname in inames :
  # print iname
  cancel_orders_account_instr (account_id, iname)

def cancel_all () :
 if hedging :
  # print "LONG:"
  cancel_all_account(long_account_id)
  # print ""
  # print "SHORT:"
  cancel_all_account(short_account_id)
 else :
  cancel_all_account(main_account_id)

def str_order(o) :
	if int(o["units"]) > 0 :
		av = "Achat"
		if o["type"] == "LIMIT" :
			cond = " <= "
		elif o["type"] == "STOP" :
			cond = " >= "
	elif int(o["units"]) < 0 :
		av = "Vente"
		if o["type"] == "LIMIT" :
			cond = " >= "
		elif o["type"] == "STOP" :
			cond = " <= "
	# s_order = o["id"] + ": " + o["type"] + " " + av + " " + str(abs(int(o["units"]))) + " " + o["instrument"] + cond + str(o["price"])
	s_order = " %6s : %-7s %5s %5s %-10s %2s %7s " % (o["id"], o["type"], av, str(abs(int(o["units"]))), o["instrument"], cond, o["price"])
	return s_order

def str_price (price, instrument) :
 pip_location = get_pip_location(instrument)
 price1 = str(round(price, -pip_location))
 return price1

class Order :

 def __init__(self, instrument, units, typ) :  
  self.instrument = instrument
  self.units = units
  self.typ = typ
  self.price = None
  self.tpof_price = None
  self.slof_price = None
  self.tslof_distance = None
  self.positionFill = None

 def build(self) :
  o = {"instrument" : self.instrument,
       "units" : str(self.units),
       "type" : self.typ,
      };
  if self.price != None :
   o["price"] = str_price(self.price,self.instrument);
  if self.tpof_price != None :
   o["takeProfitOnFill"] = { "price" : str_price(self.tpof_price,self.instrument) };
  if self.slof_price != None :
   o["stopLossOnFill"] = { "price" : str_price(self.slof_price,self.instrument) };
  if self.tslof_distance != None :
   o["trailingStopLossOnFill"] = { "distance" : str_price(self.tslof_distance,self.instrument) };
  if self.positionFill != None :
   o["positionFill"] = self.positionFill;
  return o;

 def send_account(self, account_id) :
  o = self.build()
  params = {"order" : o};
  # print "params = ", params
  # feuille.getCellByPosition(1,1).String = "params=" + str(params)
  response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
  # feuille.getCellByPosition(1,1).String = "response=" + str(response)
  # print "order(", typ, ",", price, ") -> ", response
  return response

 def send(self) :
  if hedging :
   if self.units > 0 :
    return self.send_account(long_account_id)
   elif self.units < 0 :
    return self.send_account(short_account_id)
   else :
    return None
  else :
   return self.send_account(main_account_id)


##############################################################################


# oanda = API(environment="practice", access_token=key)

# feuille.getCellRangeByName("A21").String = repr(oanda.get_instruments())

def trading1(nieme):

	global prev_orders

	# feuille.getCellByPosition(1,1).Value = nieme

	# feuille = oDoc.CurrentController.ActiveSheet

	ligne = int(feuille.getCellRangeByName("B3").Value) - 1
	dercol = int(feuille.getCellRangeByName("B4").Value)
	nc = int(feuille.getCellRangeByName("B5").Value)
	delai = feuille.getCellRangeByName("B6").Value
	ncag = int(feuille.getCellRangeByName("B7").Value)
	reel = feuille.getCellRangeByName("B8").String
	amorce = feuille.getCellRangeByName("B24").Value
	# ttlc = int(feuille.getCellRangeByName("B25").Value)
	nlignes = int(feuille.getCellRangeByName("B26").Value)

	nv = 0

	# Dim range as new com.sun.star.table.CellRangeAddress
	# range1 = com.sun.star.table.CellRangeAddress()
	range1 = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
	range1.Sheet = 0
	range1.StartColumn = 1
	range1.EndColumn = dercol+2
	range1.StartRow = ligne
	range1.EndRow = ligne

	maxcours = dercol - 2
	ncours = 0

	colcours = [0] * maxcours
	ids = [""] * maxcours

	for j in range(3,dercol+1):
		action = feuille.getCellByPosition(j,ligne-4).String
		id1 = feuille.getCellByPosition(j,ligne-3).String
		if (action == 'prices' or action == 'TRADE') and id1 != "" :
			colcours[ncours] = j
			ids[ncours] = id1
			ncours = ncours + 1

	# formule = [[""] * nlignes] * (dercol + 1)
	formule = []
	for j in range(dercol+1) :
		formule.append ([""] * nlignes)   

	for i in range(1,2): # 1,nc

		for j in range(3+nv,dercol+1):
			for k in range(0,nlignes) :
				formule[j][k] = feuille.getCellByPosition(j,ligne+k).Formula

		feuille.getCellByPosition(1,1).String = str(formule)

		# feuille.insertCells(range, com.sun.star.sheet.CellInsertMode.DOWN)
		for k in range(0,nlignes) :
			feuille.insertCells(range1, DOWN)

		for j in range(3,dercol+1):
			for k in range(0,nlignes) :
				feuille.getCellByPosition(j,ligne).NumberFormat = feuille.getCellByPosition(j,ligne+k).NumberFormat

		feuille.getCellByPosition(2,1).String = "hhh"

		now = datetime.datetime.now()
		feuille.getCellByPosition(1,ligne).String = now.strftime("%d/%m/%Y")
		feuille.getCellByPosition(2,ligne).String = now.strftime("%H:%M:%S")

		# summary = get_summary()
		try :
			details = get_details()
		except :
			details = []

		# sURL = "http://finance.yahoo.com/d/quotes.csv?s="
		# for j in range(0,ncours):
		#       if j > 0 :
		#	       sURL = sURL + ","
		#       sURL = sURL + ids[j]
		# sURL = sURL + "&f=l1"

		instruments = ""
		for j in range(0,ncours) :
			if j > 0 :
				instruments = instruments + ","
			instruments = instruments + ids[j]
		# feuille.getCellRangeByName("A21").String = instruments

		# oSimpleFileAccess = createUNOService("com.sun.star.ucb.SimpleFileAccess")
		# oInpDataStream = createUNOService ("com.sun.star.io.TextInputStream")
		# oInpDataStream.setInputStream(oSimpleFileAccess.openFileRead(sURL))

		# cours = []
		# while not oInpDataStream.isEOF() :
		#       cours = cours + [ oInpDataStream.readLine() ]

		# response = oanda.get_prices(instruments=instruments)
		# prices = response.get("prices")
		# asking_prices = []
		# bidding_prices = []
		# for price in prices :
		#	asking_prices = asking_prices + [ price.get("ask") ]
		#	bidding_prices = bidding_prices + [ price.get("bid") ]

		try :
			response = get_prices(instruments)
			feuille.getCellByPosition(1,1).String = str(response)
			prices = response
			asking_prices = []
			bidding_prices = []
			for price in prices :
				asking_prices = asking_prices + [ price["asks"][0]["price"] ]
				bidding_prices = bidding_prices + [ price["bids"][0]["price"] ]

		# asking_prices = range(0, ncours)
		# bidding_prices = range(0, ncours)

		# feuille.getCellRangeByName("A22").String = repr(asking_prices)
		# feuille.getCellRangeByName("A23").Value = asking_prices[0]
		# feuille.getCellRangeByName("A24").Value = asking_prices[1]
		# feuille.getCellRangeByName("A25").Value = asking_prices[2]

		# for j in range(0,len(cours)):
		#       feuille.getCellByPosition(colcours[j],ligne).Value = float(cours[j])

			for j in range(0,len(asking_prices)) :
				feuille.getCellByPosition(colcours[j],ligne).Value = bidding_prices[j]
				feuille.getCellByPosition(colcours[j]+1,ligne).Value = asking_prices[j]
				# if nieme == 0 :
				if "cours initial" in feuille.getCellByPosition(1,ligne+nlignes).String :
					feuille.getCellByPosition(colcours[j],ligne+nlignes).Value = bidding_prices[j]
					feuille.getCellByPosition(colcours[j]+1,ligne+nlignes).Value = asking_prices[j]

		except :
			pass

		# feuille.getCellRangeByName("A28").Value = dercol
		# feuille.getCellRangeByName("A29").String = formule[6]

		for j in range(3,dercol+1):
			action = feuille.getCellByPosition(j,ligne-4).String
			id1 = feuille.getCellByPosition(j,ligne-3).String
			if ((action != 'prices' and action != 'TRADE' and id1 != 'REQUEST' and id1 != 'RESPONSE') or id1 == "") and feuille.getCellByPosition(j,ligne).String == "" :
				for k in range(0,nlignes) : 
					feuille.getCellByPosition(j,ligne+k).Formula = formule[j][k]

		# feuille.getCellRangeByName("A30").Value = feuille.Charts.Count

		for j in range(0,feuille.Charts.Count):
			chart = feuille.Charts.getByIndex(j)
			ranges = chart.Ranges
			for range1 in ranges:
				range1.StartRow = ligne
				range1.EndRow = ligne + ncag
			chart.Ranges = ranges

		time.sleep(1)

		for j in range(3,dercol+1) :
			# feuille.getCellByPosition(0,ligne).String = '*'
			sob = feuille.getCellByPosition(j,ligne-4).String
	
			if semblable(sob,"COPY") :
				copyfrom = feuille.getCellByPosition(j,ligne-3).String
				copyto = feuille.getCellByPosition(j,ligne-2).String

				args1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
				args1.Name = "ToPoint"
				args1.Value = copyfrom

				args2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
				args2.Name = "ToPoint"
				args2.Value = copyto

				args3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
				args3.Name = "ToPoint"
				args3.Value = "A1"

				dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, tuple([args2]))
				dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, ())
				dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, tuple([args1]))
				dispatcher.executeDispatch(document, ".uno:Copy", "", 0, ())
				dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, tuple([args2]))
				dispatcher.executeDispatch(document, ".uno:Paste", "", 0, ())
				dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, tuple([args3]))

			# if sob == "sell" or sob == "buy" :
			elif (semblable(sob,"sell") or semblable(sob,"buy")) and nieme >= amorce :				
				id1 = feuille.getCellByPosition(j,ligne-3).String
				if id1 != "" :
					un = feuille.getCellByPosition(j,ligne).Value
					if un > 0 :
						if reel == 'oui' :
							# response = oanda.create_order(account_id, instrument=id1, units=int(un), side=sob, type='market')
							if sob == "buy" :
								u = int(un)
							elif sob == "sell" :
								u = - int(un)
							try :
								response = market_order(id1, u)
							except :
								response = 'error'
						else :
							response = 'fictive'
						trace = 'order: ' + sob + ' ' + ("%f" % un) + ' ' + id1
						feuille.getCellByPosition(dercol+1,ligne).String = trace
						feuille.getCellByPosition(dercol+2,ligne).String = repr(response)

			elif semblable(sob,"!POSITION") and nieme >= amorce :
				id1 = feuille.getCellByPosition(j,ligne-3).String
				if id1 != "" :
					np = int(feuille.getCellByPosition(j,ligne).Value)
					# ap = get_units_instr(id1)
					# dp = np - ap
					# if dp != 0 :
					#	response = market_order (id1, dp)
					#	feuille.getCellByPosition(dercol+2,ligne).String = repr(response)
					try :
						response = set_units_instr(id1,np)
					except :
						response = 'error'
					if response != None :
						feuille.getCellByPosition(dercol+2,ligne).String = repr(response)
			
			# if sob == "MARKET" or sob == "LIMIT" or sob == "STOP" or sob == "MARKET_IF_TOUCHED" :
			elif (semblable(sob,"MARKET") or semblable(sob,"LIMIT") or semblable(sob,"STOP") or semblable(sob,"MARKET_IF_TOUCHED")) and nieme >= amorce :
				# feuille.getCellByPosition(3,ligne).String = 'o1'
				id1 = feuille.getCellByPosition(j,ligne-3).String
				if id1 != "" :
					un = feuille.getCellByPosition(j,ligne).Value	
					if un != 0 :
						if reel == 'oui' :
							# feuille.getCellByPosition(0,ligne).String = 'o2'
							o = Order (id1, int(un), sob)
							if not semblable (sob, "MARKET") :
								o.price = feuille.getCellByPosition(j+1,ligne).Value
							positionFill = feuille.getCellByPosition(j,ligne-2).String
							if positionFill != '' :
								o.positionFill = positionFill
							if semblable (sob, 'MARKET') :
								col = j+1
							else :
								col = j+2
							col_req = -1
							col_res = -1
							items = 'items:'
							while True :
								item = feuille.getCellByPosition(col,ligne-3).String
								items += item + ';'
								value = feuille.getCellByPosition(col,ligne).Value
								# if semblable (item, 'PRICE') :
								#	o.price = value
								if semblable (item, 'TPOF') :
									o.tpof_price = value
								elif semblable (item, 'SLOF') :
									o.slof_price = value
								elif semblable (item, 'TSLOF') :
									o.tslof_distance = value
								elif semblable (item, 'REQUEST') :
								# elif item == 'REQUEST' :
									col_req = col
									feuille.getCellByPosition(col,ligne).String = 'request'
								elif semblable (item, 'RESPONSE') :
								# elif item == 'RESPONSE' :
									col_res = col
									feuille.getCellByPosition(col,ligne).String = 'response'
								elif col > j+1 :
									break
								col += 1
							# feuille.getCellByPosition(12,ligne).String = items
							# feuille.getCellByPosition(dercol+4,ligne).String = '<---'
							if col_req > 0 :
								feuille.getCellByPosition(col_req,ligne).String = repr(o.build())
							try :
								response = o.send()
							except :
								response = 'error'
							if col_res > 0 :
								feuille.getCellByPosition(col_res,ligne).String = repr(response)
							# if sob == "MARKET" :
							# 	u = int(un)
							# 	response = market_order(id1, u)
							# else :
							#	u = int(un)
							#	price = feuille.getCellByPosition(j+1,ligne).Value
							#	feuille.getCellByPosition(1,1).String = "avant order"
							#	# feuille.getCellByPosition(1,1).String = "id1=" + str(id1) 
							#	response = order(id1, u, sob, price)
							#	feuille.getCellByPosition(1,1).String = "apres order"
						else :
							response = 'fictive'
						trace = 'order: ' + sob + ' ' + ("%f" % un) + ' ' + id1
						feuille.getCellByPosition(dercol+1,ligne).String = trace
						feuille.getCellByPosition(dercol+2,ligne).String = repr(response)
						
			# if sob == "BALANCE" :
			if semblable(sob,"BALANCE") :
				try :
					feuille.getCellByPosition(j,ligne).Value = get_balance()
				except :
					pass
			# elif sob == "TRADES" :
			elif semblable(sob,"TRADES") :
				try :
					feuille.getCellByPosition(j,ligne).Value = len(get_trades())
				except :
					pass
			# elif sob == "POSITION" :
			elif semblable(sob,"POSITION") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				try :
					feuille.getCellByPosition(j,ligne).Value = get_units_instr(instr)
				except :
					pass

			elif semblable(sob,"LONG") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				feuille.getCellByPosition(j,ligne).Value = get_long_units_instr(instr)

			elif semblable(sob,"SHORT") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				feuille.getCellByPosition(j,ligne).Value = get_short_units_instr(instr)


			# elif sob == "ORDERS" :
			elif semblable(sob,"ORDERS") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				try :
					orders = get_orders_instr(instr)
				except :
					orders = []
				s_orders = ""
				pending = 0
				filled = 0
				triggered = 0
				cancelled = 0
				other = 0
				for o in orders :
					# s_orders += o["id"] + ": " + o["type"] + " " + str(o["units"]) + " " + o["instrument"] + " at " + str(o["price"]) + ", " + o["state"] + "; "
					s_orders += str_order(o) + "; "
					if o["state"] == 'PENDING' :
						pending += 1
					elif o["state"] == 'FILLED' :
						filled += 1
					elif o["state"] == 'TRIGGERED' :
						triggered += 1
					elif o["state"] == 'CANCELLED' :
						cancelled += 1
					else :
						other += 1
				feuille.getCellByPosition(j,ligne).String = str(len(orders)) + " P:" + str(pending) + " F:" + str(filled) + " T:" + str(triggered) + " C:" + str(cancelled) + " O:" + str(other) + " " + s_orders

			# elif sob == "FILLED" or sob == "FILL" :
			elif semblable(sob,"FILLED") or semblable(sob,"FILL") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				try :
					orders = get_orders_instr(instr)
				except :
					orders = []
				s_orders = ""
				n_filled = 0
				np = len(prev_orders[j])
				no = len(orders)
				for o in prev_orders[j] :
					found = False
					for o1 in orders :
						if o1["id"] == o["id"] :
							found = True
							break
					if not found :
						n_filled += 1
						# s_orders += o["id"] + ": " + o["type"] + " " + str(o["units"]) + " " + o["instrument"] + " at " + str(o["price"]) + ", " + o["state"] + "; "
						s_orders += str_order(o) + "; "
				prev_orders[j] = orders
				feuille.getCellByPosition(j,ligne).String = str(np) + "->" + str(no) + " " + str(n_filled) + " : " + s_orders
	
			elif semblable(sob,"CLOSE") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					try :
						close_instr(instr)
					except :
						pass

			elif semblable(sob,"+CLOSE") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					try :
						close_account_instr(long_account_id,instr)
					except :
						pass

			elif semblable(sob,"-CLOSE") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					try :
						close_account_instr(short_account_id,instr)
					except :
						pass

			elif semblable(sob,"CANCEL") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					try :
						cancel_orders_instr(instr)
					except :
						pass
		
			elif semblable(sob,"+CANCEL") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					try :
						cancel_orders_account_instr(long_account_id,instr)
					except :
						pass

			elif semblable(sob,"-CANCEL") and nieme >= amorce :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					try :
						cancel_orders_account_instr(short_account_id,instr)
					except :
						pass
	
			elif sob in details :
				feuille.getCellByPosition(j,ligne).Value = float(details[sob])
				
			elif len(sob) > 0 and sob[0] == '#' :
				try :
					feuille.getCellByPosition(j,ligne).Value = len(details[sob[1:]])
				except :
					feuille.getCellByPosition(j,ligne).String = "Error"

			elif len(sob) > 0 and sob[0] == '$' :
				try :
					feuille.getCellByPosition(j,ligne).String = str(details[sob[1:]])
				except :
					feuille.getCellByPosition(j,ligne).String = "Error"

		time.sleep(delai/1000)


def tradingloop() :
	global account_id, key, instruments, prev_orders
	feuille = oDoc.CurrentController.ActiveSheet
	feuille.getCellByPosition(1,1).String = "Test"

	ligne = int(feuille.getCellRangeByName("B3").Value) - 1
	nc = int(feuille.getCellRangeByName("B5").Value)

	# instruments = oanda.get_instruments(account_id)
	# instrs = []
	# for instrument in instruments["instruments"] :
	#	instrs += [ instrument["instrument"] ]
	
	"""
	feuille.getCellByPosition(1,1).String = "trace 1"
	# response = oanda("GET", "/v3/accounts/" + account_id + "/instruments", {})
	method = "GET"
	feuille.getCellByPosition(1,1).String = "trace 1.0.1"
	url = "/v3/accounts/" + account_id + "/instruments"
	feuille.getCellByPosition(1,1).String = "trace 1.0.2"
	params = {}
	feuille.getCellByPosition(1,1).String = "trace 1.0.3"
	conn = httplib.HTTPSConnection("api-fxpractice.oanda.com")
	feuille.getCellByPosition(1,1).String = "trace 1.1"
	# eparams = urllib.urlencode(params)
	eparams = json.dumps(params)
	feuille.getCellByPosition(1,1).String = "trace 1.2"
	headers = {"Content-Type" : "application/json", 
		   "Authorization" : "Bearer " + key}
	feuille.getCellByPosition(1,1).String = "trace 1.3"
	conn.request(method, url, eparams, headers)
	feuille.getCellByPosition(1,1).String = "trace 1.4"
	response = conn.getresponse()
	feuille.getCellByPosition(1,1).String = "trace 1.5"
	# print "status = ", response.status
	content = response.read()
	feuille.getCellByPosition(1,1).String = "trace 1.6"
	# print "content = ", content
	# print "msg = ", response.msg
	# print "reason = ", response.reason
	content = content.decode('utf-8')	
	feuille.getCellByPosition(1,1).String = "trace 1.7"
	try:
		feuille.getCellByPosition(1,1).String = "trace 1.8"
 		content = json.loads(content)
		feuille.getCellByPosition(1,1).String = "trace 1.9"
		response = content
		feuille.getCellByPosition(1,1).String = "trace 1.10"
	except:
		feuille.getCellByPosition(1,1).String = "trace 1.11"
		response = ""
		feuille.getCellByPosition(1,1).String = "trace 1.12"

	feuille.getCellByPosition(1,1).String = "trace 2"
	instruments = response["instruments"]
	feuille.getCellByPosition(1,1).String = "trace 3"
	instruments_names = []
	feuille.getCellByPosition(1,1).String = "trace 4"
	for instrument in instruments :
		instruments_names.append(instrument["name"])
	feuille.getCellByPosition(1,1).String = "trace 5"
	instrs = instruments_names
	feuille.getCellByPosition(1,1).String = str(instrs)
	"""

	# feuille.getCellByPosition(1,1).String = "..."

	# req = "/v3/accounts/" + account_id + "/instruments"
	# feuille.getCellByPosition(1,1).String = req
	# response = oanda("GET", req, {})
	# feuille.getCellByPosition(2,1).String = response
	# instrs = get_instruments_names()

	try :
		instrs = get_instruments_names()
	except :
		instrs = []

	try :
		instruments = get_instruments()
	except :
		instruments = []

	# feuille.getCellByPosition(1,1).String = "Essai"
	feuille.getCellByPosition(1,1).String = str(instrs)

	dercol = int(feuille.getCellRangeByName("B4").Value)
	prev_orders = [[] for i in range(dercol+1)]

	premcolcours = feuille.getCellRangeByName("B9").Value

	# ttlc = int(feuille.getCellRangeByName("B25").Value)

	if premcolcours > 0 :
		for i in range(len(instrs)):
			feuille.getCellByPosition(premcolcours+i*2,ligne-4).String = 'prices'
			feuille.getCellByPosition(premcolcours+i*2,ligne-3).String = instrs[i]
			feuille.getCellByPosition(premcolcours+i*2+1,ligne-4).String = ''
			feuille.getCellByPosition(premcolcours+i*2+1,ligne-3).String = ''

	# for i in range(nc) :
	i = 0
	while i < nc :
		# trading1(i)
		try :
			trading1(i)
		except :
			pass
		i += 1

def trading() :
	t = Thread(None, tradingloop, None, (),  {})
	t.start()

