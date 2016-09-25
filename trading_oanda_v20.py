
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


oDoc = XSCRIPTCONTEXT.getDocument()

createUNOService = (
	XSCRIPTCONTEXT
	.getComponentContext()
	.getServiceManager()
	.createInstance
		    )

feuille = oDoc.CurrentController.ActiveSheet

key = feuille.getCellRangeByName("O5").String
# account_id = int(feuille.getCellRangeByName("O4").Value)
account_id = feuille.getCellRangeByName("O4").String


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
 if sys.version_info.major < 3 :
  conn = httplib.HTTPSConnection("api-fxpractice.oanda.com")
 else :
  conn = http.client.HTTPSConnection("api-fxpractice.oanda.com")
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

def get_summary () :
 response = oanda("GET", "/v3/accounts/" + account_id + "/summary", {})
 return response.get("account")

def get_balance () :
 summary = get_summary()
 balance = summary["balance"]
 return float(balance)

def get_details () :
 response = oanda("GET", "/v3/accounts/" + account_id, {})
 return response.get("account")

def get_instruments () :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments", {})
 instruments = response["instruments"]
 return instruments

def get_instruments_names () :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments", {})
 instruments = response["instruments"]
 instruments_names = []
 for instrument in instruments :
  instruments_names.append(instrument["name"])
 return instruments_names

def get_instrument (iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/instruments?instruments=" + iname, {})
 instrument = response["instruments"][0]
 return instrument

def get_pip_location (iname) :
 global instruments
 for instrument in instruments :
  if instrument["name"] == iname :
   return instrument["pipLocation"]
 return None

def get_trades() :
 response = oanda("GET", "/v3/accounts/" + account_id + "/trades", {})
 trades = response["trades"]
 return trades

def get_prices(inames) :
 if inames == "" :
  return []
 else :
  response = oanda("GET", "/v3/accounts/" + account_id + "/pricing?instruments=" + inames, {})
  return response["prices"]

def get_ask(iname) :
 prices = get_prices(iname)
 if len(prices["asks"]) > 0:
  ask = prices["asks"][0]["price"]
  return float(ask)
 else :
  return 0.0

def get_bid(iname) :
 prices = get_prices(iname)
 if len(prices["bids"]) > 0 :
  bid = prices["bids"][0]["price"]
  return float(bid)
 else :
  return 0.0

def get_price(iname) :
 prices = get_prices(iname)
 ask = prices["asks"][0]["price"]
 bid = prices["bids"][0]["price"]
 price = (float(ask)+float(bid))/2.0
 return price

def get_positions() :
 response = oanda("GET", "/v3/accounts/" + account_id + "/positions", {})
 return response

def get_position_instr (iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/positions/" + iname, {})
 return response

def get_units_instr (iname) :
 position = get_position_instr(iname)
 if 'position' in position.keys() :
  s = int(position['position']['short']['units'])
  l = int(position['position']['long']['units'])
  u = s + l
  return u
 else :
  return 0

def get_orders_instr(iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/orders?instrument=" + iname + "&count=500", {})
 return response["orders"]

def market_order (iname, n) :
 params = {"order" : 
           {"instrument" : iname,
            "units" : str(n),
            "type" : "MARKET"
           }};
 # print "params = ", params
 response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
 return response

def order (iname, n, typ, price) :
 global feuille
 global account_id
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

 def send(self) :
  o = self.build()
  params = {"order" : o};
  # print "params = ", params
  # feuille.getCellByPosition(1,1).String = "params=" + str(params)
  response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
  # feuille.getCellByPosition(1,1).String = "response=" + str(response)
  # print "order(", typ, ",", price, ") -> ", response
  return response

# def close_all (iname) :
#  params = {"longUnits" : "ALL"}
#  response = oanda("PUT", "/v3/accounts/101-004-2088872-001/positions/EUR_USD/close", params);
#  return response

def close_long_instr (iname) :
 params = {"longUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_short_instr (iname) :
 params = {"shortUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_instr (iname) :
 u = get_units_instr(iname)
 if u > 0 :
  return close_long_instr(iname)
 elif u < 0 :
  return close_short_instr(iname) 

def str_order1(o) :
	feuille.getCellByPosition(1,1).String = repr(o)
	cond = " ?? "
	if "units" in o :
		units = int(o["units"])
	else :
		units = 0
		av = " A/V "
	if "instrument" in o :
		instrument = o["instrument"]
	else :
		instrument = "*"
	if units > 0 :
		av = "Achat"
		if o["type"] == "LIMIT" :
			cond = " <= "
		elif o["type"] == "STOP" :
			cond = " >= "
		elif o["type"] == "MARKET_IF_TOUCHED" :
			cond = " == "
	elif units < 0 :
		av = "Vente"
		if o["type"] == "LIMIT" :
			cond = " >= "
		elif o["type"] == "STOP" :
			cond = " <= "
		elif o["type"] == "MARKET_IF_TOUCHED" :
			cond = " == "
	tpof_price = None
	if "takeProfitOnFill" in o :
		if "price" in o["takeProfitOnFill"] :
			tpof_price = o["takeProfitOnFill"]["price"]
	s_order = o["id"] + ": " + o["type"] + " " + av + " " + str(abs(units)) + " " + instrument + cond + str(o["price"]) 
	if tpof_price != None :
		s_order += ", TPOF: " + str(tpof_price)
	return s_order

def str_order(o) :
	return str_order1(o) # + " " + repr(o)

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
	# ttlc = int(feuille.getCellRangeByName("B25").Value)

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

	formule = [""] * (dercol + 1)

	for i in range(1,2): # 1,nc

		for j in range(3+nv,dercol+1):
			formule[j] = feuille.getCellByPosition(j,ligne).Formula

		# feuille.insertCells(range, com.sun.star.sheet.CellInsertMode.DOWN)
		feuille.insertCells(range1, DOWN)

		for j in range(3,dercol+1):
			feuille.getCellByPosition(j,ligne).NumberFormat = feuille.getCellByPosition(j,ligne+1).NumberFormat

		now = datetime.datetime.now()
		feuille.getCellByPosition(1,ligne).String = now.strftime("%d/%m/%Y")
		feuille.getCellByPosition(2,ligne).String = now.strftime("%H:%M:%S")

		# summary = get_summary()
		details = get_details()

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
			if "cours initial" in feuille.getCellByPosition(1,ligne+1).String :
				feuille.getCellByPosition(colcours[j],ligne+1).Value = bidding_prices[j]
				feuille.getCellByPosition(colcours[j]+1,ligne+1).Value = asking_prices[j]

		# feuille.getCellRangeByName("A28").Value = dercol
		# feuille.getCellRangeByName("A29").String = formule[6]

		for j in range(3,dercol+1):
			action = feuille.getCellByPosition(j,ligne-4).String
			id1 = feuille.getCellByPosition(j,ligne-3).String
			if ((action != 'prices' and action != 'TRADE' and id1 != 'REQUEST' and id1 != 'RESPONSE') or id1 == "") and feuille.getCellByPosition(j,ligne).String == "" :
				feuille.getCellByPosition(j,ligne).Formula = formule[j]

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
			# if sob == "sell" or sob == "buy" :
			if semblable(sob,"sell") or semblable(sob,"buy") :
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
							response = market_order(id1, u)
						else :
							response = 'fictive'
						trace = 'order: ' + sob + ' ' + ("%f" % un) + ' ' + id1
						feuille.getCellByPosition(dercol+1,ligne).String = trace
						feuille.getCellByPosition(dercol+2,ligne).String = repr(response)
			
			# if sob == "MARKET" or sob == "LIMIT" or sob == "STOP" or sob == "MARKET_IF_TOUCHED" :
			if semblable(sob,"MARKET") or semblable(sob,"LIMIT") or semblable(sob,"STOP") or semblable(sob,"MARKET_IF_TOUCHED") :
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
							response = o.send()
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
				feuille.getCellByPosition(j,ligne).Value = get_balance()

			# elif sob == "TRADES" :
			elif semblable(sob,"TRADES") :
				feuille.getCellByPosition(j,ligne).Value = len(get_trades())

			# elif sob == "POSITION" :
			elif semblable(sob,"POSITION") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				feuille.getCellByPosition(j,ligne).Value = get_units_instr(instr)

			# elif sob == "ORDERS" :
			elif semblable(sob,"ORDERS") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				orders = get_orders_instr(instr)
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
				orders = get_orders_instr(instr)
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
	
			elif semblable(sob,"CLOSE") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					close_instr(instr)

			elif semblable(sob,"+CLOSE") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					close_account_instr(long_account_id,instr)

			elif semblable(sob,"-CLOSE") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				close = feuille.getCellByPosition(j,ligne).Value
				if close :
					close_account_instr(short_account_id,instr)

			elif semblable(sob,"CANCEL") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					cancel_orders_instr(instr)
			
			elif semblable(sob,"+CANCEL") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					cancel_orders_account_instr(long_account_id,instr)

			elif semblable(sob,"-CANCEL") :
				instr = feuille.getCellByPosition(j,ligne-3).String
				cancel = feuille.getCellByPosition(j,ligne).Value
				if cancel :
					cancel_orders_account_instr(short_account_id,instr)
			
			elif sob in details :
				feuille.getCellByPosition(j,ligne).Value = float(details[sob])
				
			elif len(sob) > 0 and sob[0] == '#' :
				feuille.getCellByPosition(j,ligne).Value = len(details[sob[1:]])

			elif len(sob) > 0 and sob[0] == '$' :
				feuille.getCellByPosition(j,ligne).String = str(details[sob[1:]])

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
	instrs = get_instruments_names()
	instruments = get_instruments()
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
		trading1(i)
		i += 1

def trading() :
	t = Thread(None, tradingloop, None, (),  {})
	t.start()

