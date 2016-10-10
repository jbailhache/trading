
import httplib
import urllib
import json
import datetime

long_account_id = "101-004-2088872-001"
short_account_id = "101-004-2088872-003"
main_account_id = long_account_id
key = "15c8e8e9c56b87f251c1a27c9d616c2b-1fcf01e73ffbd2992ea1c91e14ada204"
# key = "19546e57a150b42484042e0884e68a26-b1f4f3b637cbb7d186ce0e9e82e74ea6"

# Roger
# account_id = "101-004-3943284-001"
# key = "46e76556f88ed69e8bfb42b621793fc0-f70a275d2a04e5051aae22800f046477"

# account_id = "101-004-3943284-002"
# key = "b2dace0ec807b2d60e4fcf7cf72f72ed-e364407ed760c25e5fb247d324a09d2e"

hedging = True

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

def get_balance_account (account_id) :
 summary = get_summary_account(account_id)
 balance = summary["balance"]
 return float(balance)

def get_balance () :
 return get_balance_account(long_account_id) + get_balance_account(short_account_id)

def get_nav_account (account_id) :
 summary = get_summary_account(account_id)
 nav = summary['NAV']
 return float(nav)

def get_nav () :
 return get_nav_account(long_account_id) + get_nav_account(short_account_id)

def get_details_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id, {})
 return response.get("account")

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
 return get_trades_account(long_account_id) + get_trades_account(short_account_id)

def get_orders_account (account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/orders", {})
 orders = response["orders"]
 return orders 

def get_orders () :
 return get_orders_account(long_account_id) + get_orders_account(short_account_id)

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
 return get_positions_account(long_account_id) + get_positions_account(short_account_id)

def get_position_account_instr (account_id, iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/positions/" + iname, {})
 return response

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
 return get_units_account_instr(long_account_id,iname) + get_units_account_instr(short_account_id,iname)

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

def get_transactions_account (account_id, fromtime, totime) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/transactions?from=" + fromtime + "&to=" + totime, {})
 return response

def get_orders_account_instr (account_id, iname) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/orders?instrument=" + iname + "&count=500", {})
 return response["orders"]

def get_orders_instr (iname) :
 return get_orders_account_instr(long_account_id,iname) + get_orders_account_instr(short_account_id,iname)

def get_pending_orders_account(account_id) :
 response = oanda("GET", "/v3/accounts/" + account_id + "/pendingOrders", {})
 return response["orders"]

def get_pending_orders () :
 return get_pending_orders_account(long_account_id) + get_pending_orders_account(short_account_id)

def market_order_account (account_id, iname, n) :
 params = {"order" : 
           {"instrument" : iname,
            "units" : str(n),
            "type" : "MARKET"
           }};
 print "params = ", params
 response = oanda("POST", "/v3/accounts/" + account_id + "/orders", params)
 return response

def market_order (iname, n) :
 if n > 0 :
  return market_order_account (long_account_id, iname, n)
 elif n < 0 :
  return market_order_account (short_account_id, iname, n)
 else :
  return None 

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
 if n > 0 :
  return order_account (long_account_id, iname, n, typ, price)
 elif n < 0 :
  return order_account (short_account_id, iname, n, typ, price)
 else :
  return None

def close_long_account_instr (account_id, iname) :
 params = {"longUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_short_account_instr (account_id, iname) :
 params = {"shortUnits" : "ALL"}
 response = oanda("PUT", "/v3/accounts/" + account_id + "/positions/" + iname + "/close", params);
 return response

def close_account_instr (account_id, iname) :
 u = get_units_account_instr(account_id,iname)
 if u > 0 :
  return close_long_account_instr(account_id,iname)
 elif u < 0 :
  return close_short_account_instr(account_id,iname)

def close_instr (iname) :
 return [close_long_account_instr(long_account_id,iname), close_short_account_instr(short_account_id,iname)] 

def close_all_account (account_id) :
 inames = get_instruments_names_account(account_id)
 for iname in inames :
  print iname
  close_account_instr(account_id,iname)

def close_all () :
 print "LONG:"
 close_all_account(long_account_id)
 print ""
 print "SHORT:"
 close_all_account(short_account_id)

def cancel_order_account (account_id, id) :
 response = oanda("PUT", "/v3/accounts/" + account_id + "/orders/" + str(id) + "/cancel", {})
 print response
 return response

def cancel_orders_account_instr (account_id, instr) :
 orders = get_orders_account_instr (account_id, instr)
 for o in orders :
  # print "Cancel order ", o
  cancel_order_account (account_id, o["id"])

def cancel_orders_instr (instr) :
 cancel_orders_account_instr(long_account_id)
 cancel_orders_account_instr(short_account_id)

def cancel_all_account (account_id) :
 inames = get_instruments_names_account(account_id)
 for iname in inames :
  print iname
  cancel_orders_account_instr (account_id, iname)

def cancel_all () :
 print "LONG:"
 cancel_all_account(long_account_id)
 print ""
 print "SHORT:"
 cancel_all_account(short_account_id)

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
  if self.units > 0 :
   return self.send_account(long_account_id)
  elif self.units < 0 :
   return self.send_account(short_account_id)
  else :
   return None

def print_dic (dic) :
 for cle in dic.keys() :
  print " %30s : %s" % (cle, dic[cle]) 

def print_dic_list (l) :
 print len(l), " elements"
 for i in range(len(l)) :
  print i, " : "
  print_dic(l[i])

def print_orders(orders) :
 for o in orders :
  print str_order(o)

def init() :
 global instruments
 instruments = get_instruments_account(main_account_id)

"""
def test1 () :
 summary = get_summary()
 print summary
 for key in summary.keys():
  print key, " => ", summary[key]
 # print "balance = ", get_balance()
 print "balance = ", summary[u'balance']
 print "balance = ", get_balance()
 print "instruments names = ", get_instruments_names()
 prices = get_prices("EUR_USD,FR40_EUR")
 print "prices = ", prices
 ask = prices[0]["asks"]
 bid = prices[0]["bids"]
 print "ask = ", ask, " ; bid = ", bid	
 print "positions : ", get_positions()
 print "market order : ", market_order("EUR_USD", 1)
 print "positions : ", get_positions()
 # print "close : ", close_all("EUR_USD")
 print "market order : ", market_order("EUR_USD", -1)
 print "positions : ", get_positions()
 
# test()

def test_summary () :
 summary = get_summary()
 print "summary = ", summary
 if "NAV" in summary :
  print "NAV trouve = ", summary["NAV"]
 else :
  print "NAV non trouve"
 if "truc" in summary :
  print "truc trouve"
 else :
  print "truc non trouve"
 details = get_details()
 print "details = ", details
 positions = len(details["positions"])
 print "positions: ", positions, " ", details["positions"]
 trades = get_trades()
 print "trades: ", len(trades), trades

def essais () :
 global instruments
 instruments = get_instruments()
 trades = get_trades()
 print "trades = "
 for t in trades:
  print t
  print
 positions = get_positions()
 print "positions = "
 for p in positions['positions']:
  print p
  print
 position = get_position_instr("NZD_CAD")
 print "position = ", position
 units = get_units_instr("NZD_CAD")
 print "units = ", units
 position = get_position_instr("XAG_NZD")
 print "position = ", position
 units = get_units_instr("XAG_NZD")
 print "units = ", units 
 position = get_position_instr("BCO_USD")
 print "position = ", position
 units = get_units_instr("BCO_USD")
 print "units = ", units
 o = order("FR40_EUR", -10, "STOP", 4490.12345)
 print "order -> ", o
 print
 print "orders :"
 orders = get_orders()
 for o in orders :
  print o
  print

# essais()

def test_orders() :
 orders = get_orders_instr("FR40_EUR")
 print len(orders)
 for o in orders :
  s_order = o["id"] + ": " + o["type"] + " " + str(o["units"]) + " " + o["instrument"] + " at " + str(o["price"]) + ", " + o["state"] + "; "
  print s_order
 print
 orders = get_pending_orders()
 print len(orders)
 for o in orders :
  s_order = o["id"] + ": " + o["type"] + " " + str(o["units"]) + " " + o["instrument"] + " at " + str(o["price"]) + ", " + o["state"] + "; "
  print s_order

# test_orders()
"""

