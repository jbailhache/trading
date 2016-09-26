
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


# class Order :
#  def __init__(self, iname, n, typ, price) :
#   self.iname = iname
#   self.units = n
#   self.typ = typ
#   self.price = price
#   self.state = 'PENDING'

def str_price (price, instrument) :
 price1 = str(price)
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
  self.state = 'PENDING'

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


oDoc = XSCRIPTCONTEXT.getDocument()

createUNOService = (
	XSCRIPTCONTEXT
	.getComponentContext()
	.getServiceManager()
	.createInstance
		    )

feuille = oDoc.CurrentController.ActiveSheet

ligne = int(feuille.getCellRangeByName("B3").Value) - 1
dercol = int(feuille.getCellRangeByName("B4").Value)
nc = int(feuille.getCellRangeByName("B5").Value)
delai = feuille.getCellRangeByName("B6").Value
ncag = int(feuille.getCellRangeByName("B7").Value)
reel = feuille.getCellRangeByName("B8").String
iliq = feuille.getCellRangeByName("B21").Value
intervalle = feuille.getCellRangeByName("B22").Value
col = feuille.getCellRangeByName("B23").Value


def fill (o, bid, ask, pbid, pask, position) : 
 # if o.positionFill == 'OPEN_ONLY' :
 #  if o.units > 0 and position < 0 :
 #   return False
 #  if o.units < 0 and position > 0 :
 #   return False
 if o.typ == 'MARKET' :
  return True
 elif o.typ == 'LIMIT' :
  if o.units > 0 :
   return ask <= o.price
  elif o.units < 0 :
   return bid >= o.price
  else :
   return False
 elif o.typ == 'STOP' :
  if o.units > 0 :
   return ask >= o.price
  elif o.units < 0 :
   return bid <= o.price
  else :
   return False
 elif o.typ == 'MARKET_IF_TOUCHED' :
  if o.units > 0 :
   return (pask <= o.price and ask >= o.price) or (pask >= o.price and ask <= o.price) 
  elif o.units < 0 :
   return (pbid <= o.price and bid >= o.price) or (pbid >= o.price and bid <= o.price)
  else :
   return False
 else :
  return False

def number_pending(orders) :
 npending = 0
 for o in orders :
  if o.state == 'PENDING' :
   npending += 1
 return npending


def simul_thread () :

 global oDoc, feuille, ligne, dercol, nc, delai, ncag, reel, iliq, intervalle, col

 feuille = oDoc.CurrentController.ActiveSheet

 feuille.getCellByPosition(1,1).String = "Debut simulation"

 ligne = int(feuille.getCellRangeByName("B3").Value) - 1
 # ligne -= 1
 dercol = int(feuille.getCellRangeByName("B4").Value)
 nc = int(feuille.getCellRangeByName("B5").Value)
 delai = feuille.getCellRangeByName("B6").Value
 ncag = int(feuille.getCellRangeByName("B7").Value)
 reel = feuille.getCellRangeByName("B8").String
 iliq = feuille.getCellRangeByName("B21").Value
 intervalle = feuille.getCellRangeByName("B22").Value
 col = feuille.getCellRangeByName("B23").Value

 feuille.getCellByPosition(1,1).String = "Simulation en cours"
 for l in range(ligne,30000) :
  if "cours initial" in feuille.getCellByPosition(1,l).String :
   lci = l
   break
 feuille.getCellByPosition(1,1).Value = lci

 inames = []
 
 for c in range(0,dercol+1) :
  if feuille.getCellByPosition(c,ligne-4).String == 'prices' :
   iname = feuille.getCellByPosition(c,ligne-3).String
   if not iname in inames :
    inames.append(iname)

 positions = {}
 bids = {}
 asks = {}
 for iname in inames :
  positions[iname] = 0
  bids[iname] = 0
  asks[iname] = 0

 orders = [] 
 liq = iliq

 for l in range(lci-1,ligne-1,-1) :

  trade = ((lci - l) % intervalle) == 0
  if trade :
   feuille.getCellByPosition(3,l).String = "***"
  else :
   feuille.getCellByPosition(3,l).String = ""

  # feuille.getCellByPosition(3,l).String = str(trade)
  # feuille.getCellByPosition(12,l).Value = bids['FR40_EUR']

  filled_orders = ""
  for c in range(0,dercol+1) :
   op = feuille.getCellByPosition(c,ligne-4).String

   if op == '*NAV' :
    # feuille.getCellByPosition(12,l).Value = bids['FR40_EUR']
    nav = liq
    for iname in inames :
     nav += positions[iname] * bids[iname]
    feuille.getCellByPosition(c,l).Value = nav

   elif op == '*POSITION' :
    iname = feuille.getCellByPosition(c,ligne-3).String
    feuille.getCellByPosition(c,l).Value = positions[iname]

   elif op == '*FILL' :
    feuille.getCellByPosition(c,l).String = filled_orders

   elif op == '*#PENDING' :
    npending = number_pending(orders)
    feuille.getCellByPosition(c,l).Value = npending

   elif op == '*ORDERS' :
    s = ''
    for o in orders :
     s += repr(o.build()) + '; '
    feuille.getCellByPosition(c,l).String = s

   elif (op == 'MARKET' or op == 'LIMIT' or op == 'STOP' or op == 'MARKET_IF_TOUCHED') and trade :
    units = feuille.getCellByPosition(c,l).Value
    npending = number_pending(orders)
    if units != 0 and npending < 500 :
     iname = feuille.getCellByPosition(c,ligne-3).String    
     typ = op
     o = Order(iname, units, typ)
     if op != 'MARKET' :
      price = feuille.getCellByPosition(c+1,l).Value
      o.price = price
     positionFill = feuille.getCellByPosition(c,ligne-2).String
     if positionFill != '' :
      o.positionFill = positionFill

     if op == 'MARKET' :
      col = c+1
     else :
      col = c+2
     col_req = -1
     col_res = -1
     items = 'items:'
     while True :
      item = feuille.getCellByPosition(col,ligne-3).String
      items += item + ';'
      value = feuille.getCellByPosition(col,l).Value
      # if semblable (item, 'PRICE') :
      #	o.price = value
      if item == 'TPOF' :
       o.tpof_price = value
      elif item == 'SLOF' :
       o.slof_price = value
      elif item == 'TSLOF' :
       o.tslof_distance = value
      elif item == 'REQUEST' :
       col_req = col
       # feuille.getCellByPosition(col,ligne).String = 'request'
      elif item == 'RESPONSE' :
       col_res = col
       # feuille.getCellByPosition(col,ligne).String = 'response'
      elif col > c+1 :
       break
      col += 1

     orders.append(o)

   elif op == 'CLOSE' :
    iname = feuille.getCellByPosition(c,ligne-3).String
    close = feuille.getCellByPosition(c,l).Value
    if close :
     if positions[iname] > 0 :
      liq += positions[iname] * bids[iname]
      positions[iname] = 0
      feuille.getCellByPosition(0,l).String = "+C"
     elif positions[iname] < 0 :
      liq += positions[iname] * asks[iname]
      positions[iname] = 0
      feuille.getCellByPosition(0,l).String = "-C"

   elif op == 'CANCEL' :
    iname = feuille.getCellByPosition(c,ligne-3).String
    cancel = feuille.getCellByPosition(c,l).Value
    if cancel :
     orders1 = {}
     for o in orders :
      if o.instrument != iname :
       orders1.append(o)
     orders = orders1

   elif op == 'TRADE' :
    iname = feuille.getCellByPosition(c,ligne-3).String
    bid = feuille.getCellByPosition(c,l).Value
    bids[iname] = bid
    ask = feuille.getCellByPosition(c+1,l).Value
    asks[iname] = ask
    pbid = feuille.getCellByPosition(c,l+1).Value
    pask = feuille.getCellByPosition(c+1,l+1).Value
    for o in orders :
     if o.state == 'PENDING' and o.instrument == iname and fill(o,bid,ask,pbid,pask,positions[iname]) :
      o.state = 'FILLED'
      positions[iname] += o.units
      if o.units > 0 :
       liq -= o.units * ask
      elif o.units < 0 :
       liq -= o.units * bid
      filled_orders += " " + o.typ + " " + str(o.units) + " " + o.instrument + " at " + str(o.price) + ", TPOF at " + str(o.tpof_price) + " ; "
      if o.tpof_price != None :
       o1 = Order (o.instrument, -o.units, 'LIMIT')
       o1.price = o.tpof_price
       orders.append(o1)
      if o.slof_price != None :
       o1 = Order (o.instrument, -o.units, 'STOP')
       o1.price = slof_price
       orders.append(o1)

 feuille.getCellByPosition(1,1).String = "Simulation terminee"

 # s_orders = ""
 # for o in orders :
 #  s_orders += " " + o.typ + " " + str(o.units) + " " + o.iname + " at " + str(o.price) + " ; "
 # feuille.getCellByPosition(1,1).String = s_orders
 

def simul () :
	t = Thread(None, simul_thread, None, (),  {})
	t.start()


def charger1 (date, time, bid, ask, col) :

	nv = 0

	range1 = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
	range1.Sheet = 0
	range1.StartColumn = 1
	range1.EndColumn = dercol+2
	range1.StartRow = ligne
	range1.EndRow = ligne

	formule = [""] * (dercol + 1)

	for j in range(3+nv,dercol+1):
		formule[j] = feuille.getCellByPosition(j,ligne).Formula

	feuille.insertCells(range1, DOWN)

	for j in range(3,dercol+1):
		feuille.getCellByPosition(j,ligne).NumberFormat = feuille.getCellByPosition(j,ligne+1).NumberFormat

	feuille.getCellByPosition(1,ligne).String = date
	feuille.getCellByPosition(2,ligne).String = time

	feuille.getCellByPosition(col,ligne).Value = bid
	feuille.getCellByPosition(col+1,ligne).Value = ask

	if "cours initial" in feuille.getCellByPosition(1,ligne+1).String :
		feuille.getCellByPosition(col,ligne+1).Value = bid
		feuille.getCellByPosition(col+1,ligne+1).Value = ask

	for j in range(3,dercol+1):
		action = feuille.getCellByPosition(j,ligne-4).String
		id1 = feuille.getCellByPosition(j,ligne-3).String
		if (action != 'prices' or id1 == "") and feuille.getCellByPosition(j,ligne).String == "" :
				feuille.getCellByPosition(j,ligne).Formula = formule[j]

	for j in range(0,feuille.Charts.Count):
		chart = feuille.Charts.getByIndex(j)
		ranges = chart.Ranges
		for range1 in ranges:
			range1.StartRow = ligne
			range1.EndRow = ligne + ncag
		chart.Ranges = ranges

def charger_thread () :

 global ligne, dercol, nc, delai, ncag, reel, iliq, intervalle, col
 ligne = int(feuille.getCellRangeByName("B3").Value) - 1
 dercol = int(feuille.getCellRangeByName("B4").Value)
 nc = int(feuille.getCellRangeByName("B5").Value)
 delai = feuille.getCellRangeByName("B6").Value
 ncag = int(feuille.getCellRangeByName("B7").Value)
 reel = feuille.getCellRangeByName("B8").String
 iliq = feuille.getCellRangeByName("B21").Value
 intervalle = feuille.getCellRangeByName("B22").Value
 col = feuille.getCellRangeByName("B23").Value

 infile = open("/tmp/cours.txt", "r")
 # col = 4
 for line in infile :
  # time.sleep(1)
  fields = line.split(";")
  date = fields[0]
  time = fields[1]
  bid = fields[2]
  ask = fields[3]
  charger1(date, time, bid, ask, col)

def charger () :
	t = Thread(None, charger_thread, None, (),  {})
	t.start()


 
