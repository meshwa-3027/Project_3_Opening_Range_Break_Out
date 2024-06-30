########################## Row Based Approach #############################
#from pya3 import*
from alice_credentials import*
import xlwings as xw
import json
import pandas as pd
from datetime import datetime,date,time
import pdb

# AliceBlue Credentials
alice = login()

# Initialize Excel 
wb = xw.Book('Opening_Range_Break_Out.xlsx')
sht = wb.sheets['Sheet1']

current_time =datetime.now().strftime('%H:%M:%S')
break_out_time = str(time(9,15,59))
exchange_off_time = str(time(15,30,0))
intraday_trade_square_off_time = str(time(15,20,0))
investment_per_symbol = 100000
intraday_leverage = 5
margin_available = int (investment_per_symbol * intraday_leverage)
max_breakout_until_qty_increment = 5

if current_time < break_out_time or current_time > exchange_off_time:
    sht.range("C1:T200").value = None
else:
    sht.range("C1:J200").value = None
sht.range("O1").value = "Zone_Size"
sht.range("P1").value = "BO_dir"
sht.range("Q1").value = "BO_counter"
sht.range("R1").value = "Qty"
sht.range("S1").value = "Target"
sht.range("T1").value = "Trade Status"
# pdb.set_trace() # debugging 

# WebSocket Connection 
LTP = 0
socket_opened = False
subscribe_flag = False
subscribe_list = []
unsubscribe_list = []
data = {}

def socket_open():
    print("Connected")
    global socket_opened
    socket_opened = True
    if subscribe_flag:
        alice.subscribe(subscribe_list)

def socket_close():
    global socket_opened, LTP
    socket_opened = False
    LTP = 0
    print("Closed")

def socket_error(message):
    global LTP
    LTP = 0
    print("Error :", message)

def feed_data(message):
    global LTP, subscribe_flag, data
    feed_message = json.loads(message)
    if feed_message["t"] == "ck":
        print("Connection Acknowledgement status :%s (Websocket Connected)" % feed_message["s"])
        subscribe_flag = True
        print("subscribe_flag :", subscribe_flag)
        print("-------------------------------------------------------------------------------")
        pass
    elif feed_message["t"] == "tk":
        token = feed_message["tk"]
        if "ts" in feed_message:
            symbol = feed_message["ts"]
        else:   
            symbol = token  # For indices
        data[symbol] = {
            "Open": feed_message.get("o", 0),
            "High": feed_message.get("h", 0),
            "Low": feed_message.get("l", 0),
            "LTP": feed_message.get("lp", 0),
            "OI": feed_message.get("toi", 0),
            "VWAP": feed_message.get("ap", 0),
            "PrevDayClose": feed_message.get("c", 0),
                   }
        # print(f"Token Acknowledgement status for {symbol}: {feed_message}")
        print("-------------------------------------------------------------------------------")
        pass
    else:
        # print("Feed :", feed_message)
        LTP = feed_message["lp"] if "lp" in feed_message else LTP

alice.start_websocket(socket_open_callback=socket_open, socket_close_callback=socket_close,
                      socket_error_callback=socket_error, subscription_callback=feed_data, run_in_background=True, market_depth=False)

#Iterate over the rows in the sheet1
instrument = []
for row in sht.range('A2:B12').value:
    exchange,symbol = row
    if exchange and symbol:
        instrument.append((exchange,symbol))

subscribe_list = []
for exchange,symbol in instrument:
    subscribe_list.append(alice.get_instrument_by_symbol(exchange,symbol))
    new_subscribe_list = subscribe_list

def update_column_o(zone_high, zone_low, sht, row_no):
    if zone_high and zone_low is not None:
        zone_size = zone_high - zone_low
        sht.range("O"+str(row_no)).value = zone_size

def update_column_p_q_r_s_t(sht, row_no, alice, ltp, zone_high, zone_low, margin_available, max_breakout_until_qty_increment):
    """Update column O (BO_dir) based on LTP comparison with Zone High and Low."""
    zone_size = sht.range("O"+str(row_no)).value
    current_BO_dir = sht.range("P"+str(row_no)).value
    breakout_counter = sht.range("Q"+str(row_no)).value
    qty = sht.range("R"+str(row_no)).value
    target = sht.range("S"+str(row_no)).value
    trade_status = sht.range("T"+str(row_no)).value
    if trade_status != "Target Achieved":
        if current_time < intraday_trade_square_off_time:
            if current_BO_dir is None:
                if ltp > zone_high:
                    current_BO_dir = "High"
                    repeat(margin_available,max_breakout_until_qty_increment,ltp,sht, row_no, current_BO_dir,zone_high,zone_low,zone_size)
                    
                if ltp < zone_low:
                    current_BO_dir = "Low"
                    repeat(margin_available,max_breakout_until_qty_increment,ltp,sht, row_no, current_BO_dir,zone_high,zone_low,zone_size)
                    
            else:
                if current_BO_dir == "High" and ltp < zone_low:
                    current_BO_dir = "Low"
                    repeat_breakout_counter_greater_than_1(row_no,sht,alice,qty,max_breakout_until_qty_increment,margin_available,ltp,zone_high,zone_size,zone_low)
                    
                if current_BO_dir == "Low" and ltp > zone_high:
                    current_BO_dir = "High"    
                    repeat_breakout_counter_greater_than_1(row_no,sht,alice,qty,max_breakout_until_qty_increment,margin_available,ltp,zone_high,zone_size,zone_low)
                    
        else:
            place_order(row_no, sht, alice, current_BO_dir, qty)
            trade_status = "INTRADAY SQUARE-OFF TIME. TRADE CLOSED"
            sht.range("T"+str(row_no)).value = trade_status
    else:
        pass
    return current_BO_dir

def place_order(row_no, sht, alice, current_BO_dir, qty):
    exchange = sht.range("A"+str(row_no)).value 
    symbol = sht.range("B"+str(row_no)).value 
    print ("%%%%%%%%%%%%%%%%%%%%%%%%%%%%1%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
    print(
    alice.place_order(transaction_type = TransactionType.Buy if current_BO_dir == "High" else TransactionType.Sell,
                        instrument = alice.get_instrument_by_symbol(exchange, symbol),
                        quantity = int(qty),
                        order_type = OrderType.Market,
                        product_type = ProductType.Intraday,
                        price = 0.0,
                        trigger_price = None,
                        stop_loss = None,
                        square_off = None,
                        trailing_sl = None,
                        is_amo = False,
                        order_tag='order1')
    )

def repeat(margin_available,max_breakout_until_qty_increment,ltp,sht, row_no, current_BO_dir,zone_high,zone_low,zone_size  ):
    breakout_counter = 1
    qty = int (((margin_available / max_breakout_until_qty_increment)/ ltp) * breakout_counter)
    sht.range("P"+str(row_no)).value = current_BO_dir
    sht.range("Q"+str(row_no)).value = breakout_counter
    sht.range("R"+str(row_no)).value = qty
    if ltp > zone_high:
        target = float(zone_high + (zone_size * breakout_counter))
    else:
        target = float(zone_low - (zone_size * breakout_counter))
    #Trade Status
    if current_BO_dir == "High" and ltp < target:
        trade_status = "OPEN"
    if current_BO_dir == "High" and ltp > target:
        trade_status = "Target Achieved"
    if current_BO_dir == "Low" and ltp > target:
        trade_status = "OPEN"
    if current_BO_dir == "Low" and ltp < target:
        trade_status = "Target Achieved"   
        
    sht.range("S"+str(row_no)).value = target
    sht.range("T"+str(row_no)).value = trade_status
    place_order(row_no, sht, alice, current_BO_dir, qty)

def repeat_breakout_counter_greater_than_1(row_no,sht,alice,qty,max_breakout_until_qty_increment,margin_available,ltp,zone_high,zone_size,zone_low):
    sht.range("P"+str(row_no)).value = current_BO_dir
    breakout_counter = sht.range("Q"+str(row_no)).value
    place_order(row_no, sht, alice, current_BO_dir, qty)

    breakout_counter = breakout_counter + 1
    if breakout_counter <= max_breakout_until_qty_increment:
        qty = int (((margin_available / max_breakout_until_qty_increment)/ ltp) * breakout_counter)
    else:
        qty = int (((margin_available / max_breakout_until_qty_increment)/ ltp))
    sht.range("Q"+str(row_no)).value = breakout_counter
    sht.range("R"+str(row_no)).value = qty
    if ltp > zone_high:
        target = float(zone_high + (zone_size * breakout_counter))
    else:
        target = float(zone_low - (zone_size * breakout_counter))

    #Trade Status
    if current_BO_dir == "High" and ltp < target:
        trade_status = "OPEN"
    if current_BO_dir == "High" and ltp > target:
        trade_status = "Target Achieved"
    if current_BO_dir == "Low" and ltp > target:
        trade_status = "OPEN"
    if current_BO_dir == "Low" and ltp < target:
        trade_status = "Target Achieved"   

    sht.range("S"+str(row_no)).value = target
    sht.range("T"+str(row_no)).value = trade_status
    place_order(row_no, sht, alice, current_BO_dir, qty)
 
# While loop for continue data:
counter = 0
while True:
    # try:
        alice.subscribe(new_subscribe_list)

        # DataFrames
        df = pd.DataFrame.from_dict(data, orient='index')
        sht.range("C1").value = df
        print(df)
        print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
        if counter ==0:
            if 'High' in df.columns and 'Low' in df.columns:
                if current_time > break_out_time:
                    df1 = df[['High', 'Low']].copy()
                    df1.rename(columns={'High': 'Zone_High', 'Low': 'Zone_Low'}, inplace=True)
                    print(df1)
                    sht.range("L1").value = df1
                    counter = 1
                else:
                    print("'High' and/or 'Low' columns are missing from the DataFrame")

        if counter >0:
            if current_time > break_out_time:
                for row_no in range (2,12):
                    # try:
                        ltp = sht.range("G"+str(row_no)).value
                        zone_high = sht.range("M"+str(row_no)).value
                        zone_low = sht.range("N"+str(row_no)).value
                        
                        zone_size = update_column_o(zone_high, zone_low, sht, row_no)
                        current_BO_dir = update_column_p_q_r_s_t(sht, row_no, alice, ltp, zone_high, zone_low, margin_available, max_breakout_until_qty_increment)

    #                 except Exception as inner_e:
    #                     pass

    # except Exception as e:
    #     print(f"An error occurred in main loop: {e}")
