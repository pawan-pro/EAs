from ctrader_open_api import Client, Protobuf, TcpProtocol, Auth, EndPoints
from ctrader_open_api.messages.OpenApiCommonMessages_pb2 import *
from ctrader_open_api.messages.OpenApiMessages_pb2 import *
from ctrader_open_api.messages.OpenApiModelMessages_pb2 import *
from twisted.internet import reactor, defer
from twisted.internet.error import ConnectionDone, TimeoutError
import logging
import sys
import math

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ea.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

class TradingExecutor:
    def __init__(self):
        self.authenticated = False
        self.connection_attempts = 0
        self.max_connection_attempts = 3
        self.symbol_id = None
        self.order_placed = False
        self.current_market_price = None
        self.order_executed = False
        self.symbols = {}
        self.account_info = None
        self.symbol_details = {}
        
        # Embedded credentials
        self.host_type = "demo"
        self.app_client_id = "13127_QDPscTztUgs175Mge2wtfTQq7sTKRNBtHBub6glfFDEz36WdLE"
        self.app_client_secret = "R40zAVlxfvU3A6oFo6wlK3cgukkdlke35t4zpzThbdW86eS2np"
        self.access_token = "56cYMAZbn4rQGqGKJCjauI9BQjH5mKIieey2qvKwzFM"
        
        self.client = Client(
            EndPoints.PROTOBUF_DEMO_HOST,
            EndPoints.PROTOBUF_PORT,
            TcpProtocol
        )
        
        self.client.setConnectedCallback(self.on_connected)
        self.client.setDisconnectedCallback(self.on_disconnected)
        self.client.setMessageReceivedCallback(self.on_message_received)

    def normalize_pips(self, pips):
        """Normalize stop loss/take profit to symbol's precision using pipSize and digits"""
        pip_size = self.symbol_details['pipSize']
        digits = self.symbol_details['digits']
        
        # Calculate normalized pips based on symbol's pip size and precision
        normalized = round(pips / pip_size) * pip_size
        return int(normalized * (10 ** digits))  # Convert to 1/100000 units

    def calculate_position_size(self, risk_percent, stop_loss_pips):
        """Calculate position size based on risk parameters and symbol details"""
        account_balance = self.account_info.balance
        risk_amount = account_balance * (risk_percent / 100)
        pip_value = self.symbol_details['pipValue'] * 100000  # Convert to account currency
        
        if pip_value <= 0:
            return 0
            
        position_size = (risk_amount / (stop_loss_pips * pip_value)) * 1000000
        return int(max(position_size, self.symbol_details['minLot']))

    def on_message_received(self, client, message):
        msg = Protobuf.extract(message)
        
        if message.payloadType == ProtoOAGetAccountListByAccessTokenRes().payloadType:
            if len(msg.ctidTraderAccount) > 0:
                self.account_id = msg.ctidTraderAccount[0].ctidTraderAccountId
                self.send_account_auth_req()
        
        elif message.payloadType == ProtoOAAccountAuthRes().payloadType:
            self.authenticated = True
            print("\nAccount authentication successful")
            logging.info("Account authentication successful")
            self.get_symbols_list()  # Request symbols list after authentication
        
        elif message.payloadType == ProtoOASymbolsListRes().payloadType:
            # Store all available symbols and prompt for selection
            for symbol in msg.symbol:
                self.symbols[symbol.symbolName] = {
                    'id': symbol.symbolId,
                    'digits': symbol.digits,
                    'pipSize': symbol.pipSize,
                    'pipValue': symbol.pipValue,
                    'minLot': symbol.minLot
                }
            print("\nAvailable symbols:")
            for symbol_name in sorted(self.symbols.keys()):
                print(f"- {symbol_name}")
            self.select_symbol()
        
        elif message.payloadType == ProtoOASpotEvent().payloadType and not self.order_placed:
            self.current_market_price = msg.bid / 100000
            self.get_risk_parameters()

    def get_risk_parameters(self):
        """Get risk management inputs from user"""
        print("\n=== Risk Management ===")
        
        # Get risk percentage
        while True:
            try:
                risk_percent = float(input("Enter risk percentage (0.1-5%): "))
                if 0.1 <= risk_percent <= 5:
                    break
                print("Risk must be between 0.1 and 5%")
            except ValueError:
                print("Invalid input")
        
        # Get risk-reward ratio
        while True:
            try:
                rr_ratio = float(input("Enter risk-reward ratio (min 1): "))
                if rr_ratio >= 1:
                    break
                print("Ratio must be >= 1")
            except ValueError:
                print("Invalid input")
        
        self.risk_percent = risk_percent
        self.rr_ratio = rr_ratio
        self.get_order_input()

    def get_order_input(self):
        """Get target price and calculate stop loss"""
        print(f"\nCurrent price: {self.current_market_price:.5f}")
        
        while True:
            try:
                target_price = float(input("Enter target price: "))
                price_diff = abs(target_price - self.current_market_price)
                min_distance = self.symbol_details['pipSize'] * 10
                
                if price_diff < min_distance:
                    print(f"Target must be at least {min_distance:.5f} away")
                else:
                    break
            except ValueError:
                print("Invalid price")
        
        stop_loss_pips = abs(target_price - self.current_market_price) / self.symbol_details['pipSize']
        normalized_sl = self.normalize_pips(stop_loss_pips)
        normalized_tp = self.normalize_pips(stop_loss_pips * self.rr_ratio)
        
        position_size = self.calculate_position_size(self.risk_percent, stop_loss_pips)
        
        print(f"\nTrade Parameters:")
        print(f"Stop Loss: {normalized_sl/100000:.1f} pips")
        print(f"Take Profit: {normalized_tp/100000:.1f} pips")
        print(f"Position Size: {position_size/1000000:.2f} lots")
        
        self.place_trade(
            target_price, 
            normalized_sl,
            normalized_tp,
            position_size
        )

    def place_trade(self, price, sl_pips, tp_pips, volume):
        """Place trade with normalized parameters"""
        order_req = ProtoOANewOrderReq()
        order_req.ctidTraderAccountId = self.account_id
        order_req.symbolId = self.symbol_id
        order_req.tradeSide = ProtoOATradeSide.BUY
        order_req.volume = volume
        
        if price > self.current_market_price:
            order_req.orderType = ProtoOAOrderType.STOP
            order_req.stopPrice = int(price * 100000)
        else:
            order_req.orderType = ProtoOAOrderType.LIMIT
            order_req.limitPrice = int(price * 100000)
        
        order_req.relativeStopLoss = sl_pips
        order_req.relativeTakeProfit = tp_pips
        order_req.comment = f"Risk:{self.risk_percent}% RR:{self.rr_ratio}"
        
        self.client.send(order_req)
        self.order_placed = True

    def select_symbol(self):
        """Prompt user to select a symbol from the available list"""
        print("\nAvailable symbols:")
        for symbol_name in self.symbols.keys():
            print(f"- {symbol_name}")
        
        selected_symbol = input("\nEnter the symbol you want to trade: ").strip().upper()
        
        if selected_symbol in self.symbols:
            self.symbol_id = self.symbols[selected_symbol]['id']
            self.symbol_details = self.symbols[selected_symbol]
            print(f"\nSelected symbol: {selected_symbol} (ID: {self.symbol_id})")
            print(f"Symbol details: PipSize={self.symbol_details['pipSize']}, Digits={self.symbol_details['digits']}")
            self.get_market_price()
        else:
            print("\nInvalid symbol selected.")
            reactor.stop()

    def get_market_price(self):
        """Get current market price for the selected symbol"""
        if self.symbol_id is None:
            print("\nSymbol not selected yet; cannot fetch market price.")
            return  # Prevent error if symbol is not selected
        print(f"\nGetting current market price for symbol ID: {self.symbol_id}...")
        spot_req = ProtoOASubscribeSpotsReq()
        spot_req.ctidTraderAccountId = self.account_id
        spot_req.symbolId.append(self.symbol_id)
        
        deferred = self.client.send(spot_req)
        deferred.addCallbacks(lambda response: None, self.on_error)
        # Increase timeout to 10 seconds
        deferred.addTimeout(10, reactor)

    def on_connected(self, client):
        """Callback for when client connects"""
        print("\nConnected to cTrader")
        logging.info("Connected to cTrader")
        
        # Send authentication request
        print("Sending authentication request...")
        auth_req = ProtoOAApplicationAuthReq()
        auth_req.clientId = self.app_client_id
        auth_req.clientSecret = self.app_client_secret
        
        deferred = client.send(auth_req)
        deferred.addCallbacks(self.on_auth_response, self.on_error)
        deferred.addTimeout(10, reactor)  # Add 10 second timeout
    
    def on_disconnected(self, client, reason):
        """Callback for when client disconnects"""
        print(f"\nDisconnected from cTrader: {reason}")
        logging.info(f"Disconnected from cTrader: {reason}")
        
        # Only attempt reconnection if we haven't executed an order
        if not self.order_executed and not self.authenticated and self.connection_attempts < self.max_connection_attempts:
            self.connection_attempts += 1
            print(f"\nAttempting reconnection ({self.connection_attempts}/{self.max_connection_attempts})...")
            reactor.callLater(2, self.start)
        elif not self.order_executed and not self.authenticated:
            print("\nMax reconnection attempts reached. Stopping...")
            reactor.stop()
        # If order was executed, treat disconnect as normal completion
        elif self.order_executed:
            print("\nSession completed successfully.")
        
    def unsubscribe_spots(self):
        """Unsubscribe from spot price updates"""
        unsub_req = ProtoOAUnsubscribeSpotsReq()
        unsub_req.ctidTraderAccountId = self.account_id
        unsub_req.symbolId.append(self.symbol_id)
        self.client.send(unsub_req)
    
    def get_symbols_list(self):
        """Get list of available symbols"""
        print("\nRequesting symbols list...")
        symbols_req = ProtoOASymbolsListReq()
        symbols_req.ctidTraderAccountId = self.account_id
        deferred = self.client.send(symbols_req)
        deferred.addCallbacks(lambda response: None, self.on_error)
        deferred.addTimeout(10, reactor)
    
    def on_auth_response(self, response):
        """Handle authentication response"""
        if not hasattr(response, 'errorCode'):
            print("\nApplication authentication successful")
            logging.info("Application authentication successful")
            self.get_account_list()
        else:
            print(f"\nAuthentication failed: {response.errorCode}")
            logging.error(f"Authentication failed: {response.errorCode}")
            reactor.stop()
        
    def get_account_list(self):
        """Get list of available accounts"""
        print("\nRequesting account list...")
        req = ProtoOAGetAccountListByAccessTokenReq()
        req.accessToken = self.access_token
        deferred = self.client.send(req)
        deferred.addCallbacks(self.on_account_list_response, self.on_error)
        deferred.addTimeout(10, reactor)

    def on_account_list_response(self, response):
        """Handle account list response"""
        print("\nAccount list received")
    
    def send_account_auth_req(self):
        """Send account authentication request"""
        print("\nSending account authentication request...")
        account_auth_req = ProtoOAAccountAuthReq()
        account_auth_req.ctidTraderAccountId = self.account_id
        account_auth_req.accessToken = self.access_token
        deferred = self.client.send(account_auth_req)
        deferred.addCallbacks(self.on_account_auth_response, self.on_error)
        deferred.addTimeout(10, reactor)
    
    def on_account_auth_response(self, response):
        """Handle account authentication response"""
        if not hasattr(response, 'errorCode'):
            self.authenticated = True
            print("\nAccount authentication successful")
            logging.info("Account authentication successful")
        else:
            print(f"\nAccount authentication failed: {response.errorCode}")
            logging.error(f"Account authentication failed: {response.errorCode}")
            reactor.stop()
        
    def on_order_response(self, response):
        """Handle order response"""
        print(f"\nOrder response received: {response}")
        logging.info(f"Order response received: {response}")
        # Wait briefly for execution event then stop
        reactor.callLater(3, reactor.stop)
        return response
    
    def on_error(self, failure):
        """Handle errors"""
        print(f"\nError occurred: {failure}")
        logging.error(f"Error: {failure}")
        
        if isinstance(failure.value, TimeoutError):
            print("\nOrder may have been executed despite timeout. Please check your positions.")
            # Don't stop reactor immediately on timeout
            reactor.callLater(5, reactor.stop)
        elif isinstance(failure.value, ConnectionDone):
            print("Connection closed. Attempting reconnect...")
            self.on_disconnected(self.client, failure.value)
        
    def start(self):
        """Start the trading executor"""
        print("\nStarting trading executor...")
        logging.info("Starting trading executor...")
        self.client.startService()
        
        # Add a timeout to prevent infinite connection attempts
        reactor.callLater(30, self.check_connection_timeout)

    def check_connection_timeout(self):
        """If still waiting for market price, reattempt instead of stopping immediately."""
        if not self.authenticated:
            print("\nConnection/authentication timed out after 30 seconds. Stopping...")
            reactor.stop()
        elif self.authenticated and not self.order_placed:
            if self.symbol_id is None:
                print("\nNo symbol selected yet. Prompting symbol selection...")
                self.select_symbol()
                reactor.callLater(15, self.check_connection_timeout)
            elif self.current_market_price is None:
                print("\nMarket price not received yet. Retrying market price request...")
                self.get_market_price()
                reactor.callLater(15, self.check_connection_timeout)
            else:
                reactor.callLater(30, self.check_connection_timeout)

def main():
    try:
        executor = TradingExecutor()
        executor.start()
        reactor.run()
    except Exception as e:
        print(f"Main error: {str(e)}")
        logging.error(f"Main error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
