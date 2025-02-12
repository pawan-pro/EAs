from ctrader_open_api import Client, Protobuf, TcpProtocol, Auth, EndPoints
from ctrader_open_api.messages.OpenApiCommonMessages_pb2 import *
from ctrader_open_api.messages.OpenApiMessages_pb2 import *
from ctrader_open_api.messages.OpenApiModelMessages_pb2 import *
from twisted.internet import reactor, defer
from twisted.internet.error import ConnectionDone, TimeoutError
import logging
import sys
import webbrowser
import time

# Configure logging to both file and console
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
        self.symbol_id = None  # Will store selected symbol ID
        self.order_placed = False  # Track if an order has been placed
        self.current_market_price = None  # Will store current market price
        self.order_executed = False  # Add flag to track successful execution
        self.symbols = {}  # Dictionary to store symbol IDs and names
        
        # Embedded credentials
        self.host_type = "demo"  # Fixed to demo environment
        self.app_client_id = "13127_QDPscTztUgs175Mge2wtfTQq7sTKRNBtHBub6glfFDEz36WdLE"
        self.app_client_secret = "R40zAVlxfvU3A6oFo6wlK3cgukkdlke35t4zpzThbdW86eS2np"
        self.access_token = "56cYMAZbn4rQGqGKJCjauI9BQjH5mKIieey2qvKwzFM"
        
        # Initialize client with demo host
        host = EndPoints.PROTOBUF_DEMO_HOST
        self.client = Client(
            host,
            EndPoints.PROTOBUF_PORT,
            TcpProtocol
        )
        
        # Set up callbacks
        self.client.setConnectedCallback(self.on_connected)
        self.client.setDisconnectedCallback(self.on_disconnected)
        self.client.setMessageReceivedCallback(self.on_message_received)

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

    def on_message_received(self, client, message):
        """Callback for all received messages"""
        msg = Protobuf.extract(message)
        
        print(f"\nReceived message type: {message.payloadType}")
        logging.info(f"Received message type: {message.payloadType}")
        
        # Modify error handling to check error message format
        if hasattr(msg, 'errorCode'):
            error_msg = f"Error {msg.errorCode}"
            if hasattr(msg, 'description'):
                error_msg += f" - {msg.description}"
            print(f"\nError received: {error_msg}")
            logging.error(f"Error received: {error_msg}")
            if msg.errorCode == "CH_CLIENT_AUTH_FAILURE":
                print("Authentication failed. Please check your credentials.")
                reactor.stop()
                return

        # Handle execution event with more details
        if message.payloadType == ProtoOAExecutionEvent().payloadType:
            try:
                execution_event = Protobuf.extract(message)
                if hasattr(execution_event, 'order'):
                    order = execution_event.order
                    # Access order properties safely
                    volume = getattr(order, 'volume', 0) / 1000000
                    price = getattr(order, 'limitPrice', 0) / 100000
                    status = getattr(order, 'orderStatus', 'UNKNOWN')
                    
                    print(f"\nOrder execution details:")
                    print(f"Order ID: {getattr(order, 'orderId', 'N/A')}")
                    print(f"Volume: {volume:.2f} lots") 
                    print(f"Price: {price:.5f}")
                    print(f"Status: {status}")
                    
                    self.order_executed = True
                    reactor.callLater(2, reactor.stop)
            except Exception as e:
                print(f"\nError processing execution event: {str(e)}")
        
        # Handle successful responses
        if message.payloadType == ProtoOAGetAccountListByAccessTokenRes().payloadType:
            if len(msg.ctidTraderAccount) > 0:
                self.account_id = msg.ctidTraderAccount[0].ctidTraderAccountId
                print(f"\nSelected account ID: {self.account_id}")
                self.send_account_auth_req()
        elif message.payloadType == ProtoOAAccountAuthRes().payloadType:
            self.authenticated = True
            print("\nAccount authentication successful")
            logging.info("Account authentication successful")
            self.get_symbols_list()
        elif message.payloadType == ProtoOASymbolsListRes().payloadType:
            # Store all symbols in a dictionary
            for symbol in msg.symbol:
                self.symbols[symbol.symbolName] = symbol.symbolId
                print(f"Found symbol: {symbol.symbolName} (ID: {symbol.symbolId})")
            
            # Prompt user to select a symbol
            self.select_symbol()
        elif message.payloadType == ProtoOASpotEvent().payloadType and not self.order_placed:
            self.current_market_price = msg.bid / 100000
            print(f"\nSpot price received - Bid: {msg.bid/100000:.5f}, Ask: {msg.ask/100000:.5f}")
            print(f"Current market price: {self.current_market_price:.5f}")
            self.get_order_input()
            # Unsubscribe from spots after getting initial price
            self.unsubscribe_spots()
            return

    def select_symbol(self):
        """Prompt user to select a symbol from the available list"""
        print("\nAvailable symbols:")
        for symbol_name in self.symbols.keys():
            print(f"- {symbol_name}")
        
        selected_symbol = input("\nEnter the symbol you want to trade (e.g., EURUSD): ").strip().upper()
        
        if selected_symbol in self.symbols:
            self.symbol_id = self.symbols[selected_symbol]
            print(f"\nSelected symbol: {selected_symbol} (ID: {self.symbol_id})")
            self.get_market_price()
        else:
            print("\nInvalid symbol selected. Please try again.")
            reactor.stop()

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

    def get_market_price(self):
        """Get current market price for the selected symbol"""
        print(f"\nGetting current market price for symbol ID: {self.symbol_id}...")
        
        spot_req = ProtoOASubscribeSpotsReq()
        spot_req.ctidTraderAccountId = self.account_id
        spot_req.symbolId.append(self.symbol_id)
        
        deferred = self.client.send(spot_req)
        deferred.addCallbacks(lambda response: None, self.on_error)
        deferred.addTimeout(10, reactor)

    def get_order_input(self):
        """Get price input from user and determine order type"""
        print("\n=== Order Entry ===")
        print(f"Current market price: {self.current_market_price:.5f}")
        print("Please enter your target price (close to current price, e.g. 1.05000):")
        
        try:
            target_price = float(input("> "))
            
            # Validate price is reasonably close to market price
            if abs(target_price - self.current_market_price) < 0.0010:  # Within 10 pips
                print("\nPrice too close to market. Please choose a price further from current market price.")
                reactor.stop()
                return
            
            if target_price > self.current_market_price:
                print("\nExecuting BUY STOP order...")
                self.place_trade(ProtoOAOrderType.STOP, target_price, ProtoOATradeSide.BUY)
            else:
                print("\nExecuting BUY LIMIT order...")
                self.place_trade(ProtoOAOrderType.LIMIT, target_price, ProtoOATradeSide.BUY)
                
        except ValueError:
            print("\nInvalid price input. Please enter a valid number.")
            reactor.stop()

    def place_trade(self, order_type, price, trade_side):
        """Place trade with specified order type, price and side"""
        if self.order_placed:
            return
        
        # Convert price correctly based on order type
        price_int = int(price*100000)  # Use 5 decimal places for forex
        
        order_req = ProtoOANewOrderReq()
        order_req.ctidTraderAccountId = self.account_id
        order_req.symbolId = self.symbol_id
        order_req.orderType = order_type
        order_req.tradeSide = trade_side
        order_req.volume = 1000000  # 0.1 lots to match screenshot
        
        if price_int/100000 > self.current_market_price:
            order_req.orderType = ProtoOAOrderType.STOP
            order_req.stopPrice = price_int/100000
        else:
            order_req.orderType = ProtoOAOrderType.LIMIT
            order_req.limitPrice = price_int/100000

        # Set tighter SL/TP
        order_req.relativeStopLoss = 250  # 25 pips
        order_req.relativeTakeProfit = 500  # 50 pips
        order_req.comment = "Auto Trade"
        
        deferred = self.client.send(order_req)
        deferred.addCallbacks(self.on_order_response, self.on_error)
        deferred.addTimeout(30, reactor)
        self.order_placed = True

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
        """Check if we're still trying to connect after timeout period"""
        if not self.authenticated:
            print("\nConnection/authentication timed out after 30 seconds. Stopping...")
            logging.error("Connection/authentication timeout")
            reactor.stop()
        elif self.authenticated and not self.order_placed:
            if self.current_market_price is None:
                print("\nTimed out waiting for market price. Stopping...")
                reactor.stop()
            else:
                # Extend timeout if authenticated but waiting for order input
                reactor.callLater(30, self.check_connection_timeout)

def main():
    try:
        print("\nInitializing main execution...")
        executor = TradingExecutor()
        executor.start()
        
        # Start the Twisted reactor
        print("\nStarting Twisted reactor...")
        reactor.run()
        
    except Exception as e:
        print(f"\nMain execution error: {str(e)}")
        logging.error(f"Main execution error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()