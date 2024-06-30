**INTRODUCTION:**

  The Opening Range Breakout (ORB) Trading System is an automated trading solution designed to capitalize on price movements that occur after the market opens. This system integrates real-time market data, algorithmic trading logic, and automated order execution to implement a sophisticated intraday trading strategy. By leveraging the AliceBlue trading platform and Excel for data management, this project aims to provide traders with a powerful tool for executing ORB strategies efficiently and consistently.

**FEATURES:**

•	Real-time market data integration via WebSocket.
•	Dynamic calculation of opening range (high and low) for multiple instruments by creating two different DataFrame in single code.
•	Automated breakout detection and trade signal generation.
•	Scaling-in strategy with increasing position sizes on subsequent breakouts.
•	Dynamic profit target setting based on breakout direction and count.
•	Automated order placement and management
•	Intraday square-off mechanism.
•	Excel-based data visualization and management.
•	Risk management through predefined investment limits and leverage controls.


**TECHNOLIGIES USED:**

•	Programming Language: Python
•	Trading API: AliceBlue (pya3 library)
•	Excel Integration: xlwings library
•	Data Handling: pandas for data manipulation.
•	Real-time Communication: WebSocket for live market data
•	Data Storage: Excel for input/output and visualization
•	Date and Time Handling: datetime module
•	Debugging: pdb module


