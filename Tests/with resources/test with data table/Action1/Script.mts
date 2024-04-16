WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set "john" @@ hightlight id_;_-22061840_;_script infofile_;_ZIP::ssf39.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfEdit("password").SetSecure "6245952da4a22d9dd351" @@ hightlight id_;_-22079264_;_script infofile_;_ZIP::ssf42.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click
WpfWindow("OpenText MyFlight Sample").WpfComboBox("fromCity").Select Parameter("From") @@ hightlight id_;_2036014008_;_script infofile_;_ZIP::ssf47.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfComboBox("toCity").Select Parameter("To") @@ hightlight id_;_-204644568_;_script infofile_;_ZIP::ssf50.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfImage("WpfImage").Click 5,10 @@ hightlight id_;_-22056368_;_script infofile_;_ZIP::ssf53.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfCalendar("Mo").SetDate Parameter("FlightDate") @@ hightlight id_;_2038150112_;_script infofile_;_ZIP::ssf56.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfComboBox("Class").Select Parameter("Class") @@ hightlight id_;_2038149680_;_script infofile_;_ZIP::ssf59.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfComboBox("numOfTickets").Select Parameter("TicketsNumber") @@ hightlight id_;_-264367600_;_script infofile_;_ZIP::ssf62.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf65.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,1 @@ hightlight id_;_-29773832_;_script infofile_;_ZIP::ssf68.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf71.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").Set Parameter("PassengerName") @@ hightlight id_;_-29767784_;_script infofile_;_ZIP::ssf74.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2038154624_;_script infofile_;_ZIP::ssf77.xml_;_

Dim OrderNumber
OrderNumber = WpfWindow("OpenText MyFlight Sample").WpfObject("Order completed").GetROProperty("text")
DataTable("OrderNumber", dtGlobalSheet) = OrderNumber


'WpfWindow("OpenText MyFlight Sample").WpfButton("NEW SEARCH").Click @@ hightlight id_;_-232171912_;_script infofile_;_ZIP::ssf92.xml_;_
'WpfWindow("OpenText MyFlight Sample").WpfTabStrip("WpfTabStrip").Select "SEARCH ORDER" @@ hightlight id_;_-18470624_;_script infofile_;_ZIP::ssf93.xml_;_
'WpfWindow("OpenText MyFlight Sample").WpfRadioButton("byNumberRadio").Set @@ hightlight id_;_2070102336_;_script infofile_;_ZIP::ssf94.xml_;_
'WpfWindow("OpenText MyFlight Sample").WpfEdit("byNumberWatermark").Set Parameter("OrderNumber") @@ hightlight id_;_-18469568_;_script infofile_;_ZIP::ssf95.xml_;_
'WpfWindow("OpenText MyFlight Sample").WpfButton("SEARCH").Click @@ hightlight id_;_-18468032_;_script infofile_;_ZIP::ssf96.xml_;_
'
'Dim orderNumber
'Set orderNumber = WpfWindow("OpenText MyFlight Sample").WpfObject("Order .* completed")
'orderNumber.Output CheckPoint(orderNumber.GetROProperty("text"))

'WpfWindow("OpenText MyFlight Sample").WpfObject("Order completed").Output CheckPoint("Order 174 completed") @@ hightlight id_;_-3287616_;_script infofile_;_ZIP::ssf91.xml_;_

'WpfWindow("OpenText MyFlight Sample").WpfObject("Order completed").WaitProperty "enabled", true, 100000 @@ hightlight id_;_-1020184_;_script infofile_;_ZIP::ssf26.xml_;_
'WpfWindow("OpenText MyFlight Sample").WpfObject("Order completed").Output CheckPoint("Order completed") @@ hightlight id_;_-902216_;_script infofile_;_ZIP::ssf24.xml_;_
WpfWindow("OpenText MyFlight Sample").Close

