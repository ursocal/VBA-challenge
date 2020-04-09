# VBA-challenge

If you look at the PLNT values between years, it illustrates three different cases that made the coding more diffcult in terms of grabbing the openingprice value and calculating the pricechange and percentchange

case 1:  The openingprice occurs at the first date of a date array for a given ticker
case 2:  Before an openingprice occurs, each date has an openingprice of 0. These are clearly just dates where no data was recorded. The first openingprice occurs at a date in the middle of the date array for a given ticker.
case 3: The openingprice never occurs, and the entire array of dates has values of 0

The code in script.vbs takes care of all three of these cases by, for each ticker, waiting to grab the openingprice and only grabbing the openingprice if its has not been grabbed yet and if the grabbed value was not 0.
