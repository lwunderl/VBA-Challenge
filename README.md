# VBA Stock Iterator
VBA_stock_iterator.bas is a macro which iterates over an excel file with multiple sheets of the same data format<br/>
<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>
![image](https://user-images.githubusercontent.com/116906733/218816660-bae0d724-1c61-4105-91e1-048b0fed8ab9.png)<br>
Open stock ticker file.xlsm and ensure each worksheet is a single year and sorted by ticker symbol and date<br>
Import VBA_stock_iterator.bas<br>
Select the main() macro to run in the stock ticker file.xlsm<br>
The main() macro will cycle through each worksheet and return year opening and closing price, percent change, and total volume from a list of stocks<br>
