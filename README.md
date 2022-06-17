# excel-VBA-data-formatting

Author:  Erin James Wills, ejw.data@gmail.com

![Stock Price Changes](./images/stock-vba.png)  
<cite>Photo by [Annie Spratt](https://unsplash.com/@anniespratt?utm_source=unsplash&utm_medium=referral&utm_content=creditCopyText) on [Unsplash](https://unsplash.com/s/photos/stock-market?utm_source=unsplash&utm_medium=referral&utm_content=creditCopyText)</cite>  

## Overview
<hr>
Using daily stock data from 2014 - 2016, annual changes in stock price were calculated as well as total volume.  Data was programmatically extracted and added to a summary table identifying high and low performing stocks.  VBA was used to quickly search the records and generate the tables such that the code could be reused to generate similar reports.  

<br>

## Technologies  
*  Excel:  VBA Script  
<br>

## Data Source  
The origins are unknown.  The data may have orginally come from the [Yahoo Finance API](https://www.yahoofinanceapi.com/).  The dataset is the daily prices (high, low, etc) and trade volume. The dataset is about 75MB and consists of about 800,000 records per sheet with sheets for 2014, 2015, and 2016 data.  The daily stock data for about 9,000 companies is represented in the data.  A validated dataset should be obtained for a more serious analysis.    
**`The original data sources can be requested from the author.`**  
<br>  

## Data Manipulation
Below are three screens shots of what the final product looked like.  Columns A through G is the original data.  The code generated summarized information about each stock and color coded the changes in the stock value and summarized the volume of trades for the year.  To the right of the longer columns is a short summary table identifying extreme cases.  


![Data Formatting 2014](./images/2014_multi-year_screen_grab.png)    

![Data Formatting 2015](./images/2015_multi-year_screen_grab.png)    

![Data Formatting 2016](./images/2016_multi-year_screen_grab.png)  


## Search, Format, Summarize Code  
The code used is found in the repo as both .txt and .vb files.  The code is called `stock_data_code.xxx`.  It is well commented and an overview is provided at the top of the code.  


![Example Code](./images/code_structure.png)    

