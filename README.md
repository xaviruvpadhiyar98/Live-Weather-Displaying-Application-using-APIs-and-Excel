# Live Weather Displaying Application using APIs and Excel

I have use https://openweathermap.org/ for getting the weather Information
I have use Pandas for Excel Operations and requests for Open Weather API Handling in Python 

There 2 Excel Files
1. Containing Supported Cities Names
2. Containing the Actual Application - Live-Weather.xlsx

The reason for making the cities file separated is due to file reading and writing operations

To add new city
1. Open Live-Weather.xlsx
2. Add City name, Add Unit Value(K/C/F), and Add Update value (0 for no updation, 1 for updating values). Note: All of the values mentioned are required
3. Run the main.py


There is main.py file while loads the Live-Weather.xlsx and checks for all Rows and then updates all necessary rows

