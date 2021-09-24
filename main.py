import xlrd
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt


from aux import compute_performance
from aux import compute_investment
from aux import map_year_2_index
from aux import best_year_In_time_period
from aux import worst_year_In_time_period
from aux import recover_period

# Open the Workbook
workbook = xlrd.open_workbook("IBOVESPA.xlsx")

# Open the worksheet
worksheet = workbook.sheet_by_index(0)

total_years = []
index = []

# Iterate the rows and columns
for i in range(1, 55):
    
    total_years.append(worksheet.cell_value(i, 0))
    index.append(worksheet.cell_value(i, 1))
    
    #for j in range(0, 3):
        # Print the cell values with tab space
    #    print(worksheet.cell_value(i, j), end='\t')
    #print('')
    

fig, ax = plt.subplots()
ax.plot(total_years, index, 'k')
ax.set(facecolor = "lightgrey")

ax.ticklabel_format(style='plain')

plt.suptitle('Ibovespa') 
plt.grid()
plt.xlabel("Years")
plt.ylabel("Index Value")
plt.show()

# ---------------------------------------------------------------------------
# Test: print certain variation
#print(worksheet.cell_value(map_year_2_index(2021) , 2))

# ---------------------------------------------------------------------------
# Test: compute investment
performance = compute_performance(2021, 1968, worksheet)
#print("performance")
#print(performance)
performance = 0

# ---------------------------------------------------------------------------
# Test: get best year for a investment
best_year, best_variation = best_year_In_time_period(2021, 1968, worksheet)
#print(best_year)
#print(best_variation)

# ---------------------------------------------------------------------------
# Test: get worst year for a investment
worst_year, worst_variation = worst_year_In_time_period(2021, 1968, worksheet)
#print(worst_year)
#print(worst_variation)

# ---------------------------------------------------------------------------
# Test: recover period
recover = recover_period(1995, worksheet)
#print(recover)

# ---------------------------------------------------------------------------
# Test: Buy the dip every 5 years ($5000 at once) vs DCA ($1000 every year)
# $50k in total
base_year = 1968
limit_year = base_year + 5
gain = 0
money_you_end_up_with = 0

gain_over_years = []
years = []

while(limit_year < 2021):
    
    worst_year, worst_variation = worst_year_In_time_period(limit_year, base_year, worksheet)
    performance = compute_performance(2021, int(worst_year), worksheet)
    gain = compute_investment(5000, performance)
    money_you_end_up_with = money_you_end_up_with + gain
    
    years.append(base_year)
    gain_over_years.append(gain)
  
    
    base_year = base_year + 5 
    limit_year = base_year + 5
    

fig, ax = plt.subplots()
ax.plot(years, gain_over_years, marker = 'o', color = 'red')
ax.set(facecolor = "lightgrey")
ax.ticklabel_format(style='plain')
 
plt.grid()
plt.xlabel("Return of $5000 from 'x' year to 2021 (Buying the Dip every 5 years)")
plt.ylabel("Ammount of Money")
plt.show()

print("Final ammount of money buying the dip:")
print(money_you_end_up_with)


gain = 0
base_year = 1968

gain_over_years = []
years = []

while(base_year < 2018):
    
    performance = compute_performance(2021, int(base_year), worksheet)
    gain = compute_investment(1000, performance)
    
    years.append(base_year)
    gain_over_years.append(gain)
    money_you_end_up_with = money_you_end_up_with + gain
  
    base_year = base_year + 1     
    

fig, ax = plt.subplots()
ax.plot(years, gain_over_years, marker = 'o', color = 'darkgreen')
ax.set(facecolor = "lightgrey")
ax.ticklabel_format(style='plain')
 
plt.grid()

plt.xlabel("Return of $1000 from 'x' year to 2021 (Buying every year)")
plt.ylabel("Ammount of Money")
plt.show()

print("Final ammount of money with DCA")
print(money_you_end_up_with)

# results from this graphic show why often "long term investment" is something that may take from 10 to 20 years!


# ---------------------------------------------------------------------------
# Test: historical recover period
# return 0 = good year
# return = -1 never recovered
# return x = "x" years it took to recover

never_recovered = 0
good_year = 0
eventually_recovered = 0

investment_year = []
time_it_took = []

base_year = 1968
while(base_year < 2021):
    recover = recover_period(base_year, worksheet)
    
    if(recover == -1):
        never_recovered = never_recovered + 1
    else:  
        if(recover == 0):
            good_year = good_year + 1
        else:
            eventually_recovered = eventually_recovered + 1
            
            investment_year.append(base_year)
            time_it_took.append(recover)
    
    base_year = base_year + 1


year_data = [never_recovered, good_year, eventually_recovered]
names = ['Lost', 'Good Right Away', 'Eventually Recovered']

fig, ax = plt.subplots()
plt.suptitle('Investments by Year')
colors = sns.color_palette('pastel')
ax.bar(names, year_data, color=colors[:3])
ax.set(facecolor = "lightgrey")
plt.show()

fig, ax = plt.subplots()
plt.suptitle('Time to recover')
colors = sns.color_palette('Set2')
ax.scatter(investment_year, time_it_took, color=colors[3])
ax.set(facecolor = "lightgrey")
plt.xlabel("Year")
plt.ylabel("Time to make some Profit")
plt.grid()
plt.show()

# ---------------------------------------------------------------------------
