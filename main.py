import xlrd

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

# Iterate the rows and columns
for i in range(1, 55):
    for j in range(0, 3):
        # Print the cell values with tab space
        print(worksheet.cell_value(i, j), end='\t')
    print('')
    

# Test: print certain variation
print(worksheet.cell_value(map_year_2_index(2021) , 2))

# Test: compute investment
performance = compute_performance(2021, 1968, worksheet)

# Test: get best year for a investment
best_year, best_variation = best_year_In_time_period(2021, 1968, worksheet)
print(best_year)
print(best_variation)

# Test: get worst year for a investment
worst_year, worst_variation = worst_year_In_time_period(2021, 1968, worksheet)
print(worst_year)
print(worst_variation)

# Test: recover period
recover = recover_period(1995, worksheet)
print(recover)

