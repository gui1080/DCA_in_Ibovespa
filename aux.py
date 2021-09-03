from os import setxattr
import xlrd

# takes an year, outputs It's index value on the sheet
def map_year_2_index(year):
    # 1968 = 1
    # 2021 = 54
    
    offset = year - 1968
    
    return (1+offset)

# return initial value minus final value
def compute_investment(money, performance):
    return (money - (money*performance))

# returns percentage of accumulated variation
def compute_performance(finish_year, start_year,  worksheet):
    
    start_index = map_year_2_index(start_year)
    finish_index = map_year_2_index(finish_year)
    
    performance = 1
    
    for i in range(start_index, finish_index+1):
        
        current_variation = worksheet.cell_value(i, 2)
        
        if(current_variation < 0):
            current_variation = current_variation * (-1)
            performance = performance * (1-(current_variation/100))
        else:
            performance = performance * (1+(current_variation/100))
        
    return performance

# Given time period, find best performance
def best_year_In_time_period(finish_year, start_year, worksheet):
    
    start_index = map_year_2_index(start_year)
    finish_index = map_year_2_index(finish_year)
       
    
       
    for i in range(start_index, finish_index+1):
        
        current_variation = worksheet.cell_value(i, 2)
        
        if(i == start_index):
            best_year = worksheet.cell_value(i, 0)
            best_variation = current_variation
        else:
            if(current_variation > best_variation):
                best_variation = current_variation  
                best_year = worksheet.cell_value(i, 0)  
    
    return best_year, best_variation

# Given time period, find worst performance
def worst_year_In_time_period(finish_year, start_year, worksheet):
    
    start_index = map_year_2_index(start_year)
    finish_index = map_year_2_index(finish_year)
       
    for i in range(start_index, finish_index+1):
        
        current_variation = worksheet.cell_value(i, 2)
        
        if(i == start_index):
            worst_year = worksheet.cell_value(i, 0)
            worst_variation = current_variation
        else:
            if(current_variation < worst_variation):
                worst_variation = current_variation  
                worst_year = worksheet.cell_value(i, 0)  
    
    return worst_year, worst_variation

# How long It takes to recover?
def recover_period(start_year, worksheet):
    
    start_index = map_year_2_index(start_year)
    finish_index = map_year_2_index(2021)
    
    performance = 1
    time = 0
    
    for i in range(start_index, finish_index+1):
        
        current_variation = worksheet.cell_value(i, 2)
        
        if((current_variation > 0) and (i == start_index)):
            print("Pick a bad year and try again!")
            return 0
        
        
        if(current_variation < 0):
            current_variation = current_variation * (-1)
            performance = performance * (1-(current_variation/100))
        else:
            performance = performance * (1+(current_variation/100))
            
        print(performance)
        if(performance > 1):
            print("Not at a loss anymore!")
            return time
        
        time = time + 1

    
    return -1