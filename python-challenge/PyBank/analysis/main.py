# import required modules
import os
import csv
import sys

# specify the file path
csvpath = os.path.join('..', 'Resources', 'budget_data.csv')
output_file = os.path.join('..', 'analysis', 'analysis_output.txt')

with open(csvpath) as csvfile:

    # CSV reader specifies delimiter and variable that holds contents
    # next skips the header
    # declare or define variables
    csvreader = csv.reader(csvfile, delimiter=',')
    csv_header = next(csvreader)
    print(f"CSV Header: {csv_header}")
    month_count = 0
    net_amount = 0
    change = 0
    max_amount = 0
    min_amount = 0
    min_date = ''
    max_date = ''
    previous_amount = 0
    current_amount = 0
    total = 0
    average_amount = 0
    
    # for loop to read each row of the csv file and perform calculations 
    for row in csvreader:
        profit_loss = int(row[1])
        month_count = month_count + 1
        net_amount = net_amount + profit_loss
        # skip the first row value to subtract from second value
        if(previous_amount == 0):
            previous_amount = int(row[1])
            continue
        # calculate change and total amount 
        current_amount = int(row[1])
        change = current_amount - previous_amount
        previous_amount = current_amount
        total = total + change
       
        # Great increase and decrease in profits calculation
        if max_amount < change:
            max_amount = change
            max_date = row[0]
        if min_amount > change:
            min_amount = change
            min_date = row[0]
        
    # Average change calculations
    average_amount = round(float(total/(month_count-1)),2)
    print("Financial Analysis \n")
    print("--------------------------\n")
    print(f'\nTotal Months: {month_count}')
    print(f'\nTotal: ${net_amount}')
    print(f'\nAverage Change: ${average_amount}')
    print(f'\nGreatest Increase in Profits: {max_date} ${max_amount}')
    print(f'\nGreatest Decrease in Profits: {min_date} ${min_amount}')

    # print the analysis results in a specified format in a output text file   
    f = open(output_file, 'w')    
    f.write("Financial Analysis \n")
    f.write("--------------------------\n")
    f.write(f'\nTotal Months: {month_count}')
    f.write(f'\nTotal: ${net_amount}')
    f.write(f'\nAverage Change: ${average_amount}')
    f.write(f'\nGreatest Increase in Profits: {max_date} ${max_amount}')
    f.write(f'\nGreatest Decrease in Profits: {min_date} ${min_amount}')
    

            
    
    
    