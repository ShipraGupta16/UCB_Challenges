# import modules
import os
import csv
import sys 

# specify the file path
csvpath = os.path.join('..', 'Resources', 'election_data.csv')
output_file = os.path.join('..', 'analysis', 'output.txt')
# CSV reader specifies delimiter and variable that holds contents
# declare variables
with open(csvpath) as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')
    csv_header = next(csvreader)
    vote_count = 0 
    count = 0
    percent = 0   
    # Using dictionaries to define the candidate with vote counts and percent
    candidate_dict = {"Charles Casper Stockham" : {"count": 0, "percent": 0},
        "Diana DeGette" : {"count": 0, "percent": 0},
        "Raymon Anthony Doane" : {"count": 0, "percent": 0},
        }
    # for loop to read each row in the csv file 
    for row in csvreader:
        vote_count += 1
        candidate = (row[2])
        max_vote = 0
        # if condition to check if the candidate exists in the defined dictionary above
        # if candidate name exists, then increment the vote count for each candidate separately
        if candidate in candidate_dict:
            candidate_dict[candidate]["count"] += 1
        # calculate individual vote count and their %
        for candidate in candidate_dict:
            count = candidate_dict[candidate]["count"]
            percent =  (count/vote_count) * 100  
            candidate_dict[candidate]["percent"] = round(percent, 3)
            # find the winner candidate with maximum vote received 
            if count > max_vote:
                max_vote = count
                winner = candidate
                
# print the results in a specified format on the terminal                
print("Election Results")
print("\n")
print("--------------------------\n") 
print(f'Total Votes: {vote_count}')
print("\n--------------------------\n")
for candidate in candidate_dict:
    count = candidate_dict[candidate]["count"]
    percent = candidate_dict[candidate]["percent"]
    print(f'{candidate}: {percent}% ({count})')
    print("\n")
print("\n--------------------------\n")
print(f"Winner: {winner}")
print("\n--------------------------")
                
# print the results in a specified format in a output text file
f = open(output_file,'w')
f.write("Election Results")
f.write("\n")
f.write("--------------------------\n") 
f.write(f'Total Votes: {vote_count}')
f.write("\n--------------------------\n")
for candidate in candidate_dict:
    count = candidate_dict[candidate]["count"]
    percent = candidate_dict[candidate]["percent"]
    f.write(f'{candidate}: {percent}% ({count})')
    f.write("\n")
f.write("\n--------------------------\n")
f.write(f"Winner: {winner}")
f.write("\n--------------------------")