import os
import csv

# Load CSV Files
load_file = os.path.join("Resources", "budget_data.csv")
output_file = os.path.join("Resources", "budget_analysis.txt")

# Create/Initialize Variables
total_months = 0
net_total = 0

previous_revenue = 0
revenue_change = 0

# Initialize Array to hold str and int
greatest_increase = ["", 0]
greatest_decrease = ["", 9999]

monthly_changes = []
date =[]

# Read CSV File
with open(load_file) as csv_file:
	csv_reader = csv.DictReader(csv_file)

	# Loop through all the rows in the data
	for row in csv_reader:

		# Calculate total number of months in dataset
		# Calculate net total of Profit/Losses 
		total_months += 1
		net_total = net_total + int(row["Profit/Losses"])

		# Calculate change in profits from month to month
		revenue_change = int(row["Profit/Losses"]) - previous_revenue
		previous_revenue = int(row["Profit/Losses"])

		# Store changes in profits in array
		monthly_changes.append(revenue_change)

		# Calculate greatest increase and decrease
		if (revenue_change > greatest_increase[1]):
			greatest_increase[1] = revenue_change
			greatest_increase[0] = row["Date"]

		if (revenue_change < greatest_decrease[1]):
			greatest_decrease[1] = revenue_change
			greatest_decrease[0] = row["Date"]

	monthly_avg = sum(monthly_changes) / total_months

 # Show Output
print("\n")
print("Financial Analysis")
print("-------------------------")
print("Total Months: " + str(total_months))
print("Total: " + "$" + str(net_total))
print("Average Change: " + "$" + str(round(monthly_avg,2)))
print("Greatest Increase in Profits: " + str(greatest_increase[0]) + " ($" +  str(greatest_increase[1]) + ")") 
print("Greatest Decrease in Profits: " + str(greatest_decrease[0]) + " ($" +  str(greatest_decrease[1]) + ")")

# Export results to text file
with open(output_file, "w") as txt_file:
	txt_file.write("\n")
	txt_file.write("Financial Analysis\n")
	txt_file.write("-------------------------\n")
	txt_file.write("Total Months: " + str(total_months) + "\n")
	txt_file.write("Total: " + "$" + str(net_total) + "\n")
	txt_file.write("Average Change: " + "$" + str(round(monthly_avg,2)) + "\n")
	txt_file.write("Greatest Increase in Profits: " + str(greatest_increase[0]) + " ($" +  str(greatest_increase[1]) + ") \n") 
	txt_file.write("Greatest Decrease in Profits: " + str(greatest_decrease[0]) + " ($" +  str(greatest_decrease[1]) + ") \n")




