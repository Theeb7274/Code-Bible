# A scenario where the task required basic math operations in bash.
# Due to multiple commandline arguemnts and the requirement that variables persist beyond the line ran, I created a script to run.
touch script.sh

# This won't be able to run unless we provide executable permissions
chmod +x script.sh

# Now to edit the script
nano script.sh

# The goal of this script is to find the average value of the first column in the file "scores.txt"
# Firstly, we need to extract the values from the desired column, and store this as a variable
sumexpr=$(awk '{ print $2 }' scores.txt | paste -sd+)

# Add the values to get the sum, store as a variable
sum=$(echo "scale=2; $sumexpr" | bc)

# We have the sum, now we need the total amount of entries to calculate the average
columns=$(wc -l < scores.txt)

# Divide sum by total
result=$(echo "scale=2; $sum / $columns" | bc)
