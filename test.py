my_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
x = 3

# Loop through every 'x' elements using a for loop
for i in range(0, len(my_list), x):
    group = my_list[i:i + x]  # Get a slice of 'x' elements
    average = sum(group) / len(group)  # Calculate the average
    print(f"Group: {group}, Average: {average:.2f}")
