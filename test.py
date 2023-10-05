my_list = ["apple", "banana", "cherry", "date"]
search_string = "ban"

found = any(search_string in item for item in my_list)

if found:
    print(f"'{search_string}' found in the list.")
else:
    print(f"'{search_string}' not found in the list.")
