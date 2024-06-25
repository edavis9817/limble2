import pandas as pd
from collections import OrderedDict

# Load spreadsheet
df = pd.read_excel('C:\\Users\\davis\\OneDrive\\Desktop\\python\\DATA May 18th.xlsx')

# Convert the second, sixth, and twenty-sixth columns of the dataframe to a list
column_2 = df.iloc[:, 1].values.tolist()  # Here 1 is the index of the second column
column_6 = df.iloc[:, 5].values.tolist()  # Here 5 is the index of the sixth column
column_26 = df.iloc[:, 25].values.tolist()  # Here 25 is the index of the twenty-sixth column

# Create a nested list to store rows for each count and corresponding strings
nested_list = []

# Create a dictionary to store the count of lines in each category
line_counts = {}

# Go through the list
for idx, item in enumerate(column_2):
    # Check if the item is a non-empty string and has at least 3 characters
    # and if the corresponding value in the fourth column is not "PM"
    if isinstance(item, str) and len(item) >= 3 and df.iloc[idx, 3] != 'PM':
        if item[:2] == '10' and len(item) >= 4:  # Special case for strings starting with '10'
            chars = item[:4]
        else:
            chars = item[:3]
        # If all characters are digits, store the row for each count and corresponding strings
        if chars.isdigit():
            row = [chars, str(column_26[idx]), str(column_6[idx])]
            nested_list.append(row)

            # Update the count for the current category
            if chars not in line_counts:
                line_counts[chars] = 1
            else:
                line_counts[chars] += 1

# Create a list of count rows to be inserted at the beginning of nested_list
count_rows = [[category, f"Lines in Category {category}", count] for category, count in line_counts.items()]

# Insert the count rows at the beginning of nested_list
nested_list = count_rows + nested_list

# Convert nested_list to DataFrame
df_out = pd.DataFrame(nested_list, columns=['Number', 'Problem', 'Solution'])

# Write DataFrame to Excel
df_out.to_excel('C:\\Users\\davis\\OneDrive\\Desktop\\python\\Counts_output.xlsx', index=False)

import pandas as pd

def search_top_5_problems_with_most_matching_words(file_path, search_query):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Get the "Problem" and "Solution" columns as lists
    problems = df["Problem"].tolist()
    solutions = df["Solution"].tolist()

    # Create a list to store the top 5 problems and solutions
    top_results = []

    # Convert the search query to lowercase
    search_query = search_query.lower()

    # Iterate over the problems and find the ones with the most matching words
    for problem, solution in zip(problems, solutions):
        # Check if the problem is a valid string
        if isinstance(problem, str):
            # Convert the problem to lowercase
            problem_lower = problem.lower()

            # Count the number of matches with the search query
            match_count = problem_lower.count(search_query)

            # Insert the problem and solution into the top_results list at the correct position
            for i in range(len(top_results)):
                if match_count > top_results[i][2]:
                    top_results.insert(i, (problem, solution, match_count))
                    break
            else:
                if len(top_results) < 5:
                    top_results.append((problem, solution, match_count))

            # Trim the top_results list to keep only the top 5 problems and solutions
            top_results = top_results[:5]

    # Return the top 5 problems and solutions
    return top_results

def select_problem_and_get_solution(top_results):
    print("Please select a problem from the following options:")
    for i, (problem, _, _) in enumerate(top_results, 1):
        print(f"{i}. {problem}")

    user_input = input("Enter the number of the problem that fits your issue (or press Enter to skip): ")

    if user_input.isdigit():
        selected_index = int(user_input) - 1

        if 0 <= selected_index < len(top_results):
            problem, solution, _ = top_results[selected_index]
            print("Here is the solution:")
            print(solution)
        else:
            print("Invalid selection. Running the search again.")
            return None
    else:
        print("Skipping the selection. Running the search again.")
        return None

def get_work_requests_for_machine(file_path, machine):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Replace missing values (NaN or NA) in the "Problem" column with an empty string
    df["Problem"].fillna("", inplace=True)

    # Find the row with the matching machine in the "Problem" column
    row = df[df["Problem"].str.contains(f"Lines in Category {machine}", na=False)]  # Modify this condition as per your requirements

    # Extract the solution from the row
    solution = row["Solution"].values[0] if not row.empty else None

    return solution

# Example usage:
file_path = 'C:\\Users\\davis\\OneDrive\\Desktop\\python\\Counts_output.xlsx'

user_choice = input("Would you like a solution to a problem (0) or information about work requests on a machine (1)? ")

if user_choice == "0":
    while True:
        search_query = input("Enter the word to match with the problems (or press Enter to skip): ")
        if not search_query:
            print("Skipping the search.")
            break

        top_results = search_top_5_problems_with_most_matching_words(file_path, search_query)
        if top_results:
            select_problem_and_get_solution(top_results)
            if top_results is not None:
                break
        else:
            print("No matches found. Would you like to run the search again with a new word?")
            continue
            
elif user_choice == "1":
    machine = input("Enter the machine number to get work request information: ")
    solution = get_work_requests_for_machine(file_path, machine)
    if solution is not None:
        print("Work Request Information:")
        print(solution)
    else:
        print("No work request information available for the specified machine.")
else:
    print("Invalid choice. Please try again.")


