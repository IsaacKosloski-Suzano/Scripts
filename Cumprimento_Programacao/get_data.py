from datetime import datetime

# Get current date and time
now = datetime.now()

# Extract required data
week_number = now.isocalendar()[1]  # Week number of the year
month_number = now.strftime("%m")  # Month as a number (01-12)
last_two_digits_year = now.strftime("%y")  # Last two digits of the year
revisao = f"SM{week_number}{month_number}{last_two_digits_year}"

# Display the results
print(f"Week Number: {week_number}")
print(f"Month Number: {month_number}")
print(f"Last Two Digits of the Year: {last_two_digits_year}")
print(f"Revisão:-{revisao}")
