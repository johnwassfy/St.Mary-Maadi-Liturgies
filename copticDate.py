import datetime

class CopticCalendar:
    def __init__(self):
        self.coptic_default_date = [1740, 1, 1]  # Default Coptic date (year, month, day)
        self.gregorian_default_date = [2023, 9, 12]  # Default Gregorian date (year, month, day)
        self.current_gregorian_datetime = datetime.datetime.now()  # Current Gregorian datetime
        self.current_coptic_datetime = self.gregorian_to_coptic(self.current_gregorian_datetime)
        self.coptic_month_names = [
            "توت", "بابه", "هاتور", "كيهك", "طوبه", "امشير", "برمهات",
            "بارموده", "بشنس", "بؤونة", "أبيب", "مسرى", "نسيء"
        ]

    def coptic_month_name(self, month_number):
        # إرجاع اسم الشهر القبطي بناءً على رقمه
        if 1 <= month_number <= 13:
            return self.coptic_month_names[month_number - 1]
        else:
            raise ValueError("رقم الشهر القبطي غير صالح")

    def is_leap_year(self, coptic_year):
        # Check if the given year is a leap year in the Coptic calendar
        return (coptic_year % 4 == 3) or ((coptic_year % 4 == 0) and (coptic_year % 100 != 3) and ((coptic_year + 1) % 400 == 0))

    def days_since_default_date(self, current_gregorian_datetime):
        # Calculate the number of days between the default Gregorian date and the current Gregorian date
        gregorian_default_datetime = datetime.datetime(*self.gregorian_default_date)
        days = (current_gregorian_datetime - gregorian_default_datetime).days
        return days

    def gregorian_to_coptic(self, gregorian_datetime=None):
        if gregorian_datetime is None:
            gregorian_datetime = self.current_gregorian_datetime

        # Convert Gregorian date to Coptic date
        coptic_year, coptic_month, coptic_day = self.coptic_default_date
        days = self.days_since_default_date(gregorian_datetime)
        current_time = gregorian_datetime.time()

        # If it's past 5:30 PM, consider it as the next Coptic day
        if current_time.hour > 17 or (current_time.hour == 17 and current_time.minute >= 30):
            days += 1

        while days > 0:
            days_in_month = 30 if coptic_month < 13 else (6 if (coptic_month == 13 and self.is_leap_year(coptic_year)) else 5)
            if days >= days_in_month:
                days -= days_in_month
                coptic_month += 1
                if coptic_month == 14:
                    coptic_month = 1
                    coptic_year += 1
            else:
                coptic_day += days
                days = 0

        return [coptic_year, coptic_month, coptic_day]

    def coptic_to_gregorian(self, coptic_date):
        # Convert Coptic date to Gregorian date
        coptic_year, coptic_month, coptic_day = coptic_date
        default_coptic_year, default_coptic_month, default_coptic_day = self.coptic_default_date

        # Calculate the day difference between the given Coptic date and the default Coptic date
        days_difference = (coptic_year - default_coptic_year) * 365 + (coptic_month - default_coptic_month) * 30 + (coptic_day - default_coptic_day)

        # Apply the adjusted day difference to the default Gregorian date
        gregorian_datetime = datetime.datetime(*self.gregorian_default_date) + datetime.timedelta(days=days_difference)

        return gregorian_datetime

    def coptic_date_before(self, number, given_date):
        # Calculate the Coptic date before the given date by subtracting the number of days
        coptic_year, coptic_month, coptic_day = given_date
        coptic_day -= number
        while coptic_day <= 0:
            coptic_month -= 1
            if coptic_month == 0:
                coptic_month = 13
                coptic_year -= 1
            days_in_month = 30 if coptic_month < 13 else (6 if (coptic_month == 13 and self.is_leap_year(coptic_year)) else 5)
            coptic_day += days_in_month
        return [coptic_year, coptic_month, coptic_day]

    def days_between_dates(self, coptic_date):
        # Calculate the number of days between the given Coptic date and the current Coptic date
        current_coptic_year, current_coptic_month, current_coptic_day = self[0], self[1], self[2]
        given_coptic_year, given_coptic_month, given_coptic_day = coptic_date

        # Convert both dates to days since the Coptic epoch for easier calculation
        current_days = current_coptic_year * 365 + (current_coptic_month - 1) * 30 + current_coptic_day
        given_days = given_coptic_year * 365 + (given_coptic_month - 1) * 30 + given_coptic_day

        return given_days - current_days 

    def coptic_date_after(self, number, given_date):
        # Calculate the Coptic date after the given date by adding the number of days
        coptic_year, coptic_month, coptic_day = given_date
        while number > 0:
            days_in_month = 30 if coptic_month < 13 else (6 if self.is_leap_year(coptic_year + 1) else 5)
            if coptic_day + number > days_in_month:
                number -= (days_in_month - coptic_day + 1)
                coptic_day = 1
                coptic_month += 1
                if coptic_month == 14:
                    coptic_month = 1
                    coptic_year += 1
                    if coptic_month == 1 and self.is_leap_year(coptic_year):
                        days_in_month = 6
                    else:
                        days_in_month = 5
            else:
                coptic_day += number
                number = 0

        return [coptic_year, coptic_month, coptic_day]

    def get_coptic_date_range(self, coptic_date):
        # Define the ranges
        range_1_start = [coptic_date[0], 2, 2]
        range_1_end = [coptic_date[0], 5, 10]
        range_2_start = [coptic_date[0], 5, 11]
        range_2_end = [coptic_date[0], 10, 11]

        # Check if the given date falls within each range
        if range_1_start <= coptic_date <= range_1_end:
            return "Tree"
        elif range_2_start <= coptic_date <= range_2_end:
            return "Air"
        else:
            return "Water"

    def set_coptic_date(self, copticdate):
    
        self.current_coptic_datetime = [copticdate[0], copticdate[1], copticdate[2]]


# coptic_calendar = CopticCalendar()
# # Using the current Gregorian date
# coptic_date_current = coptic_calendar.gregorian_to_coptic()
# print(coptic_calendar.is_leap_year(coptic_date_current[0]))
# print("Coptic date for current Gregorian date:", coptic_date_current)

# # Using a specific Gregorian date
# specific_gregorian_datetime = datetime.datetime(2027, 1, 1, 15, 45)  # Example Gregorian date and time
# coptic_date_specific = coptic_calendar.gregorian_to_coptic(specific_gregorian_datetime)
# print("Specific Gregorian Date: ", specific_gregorian_datetime)
# print("Coptic date for specific Gregorian date:", coptic_date_specific)

# fasting = coptic_calendar.coptic_date_before(55, [1740, 8, 27])
# test = coptic_calendar.days_between_dates([1740, 6, 23])
# print(test)

# coptic_calendar = CopticCalendar()
# coptic_date = coptic_calendar.current_coptic_datetime# Example Coptic date
# print(coptic_date)
# range_number = coptic_calendar.get_coptic_date_range(coptic_date)
# print("Range:", range_number)