from datetime import date
from datetime import timedelta

class My_date:

    def __init__(self):
        self.today = date.today()
        self.days_of_the_last_week = []
        self.days_of_the_last_week2 = []
        for i in range(10, 17):
            self.day = self.today - timedelta(days=i)
            self.days_of_the_last_week.append(str(self.day))
        print(self.days_of_the_last_week)
        self.jquery_start = self.days_of_the_last_week[-1]
        self.jquery_end = self.days_of_the_last_week[0]

        for i in range(3,10):
            self.day2 = self.today - timedelta(days=i)
            self.days_of_the_last_week2.append(str(self.day2))
        print(self.days_of_the_last_week2)
        self.jquery_start2 = self.days_of_the_last_week2[-1]
        self.jquery_end2 = self.days_of_the_last_week2[0]

        #print(self.jquery_start + "////" + self.jquery_end)
    def get_query_Start(self):
        date_list =  self.jquery_start.split('-')
        yearly = str(date_list[0])
        monthly = str(date_list[1])
        daily = str(date_list[2])
        starter_date = monthly + '-' + daily + '-' + yearly
        return starter_date

    def get_query_Start2(self):
        date_list =  self.jquery_start2.split('-')
        yearly = str(date_list[0])
        monthly = str(date_list[1])
        daily = str(date_list[2])
        starter_date = monthly + '-' + daily + '-' + yearly
        return starter_date


    def get_query_end(self):
        date_list = self.jquery_end.split('-')
        yearly = str(date_list[0])
        monthly = str(date_list[1])
        daily = str(date_list[2])
        end_date = monthly + '-' + daily + '-' + yearly
        return end_date

    def get_query_end2(self):
        date_list = self.jquery_end2.split('-')
        yearly = str(date_list[0])
        monthly = str(date_list[1])
        daily = str(date_list[2])
        end_date = monthly + '-' + daily + '-' + yearly
        return end_date


    def get_todays_date(self):
        return self.today