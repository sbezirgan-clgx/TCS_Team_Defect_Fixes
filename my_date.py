from datetime import date
from datetime import timedelta
class My_date:

    def __init__(self):
        self.today = date.today()
        self.days_of_the_last_week = []
        for i in range(2, 10):
            self.day = self.today - timedelta(days=i)
            self.days_of_the_last_week.append(str(self.day))
        print(self.days_of_the_last_week)
        self.jquery_start = self.days_of_the_last_week[-1]
        self.jquery_end = self.days_of_the_last_week[0]

        #print(self.jquery_start + "////" + self.jquery_end)
    def get_query_Start(self):
        return self.jquery_start


    def get_query_end(self):
        return self.jquery_end
