import my_date

hello = my_date.My_date()
print(hello.get_query_Start())
print(hello.get_query_end())

csv_row = {}
print(str(type(csv_row)))

if str(type(csv_row)) == "<class 'dict'>":
    print("yes")
