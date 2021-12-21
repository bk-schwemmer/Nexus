#!/usr/bin/python

import sqlite3

conn = sqlite3.connect('store_ips.db')
success = 'Opened database successfully'
success_table = 'Created table successfully'
print(success)

#################################################################
# Execute the following command to make create_db.py executable #
# $ chmod +x create_db.py                                       #
# Execute the following command run script and create db        #
# $ ./create_db.py                                              #
# Upon successful creation, script will print success message   #
#################################################################
#################################################################
# The following creates the table and sets columns              #
#                                                               #
# conn.execute('''CREATE TABLE Store_ips                        #
#             (Store TEXT PRIMARY KEY,                          #
#             ip_Address TEXT,                                  #
#             Address TEXT,                                     #
#             City TEXT,                                        #
#             State TEXT,                                       #
#             Zip_Code TEXT)''')                                #
# print(success_table)                                          #
#################################################################


def insert_data(store_input, ip_input, address_input, city_input, state_input, zip_input):
    # Data insertion query
    try:
        sql = ''' INSERT INTO Store_ips(Store, ip_Address, Address, City, State, Zip_Code)
                 VALUES(store_input, ip_input, address_input, city_input, state_input, zip_input) '''
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        cur.close()

    except sqlite3.Error as error:
        print('Failed to insert data into sql table', error)
    finally:
        if conn:
            conn.close()
            print('Data inserted into sql table and connection has been closed')


def get_input():
    # Get input for new store information to add to sql database
    store_input = input("Store Name:")
    ip_input = input("ip Address:")
    address_input = input("Street Address:")
    city_input = input("City:")
    state_input = input("State:")
    zip_input = input("Zip Code:")

    # Call insert query function and pass user input
    insert_data(store_input, ip_input, address_input, city_input, state_input, zip_input)


answer = input("Add a new store? Y/N")
if answer == 'Y' or answer == 'y':
    get_input()
else:
    pass
