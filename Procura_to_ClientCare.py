import csv
import time
import sys
from operator import itemgetter

# Prompt to type in file directory
print 'Welcome!\n'
print 'Please confirm the raw data files have been put in the "input_files" folder.(Press Enter to continue)\n'
sys.stdin.readline()

# Open and read ServiceTypeDesc
csv_file = open('res/ServiceTypeDesc.csv')
csv_reader = csv.reader(csv_file)
data = list(csv_reader)
row_count = len(data)
service_type_desc = [['' for x in range(2)] for y in range(row_count)]
csv_file = open('res/ServiceTypeDesc.csv')
csv_reader = csv.reader(csv_file)
row_number = 0
for row in csv_reader:
    service_type_desc[row_number][0] = row[0]
    service_type_desc[row_number][1] = row[1]
    row_number += 1

# Open file for row counting
print "Opening source file..."
file_name = raw_input('Please type in the raw data file name, without the file extension: ')
file_dir = 'input_files/'+str(file_name)+'.csv'
csv_file = open(file_dir)
csv_reader = csv.reader(csv_file)

# Count the INVOICE row
print "Counting the rows..."
inv_count = 0
for row in csv_reader:
    if row[0] == 'INVOICE':
        inv_count += 1

# Create 2-D List
print "Creating index..."
inv_data = [['' for x in range(27)] for y in range(inv_count)]

# Open file for data mapping
print "Mapping data..."
csv_file = open(file_dir)
csv_reader = csv.reader(csv_file)

#Map data to list
row_number = 0
for row in csv_reader:
    if row[0] == 'INVOICE':
        for column_count in range(27):
            inv_data[row_number][column_count] = row[column_count]
        row_number += 1

# Change date format
print "Changing date format..."
for row_number in range(inv_count):
    from_date = inv_data[row_number][1]
    conv_date = time.strptime(from_date, "%Y%m%d")
    target_date = time.strftime("%d/%m/%Y", conv_date)
    inv_data[row_number][1] = target_date

# Transfer data into BillingSum
print "Transferring data..."
billing_sum = [['' for x in range(7)] for y in range(inv_count)]
row_number = 0
for row_number in range(inv_count):
    billing_sum[row_number][0] = inv_data[row_number][8]      # URN
    billing_sum[row_number][1] = inv_data[row_number][20]     # COST CENTRE
    if inv_data[row_number][22] == 'MOBILITYEXP' and inv_data[row_number][21] == 'PRIVATE':
        billing_sum[row_number][2] = 'RESCONP'                # MASTER ACCT
    elif inv_data[row_number][22] == 'MOBILITYEXP':
        billing_sum[row_number][2] = 'RESCON'
    else: billing_sum[row_number][2] = inv_data[row_number][22]
    billing_sum[row_number][3] = inv_data[row_number][1]      # RECORD DATE
    billing_sum[row_number][4] = inv_data[row_number][12]     # AMOUNT
    billing_sum[row_number][5] = inv_data[row_number][11]     # INVOICE NUMBER
    if inv_data[row_number][15] == 'CC':                      # NOTES
        flag_count = 0
        for row in service_type_desc:
            if inv_data[row_number][16] == row[0]:
                if inv_data[row_number][16] == 'CC':
                    billing_sum[row_number][6] = row[1] + " " + inv_data[row_number][18] + " day/s"
                else: billing_sum[row_number][6] = row[1]
            else: flag_count += 1
        if flag_count == len(service_type_desc):
            billing_sum[row_number][6] = "Please update the Billing Service Type definition lookup table"
    elif inv_data[row_number][16] == 'HCPADJCR':
        for row in service_type_desc:
            if row[0] == inv_data[row_number][16]:
                billing_sum[row_number][6] = row[1]
    elif inv_data[row_number][16] == 'HCPADJDB':
        for row in service_type_desc:
            if row[0] == inv_data[row_number][16]:
                billing_sum[row_number][6] = row[1]
    else: billing_sum[row_number][6] = inv_data[row_number][4] + " Service Fee"

# Sum-up (Excel Macro Function)
sum_up_dict = dict()
for row in billing_sum:
    key = str(row[0]) + str(row[1]) + str(row[2]) + str(row[3]) + str(row[5]) + str(row[6])
    if key in sum_up_dict.keys():
        sum_up_dict[key][4] = float("{0:.2f}".format(float(sum_up_dict[key][4]) + float(row[4])))
    else: sum_up_dict[key] = row

output_list = sum_up_dict.values()
output_list.sort(key=itemgetter(0))

# Export
print "Exproting..."
output_name = raw_input('Enter the output file name, without file extension: ') + '.csv'
output_dir = 'output_files/'+str(output_name)
csv_file = open(output_dir, 'wb')
csv_file_writerow = csv.writer(csv_file, delimiter=',', quoting=csv.QUOTE_NONE)
for item in output_list:
    csv_file_writerow.writerow(item)

inv_dir = 'output_files/INV_'+str(output_name)
csv_file = open(inv_dir, 'wb')
csv_file_writerow = csv.writer(csv_file, delimiter=',', quoting=csv.QUOTE_NONE)
for item in inv_data:
    csv_file_writerow.writerow(item)