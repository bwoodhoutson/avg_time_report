### Please forgive my formatting (self teaching)
from datetime import datetime as dt
from datetime import timedelta as tdelta
from tkinter import *
import xlsxwriter
import xlrd

def abtr_func():
    # function called by tkinter window                                 
    input_file = entry_1.get()
    wb_in = xlrd.open_workbook(input_file)
    sheet = wb_in.sheet_by_index(0) 
    lor = []
    # lor = list of whole rows
    list_of_times = []  
    # list of times only
    
    ### building current_abtr.xlsx   
    wb_out = xlsxwriter.Workbook('current_abtr.xlsx')
    ws1 = wb_out.add_worksheet("Avg Time by Bidder")
    ws2 = wb_out.add_worksheet("Avg Time by Dealer")
    bold_format = wb_out.add_format({"bold":True})
    avg_times_format = wb_out.add_format({"num_format":'[h]:mm:ss'})
    bold_avg_times_format = wb_out.add_format({"num_format":'[h]:mm:ss',"bold":True})
    ws1.write('A1', 'Bidder')
    ws1.write('B1', 'Average Time')
    ws1.write('C1', '# of bids in avg')
    ws1.set_column('A:A', 40)
    ws1.set_column('B:B', 15)
    ws1.set_column('C:C', 15)
    ws1.set_row(0, 15, bold_format)
    ws2.write('A1', 'Dealer', bold_format)
    ws2.write('B1', 'Average Time', bold_format)
    ws2.write('C1', '# of bids in avg', bold_format)
    ws2.set_column('A:A', 50)
    ws2.set_column('B:C', 15,)
    
   ### reading BTR
    for i in range(sheet.nrows):  
        index = sheet.cell_value(i,0)
        bidder = sheet.cell_value(i,4)
        dealer = sheet.cell_value(i,5)
        if index == str('InventoryId'):
            print("Reading BTR...")
        else:
            time_in = sheet.cell_value(i,6)     
            time_out = sheet.cell_value(i,7)
            t1_tup = xlrd.xldate_as_tuple(time_in, wb_in.datemode)  
            t2_tup = xlrd.xldate_as_tuple(time_out, wb_in.datemode) 
            t1_str = ','.join(str(v) for v in t1_tup)
            t2_str = ','.join(str(v) for v in t2_tup)
            # between xlrd and the join this is the new format of the times
            t1 = dt.strptime(t1_str,"%Y,%m,%d,%H,%M,%S")   
            t2 = dt.strptime(t2_str,"%Y,%m,%d,%H,%M,%S")
            t3 = t2-t1
            row_tuple = (index, bidder, dealer, t3)
            if t3.total_seconds() > 10 and t3.total_seconds() < 600: 
                #We only calculate avearges between ten seconds and ten minutes. 
                lor.append(row_tuple)
                list_of_times.append(t3)
    
    # dictionary with bidder as key
    b = {}   
    [b [t [1]].append(t [3]) if t [1] in b.keys() else b.update({t [1]: [t [3]]}) for t in lor]

    # dictionary with dealer as key
    d = {}    
    [d [t [2]].append(t [3]) if t [2] in d.keys() else d.update({t [2]: [t [3]]}) for t in lor]
    
        
    '''
    #This was how I originally built these distionaries. 
    
    for i in range(len(lor)):
        if lor[i][1] in b.keys():                             # if key is present in the list...
            b[lor[i][1]].append(lor[i][3])                    # just append the value
        else:
            b[lor[i][1]]= []                                 # else create a empty list as value for the key
            b[lor[i][1]].append(lor[i][3])                   # now append the value for that key
    print(b)

    for i in range(len(lor)):
        if lor[i][2] in d.keys():                            # if key 
            d[lor[i][2]].append(lor[i][3])                    # just append 
        else:
            d[lor[i][2]]= []                                 # else create 
            d[lor[i][2]].append(lor[i][3])                   # now append 
    print(d) 
    '''
    
    # calculate and write bidder dictionary
    this_row = 1
    for key, value in b.items():
        avg_of_list = sum(value, tdelta()) / len(value)
        number_of_bids = len(value)
        # print(key, (avg_of_list)-(tdelta(microseconds = avg_of_list.microseconds)), number_of_bids)
        # xlsxwriter uses y, x coordinates
        ws1.write(this_row, 0, str(key))                                             
        ws1.write_datetime(this_row, 1, ((avg_of_list)-(tdelta(microseconds = avg_of_list.microseconds))), avg_times_format)
        ws1.write(this_row, 2, number_of_bids) 
        this_row = this_row + 1
    
    # calculate and write dealer dictionary
    that_row = 1
    for key, value in d.items():
        avg_of_list = sum(value, tdelta()) / len(value)
        number_of_bids = len(value)
        # print(key, (avg_of_list)-(tdelta(microseconds = avg_of_list.microseconds)), number_of_bids)
        # xlsxwriter uses y, x coordinates
        ws2.write(that_row, 0, str(key))                                             
        ws2.write_datetime(that_row, 1, ((avg_of_list)-(tdelta(microseconds = avg_of_list.microseconds))), avg_times_format)
        ws2.write(that_row, 2, number_of_bids) 
        that_row = that_row + 1

    avg_time = sum(list_of_times, tdelta()) / len(list_of_times)                # avg time of all bids
    total_bids_in_avg = len(list_of_times) 
    ws1.write(this_row, 0, 'Average Time of Bids')  
    ws1.write_datetime(this_row, 1 , ((avg_time)-(tdelta(microseconds = avg_of_list.microseconds))), bold_avg_times_format)
    ws1.set_row(this_row, 15, bold_format)
    ws1.write(this_row, 2 , total_bids_in_avg)
    ws2.set_row(that_row, 15, bold_format)
    
    ### wrapping up   
    print()
    print("Total rows on BTR.xlsx: {}".format(sheet.nrows))
    print("# of bids in Average: {}".format(total_bids_in_avg))
    print("Average Time: {}".format(avg_time-(tdelta(microseconds = avg_time.microseconds))))
    print()
    print("Program complete.")
    print("Look for 'current_abtr.xlsx' in the same folder as this program.")
    print()  
    wb_out.close()
    # close xlsxwriter and save           
    my_window.destroy()
    # end tkinter                                                   
    

### Tkinter Window in A-Team colors  
my_window = Tk()
top_frame = Frame(my_window,bg="black")
top_frame.pack()
entry_1 = Entry(top_frame)
button_1 = Button(top_frame, text = "Click me to run program", command = abtr_func)    # function call
label_1 = Label(top_frame, text = "--->>   New ABTR    <<---", bg="black", fg="yellow")
label_2 = Label(top_frame, text = "This window will close when program complete.", bg="black", fg="yellow")
label_3 = Label(top_frame, text = "Look for 'current_abtr.xlsx' in the same file location as this program.", bg="black", fg="yellow")
label_4 = Label(top_frame, text = "Copy and Paste Bid Timing Report Here: ", bg="black", fg="yellow")
label_1.grid(row = 0, column = 0)
label_2.grid(row = 1, column = 0)
label_3.grid(row = 2, column = 0)
label_4.grid(row = 3, column = 0)
entry_1.grid(row = 3, column = 1)    
button_1.grid(row = 4, column = 1)
my_window.mainloop()
