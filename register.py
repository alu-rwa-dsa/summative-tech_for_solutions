from the_system import *

# create the sheet
sheet = applications_document.active
#an excel function to handle the excel document
applications=[]
details = {'form number': '', 'first name': '', 'Middle name': '', 'Last name': '', 'Gender': '',
           'Email ID': '', 'Phone number': '', 'Address': '', 'Occupation': ''}
n_applications= 0
form_n = []
global form_number
def create_form_number():
    global n_applications
    n_applications = n_applications + 1
    global form_number
    if len(form_n) <= 1:
        form_number = random.randint(0, 100)
        form_n.append(form_number)
    else:
        while True:
            try:
                form_n_calc = form_number + random.randint(form_number, form_number + 100)
                if form_n_calc not in form_n:
                    form_number = form_n_calc
                    break
            except ValueError:
                pass
app_details = {'form number': '', 'first name': '', 'Middle name': '', 'Last name': '', 'Gender': '',
               'Email ID': '', 'Phone number': '', 'Address': '', 'Occupation': ''}
def details_add():
    try:
        app_details['form number'] = form_number
        app_details['first name'] = first_name_field.get()
        app_details['Middle name'] = middle_name_field.get()
        app_details['Last name'] = last_name_field.get()
        app_details['Gender'] = gender_field.get()
        app_details['Email ID'] = Email_ID_field.get()
        app_details['Phone number'] = phone_n_field.get()
        app_details['Address'] = address_field.get()
        app_details['Occupation'] = occupation_field.get()
    except TclError:
        pass
    applications.append(app_details)
def print_details():
    print("Details of the just submitted application")
    f = 1
    for x in app_details:
        print("%s. %s: %s" %(f, x, app_details[x]))
        f = f+1

position = 0
columns_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
def excel():
    for i in details:
        global position
        position = position + 1
        sheet.cell(row=1, column=position).value = i
        applications_document.save('NSO_applications.xlsx')
    # sizing the columns width to meet the input's length
    def size_sheet():
        for value in range(1,len(columns_names)+1):
            sheet.column_dimensions[columns_names[value-1]].width = 30
    size_sheet()
    applications_document.save('NSO_members.xlsx')
    # Function to set focus (cursor)
def focus_first_name(): global first_name_field; first_name_field.focus_set()
def focus_middle_name(): middle_name_field.focus_set()
def focus_last_name(): last_name_field.focus_set()
def focus_gender(): gender_field.focus_set()
def focus_email_ID(): Email_ID_field.focus_set()
def focus_phone_n(): phone_n_field.focus_set()
def focus_address(): address_field.focus_set()
def focus_occupation(): occupation_field.focus_set()
# Clear the boxes
def clear():
    # clear the content of text entry box
    first_name_field.delete(0, END)
    middle_name_field.delete(0, END)
    last_name_field.delete(0, END)
    gender_field.delete(0,END)
    Email_ID_field.delete(0, END)
    occupation_field.delete(0, END); phone_n_field.delete(0, END); address_field.delete(0, END)
# Function to take data from GUI
# window and write to an excel file
def insert():
    details.clear()
    details_add()
    # if user not fill any entry
    # then print "empty input"

    if first_name_field.get() == "" and last_name_field.get() == "" \
            and gender_field.get()==""and Email_ID_field.get() == "" and occupation_field.get() == "" \
            and phone_n_field.get() == "" and address_field.get() == "":
        messagebox.showinfo('Data incomplete', 'Missing relevant details')
    else:
        # write methods into the excel sheet
        column = 0
        cur_row = 1
        row_view = 0
        while row_view == 0:
            cur_row = cur_row + 1
            value = sheet.cell(row=cur_row, column=1).value
            if value is None:
                row_view = 1
        current_row = cur_row
        for ins in app_details:
            column = column + 1
            sheet.cell(row=current_row, column=column).value = app_details[ins]
        # save the file
        applications_document.save('NSO_members.xlsx')
        messagebox.showinfo("Application status", "Your application has been successfully submitted")
        print_details()
        # Reset the focus to the first name field
        first_name_field.focus_set()
        # clear the field
        clear()

# Run the code
# create a GUI window and its parameters

def gui_f():
    root = Tk();
    background = 'Light blue'
    root.configure(background=background);
    root.title("NSO membership registration form");
    root.geometry("500x300")
    empty_position = Label(root, text="", background=background);
    empty_position1 = Label(root, text="", background=background)
    heading = Label(root, text="NSO Application form", background=background);
    first_name = Label(root, text="First Name", background=background);
    middle_name = Label(root, text="Middle Name", background=background);
    last_name = Label(root, text="Last Name", background=background);
    gender = Label(root, text="Gender(F/M/Other)", background=background)
    Email_ID = Label(root, text="Email ID", background=background);
    occupation = Label(root, text="Occupation", background=background);
    phone_n = Label(root, text="Phone number", background=background);
    address = Label(root, text="Address", background=background)
    # Setting widgets in position using grid method
    global pos
    pos = 0
    def new_pos():
        global pos;
        pos = pos + 1

    empty_position.grid(row=pos, column=1);
    new_pos();
    heading.grid(row=pos, column=1);
    new_pos();
    empty_position1.grid(row=pos, column=1);
    new_pos();
    global p_b; p_b = pos;
    first_name.grid(row=pos, column=0);
    new_pos();
    middle_name.grid(row=pos, column=0);
    new_pos();
    last_name.grid(row=pos, column=0);
    new_pos();
    gender.grid(row=pos, column=0)
    new_pos();
    Email_ID.grid(row=pos, column=0);
    new_pos();
    occupation.grid(row=pos, column=0);
    new_pos();
    phone_n.grid(row=pos, column=0);
    new_pos();
    address.grid(row=pos, column=0)

    # Create entry boxes
    global first_name_field
    global middle_name_field
    global last_name_field
    global gender_field
    global Email_ID_field
    global occupation_field
    global phone_n_field
    global address_field
    first_name_field = Entry(root)
    middle_name_field = Entry(root)
    last_name_field = Entry(root)
    gender_field = Entry(root)
    Email_ID_field = Entry(root)
    occupation_field = Entry(root)
    phone_n_field = Entry(root)
    address_field = Entry(root)


    # binding the events with the fields, the enter key calls the focus functions
    first_name_field.bind("<Return>", focus_first_name())
    middle_name_field.bind("<Return>", focus_middle_name())
    last_name_field.bind("<Return>", focus_last_name())
    gender_field.bind("<Return>", focus_gender())
    Email_ID_field.bind("<Return>", focus_email_ID())
    occupation_field.bind("<Return>", focus_occupation())
    phone_n_field.bind("<Return>", focus_phone_n())
    address_field.bind("<Return>", focus_address())

    # Setting the gursor to the first name field
    first_name_field.focus_set()

    def pos_box():
        global p_b
        p_b = p_b + 1

    # Setting the fields in position
    first_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    middle_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    last_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    gender_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    Email_ID_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    occupation_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    phone_n_field.grid(row=p_b, column=1, ipadx="100")
    pos_box();
    address_field.grid(row=p_b, column=1, ipadx="100")

    # set the titles for the excel sheet
    excel()

    create_form_number()

    # Set a Register button
    submit = Button(root, text="Register", fg="Black", background="Blue", command=insert)
    empty_position2 = Label(root, text="", background=background)
    new_pos();
    empty_position2.grid(row=pos, column=1)
    new_pos();
    submit.grid(row=pos, column=1)
    # start the GUI
    root.mainloop()
    exit()
gui_f()

