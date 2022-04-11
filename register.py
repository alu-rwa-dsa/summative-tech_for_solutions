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

def details_add():
    for i in details:
        try:
            details['form number'] = form_number
            details['first name'] = first_name_field.get()
            details['Middle name'] = middle_name_field.get()
            details['Last name'] = last_name_field.get()
            details['Gender'] = gender_field.get()
            details['Email ID'] = Email_ID_field.get()
            details['Phone number'] = phone_n_field.get()
            details['Address'] = address_field.get()
            details['Occupation'] = occupation_field.get()
        except TclError:
            pass
    applications.append(details)
def print_details():
    print("Details of the just submitted application")
    f = 1
    for x in details:
        print("%s. %s: %s" %(f, x, details[x]))
        f = f+1

def excel():
    # creating the titles
    sheet.cell(row=1, column=1).value = "Form Number";
    sheet.cell(row=1, column=2).value = "First Name";
    sheet.cell(row=1, column=3).value = "Middle Name"
    sheet.cell(row=1, column=4).value = "Last Name";
    sheet.cell(row=1, column=5).value = "Email ID";
    sheet.cell(row=1, column=6).value = "Phone number"
    sheet.cell(row=1, column=7).value = "Address";
    sheet.cell(row=1, column=8).value = "Occupation"
    # sizing the columns width to meet the input's length
    def size_sheet():
        sheet.column_dimensions['A'].width = 30;
        sheet.column_dimensions['B'].width = 30;
        sheet.column_dimensions['C'].width = 30
        sheet.column_dimensions['D'].width = 30;
        sheet.column_dimensions['E'].width = 40
        sheet.column_dimensions['F'].width = 20;
        sheet.column_dimensions['G'].width = 20

    size_sheet()
# Function to set focus (cursor)
def focus_first_name(): first_name_field.focus_set()
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
    # if user not fill any entry
    # then print "empty input"

    if first_name_field.get() == "" or last_name_field.get() == "" or gender_field.get()=="" \
            or Email_ID_field.get() == "" or occupation_field.get() == "" \
            or phone_n_field.get() == "" or address_field.get() == "":
        print("Relevant details missing")
    else:
        # write methods into the excel sheet
        current_row = sheet.max_row
        current_column = sheet.max_column
        global column
        column = 1
        def column_get():
            global column
            column = column + 1
        sheet.cell(row=current_row + 1, column=column).value = form_number
        column_get(); sheet.cell(row=current_row + 1, column=column).value = first_name_field.get()
        column_get(); sheet.cell(row=current_row + 1, column=column).value = middle_name_field.get()
        column_get(); sheet.cell(row=current_row + 1, column=column).value = last_name_field.get()
        column_get();sheet.cell(row=current_row + 1, column=column).value = gender_field.get()
        column_get();sheet.cell(row=current_row + 1, column=column).value = Email_ID_field.get()
        column_get();sheet.cell(row=current_row + 1, column=column).value = occupation_field.get()
        column_get();sheet.cell(row=current_row + 1, column=column).value = phone_n_field.get()
        column_get(); sheet.cell(row=current_row + 1, column=column).value = address_field.get()
        # save the file
        applications_document.save('C:NSO_assets\\NSO_applications.xlsx')
        details_add()
        print_details()
        # Reset the focus to the first name field
        first_name_field.focus_set()
        # clear the field
        clear()

# Run the code
if __name__ == "__main__":
    # create a GUI window and its parameters
    root = Tk();
    background = 'Light blue'
    root.configure(background=background); root.title("NSO membership registration form"); root.geometry("500x300")
    empty_position =  Label(root, text="", background=background); empty_position1 =  Label(root, text="", background=background)
    heading = Label(root, text="NSO Application form", background=background)
    first_name = Label(root, text="First Name", background=background)
    middle_name = Label(root, text="Middle Name", background=background)
    last_name = Label(root, text="Last Name", background=background)
    gender = Label(root, text="Gender(F/M/Other)", background=background)
    Email_ID = Label(root, text="Email ID", background=background)
    occupation = Label(root, text="Occupation", background=background)
    phone_n = Label(root, text="Phone number", background=background)
    address = Label(root, text="Address", background=background)
    # Setting widgets in position using grid method
    pos = 0
    def new_pos():
        global pos; pos = pos + 1
    empty_position.grid(row=pos, column=1);
    new_pos(); heading.grid(row=pos, column=1);
    new_pos(); empty_position1.grid(row=pos, column=1);
    new_pos(); p_b = pos; first_name.grid(row=pos, column=0);
    new_pos(); middle_name.grid(row=pos, column=0);
    new_pos(); last_name.grid(row=pos, column=0);
    new_pos(); gender.grid(row=pos, column=0)
    new_pos(); Email_ID.grid(row=pos, column=0);
    new_pos(); occupation.grid(row=pos, column=0);
    new_pos(); phone_n.grid(row=pos, column=0);
    new_pos(); address.grid(row=pos, column=0)

    # Create entry boxes
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

    #Setting the gursor to the first name field
    first_name_field.focus_set()

    def pos_box():
        global p_b; p_b = p_b + 1
    # Setting the fields in position
    first_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); middle_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); last_name_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); gender_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); Email_ID_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); occupation_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); phone_n_field.grid(row=p_b, column=1, ipadx="100")
    pos_box(); address_field.grid(row=p_b, column=1, ipadx="100")

    # set the titles for the excel sheet
    excel()

    create_form_number()

    # Set a Register button
    submit = Button(root, text="Register", fg="Black", background="Blue", command=insert)
    empty_position2 = Label(root, text="", background=background)
    new_pos(); empty_position2.grid(row=pos, column=1)
    new_pos();submit.grid(row= pos, column=1)
    # start the GUI
    root.mainloop()

