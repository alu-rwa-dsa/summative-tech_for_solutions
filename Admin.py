from the_system import *
Admin_ch = ['1. Reject application', '2. Accept Application']
# Create a clean sheet of members in the members document
sheet_members = members_document.active


def excel():
    # creating the titles
    sheet_members.cell(row=1, column=1).value = "Form Number"
    sheet_members.cell(row=1, column=2).value = "First Name"
    sheet_members.cell(row=1, column=3).value = "Middle Name"
    sheet_members.cell(row=1, column=4).value = "Last Name"
    sheet_members.cell(row=1, column=5).value = "Email ID"
    sheet_members.cell(row=1, column=6).value = "Phone number"
    sheet_members.cell(row=1, column=7).value = "Address"
    sheet_members.cell(row=1, column=8).value = "Occupation"

    # sizing the columns width to meet the input's length
    def size_sheet():
        sheet_members.column_dimensions['A'].width = 30
        sheet_members.column_dimensions['B'].width = 30
        sheet_members.column_dimensions['C'].width = 30
        sheet_members.column_dimensions['D'].width = 30
        sheet_members.column_dimensions['E'].width = 40
        sheet_members.column_dimensions['F'].width = 20
        sheet_members.column_dimensions['G'].width = 20

    size_sheet()


def accept_reject_app():
    queue_apps = []
    v_apps = Frame()
    v_apps.configure(background="light grey")
    v_apps.place(x=120, y=90, width=420, height=230)
    free_space = Label(v_apps, background='light grey')
    free_space.pack()
    font_heading = ("Book Antiqua", 15, "bold")
    heading = Label(v_apps, text='NSO submitted applications', background='light grey', font=font_heading)
    heading.pack()
    for i in range(len(applications)):
        form_n = 'Application number: ' + str(applications[i]['form number'])
        names = (applications[i]['first name'] + ' ' + applications[i]['Last name'])
        queue_apps.append(names)
        appl = Label(v_apps, text=str(i + 1) + '. ' + form_n + ', Names: ' + names, background='light grey')
        appl.pack()
    back = Button(v_apps, text="back", command=adm_nav)
    back.pack()
    back.place(x=30, y=20)
    free_space = Label(v_apps, background='light grey', font=font_heading)
    free_space.pack()
    heading = Label(v_apps, text='Below: enter the application number of the application to review',
                    background='light grey')
    heading.pack()
    app_no = Entry(v_apps, textvariable=StringVar())
    app_no.pack()

    def chec_no():
        ap_no = app_no.get()
        apps = []
        for l in range(len(applications)):
            apps.append(applications[l]['form number'])
        try:
            if int(app_no.get()) in apps:
                for i in range(len(applications)):
                    if int(app_no.get()) == int(applications[i]['form number']):
                        def acc_rej():
                            applic = applications[i]
                            queue_apps = []
                            ac_rej = Frame()
                            ac_rej.configure(background='light grey')
                            ac_rej.place(x=120, y=90, width=420, height=230)
                            font_heading = ("Book Antiqua", 15, "bold")
                            free_space = Label(ac_rej, text="", background='light grey')
                            free_space.pack()
                            heading = Label(ac_rej, text='Application being reviewed:', background='light grey',
                                            font=font_heading)
                            heading.pack()
                            ap = Label(ac_rej, text=applic, background='light grey')
                            ap.pack()

                            def accept_app():
                                current_row = sheet_members.max_row
                                current_column = sheet_members.max_column
                                global column
                                column = 1

                                def column_get():
                                    global column
                                    column = column + 1

                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'form number']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'first name']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'Middle name']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'Last name']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i]['Gender']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'Email ID']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'form number']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'Phone number']
                                column_get()
                                sheet_members.cell(row=current_row + 1, column=column).value = applications[i][
                                    'Address']
                                # save the file
                                members_document.save('NSO_assets\\NSO_members.xlsx')
                                members.append(applications[i])
                                succ = Frame()
                                succ.configure(background='light grey')
                                succ.place(x=120, y=90, width=420, height=230)
                                heading = \
                                    Label(succ, text="Application " + str(applications[i]['form number']) +
                                               ' successfully accepted.\n Contact ' + applications[i]['Phone number'] +
                                                     " to inform them of their applications's status", background='light grey')
                                heading.pack()
                                heading.place(x=40, y=50)
                                back = Button(succ, text="back", command=accept_reject_app, background='white')
                                back.pack()
                                back.place(x=30, y=20)
                                del applications[i]


                            def rej_app():
                                rejected_applications.append(applications[i])
                                messagebox.showinfo("Application " + str(applications[i]['form number']),
                                                    'Application successfully rejected.\nContact ' + applications[i][
                                                        'Phone number'] +
                                                    " to inform them of their applications's status")
                                del applications[i]

                            acc = Button(ac_rej, text="Accept application", command=accept_app, background='white')
                            acc.pack()
                            rej = Button(ac_rej, text="Reject application", command=rej_app, background='white')
                            rej.pack()

                        acc_rej()
            else:
                messagebox.showinfo('App n0', "Enter a real application number")
        except ValueError:
            messagebox.showinfo(ValueError, "Enter correct application number")

    confirm_app = Button(v_apps, text="Confirm app no:", command=chec_no, background='white')
    confirm_app.pack()

def adm_nav():
    view_t = Tk()
    view_t.geometry("700x500")
    background = 'Light blue'
    view_t.configure(background='purple')
    def view_members():
        queue_members = []
        view_m_f = Frame()
        view_m_f.configure(background='light grey')
        view_m_f.place(x=120, y=90, width=420, height=230)
        back = Button(view_m_f, text="back", command=nav, background='white')
        back.pack()
        back.place(x=30, y=20)
        font_heading = ("Book Antiqua", 15, "bold")
        free_space = Label(view_m_f, text="", background='light grey')
        free_space.pack()
        heading = Label(view_m_f, text='NSO current members', background='light grey', font=font_heading)
        heading.pack()
        for i in range(len(members)):
            names = (members[i]['first name'] + ' ' + members[i]['Last name'])
            queue_members.append(names)
            memb = Label(view_m_f, text= str(i+1) + '. ' + names,background='light grey')
            memb.pack()
    def view_applications():
        queue_applications = []
        view_app = Frame()
        view_app.configure(background="light grey")
        view_app.place(x=120, y=90, width=420, height=230)
        free_space = Label(view_app, text="", background='light grey')
        free_space.pack()
        font_heading = ("Book Antiqua", 15, "bold")
        heading = Label(view_app, text='NSO submitted applications', background='light grey', font=font_heading)
        heading.pack()
        for i in range(len(applications)):
            form_n = 'Application number: ' + str(applications[i]['form number'])
            names = (applications[i]['first name'] + ' ' + applications[i]['Last name'])
            queue_applications.append(names)
            appl = Label(view_app, text=str(i + 1) + '. ' + form_n + ', Names: ' + names, background='light grey')
            appl.pack()
        back = Button(view_app, text="back", command=nav)
        back.pack()
        back.place(x=30, y=20)

    def nav():
        adm_opt = Frame()
        adm_opt.configure(background="Orange")
        adm_opt.place(x=40, y=30, width=600, height=400)
        def fram():
            fram_opt = Frame()
            fram_opt.configure(background="purple")
            fram_opt.place(x=120, y=90, width=420, height=230)
            free_space = Label(fram_opt, text="", background='purple')
            free_space.pack()
            font_heading = ("Book Antiqua", 15, "bold")
            heading = Label(fram_opt, text="ADMIN NAVIGATION", background='orange')
            heading.configure(font=font_heading)
            heading.pack()
            free_space = Label(fram_opt, text="", background='purple')
            free_space.pack()
            view_members_b = Button(fram_opt, text="View members", command=view_members, background='grey')
            view_members_b.pack()
            View_applications_b = Button(fram_opt, text="View Applications", command=view_applications,
                                         background='grey')
            View_applications_b.pack()
            acc_rej_app_b = Button(fram_opt, text="Accept/Reject Applications", command=accept_reject_app,
                                   background='light grey')
            acc_rej_app_b.pack()
            free_space = Label(fram_opt, text="", background='purple')
            free_space.pack()
            logout = Button(fram_opt, text="Log out", command=exit, background='blue')
            logout.pack()
        fram()
    nav()
    view_t.mainloop()
    exit()
adm_nav()

