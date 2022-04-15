import the_system
from the_system import *
#the_system.fetch_app_data()
#the_system.fetch_member_data()
# create a GUI window and its parameters
new_user = 'False'
def new_member():
    print('f')
    import register
    register.gui_f()

def navigation():
    adm = Tk()
    adm.geometry("500x300")
    background = 'Light blue'
    adm.configure(background='purple')
    def admin_l():
        adm_login = Frame()
        adm_login.configure(background="purple")
        adm_login.place(x=100,y=50,width=300,height=200)
        free_space = Label(adm_login, text="", background='purple')
        free_space.pack()
        heading = Label(adm_login, text="Enter Passcode for Authorisation", background='purple');
        heading.pack()
        free_space = Label(adm_login, text="", background='purple')
        free_space.pack()
        pass_s = StringVar()
        passcode = Entry(adm_login, textvariable=pass_s, show='*')
        passcode.pack()
        def show_passcode():
            if var.get() == 1:
                passcode.configure(show = '')
            if var.get() == 0:
                passcode.configure(show='*')
        var = IntVar()
        var.set(0)
        Show_pass = Checkbutton(adm_login, text='show', variable=var, onvalue=1, offvalue=0,
                                command=show_passcode)
        Show_pass.pack()
        free_space = Label(adm_login, text="", background='purple')
        free_space.pack()
        def check_p():
            pass_c = passcode.get()
            try:
                if int(pass_c) == admin_passcode:
                    import Admin
                    Admin.adm_nav()
                else:
                    def wrong_p():
                        messagebox.showinfo('Error', 'Wrong passcode')
                        passcode.delete(0, END)
                    wrong_p()
            except ValueError:
                messagebox.showinfo(ValueError,'Hint: Passcode is only digits')
                passcode.delete(0, END)
        login = Button(adm_login, text="Login", command=check_p)
        login.pack()
    def new_member_l():
        import register
        register.gui_f()
    f_admin = Frame()
    f_admin.configure(background='light blue')
    f_admin.place(x=40, y=30, width=425, height =250)
    free_space = Label(f_admin, text="", background=background)
    free_space.pack()
    font_heading = ("Book Antiqua", 15, "bold")
    heading = Label(f_admin, text="NSO NAVIGATION SYSTEM")
    heading.configure(font=font_heading)
    heading.pack()
    free_space = Label(f_admin, text="", background=background)
    free_space.pack()
    def nav():
        adm_b = Button(f_admin, text="Admin Login", command=admin_l, background='purple')
        adm_b.pack()
        free_space = Label(f_admin, text="", background=background)
        free_space.pack()
        new_m_b = Button(f_admin, text="New member", command=new_member_l, background='purple')
        new_m_b.pack()
    nav()
    adm.mainloop()
    exit()
navigation()

