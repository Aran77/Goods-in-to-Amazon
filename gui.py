from tkinter import *
import customtkinter
import tkinter as tk
from tkinter import ttk



app = customtkinter.CTk()
# configure window
app.title("MT Clothing Goods In")
app.geometry(f"{1100}x{580}")

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"


# create sidebar frame with widgets



sidebar_frame = customtkinter.CTkFrame(app, width=140, corner_radius=0)
sidebar_frame.pack(side=LEFT, fill=Y)
data_frame = customtkinter.CTkFrame(app, corner_radius=0)
data_frame.pack(side=LEFT, expand=TRUE, fill=BOTH, padx=5, pady=5)
logo_label = customtkinter.CTkLabel(sidebar_frame, text="Goods In", font=customtkinter.CTkFont(size=20, weight="bold"))
logo_label.pack(side=TOP, padx=5, pady=5)

sidebar_button_1 = customtkinter.CTkButton(sidebar_frame, text="Import Goods In File")
sidebar_button_1.pack(side=TOP, padx=5, pady=5)
import_label = customtkinter.CTkLabel(sidebar_frame, text="Generate Import Files")
import_label.pack(side=TOP, padx=5, pady=5)
sidebar_button_2 = customtkinter.CTkButton(sidebar_frame, text="Amazon UK & US")
sidebar_button_2.pack(side=TOP, padx=5, pady=5)
sidebar_button_6 = customtkinter.CTkButton(sidebar_frame, text="Amazon Euros")
sidebar_button_6.pack(side=TOP, padx=5, pady=5)
sidebar_button_3 = customtkinter.CTkButton(sidebar_frame, text="Shopify")
sidebar_button_3.pack(side=TOP, padx=5, pady=5)
sidebar_button_4 = customtkinter.CTkButton(sidebar_frame, text="Wish")
sidebar_button_4.pack(side=TOP, padx=5, pady=5)
sidebar_button_5 = customtkinter.CTkButton(sidebar_frame, text="Priceminister")
sidebar_button_5.pack(side=TOP, padx=5, pady=5)
sidebar_button_7 = customtkinter.CTkButton(sidebar_frame, text="Cdiscount")
sidebar_button_7.pack(side=TOP, padx=5, pady=5)
sidebar_button_8 = customtkinter.CTkButton(sidebar_frame, text="Onbuy")
sidebar_button_8.pack(side=TOP, padx=5, pady=5)
quit_button = customtkinter.CTkButton(sidebar_frame, command=app.destroy, text="Quit")
quit_button.pack(side=BOTTOM, padx=5, pady=5)

columns = ("SKU","Title", "Barcode","Stock Level", "Price","UK Size", "US Size", "EU Size", "Colour")
widths =(50,100,50,25,25,25,25,25,25)
tree = ttk.Treeview(data_frame, columns=columns, show='headings')
# define headings
tree.heading('SKU', text='SKU')
tree.heading('Title', text='Title')
tree.heading('Barcode', text='Barcode')
tree.heading('Stock Level', text='Stock Level')
tree.heading('Price', text='Price')
tree.heading('UK Size', text='UK Size')
tree.heading('US Size', text='US Size')
tree.heading('EU Size', text='EU Size')
tree.heading('Colour', text='Colour')
tree.pack(expand=TRUE, fill=BOTH)

style = ttk.Style(app) 
style.theme_use("clam") # set theam to clam
style.configure('Treeview.Heading', background="Gray", font=('Helvetica', 22))

#appearance_mode_label = customtkinter.CTkLabel(sidebar_frame, text="Appearance Mode:", anchor="w")
#appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
#appearance_mode_optionemenu = customtkinter.CTkOptionMenu(sidebar_frame, values=["Light", "Dark", "System"])
#appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
#scaling_label = customtkinter.CTkLabel(sidebar_frame, text="UI Scaling:", anchor="w")
#scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
#scaling_optionemenu = customtkinter.CTkOptionMenu(sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"])
#scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

# create main entry and button


app.mainloop()