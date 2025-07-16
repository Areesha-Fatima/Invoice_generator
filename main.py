import tkinter
from tkinter import ttk  
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import datetime
import os

# Clear item inputs
def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

# Global list to store invoice items
invoice_list = []

# Edit item function
def edit_item():
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)["values"]
        
        converted_item = [int(item_values[0]), item_values[1], float(item_values[2]), float(item_values[3])]
        
        qty_spinbox.delete(0, tkinter.END)
        qty_spinbox.insert(0, converted_item[0])
        desc_entry.delete(0, tkinter.END)
        desc_entry.insert(0, converted_item[1])
        price_spinbox.delete(0, tkinter.END)
        price_spinbox.insert(0, converted_item[2])

        tree.delete(selected_item)
        if converted_item in invoice_list:
            invoice_list.remove(converted_item)
    else:
        messagebox.showwarning("No Selection", "Please select an item to edit.")

# Remove item function
def remove_item():
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)["values"]
        
        converted_item = [int(item_values[0]), item_values[1], float(item_values[2]), float(item_values[3])]

        tree.delete(selected_item)
        if converted_item in invoice_list:
            invoice_list.remove(converted_item)
    else:
        messagebox.showwarning("No Selection", "Please select an item to remove.")


#add item function 
def add_item():
    desc = desc_entry.get()
    qty = int(qty_spinbox.get())     
    price = float(price_spinbox.get()) 

    if desc == "" or qty == "" or price == "":
        messagebox.showwarning("Missing Data", "Please fill out all fields to add an item.")
        return

    total = qty * price
    item = [qty, desc, price, total]
    tree.insert('', 0, values=item)
    invoice_list.append(item)
    clear_item()


#new invoice function 
def new_invoice(): 
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

# Generate the invoice DOCX file function
def generate_invoice():
    if not invoice_list:
        messagebox.showwarning("No Items", "Please add at least one item to generate an invoice.")
        return

    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get() + " " + last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal * (1 + salestax)

    doc.render({
        "name": name,
        "phone": phone,
        "invoice_list": invoice_list,
        "subtotal": "%.2f" % subtotal,
        "salestax": str(salestax * 100) + "%",
        "total": "%.2f" % total
    })

    os.makedirs("invoices", exist_ok=True)

    doc_name = "invoices/" + name.replace(" ", "_").replace(",", "") + "_" + datetime.now().strftime("%Y-%m-%d--%H-%M-%S") + ".docx"
    
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", f"Invoice saved as: {doc_name}")
    new_invoice()


# Tkinter UI setup
window = tkinter.Tk()
window.title("Invoice Generator")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10) 

# First name
tkinter.Label(frame, text="First Name").grid(row=0, column=0)
first_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)

# Last name
tkinter.Label(frame, text="Last Name").grid(row=0, column=1)
last_name_entry = tkinter.Entry(frame)
last_name_entry.grid(row=1, column=1)

# Phone number
tkinter.Label(frame, text="Phone Number").grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=2)

# Quantity
tkinter.Label(frame, text="Qty").grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

# Item description
tkinter.Label(frame, text="Description").grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

# Unit price
tkinter.Label(frame, text="Unit Price").grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=3, column=2)

# button
edit_btn = tkinter.Button(frame,text="Edit Item", command=edit_item)
edit_btn.grid(row=4, column=0, pady=5)

remove_btn = tkinter.Button(frame, text="Remove Item", command=remove_item)
remove_btn.grid(row=4, column=1, pady=5)

add_btn = tkinter.Button(frame, text="Add Item", command=add_item)
add_btn.grid(row=4, column=2, pady=5)

# Treeview to show items
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col.capitalize())
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

# Generate and New invoice buttons
tkinter.Button(frame, text="Generate Invoice", command=generate_invoice).grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
tkinter.Button(frame, text="New Invoice", command=new_invoice).grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

window.mainloop() 