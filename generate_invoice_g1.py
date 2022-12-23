import tkinter
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate








def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0,"1")
    description_entry.delete(0, tkinter.END)
    unitprice_spinbox.delete(0, tkinter.END)
    unitprice_spinbox.insert(0,"0.00")
    
invoice_list = []


def add_item():
    qty = int(qty_spinbox.get())
    desc = description_entry.get()
    price = float(unitprice_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]
    
    tree_box.insert('',0, values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)
    


def new_invoice():
    fname_entry.delete(0, tkinter.END)
    lname_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree_box.delete(*tree_box.get_children())
    
    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("generate_invoice_template.docx")
  #  doc = DocxTemplate("invoice_template.docx")
    name = fname_entry.get()+lname_entry.get()
    phone =phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal*(1-salestax)
    
    
    doc.render({"name":name,
        "phone":phone,
        "invoice_list":invoice_list,
        "subtotal":subtotal,
        "salestax": str(salestax*100)+"%",
        "total":total})
    
    doc_name = "generated_new_invoice_" + name +".docx"
    doc.save(doc_name)
    
    #doc_name = "new_invoice" + name + datetime.now().strftime("%Y-%m-%d-%H%M%S")+".docx"
    #doc.save(doc_name)
    
    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    
    new_invoice()






window=tkinter.Tk()
window.title("Babsolut Compuworld eInvoice")
window.resizable(False,False)

frame = tkinter.Frame(window, pady=20, padx=20)
frame.pack()
#row1
fname_label = tkinter.Label(frame, text="First Name")
lname_label = tkinter.Label(frame, text="Last Name")
phone_label = tkinter.Label(frame, text="Phone")
fname_label.grid(row=0, column=0)
lname_label.grid(row=0, column=1)
phone_label.grid(row=0, column=2)

fname_entry = tkinter.Entry(frame)
lname_entry = tkinter.Entry(frame)
phone_entry = tkinter.Entry(frame)

fname_entry.grid(row=1, column=0)
lname_entry.grid(row=1, column=1)
phone_entry.grid(row=1, column=2)



#row2
qty_label = tkinter.Label(frame, text="Qty")
description_label = tkinter.Label(frame, text="Description")
unitprice_label = tkinter.Label(frame, text="Unit Price")
qty_label.grid(row=2, column=0)
description_label.grid(row=2, column=1)
unitprice_label.grid(row=2, column=2)

qty_spinbox = tkinter.Spinbox(frame, from_=1, to=500)
description_entry = tkinter.Entry(frame)
unitprice_spinbox = tkinter.Spinbox(frame, from_=0.00, to="infinity")

qty_spinbox.grid(row=3, column=0)
description_entry.grid(row=3, column=1)
unitprice_spinbox.grid(row=3, column=2)

add_button = tkinter.Button(frame, text="Add Item", command=add_item)
add_button.grid(row=4, column=2, pady=5)

#tree
columns = (("qty", "desc","uprice"))
tree_box = ttk.Treeview(frame, columns=columns, show="headings")

tree_box.grid(row=5, column=0, columnspan=3)


tree_box.heading("qty",text="Qty")
tree_box.heading("desc",text="Description")
tree_box.heading("uprice",text="Unit Price")


save_invoice_button = tkinter.Button(frame, text="Generate Invoice", fg="teal", command=generate_invoice)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)

save_invoice_button.grid(row=6, column=0,columnspan=3, sticky="news", padx=20, pady=10)
new_invoice_button.grid(row=7, column=0,columnspan=3, sticky="news", padx=20, pady=10)



window.mainloop()