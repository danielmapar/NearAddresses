#!/usr/bin/python3
from tkinter import *
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import load_workbook
import googlemaps
import threading
import subprocess

class Interface(Frame):

	left_align = 20
	left_align_field = 130

	def __init__(self, parent):

		frame = Frame.__init__(self, parent)

		self.parent = parent

		self.initUI()

	def initUI(self):

		self.parent.title("Near Addresses")

		self.pack(fill=BOTH, expand=1)
		self.center_window()

		self.my_address_text()
		self.distance_miles_text()

		self.quit_button()
		self.find_distance_button()

	def center_window(self):

		self.width = 600
		self.heigth = 200

		sw = self.parent.winfo_screenwidth()
		sh = self.parent.winfo_screenheight()

		x = (sw - self.width)/2
		y = (sh - self.heigth)/2
		self.parent.geometry('%dx%d+%d+%d' % (self.width, self.heigth, x, y))


	def my_address_label(self):

		self.my_address_label = Label(self, text="Your Address: ", font=("Helvetica", 18))
		self.my_address_label.place(x=self.left_align, y=50)

	def my_address_text(self):

		self.my_address_label()

		self.my_address_text = Text(self, height=2, width=50)
		self.my_address_text.config(bd=0, insertbackground="white", bg='black', fg="white")
		self.my_address_text.place(x=self.left_align+self.left_align_field, y=50)

	def distance_miles_label(self):

		self.distance_miles_label = Label(self, text="Distance: ", font=("Helvetica", 18))
		self.distance_miles_label.place(x=self.left_align, y=100)

	def distance_miles_text(self):

		self.distance_miles_label()

		self.distance_miles_text = Text(self, height=1, width=10)
		self.distance_miles_text.config(bd=0, insertbackground="white", bg='black', fg="white")
		self.distance_miles_text.place(x=self.left_align+self.left_align_field, y=100)

	def quit_button(self):

		self.quit_button = Button(self, text="Quit", command=self.close)
		self.quit_button.place(x=self.left_align+90, y=150)

	def find_distance_button(self):

		self.find_distance_button = Button(self, text="Find Excel", command=self.find_distance)
		self.find_distance_button.place(x=self.left_align, y=150)

	def close(self):
		self.quit()

	def find_distance(self):

		def callback():

			self.find_distance_button.config(state='disabled')
			self.quit_button.config(state='disabled')

			if self.my_address_text.get('0.0',END) and self.my_address_text.get('0.0',END).isspace() == True and\
				self.distance_miles_text.get('0.0',END) and self.distance_miles_text.get('0.0',END).isspace() == True:
				messagebox.showerror('Error', 'Provide an address and a distance!')
				return

			self.required_distance = int(self.distance_miles_text.get('0.0',END).lower().replace('\n', '').replace('\r', '').strip())

			self.filename = filedialog.askopenfilename(title = "choose your file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))

			try:
				wb = load_workbook(self.filename)
				sheet_ranges = wb[wb.get_sheet_names()[0]]

				line = 2
				address_list = []
				while sheet_ranges['C' + str(line)].value != None:
					# Clear old values
					sheet_ranges['H' + str(line)] = None
					sheet_ranges['J' + str(line)] = None

					address_list.append(str(\
					str(sheet_ranges['C' + str(line)].value) + ", " +\
					str(sheet_ranges['D' + str(line)].value) + ", " +\
					str(sheet_ranges['E' + str(line)].value) + ", " +\
					str(sheet_ranges['F' + str(line)].value))\
					)
					line = line + 1

				key = ''
				client = googlemaps.Client(key)

				my_address = [str(self.my_address_text.get('0.0',END).lower().replace('\n', '').replace('\r', '').strip())]

				distance_list = []
				address_list_size = len(address_list)
				start = 0
				end = 30
				while True:
					distance_ret = client.distance_matrix(my_address, address_list[start:end], mode="driving", units="imperial")
					distance_list.append(distance_ret)
					start = start + 30
					end = end + 30
					if start >= address_list_size:
						break

				distance_val_list = []
				for distance in distance_list:
					for element in distance['rows'][0]['elements']:
						distance_val_list.append([element['distance']['value'], element['distance']['text']])

				line = 2
				sheet_ranges['H1'] = "Miles"
				sheet_ranges['I1'] = "Miles Value"
				sheet_ranges['J1'] = "Is it near?"

				while sheet_ranges['C' + str(line)].value != None:
					conv_fac = 0.621371
					miles_distance = float(distance_val_list[line-2][0])/1000 * conv_fac
					sheet_ranges['H' + str(line)] = distance_val_list[line-2][1]
					sheet_ranges['I' + str(line)] = str(miles_distance)
					if miles_distance <= self.required_distance or "mi" not in distance_val_list[line-2][1]:
						sheet_ranges['J' + str(line)] = "YES"
					else:
						sheet_ranges['J' + str(line)] = "NO"
					line = line + 1

				wb.save(filename = self.filename)

				messagebox.showinfo('Success', 'File updated successfully!')

			except Exception as e:
				print(e)
				messagebox.showerror('Error', 'Processing fail!\n' + str(e))

			self.find_distance_button.config(state='normal')
			self.quit_button.config(state='normal')

		t = threading.Thread(target=callback)
		t.start()

def main():

	root = Tk()
	app = Interface(root)
	root.mainloop()

main()
