#!/usr/bin/python3
from tkinter import *
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import load_workbook
import googlemaps
import threading
import subprocess
from geopy.distance import vincenty
from igraph import *


class Interface(Frame):

	left_align = 20
	left_align_field = 80

	def __init__(self, parent):

		frame = Frame.__init__(self, parent)

		self.parent = parent

		self.initUI()

	def initUI(self):

		self.parent.title("Near Addresses")

		self.pack(fill=BOTH, expand=1)
		self.center_window()

		self.distance_miles_text()

		self.quit_button()
		self.find_distance_button()

	def center_window(self):

		self.width = 300
		self.heigth = 150

		sw = self.parent.winfo_screenwidth()
		sh = self.parent.winfo_screenheight()

		x = (sw - self.width)/2
		y = (sh - self.heigth)/2
		self.parent.geometry('%dx%d+%d+%d' % (self.width, self.heigth, x, y))


	def distance_miles_label(self):

		self.distance_miles_label = Label(self, text="Distance: ", font=("Helvetica", 18))
		self.distance_miles_label.place(x=self.left_align, y=50)

	def distance_miles_text(self):

		self.distance_miles_label()

		self.distance_miles_text = Text(self, height=1, width=20)
		self.distance_miles_text.config(bd=0, insertbackground="white", bg='black', fg="white")
		self.distance_miles_text.place(x=self.left_align+self.left_align_field, y=50)

	def quit_button(self):

		self.quit_button = Button(self, text="Quit", command=self.close)
		self.quit_button.place(x=self.left_align+90, y=100)

	def find_distance_button(self):

		self.find_distance_button = Button(self, text="Find Excel", command=self.find_distance)
		self.find_distance_button.place(x=self.left_align, y=100)

	def close(self):
		self.quit()

	def find_distance(self):

		def callback():

			key = ''
			client = googlemaps.Client(key)

			if self.distance_miles_text.get('0.0',END) and self.distance_miles_text.get('0.0',END).isspace() == True:
				messagebox.showerror('Error', 'Provide a distance!')
				return

			self.find_distance_button.config(state='disabled')
			self.quit_button.config(state='disabled')

			self.required_distance = int(self.distance_miles_text.get('0.0',END).lower().replace('\n', '').replace('\r', '').strip())

			self.filename = filedialog.askopenfilename(title = "choose your file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))

			graph = Graph()

			try:
				wb = load_workbook(self.filename)
				sheet_ranges = wb[wb.get_sheet_names()[0]]

				sheet_ranges['H1'] = "Lat"
				sheet_ranges['I1'] = "Long"

				line = 2
				address_list = []
				address_geolocation = None
				while sheet_ranges['C' + str(line)].value != None:

					# Append complete addresses
					address = str(\
					str(sheet_ranges['C' + str(line)].value) + ", " +\
					str(sheet_ranges['D' + str(line)].value) + ", " +\
					str(sheet_ranges['E' + str(line)].value) + ", " +\
					str(sheet_ranges['F' + str(line)].value))

					if sheet_ranges['H' + str(line)].value == None and sheet_ranges['I' + str(line)].value == None:
						address_geolocation = client.geocode(address=address)
						lat = address_geolocation[0]['geometry']['location']['lat']
						lng = address_geolocation[0]['geometry']['location']['lng']
						sheet_ranges['H' + str(line)] = lat
						sheet_ranges['I' + str(line)] = lng
						address_list.append((lat, lng))
					else:
						lat = sheet_ranges['H' + str(line)].value
						lng = sheet_ranges['I' + str(line)].value
						address_list.append((lat, lng))

					current_address = (lat, lng)

					i = 2
					while sheet_ranges['C' + str(i)].value != None:
						if i != line:
							lat = sheet_ranges['H' + str(i)].value
							lng = sheet_ranges['I' + str(i)].value
							other_address = (lat, lng)
							distance = vincenty(current_address, other_address).miles
							if distance <= self.required_distance:
								try:
									graph.vs.find(str(line))
								except Exception as e:
									graph.add_vertex(name=str(line))

								try:
									graph.vs.find(str(i))
								except Exception as e:
									graph.add_vertex(name=str(i))
								print(str(line) + " -- " + str(i))
								graph.add_edge(graph.vs.find(str(line)), graph.vs.find(str(i)), weight = distance)
						i = i + 1

					line = line + 1

				print(VertexCover(graph))
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