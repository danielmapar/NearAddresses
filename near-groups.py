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

class Vertex:
    def __init__(self, node):
        self.id = node
        self.adjacent = {}

    def __str__(self):
        return str(self.id) + ' adjacent: ' + str([x.id for x in self.adjacent])

    def add_neighbor(self, neighbor, weight=0):
        self.adjacent[neighbor] = weight

    def get_connections(self):
        return self.adjacent.keys()

    def get_id(self):
        return self.id

    def get_weight(self, neighbor):
        return self.adjacent[neighbor]

class Graph:
    def __init__(self):
        self.vert_dict = {}
        self.num_vertices = 0

    def __iter__(self):
        return iter(self.vert_dict.values())

    def get_vertex_list(self, vertexId):
    	vertex_list = []
    	for vertex in self.get_vertex(vertexId).adjacent.keys():
    		vertex_list.append(vertex.get_id())
    	return vertex_list

    def add_vertex(self, node):
        self.num_vertices = self.num_vertices + 1
        new_vertex = Vertex(node)
        self.vert_dict[node] = new_vertex
        return new_vertex

    def get_vertex(self, n):
        if n in self.vert_dict:
            return self.vert_dict[n]
        else:
            return None

    def add_edge(self, frm, to, cost = 0):
        if frm not in self.vert_dict:
            self.add_vertex(frm)
        if to not in self.vert_dict:
            self.add_vertex(to)

        self.vert_dict[frm].add_neighbor(self.vert_dict[to], cost)
        self.vert_dict[to].add_neighbor(self.vert_dict[frm], cost)

    def get_vertices(self):
        return self.vert_dict.keys()

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

				# Set Lat and Long for every address
				sheet_ranges['H1'] = "Lat"
				sheet_ranges['I1'] = "Long"

				line = 2
				address_list = []
				address_geolocation = None
				total_lines = 1
				while sheet_ranges['C' + str(line)].value != None:
					total_lines = total_lines + 1

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

					# Build a graph based on vertices distance
					i = 2
					while sheet_ranges['C' + str(i)].value != None:
						if i != line:
							lat = sheet_ranges['H' + str(i)].value
							lng = sheet_ranges['I' + str(i)].value
							other_address = (lat, lng)
							distance = vincenty(current_address, other_address).miles
							if distance <= self.required_distance:
								if graph.get_vertex(line) == None:
									graph.add_vertex(line)

								graph.add_edge(line, i, distance)
						i = i + 1

					line = line + 1


				vertices = []

				# Get vertices and neighbors
				for v in graph:
					vertices.append((v.get_id(), graph.get_vertex_list(v.get_id())))

				# Sort by vertex with more neighbors
				vertices = sorted(vertices, key=lambda x: len(x[1]), reverse=True)

				eliminated_vertices = [] 
				group_id = 1
				line = 1
				# Save optimized groups in the excel file
				sheet_ranges['J' + str(line)].value = 'Groups'
				for vertex in vertices:
					if  vertex[0] not in eliminated_vertices:
						eliminated_vertices.append(vertex[0])
						eliminated_vertices.extend(vertex[1])
						vertices_list = vertex[1]
						vertices_list.extend([vertex[0]])
						line = 2
						while sheet_ranges['C' + str(line)].value != None:
							if line in vertices_list:
								sheet_ranges['J' + str(line)].value = str(group_id)
							line = line + 1
						group_id = group_id + 1

				# Group vertices without nearby connections
				while total_lines > 0:
					if sheet_ranges['J' + str(total_lines)].value == None:
						sheet_ranges['J' + str(total_lines)].value = str(group_id)
						group_id = group_id + 1
					total_lines = total_lines - 1

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
