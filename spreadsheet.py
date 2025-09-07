import tkinter as tk
from workbook import Workbook, open_file
from tkinter import messagebox


class Spreadsheet:
    """
    A class representing a spreadsheet application.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.build_menu()

    def build_menu(self):
        """
        Initializes the Spreadsheet object.
        """
        self.menu_bg = tk.PhotoImage(file="pictures/menu_bg.png")
        self.menu_canvas = tk.Canvas(self.root, width=1920, height=1080)
        self.menu_canvas.create_image(1,1,image=self.menu_bg, anchor=tk.NW)
        self.menu_canvas.pack()
        self.menu_canvas.bind("<Button-1>", self.buttons_menu_page)
        self.menu_canvas.bind("<Motion>", self.mouse_motion)

    def buttons_menu_page(self, event):
        """
        Handles the events when buttons on the main menu are clicked.

        :param event: The event triggered by clicking a button.
        :return: None
        """
        if 1829 <= event.x <= 1891 and 20 <= event.y <= 95:
            self.root.destroy()
        elif 778 <= event.x <= 897 and 397 <= event.y <= 501:
            self.open_file()
        elif 1004 <= event.x <= 1110 and 403 <= event.y <= 503:
            self.new_file()

    def mouse_motion(self, event):
        """
        Changes the cursor appearance based on mouse position.
        :param event: The event triggered by mouse motion.
        :return: None
        """
        if 1829 <= event.x <= 1891 and 20 <= event.y <= 95 or 778 <= event.x <= 897 and 397 <= event.y <= 501\
                or 1004 <= event.x <= 1110 and 403 <= event.y <= 503:
            self.menu_canvas.config(cursor="hand2")
        else:
            self.menu_canvas.config(cursor="")

    def open_file(self):
        """
        Opens an existing spreadsheet file.
        """
        try:
            data = open_file()
            if data[0][0]:
                pass
            self.menu_canvas.destroy()
            self.work_book = Workbook(self.root, data)
        except:
            messagebox.showwarning("Failed", "Wrong file format")

    def new_file(self):
        """
        Creates a new empty spreadsheet file.
        """
        self.menu_canvas.destroy()
        self.work_book = Workbook(self.root, [[]])

    def start_spreadsheet(self):
        """
        Starts the spreadsheet application.
        """
        self.root.title("Spreadsheet")
        self.root.attributes('-fullscreen', True)
        self.root.mainloop()
