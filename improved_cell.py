import tkinter as tk
from helper import number_to_excel_column
from tkinter import font


class ImprovedCell:
    """
    A class representing an improved cell in a spreadsheet.
    """

    def __init__(self, root, i, j):
        """
        Initializes an ImprovedCell object.

        :param root: The root tkinter object.
        :param i: The row index of the cell.
        :param j: The column index of the cell.
        """
        self.entry = tk.Entry(root)
        self.row = i
        self.column = j
        self.coord_name = number_to_excel_column(j+1) + str(i + 1)
        self.function = None
        self.undo_stack = []
        self.redo_stack = []
        self.undo_redo_extension()
        self.font = font.Font(family="Arial", size=6)

    def set_font(self, font_, size):
        """
        Sets the font of the cell.

        :param font_: The name of the font family.
        :param size: The font size.
        :return: None
        """
        self.font = font.Font(family=font_, size=size)
        self.entry['font'] = self.font

    def get_font(self):
        """
        Retrieves the font of the cell.

        :return: The font object.
        """
        return self.font

    def set_function(self, function):
        """
        Sets a function for the cell.

        :param function: The function to be set.
        :return: None
        """
        self.function = function

    def clear_function(self):
        """
        Clears the function associated with the cell.

        :return: None
        """
        self.function = None

    def get_function(self):
        """
        Retrieves the function associated with the cell.

        :return: The function associated with the cell, or an empty string if no function is set.
        """
        if self.function:
            return self.function
        return ""

    def get_cell(self):
        """
        Retrieves the Entry widget representing the cell.

        :return: The Entry widget representing the cell.
        """
        return self.entry

    def get_coord_name(self):
        """
        Retrieves the coordinate name of the cell.

        :return: The coordinate name of the cell.
        """
        return self.coord_name

    def undo_redo_extension(self):
        """
        Binds events for undo and redo functionalities.

        :return: None
        """
        self.entry.bind("<Key>", self.on_change)
        self.entry.old_value = ""
        self.entry.bind("<Control-z>", self.undo)
        self.entry.bind("<Control-y>", self.redo)

    def on_change(self, event=None):
        """
        Triggered when the content of the cell changes.

        :param event: The event that triggered the change.
        :return: None
        """
        new_value = self.entry.get()
        if new_value != self.entry.old_value:
            self.undo_stack.append(self.entry.old_value)
            self.entry.old_value = new_value

    def undo(self, event=None):
        """
        Undoes the last change made to the cell.

        :param event: The event that triggered the undo action.
        :return: None
        """
        if self.undo_stack:
            current_text = self.undo_stack.pop()
            self.redo_stack.append(self.entry.old_value)
            self.entry.old_value = current_text
            self.entry.delete(0, tk.END)
            self.entry.insert(0, current_text)

    def redo(self, event=None):
        """
        Redoes the last undone change made to the cell.

        :param event: The event that triggered the redo action.
        :return: None
        """
        if self.redo_stack:
            current_text = self.redo_stack.pop()
            self.undo_stack.append(self.entry.old_value)
            self.entry.old_value = current_text
            self.entry.delete(0, tk.END)
            self.entry.insert(0, current_text)

    def change_font(self, font_):
        """
        Changes the font of the cell.

        :param font_: The name of the font family.
        :return: None
        """
        self.font.configure(family=font_)
        self.entry.config(font=self.font)

    def change_font_size(self, font_size):
        """
        Changes the font size of the cell.

        :param font_size: The size of the font.
        :return: None
        """
        self.font.configure(size=font_size)
        self.entry.config(font=self.font)

    def font_customize(self, function):
        """
        Applies font customization to the cell.

        :param function: The type of font customization to apply.
        :return: None
        """
        if function == "bold":
            if self.font.cget("weight") == "normal":
                self.font.configure(weight="bold")
            else:
                self.font.configure(weight="normal")
        elif function == "italic":
            if self.font.cget("slant") == "roman":
                self.font.configure(slant="italic")
            else:
                self.font.configure(slant="roman")
        elif function == "under":
            if self.font.cget("underline"):
                self.font.configure(underline=False)
            else:
                self.font.configure(underline=True)

    def color_customize(self, change, color):
        """
        Applies color customization to the cell.

        :param change: The type of customization to apply (text or entry).
        :param color: The color to apply.
        :return: None
        """
        if change == "text":
            self.entry.configure(fg=color)
        elif change == "entry":
            self.entry.configure(bg=color)

    def align(self, position):
        """
        Aligns the text within the cell.

        :param position: The alignment position (left, center, or right).
        :return: None
        """
        self.entry.configure(justify=position)
