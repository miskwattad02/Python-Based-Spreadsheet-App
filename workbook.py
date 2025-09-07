import tkinter as tk
from tkinter import ttk, messagebox, font, colorchooser, filedialog
from helper import *
from improved_cell import ImprovedCell
from typing import List


class Workbook:
    """
    Class representing a workbook in a spreadsheet application.
    """

    FRAME_COLOR = "Turquoise"
    EXPRESSION_EXAMPLE = ("Example: min(a1 - c23, a2 * 2, b2 + max(ac12 + av2, a13)) - avg(a1, c2) / sum(j23, x34)"
                          " OR if(A1 <= A2,<True val>,<False val>) OR countif(A1,B15, '>15')")

    def __init__(self, root, data):
        """
        Initializes a Workbook object.

        :param root: The root tkinter object.
        :param data: The initial data for the workbook.
        """
        self.root = root
        self.build_new_workbook(data)

    def build_new_workbook(self, data):
        """
        Builds a new workbook with the provided data.

        :param data: The initial data for the workbook.
        :return: None
        """
        self.data = data
        self.rows = 8 if len(data) < 2 else len(data)
        self.cols = 8 if len(data[0]) == 0 else len(data[0])
        self.build_workbook_canvas()
        self.build_sheet_frame()
        self.build_sheet()
        self.fill_sheet()
        self.rows_columns_buttons()
        self.add_functions_options()
        self.on_focus_text: ImprovedCell = None
        self.start_entry = None
        self.selected_cells: List[ImprovedCell] = []
        self.add_font_buttons()

    def build_workbook_canvas(self):
        """
        Builds the canvas for the workbook.

        :return: None
        """
        self.canvas = tk.Canvas(self.root, width=1920, height=1080)
        self.bg = tk.PhotoImage(file="pictures/bg.png")
        self.canvas.create_image(1, 1, image=self.bg, anchor=tk.NW)
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.buttons_workbook_page)
        self.canvas.bind("<Motion>", self.mouse_motion)

    def build_sheet_frame(self):
        """
        Builds the frame for the sheet within the workbook canvas.

        :return: None
        """
        self.first_canvas = tk.Canvas(self.canvas, bg=Workbook.FRAME_COLOR, highlightthickness=0)
        self.first_canvas.place(x=41, y=252, height=805, width=1840)
        self.y_scrollbar = tk.Scrollbar(self.canvas, command=self.first_canvas.yview)
        self.y_scrollbar.place(x=1897, y=255, height=800)
        self.x_scrollbar = tk.Scrollbar(self.canvas, command=self.first_canvas.xview, orient=tk.HORIZONTAL)
        self.x_scrollbar.place(x=40, y=1060, width=1840)
        self.first_canvas.configure(yscrollcommand=self.y_scrollbar.set, xscrollcommand=self.x_scrollbar.set)
        self.sheet_frame = tk.Frame(self.first_canvas, bg=Workbook.FRAME_COLOR, highlightthickness=0)
        self.sheet_frame.place(x=0, y=0, height=805, width=1840)
        self.first_canvas.create_window((0, 0), window=self.sheet_frame, anchor=tk.NW)
        self.sheet_frame.bind("<Configure>", self.reset_scrollregion)

    def reset_scrollregion(self, event):
        """
        Resets the scroll region for the canvas based on the size of the sheet frame.

        :param event: The event triggering the function.
        :return: None
        """
        self.first_canvas.configure(scrollregion=self.first_canvas.bbox("all"))

    def buttons_workbook_page(self, event):
        """
        Handles button clicks on the workbook canvas.

        :param event: The event triggering the function.
        :return: None
        """
        if 1829 <= event.x <= 1891 and 20 <= event.y <= 95:
            self.root.destroy()
        elif 1528 <= event.x <= 1606 and 89 <= event.y <= 156:
            self.open_file()
        elif 1653 <= event.x <= 1719 and 97 <= event.y <= 160:
            self.save_file()

    def mouse_motion(self, event):
        """
        Handles mouse motion events on the workbook canvas.

        :param event: The event triggering the function.
        :return: None
        """
        if (1829 <= event.x <= 1891 and 20 <= event.y <= 95 or 1528 <= event.x <= 1606 and 89 <= event.y <= 156
                or 1653 <= event.x <= 1719 and 97 <= event.y <= 160):
            self.canvas.config(cursor="hand2")
        else:
            self.canvas.config(cursor="")

    def build_sheet(self):
        """
        Builds the sheet (grid of cells) within the workbook canvas.

        :return: None
        """
        self.add_separators()
        self.sheet = []
        for i in range(self.rows):
            row = []
            self.add_row_number_label(i + 1)
            for j in range(self.cols):
                self.add_column_letters_label(j + 2)
                entry = self.build_text(i + 1, j + 2)
                row.append(entry)
            self.sheet.append(row)

    def add_separators(self):
        """
        Adds separators between rows and columns in the sheet.

        :return: None
        """
        styl = ttk.Style()
        styl.configure('TSeparator', background='black')
        self.v_separator = ttk.Separator(self.sheet_frame, orient=tk.VERTICAL, style='black.TSeparator')
        self.v_separator.grid(column=1, row=1, rowspan=self.rows + 1, sticky=tk.NS, padx=5)

    def add_row_number_label(self, i):
        """
        Adds row number labels to the sheet frame.

        :param i: The row number.
        :return: None
        """
        label = tk.Label(self.sheet_frame, text=str(i), bg=Workbook.FRAME_COLOR, font=("Arial", 12, "bold"))
        label.grid(row=i, column=0)

    def add_column_letters_label(self, j):
        """
        Adds column letter labels to the sheet frame.

        :param j: The column number.
        :return: None
        """
        label = tk.Label(self.sheet_frame, text=number_to_excel_column(j-1),
                         bg=Workbook.FRAME_COLOR, font=("Arial", 12, "bold"))
        label.grid(row=0, column=j)

    def build_text(self, i, j):
        """
        Builds a text entry widget at the specified position in the sheet frame.

        :param i: The row index.
        :param j: The column index.
        :return: The created ImprovedCell object.
        """
        cell_object = ImprovedCell(self.sheet_frame, i-1, j-2)
        entry = cell_object.get_cell()
        entry.configure(width=16)
        cell_object.set_font("Arial", 14)
        entry.grid(row=i, column=j)
        entry.config(highlightthickness=2, highlightbackground="white")
        entry.bind("<FocusIn>", lambda event: self.on_focus_in(event, cell_object))
        entry.bind("<KeyRelease>", self.on_cell_change)
        entry.bind("<Button-1>", lambda event: self.on_click(event, cell_object))
        entry.bind("<B1-Motion>", lambda event: self.on_drag(event, cell_object, i-1, j-2))
        entry.bind("<ButtonRelease-1>", lambda event: self.on_release(event, cell_object))
        return cell_object

    def on_cell_change(self, event):
        """
        Handles cell changes in the sheet.

        :param event: The event triggering the function.
        :return: None
        """
        for row in self.sheet:
            for cell in row:
                if cell.get_function():
                    solution = self.get_function_sol(cell.get_function())
                    if solution is None:
                        solution = "Error"
                    cell.get_cell().delete(0, tk.END)
                    cell.get_cell().insert(0, solution)

    def on_focus_in(self, event, entry):
        """
        Handles focus-in events on cells in the sheet.

        :param event: The event triggering the function.
        :param entry: The ImprovedCell object representing the cell.
        :return: None
        """
        self.on_focus_text = entry
        self.cell_label.configure(text=entry.get_coord_name())
        self.selected_font.set(self.on_focus_text.get_font().cget("family"))
        self.selected_size.set(self.on_focus_text.get_font().cget("size"))
        if entry.get_function():
            self.expression.delete(0, 'end')
            self.expression.insert(0, entry.get_function())

    def rows_columns_buttons(self):
        """
        Adds buttons for adding rows and columns to the workbook canvas.

        :return: None
        """
        self.row_photo = tk.PhotoImage(file="pictures/add_r.png")
        self.column_photo = tk.PhotoImage(file="pictures/add_c.png")
        buttons = tk.Button(self.canvas, image=self.row_photo, bg=Workbook.FRAME_COLOR, command=self.build_row)
        buttons.place(x=1800, y=200)
        buttons = tk.Button(self.canvas, image=self.column_photo, bg=Workbook.FRAME_COLOR, command=self.build_column)
        buttons.place(x=1840, y=200)

    def build_row(self):
        """
        Builds a new row in the sheet.

        :return: None
        """
        self.rows += 1
        self.add_row_number_label(self.rows)
        row = []
        for j in range(self.cols):
            text = self.build_text(self.rows, j + 2)
            row.append(text)
        self.sheet.append(row)
        self.v_separator.grid_configure(rowspan=self.rows)

    def build_column(self):
        """
        Builds a new column in the sheet.

        :return: None
        """
        self.cols += 1
        self.add_column_letters_label(self.cols+1)
        for i in range(self.rows):
            text = self.build_text(i + 1, self.cols+1)
            self.sheet[i].append(text)

    def add_functions_options(self):
        """
        Adds options for entering functions in cells.

        :return: None
        """
        self.expression_object = ImprovedCell(self.canvas, 0, 0)
        self.expression = self.expression_object.get_cell()
        self.expression.place(x=610, y=145, width=800, height=25)
        self.expression.insert(0, Workbook.EXPRESSION_EXAMPLE)
        self.expression.configure(state='disabled')
        self.expression.bind('<Button-1>', lambda x: self.expression_focus_in())
        self.expression.bind(
            '<FocusOut>', lambda x: self.expression_focus_out(Workbook.EXPRESSION_EXAMPLE))

        self.cell_label = tk.Label(self.canvas, text="", font=(
            "Helvetica", 10))
        self.cell_label.place(x=575, y=146, width=33, height=24)

        self.submit_im = tk.PhotoImage(file="pictures/submit.png")
        self.delete_im = tk.PhotoImage(file="pictures/delete.png")
        submit = tk.Button(self.canvas, image=self.submit_im, command=self.submit_button)
        submit.place(x=550, y=147)
        delete = tk.Button(self.canvas, image=self.delete_im, command=self.delete_button)
        delete.place(x=525, y=147)

        button = tk.Button(self.canvas, text="MIN", command=lambda: self.function_button("MIN"), font=("Helvetica", 16))
        button.place(x=540, y=87, width=60, height=42)
        button = tk.Button(self.canvas, text="MAX", command=lambda: self.function_button("MAX"), font=("Helvetica", 16))
        button.place(x=610, y=87, width=60, height=42)
        button = tk.Button(self.canvas, text="SUM", command=lambda: self.function_button("SUM"), font=("Helvetica", 16))
        button.place(x=680, y=87, width=60, height=42)
        button = tk.Button(self.canvas, text="AVERAGE", command=lambda: self.function_button("AVERAGE"), font=("Arial", 10, 'bold'))
        button.place(x=750, y=87, width=70, height=42)
        button = tk.Button(self.canvas, text="SQRT", command=lambda: self.function_button("SQRT"), font=("Helvetica", 15))
        button.place(x=830, y=87, width=60, height=42)
        button = tk.Button(self.canvas, text="IF", command=lambda: self.function_button("IF"), font=("Helvetica", 16))
        button.place(x=900, y=87, width=60, height=42)
        button = tk.Button(self.canvas, text="COUNTIF", command=lambda: self.function_button("COUNTIF"),
                           font=("Arial", 10, 'bold'))
        button.place(x=970, y=87, width=70, height=42)

    def expression_focus_in(self):
        """
        Handles focus-in events on the function expression entry widget.

        :return: None
        """
        if self.expression.cget('state') == 'disabled':
            self.expression.configure(state='normal')
            self.expression.delete(0, 'end')

    def expression_focus_out(self, placeholder):
        """
        Handles focus-out events on the function expression entry widget.

        :param placeholder: The placeholder text for the expression entry widget.
        :return: None
        """
        if self.expression.get() == "":
            self.expression.insert(0, placeholder)
            self.expression.configure(state='disabled')

    def submit_button(self):
        """
        Handles the submission of a function expression.

        :return: None
        """
        if not self.expression.get() or self.expression.cget('state') == 'disabled':
            return
        solution = self.get_function_sol(self.expression.get())
        if not self.on_focus_text:
            return
        self.on_focus_text.get_cell().delete(0, tk.END)
        self.on_focus_text.get_cell().insert(0, solution)
        self.on_focus_text.set_function(self.expression.get().upper())

    def get_function_sol(self, function):
        """
        Retrieves the solution for a given function expression.

        :param function: The function expression.
        :return: The solution for the function expression.
        """
        try:
            solution = solve_expression(function, self.get_sheet_values())
            return solution
        except:
            messagebox.showwarning("Invalid Expression", "Please enter a valid expression")
            return "Error"

    def delete_button(self):
        """
        Handles the deletion of a function expression.

        :return: None
        """
        self.expression.delete(0, 'end')
        if not self.on_focus_text:
            return
        self.on_focus_text.clear_function()

    def function_button(self, function):
        """
        Handles the addition of a function to the function expression.

        :param function: The function to be added.
        :return: None
        """
        cursor_pos = self.expression.index(tk.INSERT)
        cells = ""
        if len(self.selected_cells) > 1:
            for cell in self.selected_cells:
                cells += cell.get_coord_name() + ","
            cells = cells[:-1]
        self.expression.insert(cursor_pos, function + "(" + cells + ")")
        cursor_pos = self.expression.index(tk.INSERT)
        self.expression.icursor(cursor_pos-1)

    def get_sheet_values(self):
        """
        Retrieves the values from the cells in the sheet.

        :return: A list of lists containing the values from the cells.
        """
        sheet_values = []
        for row in self.sheet:
            row_values = []
            for entry in row:
                row_values.append(entry.get_cell().get())
            sheet_values.append(row_values)
        return sheet_values

    # ############################### drag extension ####################################
    def on_click(self, event, cell_object):
        """
        Handles mouse click events on cells in the sheet.

        :param event: The event triggering the function.
        :param cell_object: The ImprovedCell object representing the clicked cell.
        :return: None
        """
        for cell in self.selected_cells:
            cell.get_cell().config(highlightthickness=2, highlightbackground="white")
        self.selected_cells = []
        self.start_entry = cell_object
        self.selected_cells.append(cell_object)

    def on_drag(self, event, cell_object, i, j):
        """
        Handles cell dragging events in the sheet.

        :param event: The event triggering the function.
        :param cell_object: The ImprovedCell object representing the dragged cell.
        :param i: The row index of the dragged cell.
        :param j: The column index of the dragged cell.
        :return: None
        """
        cell_object.get_cell().config(highlightthickness=2, highlightbackground="black")
        if self.start_entry:
            try:
                i += event.y // 30
                j += event.x // 185
                if 0 <= i < len(self.sheet) and 0 <= j < len(self.sheet[0]):
                    if self.sheet[i][j] not in self.selected_cells:
                        self.sheet[i][j].get_cell().config(highlightthickness=2, highlightbackground="black")
                        self.selected_cells.append(self.sheet[i][j])
            except:
                pass

    def on_release(self, event, cell_object):
        """
        Handles cell release events in the sheet.

        :param event: The event triggering the function.
        :param cell_object: The ImprovedCell object representing the released cell.
        :return: None
        """
        if not self.start_entry:
            return
        if self.start_entry.get_function() == "":
            return
        if not self.selected_cells[1:]:
            return
        selected_cells = self.get_cells_names(self.selected_cells)
        ans = messagebox.askyesno("Function Addition", "Are you sure you want to add dependent functions to "
                                                 "these cells?\n" + selected_cells)
        cell_object.get_cell().config(highlightthickness=2, highlightbackground="white")
        function = self.start_entry.get_function()
        for cell in self.selected_cells[1:]:
            if ans:
                function = get_next_function(function)
                cell.set_function(function)
                cell.get_cell().delete(0, tk.END)
                cell.get_cell().insert(0, self.get_function_sol(function))
            cell.get_cell().config(highlightthickness=2, highlightbackground="white")
        self.start_entry = None
        self.selected_cells = []

    def get_cells_names(self, cells):
        """
        Retrieves the names of the selected cells.

        :param cells: List of ImprovedCell objects representing selected cells.
        :return: A string containing the names of the selected cells.
        """
        names = ""
        for cell in cells[1:]:
            names += cell.get_coord_name() + ", "
        return names

    # ############################### Font Customization ####################################

    def add_font_buttons(self):
        """
        Adds buttons for font customization to the workbook canvas.

        :return: None
        """
        all_fonts = font.families()
        self.selected_font = tk.StringVar(self.canvas)
        self.selected_font.set("Arial")
        dropdown = ttk.Combobox(self.canvas, textvariable=self.selected_font, values=all_fonts)
        dropdown.place(x=110, y=85, width=130)
        dropdown.bind("<<ComboboxSelected>>", self.change_font)

        sizes = [str(num) for num in range(8, 40)]
        self.selected_size = tk.StringVar(self.canvas)
        self.selected_size.set("14")
        dropdown = ttk.Combobox(self.canvas, textvariable=self.selected_size, values=sizes)
        dropdown.place(x=260, y=85, width=40)
        dropdown.bind("<<ComboboxSelected>>", self.change_size)
        self.add_align_buttons()
        self.add_customize_buttons()

    def add_align_buttons(self):
        """
        Adds buttons for text alignment and font style to the workbook canvas.

        :return: None
        """
        self.left_align = tk.PhotoImage(file="pictures/text_align/align-left.png")
        self.right_align = tk.PhotoImage(file="pictures/text_align/align-right.png")
        self.center_align = tk.PhotoImage(file="pictures/text_align/align-center.png")

        left_align = tk.Button(self.canvas, image=self.left_align, command=lambda: self.align_text("left"))
        left_align.place(x=310, y=85)
        center_align = tk.Button(self.canvas, image=self.center_align, command=lambda: self.align_text("center"))
        center_align.place(x=360, y=85)
        right_align = tk.Button(self.canvas, image=self.right_align, command=lambda: self.align_text("right"))
        right_align.place(x=410, y=85)

    def add_customize_buttons(self):
        self.bold = tk.PhotoImage(file="pictures/text_align/bold.png")
        self.italic = tk.PhotoImage(file="pictures/text_align/italic.png")
        self.underline = tk.PhotoImage(file="pictures/text_align/text-underline.png")
        self.text_color = tk.PhotoImage(file="pictures/text_align/gui-text-color.png")
        self.fill_color = tk.PhotoImage(file="pictures/text_align/color-fill.png")

        bold = tk.Button(self.canvas, image=self.bold, command=lambda: self.customize_font("bold"))
        bold.place(x=110, y=127)
        italic = tk.Button(self.canvas, image=self.italic, command=lambda: self.customize_font("italic"))
        italic.place(x=140, y=127)
        underline = tk.Button(self.canvas, image=self.underline, command=lambda: self.customize_font("under"))
        underline.place(x=170, y=127)
        text_color = tk.Button(self.canvas, image=self.text_color, command=lambda: self.change_color("text"))
        text_color.place(x=200, y=127)
        fill_color = tk.Button(self.canvas, image=self.fill_color, command=lambda: self.change_color("entry"))
        fill_color.place(x=230, y=127)

    def change_font(self, event):
        """
        Changes the font of the selected cell(s).

        :param event: The event triggering the function.
        :return: None
        """
        if self.on_focus_text:
            self.on_focus_text.change_font(self.selected_font.get())

    def change_size(self, event):
        """
        Changes the font size of the selected cell(s).

        :param event: The event triggering the function.
        :return: None
        """
        if self.on_focus_text:
            self.on_focus_text.change_font_size(self.selected_size.get())

    def align_text(self, position):
        """
        Aligns the text in the selected cell(s) based on the specified position.

        :param position: The alignment position ("left", "center", or "right").
        :return: None
        """
        if self.on_focus_text:
            self.on_focus_text.align(position)

    def customize_font(self, function):
        """
        Customizes the font style of the selected cell(s).

        :param function: The font customization function ("bold", "italic", or "under").
        :return: None
        """
        if self.on_focus_text:
            self.on_focus_text.font_customize(function)

    def change_color(self, change):
        """
        Changes the text or background color of the selected cell(s).

        :param change: The type of change ("text" for text color, "entry" for background color).
        :return: None
        """
        if self.on_focus_text:
            color = colorchooser.askcolor(title="Choose Color")
            if color[1]:
                self.on_focus_text.color_customize(change, color[1])

    def fill_sheet(self):
        """
        Fills the sheet with data from the provided workbook data.

        :return: None
        """
        for i in range(len(self.data)):
            for j in range(len(self.data[0])):
                if self.data[i][j]:
                    self.sheet[i][j].get_cell().delete(0, tk.END)
                    self.sheet[i][j].get_cell().insert(0, self.data[i][j])

    def open_file(self):
        """
        Opens a file dialog for selecting a workbook file to open.

        :return: None
        """
        try:
            data = open_file()
            if data[0][0]:
                pass
            self.canvas.destroy()
            self.build_new_workbook(data)
        except:
            messagebox.showwarning("Failed", "Wrong file format")

    def save_file(self):
        """
        Opens a file dialog for selecting a location to save the workbook data.

        :return: None
        """
        filetypes = (
            ("JSON files", "*.json"), ("YAML files", "*.yaml"), ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv"), ("PDF files", "*.pdf")
        )
        data = self.get_sheet_data()
        filepath = filedialog.asksaveasfilename(title="Save File", filetypes=filetypes, defaultextension=".json")
        if filepath:
            try:
                file_type = filepath.split(".")[-1]
                if file_type == "json":
                    write_json_file(filepath, data)
                elif file_type == "yaml":
                    write_yaml_file(filepath, data)
                elif file_type == "xlsx":
                    write_excel_file(filepath, data)
                elif file_type == "csv":
                    write_csv_file(filepath, data)
                elif file_type == "pdf":
                    write_pdf_file(filepath, data)
                messagebox.showinfo("Success", "The file has been saved successfully")
            except:
                messagebox.showwarning("Failed", "Can't save file")

    def get_sheet_data(self):
        """
        Retrieves the data from the cells in the sheet.

        :return: A list of lists containing the data from the cells.
        """
        data = [[self.sheet[i][j].get_cell().get() for j in range(len(self.sheet[0]))] for i in range(len(self.sheet))]
        return data


def open_file():
    """
    Opens a file dialog for selecting a workbook file to open.

    :return: None
    """
    filetypes = (
        ("JSON files", "*.json"), ("YAML files", "*.yaml"), ("Excel files", "*.xlsx"),
        ("CSV files", "*.csv"), ("PDF files", "*.pdf")
    )
    filepath = filedialog.askopenfilename(title="Open File", filetypes=filetypes)
    if filepath:
        data = [[]]
        file_type = filepath.split(".")[-1]
        if file_type == "json":
            data = read_json_file(filepath)
        elif file_type == "yaml":
            data = read_yaml_file(filepath)
        elif file_type == "xlsx":
            data = read_excel_file(filepath)
        elif file_type == "csv":
            data = read_csv_file(filepath)
        elif file_type == "pdf":
            data = read_pdf_file(filepath)
        return data

