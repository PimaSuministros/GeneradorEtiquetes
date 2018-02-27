import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from os.path import join
import docx

class Application():

    TOP_SPACE, LEFT_SPACE = 15, 30
    TEXT_SIZE = 9
    UPPERCASE, LOWERCASE = None, None

    def __init__(self, master=None):
        self.master = master
        self.build_gui()

    def build_gui(self, width=20, height=30, cols=4):
        # Headers
        self.headers = []
        for h in range(4):
            e = tk.Entry(self.master, width=width+10)
            e.grid(row=0, column=h)
            self.headers.append(e)
            e = None
        
        # Rows
        self.rows = []
        for t in range(4):
            r = ScrolledText(self.master, width=width, height=height)
            r.grid(row=1, column=t)
            self.rows.append(r)
            r = None
        
        # Checkboxes
        self.ROWS_UPPERCASE = tk.IntVar()
        self.ROWS_LOWERCASE = tk.IntVar()
        tk.Checkbutton(
            self.master, 
            text="en MAJUSCULES", 
            variable=self.UPPERCASE
        ).grid(row=1, column=cols, sticky='N')
        tk.Checkbutton(
            self.master, 
            text="en minuscules", 
            variable=self.LOWERCASE
        ).grid(row=1, column=cols+1, sticky='N')
        
        # Button
        self.do = tk.Button(text='Generar etiquetes', command=self.do_action)
        self.do.grid(row=2, column=cols+1, sticky='EW')
        
    
    def do_action(self):
        self.generate_excel(
            data = self.prepare_data(), 
            entry_name = join(
                join('.','input'),
                'Etiquetes.docx'
            ),
            output_name = join(
                join('.', 'output'),
                'Result.docx'
            )
        )
        self.master.destroy()
    
    
    def process_text(self, text):
        replacements = [
            ('\t', ' ')
            #, ('  ',' ')
        ]
        for search, replace in replacements:
            text = text.replace(search, replace)
        if self.ROWS_UPPERCASE.get():
            text = text.upper()
        if self.ROWS_LOWERCASE.get():
            text = text.lower()
        return text
    
    
    def prepare_data(self):
        headers_text = [ h.get() for h in self.headers ]
        rows_text = [ r.get(1.0, tk.END) for r in self.rows ]
        
        cells = [""]*(len(headers_text) * len(rows_text[0].splitlines()))
        for header, col in zip(headers_text, rows_text):
            for i in range(len(col.splitlines())):
                if not col=="":
                    cells[i] += "{header}{sep}{text}{newline}" \
                        .format(
                            header=header if len(col.splitlines()[i].strip())>0 else "",
                            text=self.process_text(col.splitlines()[i]),
                            sep=": " if header and len(col.splitlines()[i].strip())>0 else "",
                            newline="\n"
                        )
        return [cell for cell in cells if cell]

        
    def generate_excel(self, data, entry_name, output_name):
        document = docx.Document(entry_name)
        
        i=0
        for row_i, row in enumerate(document.tables[0].rows):
            for col_i, cell in enumerate(row.cells):
                # add text
                if i<len(data):
                    cell.text = data[i]
                    i+=1
                elif col_i==0:
                    # TODO: remove row or cells
                    pass
                
                # format cell
                for paragraph in cell.paragraphs:
                    p_format = paragraph.paragraph_format
                    p_format.left_indent = docx.shared.Pt(self.LEFT_SPACE)
                    p_format.space_before = docx.shared.Pt(self.TOP_SPACE)
                    
                    for run in paragraph.runs:
                        font = run.font
                        font.size = docx.shared.Pt(self.TEXT_SIZE)

        document.save(output_name)

if __name__ == '__main__':
    root = tk.Tk()
    app = Application(root)
    root.mainloop()


