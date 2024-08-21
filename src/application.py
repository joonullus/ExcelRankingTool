import os
import pandas as pd
import ttkbootstrap as tk
from ttkbootstrap.constants import *

class App(tk.Window):
    def __init__(self):
        super().__init__()
        self.title("RankingTool")

        # The width and height of the app
        window_width = 1000
        window_height = 600

        # The width and height of the user's screen
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculates position x, y to center the window
        position_x = int((screen_width - window_width) / 2)
        position_y = int((screen_height - window_height) / 2)

        # Sets the geometry of the window to center it
        self.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        self.resizable(False, False)
        tk.Style("vapor")

        # Creation of frames
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)
        self.frames = {}

        for F in (StartPage, ComparisonPage, EndPage):
            frame = F(container, self)
            self.frames[F] = frame

        # Show the initial frame
        self.show_frame(StartPage)

        # Creates a Tkinter variable to monitor button actions
        self.choice = tk.StringVar()

    def show_frame(self, page_name):
        for frame in self.frames.values():
            frame.pack_forget()

        frame = self.frames[page_name]
        frame.pack(fill="both", expand=True)
        
        signature = tk.Label(frame, text="github.com/joonullus", font=("Arial", 10))
        signature.place(relx=1.0, rely=1.0, anchor="se")

    def ranking(self):
        input_df = pd.read_excel('test.xlsx', sheet_name='Sheet1')

        # Checks if the output file already exists.
        # If it does, the process continues from the already existing file.
        # Otherwise, a new output file is created.
        # This allows the user to keep the progress if the app is closed before the process is over.
        if os.path.isfile('output.xlsx'):
            output_df = pd.read_excel('output.xlsx', sheet_name='Sheet1')
        else:
            output_df = pd.DataFrame({
                'Openings': [str(input_df.loc[0, 'Anime Name']) + ' Opening ' + str(input_df.loc[0, 'Opening Number'])]
            })
            output_df.to_excel('output.xlsx', index=False)
        
        # Gets the total number of rows in the input and the output.
        no_of_total_rows_input = input_df.shape[0]
        no_of_rows_output = output_df.shape[0]

        # Binary insertion sort algorithm to reduce the number of user comparisons.
        # The progress can be tracked using the number of rows in the output file.
        for i in range(no_of_rows_output, no_of_total_rows_input):
            choice2 = str(input_df.loc[i, 'Anime Name']) + ' Opening ' + str(input_df.loc[i, 'Opening Number'])

            low = 0
            high = output_df.shape[0] - 1
            position = -1
            get_position = False

            while low <= high:
                mid = (high + low) // 2
                choice1 = str(output_df.loc[mid, 'Openings'])

                self.frames[ComparisonPage].create_buttons(choice1, choice2)
                self.wait_variable(self.choice)

                if low == high:
                    get_position = True
                
                if self.choice.get() == choice1:
                    low = mid + 1
                    if get_position:
                        position = low
                else:
                    high = mid - 1
                    if get_position:
                        position = mid

            new_row = {'Openings': choice2}
            new_row_df = pd.DataFrame([new_row], index=[position])
            output_df = pd.concat([output_df.iloc[:position], new_row_df, output_df.iloc[position:]]).reset_index(drop=True)
            output_df.to_excel('output.xlsx', index=False)
        
        self.show_frame(EndPage)
            
class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller

        label = tk.Label(self, text="Excel RankingTool v1.0", font=("Arial", 20, "bold"))
        label.pack(expand=True, pady=(50, 0))

        button = tk.Button(self, text="Start", bootstyle=PRIMARY, width=20, command=self.goto_comparison_page)
        button.pack(expand=True, pady=(0, 100))

    def goto_comparison_page(self):
        self.controller.show_frame(ComparisonPage)
        self.controller.ranking()

class ComparisonPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller

        label = tk.Label(self, text="Pick the greater one", font=("Arial", 20, "bold"))
        label.pack(expand=True)

    def create_buttons(self, choice1, choice2):
        self.button1 = tk.Button(self, text=choice1, command=lambda: self.pick_choice(choice1))
        self.button1.pack(expand=True, pady=(0, 50))

        self.button2 = tk.Button(self, text=choice2, command=lambda: self.pick_choice(choice2))
        self.button2.pack(expand=True, pady=(0, 100))

    def pick_choice(self, picked_choice):
        self.controller.choice.set(picked_choice)
        self.button1.destroy()
        self.button2.destroy()
    
class EndPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller

        label = tk.Label(self, text="Ranking completed", font=("Arial", 20, "bold"))
        label.pack(expand=True)

if __name__ == "__main__":
    app = App()
    app.mainloop()
