import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageGrab
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import io
from datetime import datetime

class DataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("APP NAME") #Change this for the name you want.
        self.root.configure(bg="#bbe6f6")  # Set background color

        # Load and place the logo IF NEEDED
        #self.logo_path = "logo.png"  # Replace with your logo file path
        #self.logo_image = Image.open(self.logo_path)
        #self.logo_image = self.logo_image.resize((40, 40), Image.LANCZOS)  # Resize if necessary
        #self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        #self.logo_label = tk.Label(root, image=self.logo_photo, bg="#f0f0f0")
        #self.logo_label.grid(row=0, column=0, pady=10, padx=10)

        # Initialize data attributes
        self.data = None
        self.id_matrix = []

        # Create and place the browse button
        self.browse_button = tk.Button(root, text="Generate Graph", command=self.load_file, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.browse_button.grid(row=0, column=1, pady=10, padx=10)
        self.browse_button.config(state=tk.DISABLED)

        # Create and place the ID entry
        self.id_entry_label = tk.Label(root, text="Serial Number:", bg="#bbe6f6", font=("Arial", 10))
        self.id_entry_label.grid(row=1, column=0, pady=5, padx=10, sticky="e")
        self.id_entry = tk.Entry(root, font=("Arial", 10))
        self.id_entry.grid(row=1, column=1, pady=5, padx=10, sticky="w")

        # Create and place the update button
        self.update_button = tk.Button(root, text="Add", command=self.update_ids, bg="#2196F3", fg="white", font=("Arial", 12, "bold"))
        self.update_button.grid(row=1, column=2, pady=10, padx=10)

        # Create a label to display the IDs
        self.id_display_label = tk.Label(root, text="Tested Serial Numbers:", bg="#bbe6f6", font=("Arial", 10))
        self.id_display_label.grid(row=2, column=0, pady=5, padx=10, sticky="w")

        self.id_display_text = tk.Text(root, height=10, width=50, font=("Arial", 10))
        self.id_display_text.grid(row=3, column=0, columnspan=3, pady=5, padx=10)

    def update_ids(self):
        id_text = self.id_entry.get().strip()
        if id_text:
            id_list = id_text.split(',')
            self.id_matrix.append(id_list)
            self.id_entry.delete(0, tk.END)
            self.update_id_display()
            self.browse_button.config(state=tk.NORMAL)
        else:
            messagebox.showwarning("Input Error", "Please enter at least one Serial Number.")

    def update_id_display(self):
        self.id_display_text.delete(1.0, tk.END)
        for row in self.id_matrix:
            self.id_display_text.insert(tk.END, ', '.join(row) + '\n')

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            df = None
            try:
                if file_path.endswith('.xls'):
                    try:
                        # Attempt to read .xls file using xlrd
                        df = pd.read_excel(file_path, engine='xlrd', header=None)
                    except Exception as e1:
                        # Attempt to read .xls file using pyxlsb
                        import pyxlsb
                        from pyxlsb import open_workbook
                        try:
                            with open_workbook(file_path) as wb:
                                with wb.get_sheet(0) as sheet:
                                    data = [row for row in sheet.rows()]
                                    df = pd.DataFrame(data[1:], columns=data[0])
                            print("File read using pyxlsb")
                        except Exception as e2:
                            raise ValueError(f"Failed to read .xls file using both xlrd and pyxlsb: {e1} | {e2}")
                elif file_path.endswith('.xlsx'):
                    # Read .xlsx file
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                else:
                    raise ValueError("Unsupported file format. Please select a .xls or .xlsx file.")
                
                # Check if the file has at least 6 columns, this can be modify or changed, depends on your log file.
                if df.shape[1] < 6:
                    raise ValueError("The file must have at least 6 columns.")

		# Drop the first row
            	df = df.drop(index=0)
                
                # Extract data from specific columns, see your .xls file to get the right columns
                dates = df[:, 1]  # Date column (second column, index 1)
                times = df[:, 2]  # Time column (third column, index 2)
                temps_values = df[:, 5]  # Temperature column column (sixth column, index 5)

                # Convert dates and times to string format
                dates = dates.astype(str)
                times = times.astype(str)

                # Convert Ch2_Value to numeric, forcing errors to NaN
                temps_values = pd.to_numeric(temps_values, errors='coerce')

                # Prepare for plotting
                self.plot_data(dates, times, temps_values, self.id_matrix)

            except Exception as e:
                # Show an error message if loading the file fails
                messagebox.showerror("Error", f"Failed to read the file: {e}")
            

    def plot_data(self, dates, times, temps_values, id_matrix):
        print(dates)
        
        # Create a new window for the plot
        plot_window = tk.Toplevel(self.root)
        plot_window.title("Burning Process NAME") #Change Name for your burning process
        plot_window.geometry("800x600")
        plot_window.configure(bg="#bbe6f6")
        plot_frame = tk.Frame(plot_window, bg="#bbe6f6")
        plot_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        id_frame = tk.Frame(plot_window, bg="#bbe6f6")
        id_frame.pack(side=tk.RIGHT, fill=tk.Y)

            # Add the logo to the ID frame
        logo_image = Image.open(self.logo_path)
        logo_image = logo_image.resize((40, 40), Image.LANCZOS)  # Resize if necessary
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(id_frame, image=logo_photo, bg="#f0f0f0")
        logo_label.image = logo_photo  # Keep a reference to avoid garbage collection
        logo_label.pack(pady=10, padx=10)

        #Downsample data to improve clarity, IF NEEDED, the equipment measures every 600 sec
        #downsample_factor = 120  # Adjust this factor to control the number of points displayed
        #times_downsampled = times[::downsample_factor]
        #temps_values_downsampled = temps_values[::downsample_factor]

        # Create a plot

        fig, ax = plt.subplots()
        
        # Plot the data
        ax.plot(times, temps_values, marker='o', linestyle='-')
        ax.set_xlabel("Time")
        ax.set_ylabel("Temperature (°C)")

        # Set title with the date range
        if len(dates.unique()) > 1:
            date_range = f"{dates.iloc[1]} - {dates.iloc[-1]}"
        else:
            date_range = str(dates.iloc[1])
        ax.set_title(f"Burning process from: {date_range}")

        # Adjust x-axis to show a limited number of significant time labels
        num_ticks = 10
        times_list = times.tolist()
        if len(times_list) > num_ticks:
            # Generate equally spaced indices for tick placement
            tick_indices = range(0, len(times_list), len(times_list) // num_ticks)
            tick_labels = [times_list[i] for i in tick_indices]
            ax.set_xticks(tick_labels)
            ax.set_xticklabels(tick_labels, rotation=45, ha='right')
        else:
            # If there are fewer data points than num_ticks, show all of them
            ax.set_xticks(times_list)
            ax.set_xticklabels(times_list, rotation=45, ha='right')

        # Embed the plot in the Tkinter window
        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


       # Display the ID matrix horizontally next to the plot
        id_text = '\n'.join([f"• {', '.join(row)}" for row in id_matrix])
        id_label = tk.Label(id_frame, text="Tested Serial Numbers:", font=("Arial", 15, "bold"), bg="#bbe6f6")
        id_label.pack(pady=10)
        id_display = tk.Label(id_frame, text=id_text, font=("Arial", 12, "bold"), bg="#bbe6f6", justify=tk.LEFT)
        id_display.pack(padx=10, pady=10, anchor='w')
     
        # Add a close button
        close_button = tk.Button(id_frame, text="Close", command=plot_window.destroy, bg="#f44336", fg="white", font=("Arial", 12, "bold"))
        close_button.pack(pady=10)
        
        # Add a screenshot button
        screenshot_button = tk.Button(id_frame, text="Screenshot", command=lambda: self.save_screenshot(plot_window), bg="#00b300", fg="white", font=("Arial", 12, "bold"))
        screenshot_button.pack(side=tk.BOTTOM)


    def save_screenshot(self, plot_window):
        plot_window.update_idletasks()
        # Get the plot area and Serial Number frame
        x = plot_window.winfo_rootx()
        y = plot_window.winfo_rooty()
        width = plot_window.winfo_width()
        height = plot_window.winfo_height()

        # Capture the screenshot
        image = ImageGrab.grab((x, y, x + width, y + height))

        # Save the image
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if file_path:
            image.save(file_path)
            messagebox.showinfo("Screenshot", f"Screenshot saved as {file_path}")


    
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()
