import customtkinter as ctk
import tkinter.messagebox as messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.patches as patches
import ezdxf
import math
import os
import csv
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TCGBlockApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Rectangular TCG Calibration Block Viewer")
        self.geometry("1400x800")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # --- Sidebar ---
        self.sidebar_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Rectangular Block", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 20))
        
        self.inputs = {}
        row_idx = 1
        self.add_input_field("Total Length (L) [mm]", "length", "300", row_idx); row_idx += 1
        self.add_input_field("Total Height (H) [mm]", "height", "50", row_idx); row_idx += 1
        self.add_input_field("Total Width (W) [mm]", "width", "80", row_idx); row_idx += 1
        
        ctk.CTkLabel(self.sidebar_frame, text="SDH (Holes) Spec", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, pady=(15, 5)); row_idx += 1
        self.add_input_field("Hole Depth [mm]", "sdh_depth", "40", row_idx); row_idx += 1
        self.add_input_field("SDH1 X Position (Front)", "sdh_x", "200", row_idx); row_idx += 1
        self.add_input_field("SDH1 Diameter [mm]", "sdh_diameter", "3", row_idx); row_idx += 1
        self.add_input_field("1st Hole Y (bottom)", "sdh_start_y", "10", row_idx); row_idx += 1
        self.add_input_field("Hole Pitch [mm]", "sdh_pitch", "8", row_idx); row_idx += 1
        
        self.add_input_field("SDH2 X Position (Back)", "sdh2_x", "100", row_idx); row_idx += 1
        self.add_input_field("SDH2 Diameter [mm]", "sdh2_diameter", "2.5", row_idx); row_idx += 1
        
        self.flip_var = ctk.BooleanVar(value=True)
        self.chk_flip = ctk.CTkCheckBox(self.sidebar_frame, text="Flip Horizontal (좌우 반전)", variable=self.flip_var, command=self.draw_block)
        self.chk_flip.grid(row=row_idx, column=0, padx=20, pady=10); row_idx += 1
        
        self.show_dim_var = ctk.BooleanVar(value=True)
        self.chk_dim = ctk.CTkCheckBox(self.sidebar_frame, text="Show Dimensions (치수 표시)", variable=self.show_dim_var, command=self.draw_block)
        self.chk_dim.grid(row=row_idx, column=0, padx=20, pady=0); row_idx += 1
        
        self.btn_draw = ctk.CTkButton(self.sidebar_frame, text="Draw / Update", command=self.draw_block)
        self.btn_draw.grid(row=row_idx, column=0, padx=20, pady=20); row_idx += 1
        
        self.btn_dxf = ctk.CTkButton(self.sidebar_frame, text="Export DXF", fg_color="forestgreen", hover_color="darkgreen", command=self.export_dxf)
        self.btn_dxf.grid(row=row_idx, column=0, padx=20, pady=10); row_idx += 1
        
        self.btn_excel = ctk.CTkButton(self.sidebar_frame, text="Export Excel (XLSX)", fg_color="#107c41", hover_color="#0c5e31", command=self.export_excel)
        self.btn_excel.grid(row=row_idx, column=0, padx=20, pady=10); row_idx += 1

        # --- Main Area ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.fig.patch.set_facecolor('#2b2b2b')
        self.ax.set_facecolor('#2b2b2b')
        self.ax.tick_params(colors='white')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.main_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.main_frame)
        self.toolbar.update()
        
        self.draw_block()

    def add_input_field(self, label_text, key, default_val, row):
        frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        frame.grid(row=row, column=0, padx=20, pady=5, sticky="ew")
        lbl = ctk.CTkLabel(frame, text=label_text, width=150, anchor="w")
        lbl.pack(side="left")
        ent = ctk.CTkEntry(frame, width=80)
        ent.insert(0, default_val)
        ent.pack(side="right")
        self.inputs[key] = ent

    def get_values(self):
        try:
            return {
                "L": float(self.inputs["length"].get()),
                "H": float(self.inputs["height"].get()),
                "W": float(self.inputs["width"].get()),
                "sdh_depth": float(self.inputs["sdh_depth"].get()),
                "sdh_x": float(self.inputs["sdh_x"].get()),
                "sdh_start_y": float(self.inputs["sdh_start_y"].get()),
                "sdh_pitch": float(self.inputs["sdh_pitch"].get()),
                "sdh_diameter": float(self.inputs["sdh_diameter"].get()),
                "sdh2_x": float(self.inputs["sdh2_x"].get()),
                "sdh2_diameter": float(self.inputs["sdh2_diameter"].get()),
            }
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values.")
            return None

    def draw_block(self):
        vals = self.get_values()
        if not vals: return
        
        self.ax.clear()
        
        y_bottom = 0.0
        y_top = vals["H"]
        x_min = -vals["L"] if self.flip_var.get() else 0.0
        x_max = 0.0 if self.flip_var.get() else vals["L"]
        
        # 1. Side View Block
        rect_side = patches.Rectangle((x_min, y_bottom), vals["L"], vals["H"], facecolor='#b0c4de', edgecolor='black', lw=2, zorder=2)
        self.ax.add_patch(rect_side)
        
        # 2. SDHs in Side View
        y_first_sdh = y_bottom + vals["sdh_start_y"]
        cx_sdh = -vals["sdh_x"] if self.flip_var.get() else vals["sdh_x"]
        cx2_sdh = -vals["sdh2_x"] if self.flip_var.get() else vals["sdh2_x"]
        
        for i in range(5):
            cy = y_first_sdh + i * vals["sdh_pitch"]
            # Front face holes (solid)
            self.ax.add_patch(patches.Circle((cx_sdh, cy), vals["sdh_diameter"]/2, facecolor='black', edgecolor='none', zorder=3))
            # Back face holes (dashed)
            self.ax.add_patch(patches.Circle((cx2_sdh, cy), vals["sdh2_diameter"]/2, facecolor='#b0c4de', edgecolor='black', linestyle='--', linewidth=1.5, zorder=3))
            
        # 3. Top View Block
        y_top_view = vals["H"] + 40
        rect_top = patches.Rectangle((x_min, y_top_view), vals["L"], vals["W"], facecolor='#aaaaaa', edgecolor='white', lw=1.5)
        self.ax.add_patch(rect_top)
        self.ax.annotate("TOP VIEW", xy=(x_min, y_top_view + vals["W"] + 15), color='white', fontweight='bold', ha='left', va='bottom')
        
        text_bbox = dict(boxstyle='round,pad=0.2', fc='#2b2b2b', ec='none', alpha=0.8)
        depth = vals["sdh_depth"]
        
        # SDHs (Front Face in Top View)
        self.ax.add_patch(patches.Rectangle((cx_sdh - vals["sdh_diameter"]/2, y_top_view), vals["sdh_diameter"], depth, facecolor='yellow', edgecolor='black'))
        self.ax.annotate(f"Front Face\n(Depth {depth:g})", xy=(cx_sdh, y_top_view - 5), color='yellow', ha='center', va='top', fontsize=9, bbox=text_bbox)
        
        # SDHs 2 (Back Face in Top View)
        self.ax.add_patch(patches.Rectangle((cx2_sdh - vals["sdh2_diameter"]/2, y_top_view + vals["W"] - depth), vals["sdh2_diameter"], depth, facecolor='cyan', edgecolor='black'))
        self.ax.annotate(f"Back Face\n(Depth {depth:g})", xy=(cx2_sdh, y_top_view + vals["W"] + 5), color='cyan', ha='center', va='bottom', fontsize=9, bbox=text_bbox)

        # 4. Axes limits and Dimensions
        self.ax.set_aspect('equal')
        self.ax.set_xlim(x_min - 60 if self.flip_var.get() else x_min - 20, x_max + 20 if self.flip_var.get() else x_max + 60)
        self.ax.set_ylim(-20, y_top_view + vals["W"] + 40)
        self.ax.set_title("Rectangular TCG Block Schematic (Side & Top View)", color='white', fontweight='bold')
        self.ax.yaxis.set_visible(False) # Hide Y-axis numbers to prevent overlap
        
        if getattr(self, 'show_dim_var', None) and self.show_dim_var.get():
            def draw_dim(x1, y1, x2, y2, text, text_offset_x=0, text_offset_y=5, ha='center', va='center'):
                self.ax.annotate(text, xy=((x1+x2)/2, (y1+y2)/2), xytext=(text_offset_x, text_offset_y),
                                 textcoords="offset points", ha=ha, va=va, color='white', fontsize=10, fontweight='bold')
                self.ax.annotate('', xy=(x1, y1), xytext=(x2, y2), arrowprops=dict(arrowstyle='<->', color='white', lw=1.5))
                                 
            # L
            draw_dim(x_min, -12, x_max, -12, f'L = {vals["L"]:g}', text_offset_y=-10, va='top')
            # H
            x_h = x_min - 30 if self.flip_var.get() else x_max + 30
            ha_h = 'right' if self.flip_var.get() else 'left'
            draw_dim(x_h, 0, x_h, vals["H"], f'H = {vals["H"]:g}', text_offset_x=-20 if self.flip_var.get() else 20, text_offset_y=0, ha=ha_h)
            # W
            draw_dim(x_h, y_top_view, x_h, y_top_view + vals["W"], f'W = {vals["W"]:g}', text_offset_x=-20 if self.flip_var.get() else 20, text_offset_y=0, ha=ha_h)
            
            # Annotations for holes
            mid_sdh_y = y_first_sdh + 2 * vals["sdh_pitch"]
            top_sdh_y = y_first_sdh + 4 * vals["sdh_pitch"]
            
            self.ax.annotate(f'5-SDH Ø{vals["sdh_diameter"]:g}\n(T > 25.4mm)\n(Front Face, Depth {depth:g})', 
                             xy=(cx_sdh, mid_sdh_y), xytext=(20 if self.flip_var.get() else -20, 0),
                             textcoords='offset points', color='yellow', arrowprops=dict(arrowstyle='->', color='yellow', lw=1.5),
                             ha='left' if self.flip_var.get() else 'right', va='center', fontweight='bold', bbox=text_bbox)
                             
            self.ax.annotate(f'5-SDH Ø{vals["sdh2_diameter"]:g}\n(T ≤ 25.4mm)\n(Back Face, Depth {depth:g})', 
                             xy=(cx2_sdh, top_sdh_y), xytext=(0, 30),
                             textcoords='offset points', color='cyan', arrowprops=dict(arrowstyle='->', color='cyan', lw=1.5),
                             ha='center', va='bottom', fontweight='bold', bbox=text_bbox)
        
        self.canvas.draw()

    def export_dxf(self):
        vals = self.get_values()
        if not vals: return
        filename = "TCG_Rectangular_Block.dxf"
        try:
            doc = ezdxf.new('R2010')
            msp = doc.modelspace()
            
            y_bottom = 0.0
            y_top = vals["H"]
            x_min = -vals["L"] if self.flip_var.get() else 0.0
            x_max = 0.0 if self.flip_var.get() else vals["L"]
            
            doc.layers.add("BLOCK_OUTLINE", color=7)
            doc.layers.add("HOLES", color=1)
            
            msp.add_line((x_min, y_bottom), (x_max, y_bottom), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((x_max, y_bottom), (x_max, y_top), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((x_max, y_top), (x_min, y_top), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((x_min, y_top), (x_min, y_bottom), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            
            cx_sdh = -vals["sdh_x"] if self.flip_var.get() else vals["sdh_x"]
            cx2_sdh = -vals["sdh2_x"] if self.flip_var.get() else vals["sdh2_x"]
            y_first_sdh = y_bottom + vals["sdh_start_y"]
            
            for i in range(5):
                cy = y_first_sdh + i * vals["sdh_pitch"]
                msp.add_circle((cx_sdh, cy), vals["sdh_diameter"]/2, dxfattribs={'layer': 'HOLES'})
                msp.add_circle((cx2_sdh, cy), vals["sdh2_diameter"]/2, dxfattribs={'layer': 'HOLES'})
            
            doc.saveas(filename)
            messagebox.showinfo("Export Successful", f"Saved DXF to:\n{os.path.abspath(filename)}")
            
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred:\n{str(e)}")

    def export_excel(self):
        vals = self.get_values()
        if not vals: return
        filename = "TCG_Rectangular_Block.xlsx"
        img_temp = "temp_plot.png"
        try:
            # Save the current plot as an image
            self.fig.savefig(img_temp, dpi=150, bbox_inches='tight', facecolor=self.fig.get_facecolor())
            
            # Create a new Excel workbook and select the active worksheet
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Block Data"
            
            # Left table: Parameters (Row 1-11)
            params = [
                ["Parameter", "Value", "Unit"],
                ["Total Length (L)", vals["L"], "mm"],
                ["Total Height (H)", vals["H"], "mm"],
                ["Total Width (W)", vals["W"], "mm"],
                ["Hole Depth", vals["sdh_depth"], "mm"],
                ["SDH1 X Position (Front)", vals["sdh_x"], "mm"],
                ["SDH1 Diameter", vals["sdh_diameter"], "mm"],
                ["1st Hole Y (bottom)", vals["sdh_start_y"], "mm"],
                ["Hole Pitch", vals["sdh_pitch"], "mm"],
                ["SDH2 X Position (Back)", vals["sdh2_x"], "mm"],
                ["SDH2 Diameter", vals["sdh2_diameter"], "mm"]
            ]
            
            # Right table: Hole ID (Row 1-11)
            holes = [["Hole ID", "X Position", "Y Position", "Diameter", "Face", "Depth"]]
            cx_sdh = -vals["sdh_x"] if self.flip_var.get() else vals["sdh_x"]
            y_first_sdh = vals["sdh_start_y"]
            for i in range(5):
                cy = y_first_sdh + i * vals["sdh_pitch"]
                holes.append([f"SDH1-{i+1}", cx_sdh, cy, vals["sdh_diameter"], "Front", vals["sdh_depth"]])
                
            cx2_sdh = -vals["sdh2_x"] if self.flip_var.get() else vals["sdh2_x"]
            for i in range(5):
                cy = y_first_sdh + i * vals["sdh_pitch"]
                holes.append([f"SDH2-{i+1}", cx2_sdh, cy, vals["sdh2_diameter"], "Back", vals["sdh_depth"]])
                
            # Now insert them row by row side-by-side
            for r in range(max(len(params), len(holes))):
                row_data = []
                if r < len(params):
                    row_data.extend(params[r])
                else:
                    row_data.extend(["", "", ""])
                    
                row_data.append("") # Empty spacer column D
                
                if r < len(holes):
                    row_data.extend(holes[r])
                    
                ws.append(row_data)
                
            # Adjust column widths for better visibility
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 18
            ws.column_dimensions['C'].width = 18
            ws.column_dimensions['D'].width = 5   # spacer
            ws.column_dimensions['E'].width = 18
            ws.column_dimensions['F'].width = 18
            ws.column_dimensions['G'].width = 18
            ws.column_dimensions['H'].width = 18
            ws.column_dimensions['I'].width = 18
            ws.column_dimensions['J'].width = 18
            
            # Increase font size and center text for all data cells
            large_font = Font(size=18)
            bold_font = Font(size=18, bold=True)
            for row in ws.iter_rows(min_row=1, max_row=11, min_col=1, max_col=10):
                for cell in row:
                    if cell.value is not None:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        if cell.row == 1:
                            cell.font = bold_font
                        else:
                            cell.font = large_font
                
            # Insert the plot image into the Excel sheet below the data
            img = OpenpyxlImage(img_temp)
            # Increase scale to 0.9 to fill more horizontal space
            scale = 0.9
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
            ws.add_image(img, "A14")
            
            # Set up Page Layout for A4 printing (Fit to 1 page wide)
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            
            # Center the content horizontally and vertically on the printed page
            ws.print_options.horizontalCentered = True
            ws.print_options.verticalCentered = True
            
            # Save the workbook
            wb.save(filename)
            
            # Clean up the temporary image
            if os.path.exists(img_temp):
                os.remove(img_temp)
                
            messagebox.showinfo("Export Successful", f"Saved Excel (XLSX) with Image to:\n{os.path.abspath(filename)}")
            
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred:\n{str(e)}")
            if os.path.exists(img_temp):
                os.remove(img_temp)

if __name__ == "__main__":
    app = TCGBlockApp()
    app.mainloop()
