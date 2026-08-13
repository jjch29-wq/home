import customtkinter as ctk
import tkinter.messagebox as messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.patches as patches
import numpy as np
import ezdxf
import math
import os

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class CalibrationBlockApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("#3 Calibration Block Viewer")
        self.geometry("1400x800")
        
        # Grid layout (1 row, 2 columns)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # --- Sidebar (Input Panel) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Block Dimensions", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 20))
        
        # Inputs dictionary to store CTkEntry objects
        self.inputs = {}
        
        row_idx = 1
        self.add_input_field("Total Length (L) [mm]", "length", "297", row_idx); row_idx += 1
        self.add_input_field("Total Height (H) [mm]", "height", "50", row_idx); row_idx += 1
        self.add_input_field("Top Edge (from 1mm hole Y)", "y_top", "35", row_idx); row_idx += 1
        self.add_input_field("Bevel Angle [deg]", "bevel", "45", row_idx); row_idx += 1
        
        ctk.CTkLabel(self.sidebar_frame, text="SDH (Holes) Spec", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, pady=(15, 5)); row_idx += 1
        self.add_input_field("SDH X Position", "sdh_x", "145", row_idx); row_idx += 1
        self.add_input_field("1st Hole Y (from bottom)", "sdh_start_y", "10", row_idx); row_idx += 1
        self.add_input_field("Hole Pitch [mm]", "sdh_pitch", "8", row_idx); row_idx += 1
        self.add_input_field("Hole Diameter [mm]", "sdh_diameter", "3", row_idx); row_idx += 1
        self.add_input_field("1mm Hole X pos", "hole_1mm_x", "37", row_idx); row_idx += 1
        self.add_input_field("1mm Hole Y pos", "hole_1mm_y", "30", row_idx); row_idx += 1
        
        self.flip_var = ctk.BooleanVar(value=True)
        self.chk_flip = ctk.CTkCheckBox(self.sidebar_frame, text="Flip Horizontal (좌우 반전)", variable=self.flip_var, command=self.draw_block)
        self.chk_flip.grid(row=row_idx, column=0, padx=20, pady=10); row_idx += 1
        
        self.show_dim_var = ctk.BooleanVar(value=True)
        self.chk_dim = ctk.CTkCheckBox(self.sidebar_frame, text="Show Dimensions (치수 표시)", variable=self.show_dim_var, command=self.draw_block)
        self.chk_dim.grid(row=row_idx, column=0, padx=20, pady=0); row_idx += 1
        
        # Buttons
        self.btn_draw = ctk.CTkButton(self.sidebar_frame, text="Draw / Update", command=self.draw_block)
        self.btn_draw.grid(row=row_idx, column=0, padx=20, pady=20); row_idx += 1
        
        self.btn_dxf = ctk.CTkButton(self.sidebar_frame, text="Export DXF", fg_color="forestgreen", hover_color="darkgreen", command=self.export_dxf)
        self.btn_dxf.grid(row=row_idx, column=0, padx=20, pady=10); row_idx += 1

        # --- Main Area (Matplotlib Canvas) ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.fig.patch.set_facecolor('#2b2b2b')
        self.ax.set_facecolor('#2b2b2b')
        self.ax.tick_params(colors='white')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.main_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        
        # Add matplotlib toolbar
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.main_frame)
        self.toolbar.update()
        
        # Initial Draw
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
                "y_top": float(self.inputs["y_top"].get()),
                "bevel": float(self.inputs["bevel"].get()),
                "sdh_x": float(self.inputs["sdh_x"].get()),
                "sdh_start_y": float(self.inputs["sdh_start_y"].get()),
                "sdh_pitch": float(self.inputs["sdh_pitch"].get()),
                "sdh_diameter": float(self.inputs["sdh_diameter"].get()),
                "hole_1mm_x": float(self.inputs["hole_1mm_x"].get()),
                "hole_1mm_y": float(self.inputs["hole_1mm_y"].get())
            }
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values.")
            return None

    def draw_block(self):
        vals = self.get_values()
        if not vals: return
        
        # Center of radii is (0,0) at bottom right
        y_bottom = 0.0
        y_top = vals["H"]
        r30, r50 = 30.0, 50.0
        
        # Left edge is based on total length
        x_left_bottom = r50 - vals["L"]
        bevel_width = vals["H"] * np.tan(np.radians(vals["bevel"]))
        x_min = x_left_bottom + bevel_width
        
        self.ax.clear()
        
        # 1. Back Face (R50, smooth curve)
        theta50 = np.linspace(np.pi/2, 0, 50)
        arc_x_50 = r50 * np.cos(theta50)
        arc_y_50 = r50 * np.sin(theta50)
        
        verts50 = [(x_min, y_top), (0, y_top)]
        for x, y in zip(arc_x_50, arc_y_50): verts50.append((x, y))
        verts50.append((x_left_bottom, y_bottom))
        verts50.append((x_min, y_top))
        
        if self.flip_var.get():
            verts50 = [(-x, y) for x, y in verts50]
        
        poly50 = patches.Polygon(verts50, facecolor='#dddddd', edgecolor='gray', lw=1, zorder=1)
        self.ax.add_patch(poly50)

        # 2. Front Face (R30, has vertical cut)
        theta30 = np.linspace(np.pi/2, 0, 50)
        arc_x_30 = r30 * np.cos(theta30)
        arc_y_30 = r30 * np.sin(theta30)
        
        verts30 = [(x_min, y_top), (0, y_top), (0, r30)]
        for x, y in zip(arc_x_30, arc_y_30): verts30.append((x, y))
        verts30.append((x_left_bottom, y_bottom))
        verts30.append((x_min, y_top))
        
        if self.flip_var.get():
            verts30 = [(-x, y) for x, y in verts30]
        
        poly30 = patches.Polygon(verts30, facecolor='#b0c4de', edgecolor='black', lw=2, zorder=2)
        self.ax.add_patch(poly30)
        
        # 3. Holes
        # 1mm hole (from user input: 20mm from cut, 37mm from bottom)
        cx_hole = vals["hole_1mm_x"] if self.flip_var.get() else -vals["hole_1mm_x"]
        cy_hole = vals["hole_1mm_y"]
        self.ax.add_patch(patches.Circle((cx_hole, cy_hole), 0.5, facecolor='red', edgecolor='darkred', zorder=4))
        
        y_first_sdh = y_bottom + vals["sdh_start_y"]
        for i in range(5):
            cy = y_first_sdh + i * vals["sdh_pitch"]
            cx = vals["sdh_x"] if self.flip_var.get() else -vals["sdh_x"]
            self.ax.add_patch(patches.Circle((cx, cy), vals["sdh_diameter"]/2, facecolor='black', edgecolor='none', zorder=3))

        self.ax.set_aspect('equal')
        if self.flip_var.get():
            self.ax.set_xlim(-80, -x_left_bottom + 20)
        else:
            self.ax.set_xlim(x_left_bottom - 20, 80)
        self.ax.set_ylim(-15, vals["H"] + 15)
        self.ax.set_title("Calibration Block Schematic", color='white', fontweight='bold')
        
        # Draw dimensions if checked
        if getattr(self, 'show_dim_var', None) and self.show_dim_var.get():
            def draw_dim(x1, y1, x2, y2, text, text_offset_x=0, text_offset_y=5, ha='center', va='center'):
                self.ax.annotate(text, xy=((x1+x2)/2, (y1+y2)/2), xytext=(text_offset_x, text_offset_y),
                                 textcoords="offset points", ha=ha, va=va, color='white', fontsize=10, fontweight='bold')
                self.ax.annotate('', xy=(x1, y1), xytext=(x2, y2),
                                 arrowprops=dict(arrowstyle='<->', color='white', lw=1.5))
                                 
            x_min_real = -r50 if self.flip_var.get() else x_left_bottom
            x_max_real = -x_left_bottom if self.flip_var.get() else r50
            
            # L
            draw_dim(x_min_real, -8, x_max_real, -8, f'L = {vals["L"]:g}', text_offset_y=-10, va='top')
            
            # H
            x_h = x_min_real - 10 if self.flip_var.get() else x_max_real + 10
            ha_h = 'right' if self.flip_var.get() else 'left'
            draw_dim(x_h, 0, x_h, vals["H"], f'H = {vals["H"]:g}', text_offset_x=-10 if self.flip_var.get() else 10, text_offset_y=0, ha=ha_h)
            
            # 1mm hole
            self.ax.annotate(f'1mm Hole\n(X:{vals["hole_1mm_x"]:g}, Y:{vals["hole_1mm_y"]:g})', 
                             xy=(cx_hole, cy_hole), xytext=(20 if self.flip_var.get() else -20, 15),
                             textcoords='offset points', color='yellow', arrowprops=dict(arrowstyle='->', color='yellow', lw=1.5),
                             ha='left' if self.flip_var.get() else 'right', fontweight='bold')
                             
            # SDHs
            mid_sdh_y = y_first_sdh + 2 * vals["sdh_pitch"]
            cx_sdh = vals["sdh_x"] if self.flip_var.get() else -vals["sdh_x"]
            self.ax.annotate(f'5-SDH Ø{vals["sdh_diameter"]:g}\nPitch: {vals["sdh_pitch"]:g}\nX: {vals["sdh_x"]:g}', 
                             xy=(cx_sdh, mid_sdh_y), xytext=(20 if self.flip_var.get() else -20, 0),
                             textcoords='offset points', color='yellow', arrowprops=dict(arrowstyle='->', color='yellow', lw=1.5),
                             ha='left' if self.flip_var.get() else 'right', va='center', fontweight='bold')
                             
            # Curves
            self.ax.annotate('R50', xy=(-r50 * 0.707 if self.flip_var.get() else r50 * 0.707, r50 * 0.707), 
                             xytext=(-20 if self.flip_var.get() else 20, 20), textcoords='offset points', color='cyan',
                             arrowprops=dict(arrowstyle='->', color='cyan', lw=1.5), ha='right' if self.flip_var.get() else 'left', fontweight='bold')
            self.ax.annotate('R30', xy=(-r30 * 0.707 if self.flip_var.get() else r30 * 0.707, r30 * 0.707), 
                             xytext=(20 if self.flip_var.get() else -20, -10), textcoords='offset points', color='cyan',
                             arrowprops=dict(arrowstyle='->', color='cyan', lw=1.5), ha='left' if self.flip_var.get() else 'right', fontweight='bold')
        
        self.canvas.draw()

    def export_dxf(self):
        vals = self.get_values()
        if not vals: return
        
        filename = "Calibration_Block_3.dxf"
        try:
            doc = ezdxf.new('R2010')
            msp = doc.modelspace()
            
            y_bottom = 0.0
            y_top = vals["H"]
            r30, r50 = 30.0, 50.0
            
            x_bottom_left = r50 - vals["L"]
            bevel_width = vals["H"] * math.tan(math.radians(vals["bevel"]))
            x_min = x_bottom_left + bevel_width
            
            doc.layers.add("BLOCK_OUTLINE", color=7)
            doc.layers.add("HOLES", color=1)
            
            # Define helper function to flip X if needed
            def fx(x_val):
                return -x_val if self.flip_var.get() else x_val
            
            # Back Face (R50, smooth curve)
            msp.add_line((fx(x_min), y_top), (0, y_top), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            if self.flip_var.get():
                msp.add_arc((0,0), r50, 90.0, 180.0, dxfattribs={'layer': 'BLOCK_OUTLINE'})
            else:
                msp.add_arc((0,0), r50, 0.0, 90.0, dxfattribs={'layer': 'BLOCK_OUTLINE'})
                
            msp.add_line((fx(r50), y_bottom), (fx(x_bottom_left), y_bottom), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((fx(x_bottom_left), y_bottom), (fx(x_min), y_top), dxfattribs={'layer': 'BLOCK_OUTLINE'})

            # Front Face (R30, has vertical cut)
            msp.add_line((fx(x_min), y_top), (0, y_top), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((0, y_top), (0, r30), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            if self.flip_var.get():
                msp.add_arc((0,0), r30, 90.0, 180.0, dxfattribs={'layer': 'BLOCK_OUTLINE'})
            else:
                msp.add_arc((0,0), r30, 0.0, 90.0, dxfattribs={'layer': 'BLOCK_OUTLINE'})
            msp.add_line((fx(r30), y_bottom), (fx(x_bottom_left), y_bottom), dxfattribs={'layer': 'BLOCK_OUTLINE'})
            
            # 1mm Hole
            msp.add_circle((fx(-vals["hole_1mm_x"]), vals["hole_1mm_y"]), 0.5, dxfattribs={'layer': 'HOLES'})
            
            # SDHs
            y_first_sdh = y_bottom + vals["sdh_start_y"]
            for i in range(5):
                cy = y_first_sdh + i * vals["sdh_pitch"]
                msp.add_circle((fx(-vals["sdh_x"]), cy), vals["sdh_diameter"]/2, dxfattribs={'layer': 'HOLES'})
            
            doc.saveas(filename)
            messagebox.showinfo("Export Successful", f"Saved DXF to:\n{os.path.abspath(filename)}")
            
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    app = CalibrationBlockApp()
    app.mainloop()
