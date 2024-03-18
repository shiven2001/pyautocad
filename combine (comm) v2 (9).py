# importing pyautocad
from pyautocad import Autocad, APoint
import comtypes.client
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import asksaveasfilename
import os

# Define the layer names
target_layers = {
    "A_Point", "B_Point", "C_Point", "D_Point", "E_Point", "F_Point",
    "G_Point", "H_Point", "J_Point", "K_Point", "L_Point", "M_Point",
    "N_Point", "P_Point", "S_Point", "T_Point", "U_Point", "W_Point",
    "X_Point", "A_Depth", "B_Depth", "C_Depth", "D_Depth", "E_Depth",
    "F_Depth", "G_Depth", "H_Depth", "J_Depth", "K_Depth", "L_Depth",
    "M_Depth", "N_Depth", "P_Depth", "S_Depth", "T_Depth", "U_Depth",
    "W_Depth", "X_Depth", "A_Dia", "B_Dia", "C_Dia", "D_Dia", "E_Dia", "F_Dia",
    "G_Dia", "H_Dia", "J_Dia", "K_Dia", "L_Dia", "M_Dia", "N_Dia", "P_Dia",
    "S_Dia", "T_Dia", "U_Dia", "W_Dia", "X_Dia", "_Feature", "Defpoints",
    "_UU Boundary", "_DX_DH", "_Cutline", "A_Line", "B_Line", "C_Line",
    "D_Line", "E_Line", "F_Line", "G_Line", "H_Line", "J_Line", "K_Line",
    "L_Line", "M_Line", "N_Line", "P_Line", "S_Line", "T_Line", "U_Line",
    "W_Line", "X_Line", "_Manhole", "_Section", "_Trial Pit"
}

# Create a COM connection to AutoCAD
acad = comtypes.client.GetActiveObject("AutoCAD.Application")

# Make AutoCAD visible (optional)
acad.Visible = True

# Open the file explorer dialog to select an AutoCAD drawings
Tk().withdraw()
drawing_paths = askopenfilenames(title='Select AutoCAD Drawing',
                                 filetypes=[('AutoCAD Drawings', '*.dwg')])

# Open the selected AutoCAD drawings
documents = acad.Documents
drawingslist = []
for path in drawing_paths:
  drawings = acad.Documents.Open(path)
  drawingslist.append(drawings)
# Check if there are any drawings in the list
if len(drawingslist) > 0:
  # Extract the directory path from the original drawing file
  original_path = drawing_paths[0]
  directory = os.path.dirname(original_path)
  # Prompt the user for a new filename
  new_filename = asksaveasfilename(title='Create Full Drawing',
                                   initialdir=directory,
                                   defaultextension='.dwg',
                                   filetypes=[('AutoCAD Drawings', '*.dwg')])
  # Combine the directory path with the new filename
  save_path = os.path.join(directory, new_filename)
  # Save the first drawing as a new drawing
  drawingslist[0].SaveAs(save_path)
  # Open the destination drawing
  destination_drawing = acad.Application.Documents.Open(save_path)
  # Copy model space of other drawings
  for drawing in drawingslist[1:]:
    main_drawing = drawing
    acad.ActiveDocument = main_drawing
    print(main_drawing)
    #Select all entities in the drawing
    source_model_space = main_drawing.ModelSpace
    destination_model_space = destination_drawing.ModelSpace

    # if you get the best interface, you can investigate its properties with 'dir()'
    m = comtypes.client.GetBestInterface(source_model_space)
    handle_string = 'COPYBASE\n'
    handle_string += '0,0,0\n'
    for entity in m:
      if entity.Layer in target_layers:
        handle_string += '(handent "' + entity.Handle + '")\n'
    handle_string += '\n\n'
    acad.ActiveDocument.SendCommand(handle_string)

    # Paste the objects at the same location in the target drawing
    acad.ActiveDocument = destination_drawing
    handle_string = 'PASTECLIP\n'
    handle_string += '0,0,0\n'
    handle_string += '\n\n'
    acad.ActiveDocument.SendCommand(handle_string)

# Close the drawing
