# importing pyautocad
import pyautocad
from pyautocad import Autocad, APoint
import math
from itertools import product

acad = pyautocad.Autocad()
doc = acad.app.ActiveDocument
print(acad.doc.Name)
input('Press ENTER to continue')
# Define the layer names
spacing_valuex = 0.35
spacing_valuey = 0.18

target_layers = {
    "A_Point", "B_Point", "C_Point", "D_Point", "E_Point", "F_Point",
    "G_Point", "H_Point", "J_Point", "K_Point", "L_Point", "M_Point",
    "N_Point", "P_Point", "S_Point", "T_Point", "U_Point", "W_Point",
    "X_Point", "A_Depth", "B_Depth", "C_Depth", "D_Depth", "E_Depth",
    "F_Depth", "G_Depth", "H_Depth", "J_Depth", "K_Depth", "L_Depth",
    "M_Depth", "N_Depth", "P_Depth", "S_Depth", "T_Depth", "U_Depth",
    "W_Depth", "X_Depth"
}

target_layers_point = {
    "A_Point", "B_Point", "C_Point", "D_Point", "E_Point", "F_Point",
    "G_Point", "H_Point", "J_Point", "K_Point", "L_Point", "M_Point",
    "N_Point", "P_Point", "S_Point", "T_Point", "U_Point", "W_Point", "X_Point"
}

target_layers_depth = {
    "A_Depth", "B_Depth", "C_Depth", "D_Depth", "E_Depth", "F_Depth",
    "G_Depth", "H_Depth", "J_Depth", "K_Depth", "L_Depth", "M_Depth",
    "N_Depth", "P_Depth", "S_Depth", "T_Depth", "U_Depth", "W_Depth", "X_Depth"
}

# Iterate through text entities
texts = [
    text for text in acad.iter_objects("Text")
    if text.Layer in target_layers_depth
]
alpha = product(texts, repeat=2)
overlapping_found = False
for text1, text2 in alpha:
  # Check if text1 and text2 are not erased
  if text1 is not text2:
    pos1 = text1.InsertionPoint
    pos2 = text2.InsertionPoint
    # Check for overlapping text
    if abs(pos1[0] - pos2[0]) < 0.42 and abs(pos1[1] - pos2[1]) < 0.16:
      print(text1.TextString, " and ", text2.TextString)
      #Adjust the position of text2
      text2.Move(APoint(pos2[0], pos2[1]),
                 APoint(pos2[0], pos2[1] + spacing_valuey))
      overlapping_found = True

# Iterate through text entities
texts = [
    text for text in acad.iter_objects("Text") if text.Layer in target_layers
]
alpha = product(texts, repeat=2)
overlapping_found = False
for text1, text2 in alpha:
  # Check if text1 and text2 are not erased
  if text1 is not text2 and text1.Layer in target_layers_depth and text2.Layer in target_layers_point:
    pos1 = text1.InsertionPoint
    pos2 = text2.InsertionPoint
    # Check for overlapping text
    if abs(pos1[0] - pos2[0]) < 0.38 and abs(pos1[1] - pos2[1]) < 0.18:
      print(text1.TextString, " and ", text2.TextString)
      #Adjust the position of text2
      text2.Move(APoint(pos2[0], pos2[1]),
                 APoint(pos2[0] - spacing_valuex, pos2[1]))
      overlapping_found = True
  #input('Press ENTER to continue')

# Iterate through text entities
texts = [
    text for text in acad.iter_objects("Text") if text.Layer in target_layers
]
alpha = product(texts, repeat=2)
overlapping_found = False
for text1, text2 in alpha:
  # Check if text1 and text2 are not erased
  if text1 is not text2:
    pos1 = text1.InsertionPoint
    pos2 = text2.InsertionPoint
    # Check for overlapping text
    if abs(pos1[0] - pos2[0]) < 0.38 and abs(pos1[1] - pos2[1]) < 0.18:
      print(text1.TextString, " and ", text2.TextString)
      #Adjust the position of text2
      text2.Move(APoint(pos2[0], pos2[1]),
                 APoint(pos2[0], pos2[1] - spacing_valuey))
      overlapping_found = True
#input('Press ENTER to continue')
