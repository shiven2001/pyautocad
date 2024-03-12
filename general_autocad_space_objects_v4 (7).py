# importing pyautocad
import pyautocad
from pyautocad import Autocad, APoint
import math
from itertools import product

acad = pyautocad.Autocad()
active_docs = acad.app.Documents
for i, doc in enumerate(active_docs):
  print(f"{i+1}. {doc.Name}")
selected_index = int(
    input("Enter the index of the document you want to select: ")) - 1
if selected_index < 0 or selected_index >= active_docs.Count:
  print("Invalid index. Exiting...")
  exit()
selected_doc = active_docs[selected_index]
print(f"Selected document: {selected_doc.Name}")
input('Press ENTER to continue')
i = 1

# Define the layer names
target_layers = {
    "A_Point", "B_Point", "C_Point", "D_Point", "E_Point", "F_Point",
    "G_Point", "H_Point", "J_Point", "K_Point", "L_Point", "M_Point",
    "N_Point", "P_Point", "S_Point", "T_Point", "U_Point", "W_Point",
    "X_Point", "A_Depth", "B_Depth", "C_Depth", "D_Depth", "E_Depth",
    "F_Depth", "G_Depth", "H_Depth", "J_Depth", "K_Depth", "L_Depth",
    "M_Depth", "N_Depth", "P_Depth", "S_Depth", "T_Depth", "U_Depth",
    "W_Depth", "X_Depth", "_Feature"
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

"""
#check block refernece and text spacing in point layer
block_refs = [
    block_ref for block_ref in acad.iter_objects("AcDbBlockReference")
    if block_ref.Layer in target_layers_point
]
texts = [
    text for text in acad.iter_objects("Text") if text.Layer in target_layers
]

for block1 in block_refs:
  for text1 in texts:
    pos2 = block1.InsertionPoint
    pos1 = text1.InsertionPoint
    sx = 0
    sy = 0
    textlayer = text1.Layer
    if textlayer in target_layers_depth: 
      if abs(pos1[1] - pos2[1]) < 0.02:
        sy = 0
      elif abs(pos1[1] - pos2[1]) < 0.38:
        sy = 0.20
      if abs(pos1[0] - pos2[0]) < 0.02:
        sx = 0
      elif abs(pos1[0] - pos2[0]) < 0.38:
        sx = 0.45
      
    if abs(pos1[0] - pos2[0]) < 0.38 and abs(pos1[1] - pos2[1]) < 0.38:
        print("Overlapped on manhole: ", text1.TextString)
        if pos2[0] - pos1[0] < 0:
          text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] - sx,
                                                      pos2[1] + sy))
        elif pos2[1] - pos1[1] > 0:
          text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] + sx,
                                                      pos2[1] + sy))
          
        i = 2
      elif i == 2:
        print("Overlapped on manhole, move up: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] - 0.10, pos2[1] + 0.56))
        i = 1
    #point is down
    elif abs(pos1[0] - pos2[0]) < 0.20 and 0.01 > pos1[1] - pos2[1] > -0.20:
      overlapping_found = True
      if i == 1:
        print("Overlapped on manhole, move down: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] + 0.10, pos2[1] - 0.38))
        i = 2
      elif i == 2:
        print("Overlapped on manhole, move down: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] - 0.10, pos2[1] - 0.38))
        i = 1
    #point is left
    elif abs(pos1[1] - pos2[1]) < 0.20 and 0.01 > pos1[0] - pos2[0] > -0.20:
      overlapping_found = True
      if i == 1:
        print("Overlapped on manhole, move left: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] - 0.56, pos2[1] + 0.10))
        i = 2
      elif i == 2:
        print("Overlapped on manhole, move left: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] - 0.56, pos2[1] - 0.10))
        i = 1
    #point is right
    elif abs(pos1[1] - pos2[1]) < 0.20 and 0.01 < pos1[0] - pos2[0] < 0.20:
      overlapping_found = True
      if i == 1:
        print("Overlapped on manhole, move right: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] + 0.38, pos2[1] + 0.10))
        i = 2
      elif i == 2:
        print("Overlapped on manhole, move right: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] + 0.38, pos2[1] - 0.10))
        i = 1
    elif abs(pos1[1] -
             pos2[1]) < 0.01 and 0.38 > abs(pos1[0] - pos2[0]) > 0.01:
      overlapping_found = True
      if pos1[0] - pos2[0] < 0:
        print("Overlapped on manhole, move left: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] - 0.56, pos2[1]))
      elif pos1[0] - pos2[0] > 0:
        print("Overlapped on manhole, move right: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] + 0.38, pos2[1]))
    elif abs(pos1[0] -
             pos2[0]) < 0.01 and 0.38 > abs(pos1[1] - pos2[1]) > 0.01:
      overlapping_found = True
      if pos1[1] - pos2[1] > 0:
        print("Overlapped on manhole, move up: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0], pos2[1] + 0.56))
      elif pos1[1] - pos2[1] < 0:
        print("Overlapped on manhole, move down: ", text1.TextString)
        text1.Move(APoint(pos2[0], pos2[1]),
                   APoint(pos2[0] + 0.38, pos2[1] - 0.38))
    elif abs(pos1[0] - pos2[0]) < 0.01 and abs(pos1[1] - pos2[1]) < 0.01:
      overlapping_found = True
      print("Overlapped on manhole, move up: ", text1.TextString)
      text1.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0], pos2[1] - 0.38))

"""


# text spacing
texts = [
    text for text in acad.iter_objects("Text") if text.Layer in target_layers
]
alpha = product(texts, repeat=2)
for text1, text2 in alpha:
  # Check if text1 and text2 are not erased
  if text1 is not text2:
    pos1 = text1.InsertionPoint
    pos2 = text2.InsertionPoint
    sx = 0
    sy = 0
    if abs(pos1[1] - pos2[1]) < 0.02:
      sy = 0.22
    elif abs(pos1[1] - pos2[1]) < 0.12:
      sy = 0.12
    elif abs(pos1[1] - pos2[1]) < 0.18:
      sy = 0.05

    if abs(pos1[0] - pos2[0]) < 0.02:
      sx = 0.42
    elif abs(pos1[0] - pos2[0]) < 0.25:
      sx = 0.22
    elif abs(pos1[0] - pos2[0]) < 0.42:
      sx = 0.05
    # Check for overlapping text
    if abs(pos1[0] - pos2[0]) < 0.42 and abs(pos1[1] - pos2[1]) < 0.18:
      print(text1.TextString, " and ", text2.TextString)
      #UP RIGHT
      if pos2[0] - pos1[0] > 0 and pos2[1] - pos1[1] > 0:
        text2.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] + sx,
                                                    pos2[1] + sy))
      elif pos2[0] - pos1[0] > 0 and pos2[1] - pos1[1] < 0:
        text2.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] + sx,
                                                    pos2[1] - sy))
      elif pos2[0] - pos1[0] < 0 and pos2[1] - pos1[1] < 0:
        text2.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] - sx,
                                                    pos2[1] - sy))
      elif pos2[0] - pos1[0] < 0 and pos2[1] - pos1[1] > 0:
        text2.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] - sx,
                                                    pos2[1] + sy))
      elif pos2[0] - pos1[0] == 0 and pos2[1] - pos1[1] == 0:
        text2.Move(APoint(pos2[0], pos2[1]), APoint(pos2[0] - sx,
                                                    pos2[1] - sy))
