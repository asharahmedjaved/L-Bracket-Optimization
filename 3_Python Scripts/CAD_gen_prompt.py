import win32com.client
import pythoncom
import time
import pySldWrap.sw_tools as sw_tools
import random
import os

# Launch SolidWorks
swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible = True

template = "C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2022\\templates\\Part.prtdot"
save_dir = r"D:\\New Life\\Heavy Projects\\L-Bracket Optimization Project\\1_CAD\\Python CAD"

# Function to create random L-bracket dimensions
def random_l_bracket():
    length = random.randint(40,60)/1000 #length tolerance
    width = random.randint(40, 60)/1000  # Width tolerance
    height = random.randint(40, 60)/1000  # Height tolerance
    thickness = random.randint(5, 10)/1000  # Thickness tolerance
    fillet_radius = random.randint(1, 3)/1000  # Fillet radius tolerance
    hole_radius = random.randint(2, 8)/1000  # Hole radius tolerance

    if fillet_radius > (thickness/2):
        fillet_radius = thickness/2-0.5

    print(length, width, height, thickness, fillet_radius, hole_radius)

    return length, width, height, thickness, fillet_radius, hole_radius

# Generate 10 randomized parts
for i in range(1, 6):
    length, width, height, thickness, fillet_radius, hole_radius = random_l_bracket()
    
    # Create a new part document
    doc = swApp.NewDocument(template, 0, 0, 0)
    if not doc:
        raise RuntimeError("Failed to create a new SolidWorks part.")

    part = swApp.ActiveDoc
    if not part:
        raise RuntimeError("Failed to activate SolidWorks part document.")
    
    # Start a new sketch on the front plane
    sketchMgr = part.SketchManager
    sketchMgr.InsertSketch(True)
    
    # Create randomized L-bracket shape 
    sketchMgr.CreateLine(0, 0, 0, length, 0, 0)
    sketchMgr.CreateLine(length, 0, 0, length, height, 0)
    sketchMgr.CreateLine(length, height, 0, length-thickness, height, 0)
    sketchMgr.CreateLine(length-thickness, height, 0, length-thickness, thickness, 0)
    sketchMgr.CreateLine(length-thickness, thickness, 0, 0, thickness, 0)
    sketchMgr.CreateLine(0, thickness, 0, 0, 0, 0)
    
    # Extrude the active sketch
    features = part.FeatureManager
    # Allow SolidWorks to process
    extrude = features.FeatureExtrusion3(
        True, False, True, 0, 0, width, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, False, 0, 0, False
    )

    if not extrude:
        print(f"❌ Extrusion failed for part {i}.")
        continue
    
    # Apply fillets
    bodies = part.GetBodies2(0, False)
    
    if not bodies:
        raise RuntimeError("No solid bodies found.")

    # Get all edges from the first solid body
    edges = bodies[0].GetEdges()
    if not edges:
        raise RuntimeError("No edges found for fillet operation.")
    
    swSelMgr = part.SelectionManager
    swSelData = swSelMgr.CreateSelectData
    
    time.sleep(1)  # Wait before closing
    # Select all edges
    
    
    
    
    
    for edge in edges:
        edge.Select4(True, swSelData)  # Select edges one by one
        print(edge)

    time.sleep(1)  # Wait before closing
    fillet = features.FeatureFillet3(195, fillet_radius, fillet_radius, 0, 0, 0, 0, 0)

    if fillet:
        print("✅ Fillet applied to all edges!")
    else:
        print("❌ Fillet application failed.")

    # Create holes on Top Plane
        # Select the Top Plane 
    top_plane = part.FeatureByName("Top Plane")
    top_plane.Select2(False, 0)
    sketchMgr.InsertSketch(True)
    
    sketchMgr.CreateCircleByRadius((length-thickness)/2, width/2, 0, hole_radius) 
    
    # Cut through all
    Cut_Extrude_1 = features.FeatureCut4(
        True, False, True, 1, 0, 0, 0, False, False, False, False, 0, 0, False, False, False, False, False, False, False, False, False, False, 0, 0, False, True
        )
    
    if Cut_Extrude_1:
        print("✅ Cut Extrude 1 successful!")
    else:
        print("❌ Cut Extrude 1 application failed.")
    
    # Create holes on Right Plane
        # Select the Right Plane 
    right_plane = part.FeatureByName("Right Plane")
    right_plane.Select2(False, 0)
    sketchMgr.InsertSketch(True)
    
    sketchMgr.CreateCircleByRadius(width/2, (height/2)+thickness, 0, hole_radius)  
    
    # Cut through all
    Cut_Extrude_2 = features.FeatureCut4(
        True, False, True, 9, 0, 0, 0, False, False, False, False, 0, 0, False, False, False, False, False, False, False, False, False, False, 0, 0, False, True
        )
    
    if(Cut_Extrude_2):
        print("✅ Holes created successfully on Top & Right planes!")
    else:
        print("Fuck you Tony!")
    
    # Save part
    part_name = f"L-Bracket_{i}.SLDPRT"
    part.SaveAs(os.path.join(save_dir, part_name))
    part_name_STEP = f"L-Bracket_{i}.STEP"
    part.SaveAs(os.path.join(save_dir, part_name_STEP))
        
    # Release part reference
    part = None
    time.sleep(2)  # Wait before closing
    swApp.CloseDoc(part_name)  # Use CloseDoc instead of part.Close()

    print(f"✅ Part {i} saved successfully as {part_name}!")
    