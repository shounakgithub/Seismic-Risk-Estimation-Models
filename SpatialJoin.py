import arcpy
import os

A = arcpy.GetParameterAsText(0) # Import Cluster FC
B = arcpy.GetParameterAsText(1) # Import BFP FC
C = arcpy.GetParameterAsText(2) # Import Surveyed Building FC
D = arcpy.GetParameterAsText(3) # Output Location OF Join OF A with B
E = arcpy.GetParameterAsText(4) # Output Location OF Join OF B with C

newD = D+".shp"
newE = E+".shp"

arcpy.SpatialJoin_analysis(A, B, D,"JOIN_ONE_TO_MANY")
arcpy.SpatialJoin_analysis(B, C, E,"JOIN_ONE_TO_ONE")
arcpy.JoinField_management (newD,"JOIN_FID",newE,"TARGET_FID")








