# ----------------------------------------------------------------------------
# RiskResiliency_Batch.py
# Written by: Jennifer McCall
# Description: Application to run Risk&Resiliency model using StreamStats batch processor output
# ----------------------------------------------------------------------------


from __future__ import division
import os, sys, string, tempfile, errno, shutil, datetime, xlwt, requests, time, zipfile, shutil, streamstats, math
import arcpy
from arcpy import env
from datetime import date, datetime, timedelta
from collections import Counter



arcpy.env.overwriteOutput = True
times1 = list()
times1.append(time.time())

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Risk and Resiliency Minor Culvert Batch Tool"
        self.alias = "Risk and Resiliency Minor Culvert Batch Tool"

        # List of tool classes associated with this toolbox
        self.tools = [RiskResiliencyBatch]

class RiskResiliencyBatch(object):
    def __init__(self):
        
        
        self.label = "RiskResiliencyBatch"
        self.description = "Risk and Resiliency Minor Culvert Batch Tool"
        self.canRunInBackground = False

    def getParameterInfo(self):
        # Toolbox input parameters

        param0 = arcpy.Parameter(
            displayName="Culvert Dataset",
            name="Culverts",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")

        param1 = arcpy.Parameter(
            displayName="Detour Dataset",
            name="Detour",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")

        param2 = arcpy.Parameter(
            displayName="Batch Processor Flows Dataset",
            name="Flows",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")            

        param3 = arcpy.Parameter(
            displayName="Risk and Resiliency Output Folder",
            name="Output Folder",
            datatype="DEworkspace",
            parameterType="Required",
            direction="Input")



        params = [param0, param1, param2, param3]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def execute(self, parameters, messages):
                # R&R tool execution code
        path = parameters[3].valueAsText
        culverts = parameters[0].valueAsText
        detour = parameters[1].valueAsText
        flows = parameters[2].valueAsText
        rowNbr = 0
        count = 0
        latitudeList = []
        longitudeList = []
        slope = ""
        water_snow = "WaterSnow"
        trees = "Trees"
        shrub = "Shrub"
        urban = "Urban"
        landcover = ""
        debrisPotential = ""
        flocCol = []
        objCol = []
        regionCol = []
        landcoverCol = []
        drainCol = []
        slopeCol = []
        debrisCol = []
        qdEvent25Col = []
        qdEvent50Col = []
        qdEvent100Col = []
        qdDesign25Col = []
        qdDesign50Col = []
        qdDesign100Col = []
        qdRatio25Col = []
        qdRatio50Col = []
        qdRatio100Col = []
        ownerRisk25Col = []
        ownerRisk50Col = []
        ownerRisk100Col = []
        totalOwnerRiskCol = []
        userRisk25Col = []
        userRisk50Col = []
        userRisk100Col = []
        totalUserRiskCol = []
        totalRiskCol = []
        equation25Col = []
        culvertsFlows_joined = arcpy.AddJoin_management(flows, "Name",culverts, "FLOC", "KEEP_COMMON")
        culvertName = os.path.basename(culverts)
        culvertBase = os.path.splitext(culvertName)[0]
        flowName = os.path.basename(flows)
        flowBase = os.path.splitext(flowName)[0]



        #Pull needed values from JSON
        def extract_values(obj, key):
            """Pull all values of specified key from nested JSON."""
            arr = []

            def extract(obj, arr, key):
                """Recursively search for values of key in JSON tree."""
                if isinstance(obj, dict):
                    for k, v in obj.items():
                        if isinstance(v, (dict, list)):
                            extract(v, arr, key)
                        elif k == key:
                            arr.append(v)
                elif isinstance(obj, list):
                    for item in obj:
                        extract(item, arr, key)
                return arr

            results = extract(obj, arr, key)
            return results


        def timestamp():
            #Timestamp for printing on output, formatted as m/d h:m:s e.g. 3/4 13:58:24 
            times1.append(time.time())  # float format; datetime.now() is datetime format
            return " ["+"{dt.month}/{dt.day} {dt:%H}:{dt:%M}:{dt:%S}".format(dt=datetime.now())+"] "

        def addMsg(msg, ts=True, warning=False): 
            #Print time-stamped message to results and to log file. 
            #(Can use 2nd parameter to optionally turn off timestamp.)
            if ts: msg = timestamp() + msg
            # arcpy.AddMessage(msg) for running as a Python Toolbox; print(msg) for standalone
            
            arcpy.AddWarning(msg) if warning else arcpy.AddMessage(msg)
            # Print to logFile also if it exists
            # Set log folder to contain log files
            logFldr = path
            logFile = logFldr + "\\RnR_Log_Batch_"+time.strftime("%Y%m%d")+".log"
            if not os.path.exists(logFldr):  
                addMsg("Creating folder to contain script results logs,\n  {0}...".format(logFldr))
                os.makedirs(logFldr)

            try: logFile
            except NameError: logFile_exists = False
            else: logFile_exists = True
            if logFile_exists:
                with open(logFile, "a") as myFile: myFile.write (msg+"\n")

        addMsg("Starting Risk & Resiliency Script")     
        #get lat/long for each culvert
        with arcpy.da.SearchCursor(culvertsFlows_joined, ["{}.FID".format(flowBase),"{}.Latitude".format(flowBase),"{}.Longitude".format(flowBase), "{}.Name".format(flowBase),"{}.RegionID".format(flowBase)]) as cursor:
            for row in cursor:
                objCol.append(row[0])
                latitudeList.append(row[1])
                longitudeList.append(row[2])   
                flocCol.append(row[3])  
                regionCol.append(row[4])
                

        count=0

        culvertCoords = list(zip(latitudeList, longitudeList))

        #runs calculations culvert by culvert
        for coords in culvertCoords:
            
            FLOC = flocCol[count]
            FID = objCol[count]
            regionID = regionCol[count]
            addMsg("Calculating Risk & Resiliency for culvert {} at {}".format(FLOC, coords) ) 
            
            #get dimensions and material of culvert

            

            with arcpy.da.SearchCursor(culvertsFlows_joined, ["{}.BoxHeight_".format(culvertBase),"{}.BoxWidth_I".format(culvertBase), "{}.Diameter_I".format(culvertBase), "{}.Drain_Mate".format(culvertBase), "{}.CulvertLen".format(culvertBase), "{}.AADT".format(culvertBase), "{}.AADTTRUCKS".format(culvertBase),"{}.RegionName".format(flowBase)], where_clause="{}.Name='{}' AND {}.FID={}".format(flowBase,FLOC,flowBase,FID)) as cursor:
                for row in cursor:
                    boxHeight = row[0]
                    boxWidth = row[1]
                    diameterWidth = row[2]
                    drainMaterial = row[3]
                    culvertLength = row[4]
                    aadtTotal = row[5]
                    aadtTruck = row[6]
                    aadtVehicle = aadtTotal - aadtTruck
                    equation25 = row[7]  
            #get detour time and distance
            with arcpy.da.SearchCursor(detour, ["TM_ADDTIME", "TM_ADDDIST"], where_clause="FLOC='{}'".format(FLOC)) as cursor:
                for row in cursor:
                    addTime = row[0]
                    addDist = row[1]        
            
            #Part of QDesign routine
            try:
                addMsg("Box Height is " + str(boxHeight))
            except:
                boxHeight = 0  
                addMsg("Box Height is " + str(boxHeight)) 
            try:
                addMsg("Box width is " + str(boxWidth))
            except:
                boxWidth = 0
                addMsg("Box width is " + str(boxWidth)) 
            try:
                addMsg("Diameter is " + str(diameterWidth))  
            except:
                diameterWidth = 0 
                addMsg("Diameter is " + str(diameterWidth)) 
            try:
                addMsg("Drain Material is " + str(drainMaterial))  
            except:
                drainMaterial = "N/A"   
                addMsg("Drain Material is " + str(drainMaterial)) 
            try:
                addMsg("Culvert Length is " + str(culvertLength))  
            except:
                culvertLength = 0   
                addMsg("Culvert Length is " + str(culvertLength)) 
            try:
                addMsg("AADT Total is " + str(aadtTotal))  
            except:
                aadtTotal = "N/A"   
                addMsg("AADT Total is " + str(aadtTotal)) 
            try:
                addMsg("AADT Truck is " + str(aadtTruck))  
            except:
                aadtTruck = "N/A"   
                addMsg("AADT Truck is " + str(aadtTruck)) 
            try:
                addMsg("AADT Vehicle is " + str(aadtVehicle))  
            except:
                aadtVehicle = "N/A"   
                addMsg("AADT Vehicle is " + str(aadtVehicle)) 
            try:
                addMsg("Equation is " + str(equation25))  
                equation25Col.append(equation25)       
            except:
                equation25 = "N/A"   
                addMsg("Equation is " + str(equation25)) 
                
            
        
            #calculate unit cost for owner risk calculation

            #Find USGS watershed info
            watershedList = []
            watershed = streamstats.watershed.Watershed(coords[0],coords[1])
            watershedList.append(watershed)
            watershedJSON = watershed._delineate()
                #determine drainage area
            drainDict = watershed.get_characteristic(code="DRNAREA")
            drainList=[]
            drainValues = extract_values(drainDict, 'value')
            for i in drainValues:
                drainList.append(i)
            for j in drainList:
                drainCol.append(j)  

            #determine slope of basin    
            slopeDict = watershed.get_characteristic(code="BSLDEM10M")
            slopeList=[]
            slopeValues = extract_values(slopeDict, 'value')
            for i in slopeValues:
                slopeList.append(i)
            for j in slopeList:
                slopePCT = j    
            print(slopePCT)
            

            if float(slopePCT) >= 0 and float(slopePCT) <= 8:
                slope = "Low"
            elif float(slopePCT) >= 9 and float(slopePCT) <= 16:
                slope = "Moderate"    
            elif float(slopePCT) > 16:
                slope = "High"        
            addMsg("slope is {}".format(slope)) 
            slopeCol.append(slope)
            
            # find landcover classes in basin
            waterDict = watershed.get_characteristic(code="LC11WATER")
            waterList=[]
            waterValues = extract_values(waterDict, 'value')
            for i in waterValues:
                waterList.append(i)
            for j in waterList:
                waterPCT = j    
            snowDict = watershed.get_characteristic(code="LC11SNOIC")
            snowList=[]
            snowValues = extract_values(snowDict, 'value')
            for i in snowValues:
                snowList.append(i)
            for j in snowList:
                snowPCT = j    
            water_snowPCT = waterPCT + snowPCT
            shrubDict = watershed.get_characteristic(code="LC11SHRUB")
            shrubList=[]
            shrubValues = extract_values(shrubDict, 'value')
            for i in shrubValues:
                shrubList.append(i)
            for j in shrubList:
                shrubPCT = j  
            treesDict = watershed.get_characteristic(code="LC11FOREST")
            treesList=[]
            treesValues = extract_values(treesDict, 'value')
            for i in treesValues:
                treesList.append(i)
            for j in treesList:
                treesPCT = j    
            urbanDict = watershed.get_characteristic(code="LC11DEV")
            urbanList=[]
            urbanValues = extract_values(urbanDict, 'value')
            for i in urbanValues:
                urbanList.append(i)
            for j in urbanList:
                urbanPCT = j    
            #find majority landcover class
            landList = [water_snowPCT, shrubPCT, treesPCT, urbanPCT]
            landcoverMax = max(landList)  

            if landcoverMax == water_snowPCT:
                landcover = water_snow
            elif landcoverMax == shrubPCT:
                landcover = shrub
            elif landcoverMax == urbanPCT:
                landcover = urban
            elif landcoverMax == treesPCT:
                landcover = trees   
            addMsg("landcover class is {}".format(landcover))        
            landcoverCol.append(landcover)

            if landcover == water_snow:
                debrisPotential = "Very Low"
            if landcover == urban and slope == "Low":
                debrisPotential = "Low"
            if landcover == urban and slope == "Moderate":
                debrisPotential = "Moderate"
            if landcover == urban and slope == "High":
                debrisPotential = "High"    
            if landcover == shrub and slope == "Low":
                debrisPotential = "Moderate"
            if landcover == shrub and slope == "Moderate" and slope == "High":
                debrisPotential = "High"    
            if landcover == trees and slope == "Low":
                debrisPotential = "Moderate"
            if landcover == trees and slope == "Moderate":
                debrisPotential = "High"
            if landcover == trees and slope == "High":
                debrisPotential = "Very High"    
            debrisCol.append(debrisPotential)    

            addMsg("debris potential is {}".format(debrisPotential))   

            equation25 = "N/A No stream in basin"  
            peakFlow25 = "N/A No stream in basin" 
            peakFlow50 = "N/A No stream in basin" 
            peakFlow100 = "N/A No stream in basin" 
                

            with arcpy.da.SearchCursor(culvertsFlows_joined, ["{}.Value".format(flowBase)], where_clause="{}.Name='{}' AND StatName='25 Year Peak Flood'".format(flowBase,FLOC)) as cursor:
                for row in cursor:
                    peakFlow25 = float(row[0])  
            with arcpy.da.SearchCursor(culvertsFlows_joined, ["{}.Value".format(flowBase)], where_clause="{}.Name='{}' AND StatName='50 Year Peak Flood'".format(flowBase,FLOC)) as cursor:
                for row in cursor:
                    peakFlow50 = float(row[0])                    
            with arcpy.da.SearchCursor(culvertsFlows_joined, ["{}.Value".format(flowBase)], where_clause="{}.Name='{}' AND StatName='100 Year Peak Flood'".format(flowBase,FLOC)) as cursor:
                for row in cursor:
                    peakFlow100 = float(row[0])   

                    
            try:
                addMsg("25 year discharge is " + str(peakFlow25)) 
                qdEvent25Col.append(peakFlow25)   
            except:
                peakFlow25 = "N/A"   
                addMsg("25 year discharge is " + str(peakFlow25)) 
            try:
                addMsg("50 year discharge is " + str(peakFlow50))   
                qdEvent50Col.append(peakFlow50)
            except:
                peakFlow50 = "N/A"   
                addMsg("50 year discharge is " + str(peakFlow50))
            try:
                addMsg("100 year discharge is " + str(peakFlow100))  
                qdEvent100Col.append(peakFlow100)    
            except:
                peakFlow100 = "N/A"   
                addMsg("100 year discharge is " + str(peakFlow100))
                
            
            

            #Routine to determine QDesign

            if boxHeight == 0 and diameterWidth == 0 and boxWidth == 0:
                            
                qdDesign25Col.append("N/A no dimensions")  
                qdDesign50Col.append("N/A no dimensions")  
                qdDesign100Col.append("N/A no dimensions")#25-year QDesign Ratio
                qDesignRatio25 = "N/A no dimensions"
                qdRatio25Col.append(qDesignRatio25)  
                qDesignRatio50 = "N/A no dimensions"
                qdRatio50Col.append(qDesignRatio50)  
                qDesignRatio100 = "N/A no dimensions"
                qdRatio100Col.append(qDesignRatio100)  
                addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                addMsg("100 year qDesign ratio is " + str(qDesignRatio100))  

                
                        
                #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                #Start User Risk routine
                #determine QD Val
                

                ownerRisk25Col.append("N/A no dimensions")
                ownerRisk50Col.append("N/A no dimensions")
                ownerRisk100Col.append("N/A no dimensions")
                totalOwnerRiskCol.append("N/A no dimensions")
                userRisk25Col.append("N/A no dimensions")
                userRisk50Col.append("N/A no dimensions")
                userRisk100Col.append("N/A no dimensions")
                totalUserRiskCol.append("N/A no dimensions")
                totalRiskCol.append("N/A no dimensions")

            elif boxHeight != 0 and boxWidth != 0 and diameterWidth == 0:    
                r1 = boxHeight/2
                r2 = boxWidth/2
                print(r2)
                A = (r1*r2)/144
                addMsg("Culvert cross-sectional area is " + str(A))
                Ku = 1
                D = boxHeight/12

                #calculate unit cost for owner risk calculation
                if boxWidth <48:
                    unitCost=2205
                if boxWidth ==48:
                    unitCost=2225 
                if boxWidth ==54:
                    unitCost=2660   
                if boxWidth ==60:
                    unitCost=3135
                if boxWidth ==66:
                    unitCost=3660 
                if boxWidth ==72:
                    unitCost=4235
                if boxWidth ==78:
                    unitCost=4865
                if boxWidth ==84:
                    unitCost=5550 
                if boxWidth ==90:
                    unitCost=10325 
                if boxWidth ==96:
                    unitCost=11690 
                if boxWidth ==102:
                    unitCost=13160
                if boxWidth ==108:
                    unitCost=14770 
                if boxWidth ==120:
                    unitCost=18325
                if boxWidth ==138:
                    unitCost=42695 

                if drainMaterial == 'CON':
                    c = 0.0347
                    Y = 0.81  
                elif drainMaterial == 'PL':
                    c = 0.0347
                    Y = 0.81      
                elif drainMaterial == '':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'ME':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'WO':
                    c = 0.0496
                    Y = 0.57          
                elif drainMaterial == 'OT':
                    c = 0.0496
                    Y = 0.57           

                if float(boxHeight) < 36:
                    hwd = 2.0
                elif float(boxHeight) >= 36 and float(boxHeight) <= 60:
                    hwd = 1.7    
                elif float(boxHeight) > 60 and float(boxHeight) <= 84:
                    hwd = 1.5
                elif float(boxHeight) > 84 and float(boxHeight) <= 120:
                    hwd = 1.2
                elif float(boxHeight) > 120:
                    hwd = 1.00                    

                addMsg("HWD is " + str(hwd))    

                designFlow = math.sqrt((hwd-Y)/c)*A*pow(D,0.5)
                addMsg("Design Flow is " + str(designFlow))
                
                qdDesign25Col.append(designFlow)  
                qdDesign50Col.append(designFlow)  
                qdDesign100Col.append(designFlow)           
            
                if peakFlow25 == "N/A No stream in basin":    
                    qdRatio25Col.append("N/A No stream in basin")  
                        #50-year QDesign Ratio
                    qdRatio50Col.append("N/A No stream in basin")  
                        #100-year QDesign Ratio
                    qdRatio100Col.append("N/A No stream in basin")        
                    ownerRisk25Col.append("N/A No stream in basin")
                    ownerRisk50Col.append("N/A No stream in basin")
                    ownerRisk100Col.append("N/A No stream in basin")
                    totalOwnerRiskCol.append("N/A No stream in basin")
                    userRisk25Col.append("N/A No stream in basin")
                    userRisk50Col.append("N/A No stream in basin")
                    userRisk100Col.append("N/A No stream in basin")
                    totalUserRiskCol.append("N/A No stream in basin")
                    totalRiskCol.append("N/A No stream in basin")
                else: 

                        #25-year QDesign Ratio
                    qDesignRatio25 = peakFlow25/designFlow
                    qdRatio25Col.append(qDesignRatio25)  
                        #50-year QDesign Ratio
                    qDesignRatio50 = peakFlow50/designFlow
                    qdRatio50Col.append(qDesignRatio50)  
                        #100-year QDesign Ratio
                    qDesignRatio100 = peakFlow100/designFlow
                    qdRatio100Col.append(qDesignRatio100)  
                    addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                    addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                    addMsg("100 year qDesign ratio is " + str(qDesignRatio100))  

                            
                    #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                    #Start User Risk routine
                    #determine QD Val
                    
                    if qDesignRatio25 <= 2:
                        qdval25 = "1to2"
                    elif qDesignRatio25 > 2 and qDesignRatio25 <= 3:
                        qdval25 = "2.1to3"   
                    elif qDesignRatio25 > 3 and qDesignRatio25 <= 4:
                        qdval25 = "3.1to4"   
                    elif qDesignRatio25 > 4:
                        qdval25 = "gt4"              
                    else:
                        qdval25 = "N/A"

                    if qDesignRatio50 <= 2:
                        qdval50 = "1to2"
                    elif qDesignRatio50 > 2 and qDesignRatio50 <= 3:
                        qdval50 = "2.1to3"   
                    elif qDesignRatio50 > 3 and qDesignRatio50 <= 4:
                        qdval50 = "3.1to4"   
                    elif qDesignRatio50 > 4:
                        qdval50 = "gt4"              
                    else:
                        qdval50 = "N/A"

                    if qDesignRatio100 <= 2:
                        qdval100 = "1to2"
                    elif qDesignRatio100 > 2 and qDesignRatio100 <= 3:
                        qdval100 = "2.1to3"   
                    elif qDesignRatio100 > 3 and qDesignRatio100 <= 4:
                        qdval100 = "3.1to4"   
                    elif qDesignRatio100 > 4:
                        qdval100 = "gt4"           
                    else:
                        qdval100 = "N/A"    

                    #Routine to determine vulnerable value
                    if qdval25 == "1to2" and debrisPotential == "Very Low":
                        vulnerable25 = 0.25
                    elif qdval25 == "1to2" and debrisPotential == "Low":
                        vulnerable25 = 0.30
                    elif qdval25 == "1to2" and debrisPotential == "Moderate":
                        vulnerable25 = 0.42
                    elif qdval25 == "1to2" and debrisPotential == "High":
                        vulnerable25 = 0.64
                    elif qdval25 == "1to2" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable25 = 0.38
                    elif qdval25 == "2.1to3" and debrisPotential == "Low":
                        vulnerable25 = 0.47
                    elif qdval25 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable25 = 0.64
                    elif qdval25 == "2.1to3" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.89
                    elif qdval25 == "3.1to4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 =="3.1to4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99  
                    else:
                        vulnerable25 = "N/A"   
                        

                    if qdval50 == "1to2" and debrisPotential == "Very Low":
                        vulnerable50 = 0.25
                    elif qdval50 == "1to2" and debrisPotential == "Low":
                        vulnerable50 = 0.30
                    elif qdval50 == "1to2" and debrisPotential == "Moderate":
                        vulnerable50 = 0.42
                    elif qdval50 == "1to2" and debrisPotential == "High":
                        vulnerable50 = 0.64
                    elif qdval50 == "1to2" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable50 = 0.38
                    elif qdval50 == "2.1to3" and debrisPotential == "Low":
                        vulnerable50 = 0.47
                    elif qdval50 =="2.1to3" and debrisPotential == "Moderate":
                        vulnerable50 = 0.64
                    elif qdval50 == "2.1to3" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.89
                    elif qdval50 == "3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.99
                    elif qdval50 =="3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential =="Very High":
                        vulnerable50 = 0.99     
                    else:
                        vulnerable50 = "N/A"   

                    if qdval100 == "1to2" and debrisPotential == "Very Low":
                        vulnerable100 = 0.25
                    elif qdval100 == "1to2" and debrisPotential == "Low":
                        vulnerable100 = 0.30
                    elif qdval100 == "1to2" and debrisPotential == "Moderate":
                        vulnerable100 = 0.42
                    elif qdval100 == "1to2" and debrisPotential == "High":
                        vulnerable100 = 0.64
                    elif qdval100 == "1to2" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable100 = 0.38
                    elif qdval100 == "2.1to3" and debrisPotential == "Low":
                        vulnerable100 = 0.47
                    elif qdval100 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable100 = 0.64
                    elif qdval100 == "2.1to3" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.89
                    elif qdval100 == "3.1to4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99    
                    else:
                        vulnerable100 = "N/A"     

                    addMsg("25 year vulnerability value is {}".format(vulnerable25))       
                    addMsg("50 year vulnerability value is {}".format(vulnerable50))       
                    addMsg("100 year vulnerability value is {}".format(vulnerable100))      

                    #User consequence calculations 

                    vocFC = ((.59*aadtVehicle)+(.96*aadtTruck))*3*addDist
                    addMsg("vocFC is {}".format(vocFC))     
                    lwFC = (10.62*1.77*aadtVehicle+25.31*aadtTruck)*3*(addTime/60)
                    addMsg("lwFC is {}".format(lwFC))     
                    totalUserConsequence = vocFC+lwFC
                    addMsg("Total User Consequence is {}".format(totalUserConsequence))  

                    #Owner + User Risk Calculations
                    if vulnerable25 == "N/A":
                        OwnerRisk25 = "N/A"
                        UserRisk25 = "N/A"
                    else:    
                        OwnerRisk25 = (unitCost*culvertLength+5000)*vulnerable25*0.04
                        UserRisk25 = totalUserConsequence*vulnerable25*0.04
                    addMsg("25 year vulnerability value is ${}".format(OwnerRisk25))  
                    if vulnerable50 == "N/A":
                        OwnerRisk50 = "N/A"
                        UserRisk50 = "N/A"
                    else:
                        OwnerRisk50 = (unitCost*culvertLength+5000)*vulnerable50*0.02
                        UserRisk50 = totalUserConsequence*vulnerable50*0.02
                    addMsg("50 year vulnerability value is ${}".format(OwnerRisk50))  
                    if vulnerable100 == "N/A":
                        OwnerRisk100 = "N/A"
                        UserRisk100 = "N/A"
                    else:
                        OwnerRisk100 = (unitCost*culvertLength+5000)*vulnerable100*0.01
                        UserRisk100 = totalUserConsequence*vulnerable100*0.01
                        addMsg("100 year vulnerability value is ${}".format(OwnerRisk100))  

                    if vulnerable25 == "N/A" or vulnerable50 == "N/A" or vulnerable100 == "N/A":
                        TotalOwnerRisk = "N/A"
                        TotalUserRisk = "N/A"
                        TotalRisk = "N/A"
                    else:
                        TotalOwnerRisk = OwnerRisk25+OwnerRisk50+OwnerRisk100
                        TotalUserRisk = UserRisk25+UserRisk50+UserRisk100
                        TotalRisk = TotalOwnerRisk+TotalUserRisk

                    addMsg("Total Owner Risk is ${}".format(TotalOwnerRisk))  
                    addMsg("Total User Risk is ${}".format(TotalUserRisk))  
                    addMsg("Total Risk is ${}".format(TotalRisk))  

                    ownerRisk25Col.append(OwnerRisk25)
                    ownerRisk50Col.append(OwnerRisk50)
                    ownerRisk100Col.append(OwnerRisk100)
                    totalOwnerRiskCol.append(TotalOwnerRisk)
                    userRisk25Col.append(UserRisk25)
                    userRisk50Col.append(UserRisk50)
                    userRisk100Col.append(UserRisk100)
                    totalUserRiskCol.append(TotalUserRisk)
                    totalRiskCol.append(TotalRisk)

            elif boxHeight == 0 and diameterWidth != 0:    
                r1 = diameterWidth/2
                A = (math.pi*(pow(r1,2)))/144
                addMsg("Culvert cross-sectional area is " + str(A))
                Ku = 1
                D = diameterWidth/12

                #calculate unit cost for owner risk calculation
                if diameterWidth <48:
                    unitCost=2205
                if diameterWidth ==48:
                    unitCost=2225 
                if diameterWidth ==54:
                    unitCost=2660   
                if diameterWidth ==60:
                    unitCost=3135
                if diameterWidth ==66:
                    unitCost=3660 
                if diameterWidth ==72:
                    unitCost=4235
                if diameterWidth ==78:
                    unitCost=4865
                if diameterWidth ==84:
                    unitCost=5550 
                if diameterWidth ==90:
                    unitCost=10325 
                if diameterWidth ==96:
                    unitCost=11690 
                if diameterWidth == 102:
                    unitCost=13160
                if diameterWidth ==108:
                    unitCost=14770 
                if diameterWidth ==120:
                    unitCost=18325
                if diameterWidth ==138:
                    unitCost=42695 

                if drainMaterial == 'CON':
                    c = 0.0317
                    Y = 0.69
                elif drainMaterial == 'PL':
                    c = 0.0317
                    Y = 0.69   
                elif drainMaterial == '':
                    c = 0.0553
                    Y = 0.54         
                elif drainMaterial == 'ME':
                    c = 0.0553
                    Y = 0.54       
                elif drainMaterial == 'WO':
                    c = 0.0553
                    Y = 0.54        
                elif drainMaterial == 'OT':
                    c = 0.0553
                    Y = 0.54         

                if float(diameterWidth) < 36:
                    hwd = 2.0
                elif float(diameterWidth) >= 36 and float(diameterWidth) <= 60:
                    hwd = 1.7    
                elif float(diameterWidth) > 60 and float(diameterWidth) <= 84:
                    hwd = 1.5
                elif float(diameterWidth) > 84 and float(diameterWidth) <= 120:
                    hwd = 1.2
                elif float(diameterWidth) > 120:
                    hwd = 1.00                     

                addMsg("HWD is " + str(hwd))    
                    

                designFlow = math.sqrt((hwd-Y)/c)*A*pow(D,0.5)
                addMsg("Design Flow is " + str(designFlow))
                
                qdDesign25Col.append(designFlow)  
                qdDesign50Col.append(designFlow)  
                qdDesign100Col.append(designFlow)           
            
                if peakFlow25 == "N/A No stream in basin":     
                    qdRatio25Col.append("N/A No stream in basin")  
                        #50-year QDesign Ratio
                    qdRatio50Col.append("N/A No stream in basin")  
                        #100-year QDesign Ratio
                    qdRatio100Col.append("N/A No stream in basin")        
                    ownerRisk25Col.append("N/A No stream in basin")
                    ownerRisk50Col.append("N/A No stream in basin")
                    ownerRisk100Col.append("N/A No stream in basin")
                    totalOwnerRiskCol.append("N/A No stream in basin")
                    userRisk25Col.append("N/A No stream in basin")
                    userRisk50Col.append("N/A No stream in basin")
                    userRisk100Col.append("N/A No stream in basin")
                    totalUserRiskCol.append("N/A No stream in basin")
                    totalRiskCol.append("N/A No stream in basin")      
                else: 

                        #25-year QDesign Ratio
                    qDesignRatio25 = peakFlow25/designFlow
                    qdRatio25Col.append(qDesignRatio25)  
                        #50-year QDesign Ratio
                    qDesignRatio50 = peakFlow50/designFlow
                    qdRatio50Col.append(qDesignRatio50)  
                        #100-year QDesign Ratio
                    qDesignRatio100 = peakFlow100/designFlow
                    qdRatio100Col.append(qDesignRatio100)  
                    addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                    addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                    addMsg("100 year qDesign ratio is " + str(qDesignRatio100))      

                    
                            
                    #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                    #Start User Risk routine
                    #determine QD Val
                    
                    if qDesignRatio25 <= 2:
                        qdval25 = "1to2"
                    elif qDesignRatio25 > 2 and qDesignRatio25 <= 3:
                        qdval25 = "2.1to3"   
                    elif qDesignRatio25 > 3 and qDesignRatio25 <= 4:
                        qdval25 = "3.1to4"   
                    elif qDesignRatio25 > 4:
                        qdval25 = "gt4"              
                    else:
                        qdval25 = "N/A"
                    
                    addMsg("25 qdval is " + str(qdval25))    

                    if qDesignRatio50 <= 2:
                        qdval50 = "1to2"
                    elif qDesignRatio50 > 2 and qDesignRatio50 <= 3:
                        qdval50 = "2.1to3"   
                    elif qDesignRatio50 > 3 and qDesignRatio50 <= 4:
                        qdval50 = "3.1to4"   
                    elif qDesignRatio50 > 4:
                        qdval50 = "gt4"              
                    else:
                        qdval50 = "N/A"

                    if qDesignRatio100 <= 2:
                        qdval100 = "1to2"
                    elif qDesignRatio100 > 2 and qDesignRatio100 <= 3:
                        qdval100 = "2.1to3"   
                    elif qDesignRatio100 > 3 and qDesignRatio100 <= 4:
                        qdval100 = "3.1to4"   
                    elif qDesignRatio100 > 4:
                        qdval100 = "gt4"           
                    else:
                        qdval100 = "N/A"    

                    #Routine to determine vulnerable value
                    if qdval25 == "1to2" and debrisPotential == "Very Low":
                        vulnerable25 = 0.25
                    elif qdval25 == "1to2" and debrisPotential == "Low":
                        vulnerable25 = 0.30
                    elif qdval25 == "1to2" and debrisPotential == "Moderate":
                        vulnerable25 = 0.42
                    elif qdval25 == "1to2" and debrisPotential == "High":
                        vulnerable25 = 0.64
                    elif qdval25 == "1to2" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable25 = 0.38
                    elif qdval25 == "2.1to3" and debrisPotential == "Low":
                        vulnerable25 = 0.47
                    elif qdval25 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable25 = 0.64
                    elif qdval25 == "2.1to3" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.89
                    elif qdval25 == "3.1to4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 =="3.1to4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99  
                    else:
                        vulnerable25 = "N/A"   

                    if qdval50 == "1to2" and debrisPotential == "Very Low":
                        vulnerable50 = 0.25
                    elif qdval50 == "1to2" and debrisPotential == "Low":
                        vulnerable50 = 0.30
                    elif qdval50 == "1to2" and debrisPotential == "Moderate":
                        vulnerable50 = 0.42
                    elif qdval50 == "1to2" and debrisPotential == "High":
                        vulnerable50 = 0.64
                    elif qdval50 == "1to2" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable50 = 0.38
                    elif qdval50 == "2.1to3" and debrisPotential == "Low":
                        vulnerable50 = 0.47
                    elif qdval50 =="2.1to3" and debrisPotential == "Moderate":
                        vulnerable50 = 0.64
                    elif qdval50 == "2.1to3" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.89
                    elif qdval50 == "3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.99
                    elif qdval50 =="3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential =="Very High":
                        vulnerable50 = 0.99     
                    else:
                        vulnerable50 = "N/A"   

                    if qdval100 == "1to2" and debrisPotential == "Very Low":
                        vulnerable100 = 0.25
                    elif qdval100 == "1to2" and debrisPotential == "Low":
                        vulnerable100 = 0.30
                    elif qdval100 == "1to2" and debrisPotential == "Moderate":
                        vulnerable100 = 0.42
                    elif qdval100 == "1to2" and debrisPotential == "High":
                        vulnerable100 = 0.64
                    elif qdval100 == "1to2" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable100 = 0.38
                    elif qdval100 == "2.1to3" and debrisPotential == "Low":
                        vulnerable100 = 0.47
                    elif qdval100 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable100 = 0.64
                    elif qdval100 == "2.1to3" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.89
                    elif qdval100 == "3.1to4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99    
                    else:
                        vulnerable100 = "N/A"     

                    addMsg("25 year vulnerability value is {}".format(vulnerable25))       
                    addMsg("50 year vulnerability value is {}".format(vulnerable50))       
                    addMsg("100 year vulnerability value is {}".format(vulnerable100))      

                    #User consequence calculations 

                    vocFC = ((.59*aadtVehicle)+(.96*aadtTruck))*3*addDist
                    addMsg("vocFC is {}".format(vocFC))     
                    lwFC = (10.62*1.77*aadtVehicle+25.31*aadtTruck)*3*(addTime/60)
                    addMsg("lwFC is {}".format(lwFC))   
                    totalUserConsequence = vocFC+lwFC
                    addMsg("Total User Consequence is {}".format(totalUserConsequence))  

                    #Owner + User Risk Calculations
                    if vulnerable25 == "N/A":
                        OwnerRisk25 = "N/A"
                        UserRisk25 = "N/A"
                    else:    
                        OwnerRisk25 = (unitCost*culvertLength+5000)*vulnerable25*0.04
                        UserRisk25 = totalUserConsequence*vulnerable25*0.04
                    addMsg("25 year vulnerability value is ${}".format(OwnerRisk25))  
                    if vulnerable50 == "N/A":
                        OwnerRisk50 = "N/A"
                        UserRisk50 = "N/A"
                    else:
                        OwnerRisk50 = (unitCost*culvertLength+5000)*vulnerable50*0.02
                        UserRisk50 = totalUserConsequence*vulnerable50*0.02
                    addMsg("50 year vulnerability value is ${}".format(OwnerRisk50))  
                    if vulnerable100 == "N/A":
                        OwnerRisk100 = "N/A"
                        UserRisk100 = "N/A"
                    else:
                        OwnerRisk100 = (unitCost*culvertLength+5000)*vulnerable100*0.01
                        UserRisk100 = totalUserConsequence*vulnerable100*0.01
                        addMsg("100 year vulnerability value is ${}".format(OwnerRisk100))  

                    if vulnerable25 == "N/A" or vulnerable50 == "N/A" or vulnerable100 == "N/A":
                        TotalOwnerRisk = "N/A"
                        TotalUserRisk = "N/A"
                        TotalRisk = "N/A"
                    else:
                        TotalOwnerRisk = OwnerRisk25+OwnerRisk50+OwnerRisk100
                        TotalUserRisk = UserRisk25+UserRisk50+UserRisk100
                        TotalRisk = TotalOwnerRisk+TotalUserRisk

                    addMsg("Total Owner Risk is ${}".format(TotalOwnerRisk))  
                    addMsg("Total User Risk is ${}".format(TotalUserRisk))  
                    addMsg("Total Risk is ${}".format(TotalRisk))  

                    ownerRisk25Col.append(OwnerRisk25)
                    ownerRisk50Col.append(OwnerRisk50)
                    ownerRisk100Col.append(OwnerRisk100)
                    totalOwnerRiskCol.append(TotalOwnerRisk)
                    userRisk25Col.append(UserRisk25)
                    userRisk50Col.append(UserRisk50)
                    userRisk100Col.append(UserRisk100)
                    totalUserRiskCol.append(TotalUserRisk)
                    totalRiskCol.append(TotalRisk)
                
            elif boxHeight != 0 and boxWidth != 0 and diameterWidth == 0:    
                r1 = boxHeight/2
                r2 = boxWidth/2
                print(r2)
                A = (r1*r2)/144
                addMsg("Culvert cross-sectional area is " + str(A))
                Ku = 1
                D = boxHeight/12

                #calculate unit cost for owner risk calculation
                if boxWidth <48:
                    unitCost=2205
                if boxWidth ==48:
                    unitCost=2225 
                if boxWidth ==54:
                    unitCost=2660   
                if boxWidth ==60:
                    unitCost=3135
                if boxWidth ==66:
                    unitCost=3660 
                if boxWidth ==72:
                    unitCost=4235
                if boxWidth ==78:
                    unitCost=4865
                if boxWidth ==84:
                    unitCost=5550 
                if boxWidth ==90:
                    unitCost=10325 
                if boxWidth ==96:
                    unitCost=11690 
                if boxWidth ==102:
                    unitCost=13160
                if boxWidth ==108:
                    unitCost=14770 
                if boxWidth ==120:
                    unitCost=18325
                if boxWidth ==138:
                    unitCost=42695 

                if drainMaterial == 'CON':
                    c = 0.0347
                    Y = 0.81  
                elif drainMaterial == 'PL':
                    c = 0.0347
                    Y = 0.81      
                elif drainMaterial == '':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'ME':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'WO':
                    c = 0.0496
                    Y = 0.57          
                elif drainMaterial == 'OT':
                    c = 0.0496
                    Y = 0.57    

                if float(boxHeight) < 36:
                    hwd = 2.0
                elif float(boxHeight) >= 36 and float(boxHeight) <= 60:
                    hwd = 1.7    
                elif float(boxHeight) > 60 and float(boxHeight) <= 84:
                    hwd = 1.5
                elif float(boxHeight) > 84 and float(boxHeight) <= 120:
                    hwd = 1.2
                elif float(boxHeight) > 120:
                    hwd = 1.00                        

                addMsg("HWD is " + str(hwd))                   

                designFlow = math.sqrt((hwd-Y)/c)*A*pow(D,0.5)
                addMsg("Design Flow is " + str(designFlow))
                
                qdDesign25Col.append(designFlow)  
                qdDesign50Col.append(designFlow)  
                qdDesign100Col.append(designFlow)         
            
                if peakFlow25 == "N/A No stream in basin":    
                    qdRatio25Col.append("N/A No stream in basin")  
                        #50-year QDesign Ratio
                    qdRatio50Col.append("N/A No stream in basin")  
                        #100-year QDesign Ratio
                    qdRatio100Col.append("N/A No stream in basin")        
                    ownerRisk25Col.append("N/A No stream in basin")
                    ownerRisk50Col.append("N/A No stream in basin")
                    ownerRisk100Col.append("N/A No stream in basin")
                    totalOwnerRiskCol.append("N/A No stream in basin")
                    userRisk25Col.append("N/A No stream in basin")
                    userRisk50Col.append("N/A No stream in basin")
                    userRisk100Col.append("N/A No stream in basin")
                    totalUserRiskCol.append("N/A No stream in basin")
                    totalRiskCol.append("N/A No stream in basin")      
                else: 

                        #25-year QDesign Ratio
                    qDesignRatio25 = peakFlow25/designFlow
                    qdRatio25Col.append(qDesignRatio25)  
                        #50-year QDesign Ratio
                    qDesignRatio50 = peakFlow50/designFlow
                    qdRatio50Col.append(qDesignRatio50)  
                        #100-year QDesign Ratio
                    qDesignRatio100 = peakFlow100/designFlow
                    qdRatio100Col.append(qDesignRatio100)  
                    addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                    addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                    addMsg("100 year qDesign ratio is " + str(qDesignRatio100))  

                            
                    #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                    #Start User Risk routine
                    #determine QD Val
                    
                    if qDesignRatio25 <= 2:
                        qdval25 = "1to2"
                    elif qDesignRatio25 > 2 and qDesignRatio25 <= 3:
                        qdval25 = "2.1to3"   
                    elif qDesignRatio25 > 3 and qDesignRatio25 <= 4:
                        qdval25 = "3.1to4"   
                    elif qDesignRatio25 > 4:
                        qdval25 = "gt4"              
                    else:
                        qdval25 = "N/A"

                    if qDesignRatio50 <= 2:
                        qdval50 = "1to2"
                    elif qDesignRatio50 > 2 and qDesignRatio50 <= 3:
                        qdval50 = "2.1to3"   
                    elif qDesignRatio50 > 3 and qDesignRatio50 <= 4:
                        qdval50 = "3.1to4"   
                    elif qDesignRatio50 > 4:
                        qdval50 = "gt4"              
                    else:
                        qdval50 = "N/A"

                    if qDesignRatio100 <= 2:
                        qdval100 = "1to2"
                    elif qDesignRatio100 > 2 and qDesignRatio100 <= 3:
                        qdval100 = "2.1to3"   
                    elif qDesignRatio100 > 3 and qDesignRatio100 <= 4:
                        qdval100 = "3.1to4"   
                    elif qDesignRatio100 > 4:
                        qdval100 = "gt4"           
                    else:
                        qdval100 = "N/A"    

                    #Routine to determine vulnerable value
                    if qdval25 == "1to2" and debrisPotential == "Very Low":
                        vulnerable25 = 0.25
                    elif qdval25 == "1to2" and debrisPotential == "Low":
                        vulnerable25 = 0.30
                    elif qdval25 == "1to2" and debrisPotential == "Moderate":
                        vulnerable25 = 0.42
                    elif qdval25 == "1to2" and debrisPotential == "High":
                        vulnerable25 = 0.64
                    elif qdval25 == "1to2" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable25 = 0.38
                    elif qdval25 == "2.1to3" and debrisPotential == "Low":
                        vulnerable25 = 0.47
                    elif qdval25 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable25 = 0.64
                    elif qdval25 == "2.1to3" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.89
                    elif qdval25 == "3.1to4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 =="3.1to4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99  
                    else:
                        vulnerable25 = "N/A"   

                    if qdval50 == "1to2" and debrisPotential == "Very Low":
                        vulnerable50 = 0.25
                    elif qdval50 == "1to2" and debrisPotential == "Low":
                        vulnerable50 = 0.30
                    elif qdval50 == "1to2" and debrisPotential == "Moderate":
                        vulnerable50 = 0.42
                    elif qdval50 == "1to2" and debrisPotential == "High":
                        vulnerable50 = 0.64
                    elif qdval50 == "1to2" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable50 = 0.38
                    elif qdval50 == "2.1to3" and debrisPotential == "Low":
                        vulnerable50 = 0.47
                    elif qdval50 =="2.1to3" and debrisPotential == "Moderate":
                        vulnerable50 = 0.64
                    elif qdval50 == "2.1to3" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.89
                    elif qdval50 == "3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.99
                    elif qdval50 =="3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential =="Very High":
                        vulnerable50 = 0.99     
                    else:
                        vulnerable50 = "N/A"   

                    if qdval100 == "1to2" and debrisPotential == "Very Low":
                        vulnerable100 = 0.25
                    elif qdval100 == "1to2" and debrisPotential == "Low":
                        vulnerable100 = 0.30
                    elif qdval100 == "1to2" and debrisPotential == "Moderate":
                        vulnerable100 = 0.42
                    elif qdval100 == "1to2" and debrisPotential == "High":
                        vulnerable100 = 0.64
                    elif qdval100 == "1to2" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable100 = 0.38
                    elif qdval100 == "2.1to3" and debrisPotential == "Low":
                        vulnerable100 = 0.47
                    elif qdval100 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable100 = 0.64
                    elif qdval100 == "2.1to3" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.89
                    elif qdval100 == "3.1to4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99    
                    else:
                        vulnerable100 = "N/A"     

                    addMsg("25 year vulnerability value is {}".format(vulnerable25))       
                    addMsg("50 year vulnerability value is {}".format(vulnerable50))       
                    addMsg("100 year vulnerability value is {}".format(vulnerable100))      

                    #User consequence calculations 

                    vocFC = ((.59*aadtVehicle)+(.96*aadtTruck))*3*addDist
                    addMsg("vocFC is {}".format(vocFC))     
                    lwFC = (10.62*1.77*aadtVehicle+25.31*aadtTruck)*3*(addTime/60)
                    addMsg("lwFC is {}".format(lwFC))   
                    totalUserConsequence = vocFC+lwFC
                    addMsg("Total User Consequence is {}".format(totalUserConsequence))  

                    #Owner + User Risk Calculations
                    if vulnerable25 == "N/A":
                        OwnerRisk25 = "N/A"
                        UserRisk25 = "N/A"
                    else:    
                        OwnerRisk25 = (unitCost*culvertLength+5000)*vulnerable25*0.04
                        UserRisk25 = totalUserConsequence*vulnerable25*0.04
                    addMsg("25 year vulnerability value is ${}".format(OwnerRisk25))  
                    if vulnerable50 == "N/A":
                        OwnerRisk50 = "N/A"
                        UserRisk50 = "N/A"
                    else:
                        OwnerRisk50 = (unitCost*culvertLength+5000)*vulnerable50*0.02
                        UserRisk50 = totalUserConsequence*vulnerable50*0.02
                    addMsg("50 year vulnerability value is ${}".format(OwnerRisk50))  
                    if vulnerable100 == "N/A":
                        OwnerRisk100 = "N/A"
                        UserRisk100 = "N/A"
                    else:
                        OwnerRisk100 = (unitCost*culvertLength+5000)*vulnerable100*0.01
                        UserRisk100 = totalUserConsequence*vulnerable100*0.01
                        addMsg("100 year vulnerability value is ${}".format(OwnerRisk100))  

                    if vulnerable25 == "N/A" or vulnerable50 == "N/A" or vulnerable100 == "N/A":
                        TotalOwnerRisk = "N/A"
                        TotalUserRisk = "N/A"
                        TotalRisk = "N/A"
                    else:
                        TotalOwnerRisk = OwnerRisk25+OwnerRisk50+OwnerRisk100
                        TotalUserRisk = UserRisk25+UserRisk50+UserRisk100
                        TotalRisk = TotalOwnerRisk+TotalUserRisk

                    addMsg("Total Owner Risk is ${}".format(TotalOwnerRisk))  
                    addMsg("Total User Risk is ${}".format(TotalUserRisk))  
                    addMsg("Total Risk is ${}".format(TotalRisk))  

                    ownerRisk25Col.append(OwnerRisk25)
                    ownerRisk50Col.append(OwnerRisk50)
                    ownerRisk100Col.append(OwnerRisk100)
                    totalOwnerRiskCol.append(TotalOwnerRisk)
                    userRisk25Col.append(UserRisk25)
                    userRisk50Col.append(UserRisk50)
                    userRisk100Col.append(UserRisk100)
                    totalUserRiskCol.append(TotalUserRisk)
                    totalRiskCol.append(TotalRisk)
                
            elif boxHeight != 0 and diameterWidth == 0:    
                r1 = boxHeight/2
                A = (math.pi*(pow(r1,2)))/144
                addMsg("Culvert cross-sectional area is " + str(A))
                Ku = 1
                D = boxHeight/12

                #calculate unit cost for owner risk calculation
                if boxWidth <48:
                    unitCost=2205
                if boxWidth ==48:
                    unitCost=2225 
                if boxWidth ==54:
                    unitCost=2660   
                if boxWidth ==60:
                    unitCost=3135
                if boxWidth ==66:
                    unitCost=3660 
                if boxWidth ==72:
                    unitCost=4235
                if boxWidth ==78:
                    unitCost=4865
                if boxWidth ==84:
                    unitCost=5550 
                if boxWidth ==90:
                    unitCost=10325 
                if boxWidth ==96:
                    unitCost=11690 
                if boxWidth ==102:
                    unitCost=13160
                if boxWidth ==108:
                    unitCost=14770 
                if boxWidth ==120:
                    unitCost=18325
                if boxWidth ==138:
                    unitCost=42695 

                if drainMaterial == 'CON':
                    c = 0.0347
                    Y = 0.81  
                elif drainMaterial == 'PL':
                    c = 0.0347
                    Y = 0.81      
                elif drainMaterial == '':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'ME':
                    c = 0.0496
                    Y = 0.57         
                elif drainMaterial == 'WO':
                    c = 0.0496
                    Y = 0.57          
                elif drainMaterial == 'OT':
                    c = 0.0496
                    Y = 0.57         

                if float(boxHeight) < 36:
                    hwd = 2.0
                elif float(boxHeight) >= 36 and float(boxHeight) <= 60:
                    hwd = 1.7    
                elif float(boxHeight) > 60 and float(boxHeight) <= 84:
                    hwd = 1.5
                elif float(boxHeight) > 84 and float(boxHeight) <= 120:
                    hwd = 1.2
                elif float(boxHeight) > 120:
                    hwd = 1.00                            

                addMsg("HWD is " + str(hwd))                

                designFlow = math.sqrt((hwd-Y)/c)*A*pow(D,0.5)
                addMsg("Design Flow is " + str(designFlow))
                
                qdDesign25Col.append(designFlow)  
                qdDesign50Col.append(designFlow)  
                qdDesign100Col.append(designFlow)         
            
                if peakFlow25 == "N/A No stream in basin":     
                    qdRatio25Col.append("N/A No stream in basin")  
                        #50-year QDesign Ratio
                    qdRatio50Col.append("N/A No stream in basin")  
                        #100-year QDesign Ratio
                    qdRatio100Col.append("N/A No stream in basin")        
                    ownerRisk25Col.append("N/A No stream in basin")
                    ownerRisk50Col.append("N/A No stream in basin")
                    ownerRisk100Col.append("N/A No stream in basin")
                    totalOwnerRiskCol.append("N/A No stream in basin")
                    userRisk25Col.append("N/A No stream in basin")
                    userRisk50Col.append("N/A No stream in basin")
                    userRisk100Col.append("N/A No stream in basin")
                    totalUserRiskCol.append("N/A No stream in basin")
                    totalRiskCol.append("N/A No stream in basin")       
                else: 

                        #25-year QDesign Ratio
                    qDesignRatio25 = peakFlow25/designFlow
                    qdRatio25Col.append(qDesignRatio25)  
                        #50-year QDesign Ratio
                    qDesignRatio50 = peakFlow50/designFlow
                    qdRatio50Col.append(qDesignRatio50)  
                        #100-year QDesign Ratio
                    qDesignRatio100 = peakFlow100/designFlow
                    qdRatio100Col.append(qDesignRatio100)  
                    addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                    addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                    addMsg("100 year qDesign ratio is " + str(qDesignRatio100))      

                    
                            
                    #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                    #Start User Risk routine
                    #determine QD Val
                    
                    if qDesignRatio25 <= 2:
                        qdval25 = "1to2"
                    elif qDesignRatio25 > 2 and qDesignRatio25 <= 3:
                        qdval25 = "2.1to3"   
                    elif qDesignRatio25 > 3 and qDesignRatio25 <= 4:
                        qdval25 = "3.1to4"   
                    elif qDesignRatio25 > 4:
                        qdval25 = "gt4"              
                    else:
                        qdval25 = "N/A"

                    if qDesignRatio50 <= 2:
                        qdval50 = "1to2"
                    elif qDesignRatio50 > 2 and qDesignRatio50 <= 3:
                        qdval50 = "2.1to3"   
                    elif qDesignRatio50 > 3 and qDesignRatio50 <= 4:
                        qdval50 = "3.1to4"   
                    elif qDesignRatio50 > 4:
                        qdval50 = "gt4"              
                    else:
                        qdval50 = "N/A"

                    if qDesignRatio100 <= 2:
                        qdval100 = "1to2"
                    elif qDesignRatio100 > 2 and qDesignRatio100 <= 3:
                        qdval100 = "2.1to3"   
                    elif qDesignRatio100 > 3 and qDesignRatio100 <= 4:
                        qdval100 = "3.1to4"   
                    elif qDesignRatio100 > 4:
                        qdval100 = "gt4"           
                    else:
                        qdval100 = "N/A"    

                    #Routine to determine vulnerable value
                    if qdval25 == "1to2" and debrisPotential == "Very Low":
                        vulnerable25 = 0.25
                    elif qdval25 == "1to2" and debrisPotential == "Low":
                        vulnerable25 = 0.30
                    elif qdval25 == "1to2" and debrisPotential == "Moderate":
                        vulnerable25 = 0.42
                    elif qdval25 == "1to2" and debrisPotential == "High":
                        vulnerable25 = 0.64
                    elif qdval25 == "1to2" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable25 = 0.38
                    elif qdval25 == "2.1to3" and debrisPotential == "Low":
                        vulnerable25 = 0.47
                    elif qdval25 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable25 = 0.64
                    elif qdval25 == "2.1to3" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.89
                    elif qdval25 == "3.1to4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 =="3.1to4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99  
                    else:
                        vulnerable25 = "N/A"   

                    if qdval50 == "1to2" and debrisPotential == "Very Low":
                        vulnerable50 = 0.25
                    elif qdval50 == "1to2" and debrisPotential == "Low":
                        vulnerable50 = 0.30
                    elif qdval50 == "1to2" and debrisPotential == "Moderate":
                        vulnerable50 = 0.42
                    elif qdval50 == "1to2" and debrisPotential == "High":
                        vulnerable50 = 0.64
                    elif qdval50 == "1to2" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable50 = 0.38
                    elif qdval50 == "2.1to3" and debrisPotential == "Low":
                        vulnerable50 = 0.47
                    elif qdval50 =="2.1to3" and debrisPotential == "Moderate":
                        vulnerable50 = 0.64
                    elif qdval50 == "2.1to3" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.89
                    elif qdval50 == "3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.99
                    elif qdval50 =="3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential =="Very High":
                        vulnerable50 = 0.99     
                    else:
                        vulnerable50 = "N/A"   

                    if qdval100 == "1to2" and debrisPotential == "Very Low":
                        vulnerable100 = 0.25
                    elif qdval100 == "1to2" and debrisPotential == "Low":
                        vulnerable100 = 0.30
                    elif qdval100 == "1to2" and debrisPotential == "Moderate":
                        vulnerable100 = 0.42
                    elif qdval100 == "1to2" and debrisPotential == "High":
                        vulnerable100 = 0.64
                    elif qdval100 == "1to2" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable100 = 0.38
                    elif qdval100 == "2.1to3" and debrisPotential == "Low":
                        vulnerable100 = 0.47
                    elif qdval100 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable100 = 0.64
                    elif qdval100 == "2.1to3" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.89
                    elif qdval100 == "3.1to4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99    
                    else:
                        vulnerable100 = "N/A"     

                    addMsg("25 year vulnerability value is {}".format(vulnerable25))       
                    addMsg("50 year vulnerability value is {}".format(vulnerable50))       
                    addMsg("100 year vulnerability value is {}".format(vulnerable100))      

                    #User consequence calculations 

                    vocFC = ((.59*aadtVehicle)+(.96*aadtTruck))*3*addDist
                    addMsg("vocFC is {}".format(vocFC))     
                    lwFC = (10.62*1.77*aadtVehicle+25.31*aadtTruck)*3*(addTime/60)
                    addMsg("lwFC is {}".format(lwFC))   
                    totalUserConsequence = vocFC+lwFC
                    addMsg("Total User Consequence is {}".format(totalUserConsequence))  

                    #Owner + User Risk Calculations
                    if vulnerable25 == "N/A":
                        OwnerRisk25 = "N/A"
                        UserRisk25 = "N/A"
                    else:    
                        OwnerRisk25 = (unitCost*culvertLength+5000)*vulnerable25*0.04
                        UserRisk25 = totalUserConsequence*vulnerable25*0.04
                    addMsg("25 year vulnerability value is ${}".format(OwnerRisk25))  
                    if vulnerable50 == "N/A":
                        OwnerRisk50 = "N/A"
                        UserRisk50 = "N/A"
                    else:
                        OwnerRisk50 = (unitCost*culvertLength+5000)*vulnerable50*0.02
                        UserRisk50 = totalUserConsequence*vulnerable50*0.02
                    addMsg("50 year vulnerability value is ${}".format(OwnerRisk50))  
                    if vulnerable100 == "N/A":
                        OwnerRisk100 = "N/A"
                        UserRisk100 = "N/A"
                    else:
                        OwnerRisk100 = (unitCost*culvertLength+5000)*vulnerable100*0.01
                        UserRisk100 = totalUserConsequence*vulnerable100*0.01
                        addMsg("100 year vulnerability value is ${}".format(OwnerRisk100))  

                    if vulnerable25 == "N/A" or vulnerable50 == "N/A" or vulnerable100 == "N/A":
                        TotalOwnerRisk = "N/A"
                        TotalUserRisk = "N/A"
                        TotalRisk = "N/A"
                    else:
                        TotalOwnerRisk = OwnerRisk25+OwnerRisk50+OwnerRisk100
                        TotalUserRisk = UserRisk25+UserRisk50+UserRisk100
                        TotalRisk = TotalOwnerRisk+TotalUserRisk

                    addMsg("Total Owner Risk is ${}".format(TotalOwnerRisk))  
                    addMsg("Total User Risk is ${}".format(TotalUserRisk))  
                    addMsg("Total Risk is ${}".format(TotalRisk))  

                    ownerRisk25Col.append(OwnerRisk25)
                    ownerRisk50Col.append(OwnerRisk50)
                    ownerRisk100Col.append(OwnerRisk100)
                    totalOwnerRiskCol.append(TotalOwnerRisk)
                    userRisk25Col.append(UserRisk25)
                    userRisk50Col.append(UserRisk50)
                    userRisk100Col.append(UserRisk100)
                    totalUserRiskCol.append(TotalUserRisk)
                    totalRiskCol.append(TotalRisk)    

            elif boxWidth != 0 and boxHeight != 0 and diameterWidth != 0:    
                r1 = diameterWidth/2
                A = (math.pi*(pow(r1,2)))/144
                addMsg("Culvert cross-sectional area is " + str(A))
                Ku = 1
                D = diameterWidth/12

                #calculate unit cost for owner risk calculation
                if diameterWidth <48:
                    unitCost=2205
                if diameterWidth ==48:
                    unitCost=2225 
                if diameterWidth ==54:
                    unitCost=2660   
                if diameterWidth ==60:
                    unitCost=3135
                if diameterWidth ==66:
                    unitCost=3660 
                if diameterWidth ==72:
                    unitCost=4235
                if diameterWidth ==78:
                    unitCost=4865
                if diameterWidth ==84:
                    unitCost=5550 
                if diameterWidth ==90:
                    unitCost=10325 
                if diameterWidth ==96:
                    unitCost=11690 
                if diameterWidth == 102:
                    unitCost=13160
                if diameterWidth ==108:
                    unitCost=14770 
                if diameterWidth ==120:
                    unitCost=18325
                if diameterWidth ==138:
                    unitCost=42695 

                if drainMaterial == 'CON':
                    c = 0.0317
                    Y = 0.69
                elif drainMaterial == 'PL':
                    c = 0.0317
                    Y = 0.69    
                elif drainMaterial == '':
                    c = 0.0553
                    Y = 0.54         
                elif drainMaterial == 'ME':
                    c = 0.0553
                    Y = 0.54       
                elif drainMaterial == 'WO':
                    c = 0.0553
                    Y = 0.54        
                elif drainMaterial == 'OT':
                    c = 0.0553
                    Y = 0.54         


                if float(diameterWidth) < 36:
                    hwd = 2.0
                elif float(diameterWidth) >= 36 and float(diameterWidth) <= 60:
                    hwd = 1.7    
                elif float(diameterWidth) > 60 and float(diameterWidth) <= 84:
                    hwd = 1.5
                elif float(diameterWidth) > 84 and float(diameterWidth) <= 120:
                    hwd = 1.2
                elif float(diameterWidth) > 120:
                    hwd = 1.00                      

                addMsg("HWD is " + str(hwd))                           

                designFlow = math.sqrt((hwd-Y)/c)*A*pow(D,0.5)
                addMsg("Design Flow is " + str(designFlow))
                
                qdDesign25Col.append(designFlow)  
                qdDesign50Col.append(designFlow)  
                qdDesign100Col.append(designFlow)        
            
                if peakFlow25 == "N/A No stream in basin":       
                    qdRatio25Col.append("N/A No stream in basin")  
                        #50-year QDesign Ratio
                    qdRatio50Col.append("N/A No stream in basin")  
                        #100-year QDesign Ratio
                    qdRatio100Col.append("N/A No stream in basin")        
                    ownerRisk25Col.append("N/A No stream in basin")
                    ownerRisk50Col.append("N/A No stream in basin")
                    ownerRisk100Col.append("N/A No stream in basin")
                    totalOwnerRiskCol.append("N/A No stream in basin")
                    userRisk25Col.append("N/A No stream in basin")
                    userRisk50Col.append("N/A No stream in basin")
                    userRisk100Col.append("N/A No stream in basin")
                    totalUserRiskCol.append("N/A No stream in basin")
                    totalRiskCol.append("N/A No stream in basin")    
                else:  

                        #25-year QDesign Ratio
                    qDesignRatio25 = peakFlow25/designFlow
                    qdRatio25Col.append(qDesignRatio25)  
                        #50-year QDesign Ratio
                    qDesignRatio50 = peakFlow50/designFlow
                    qdRatio50Col.append(qDesignRatio50)  
                        #100-year QDesign Ratio
                    qDesignRatio100 = peakFlow100/designFlow
                    qdRatio100Col.append(qDesignRatio100)  
                    addMsg("25 year qDesign ratio is " + str(qDesignRatio25))
                    addMsg("50 year qDesign ratio is " + str(qDesignRatio50))
                    addMsg("100 year qDesign ratio is " + str(qDesignRatio100))      

                    
                            
                    #=IF(qDesignRatio25<=2,"1to2",IF(qDesignRatio25<=3,"2.1to3",IF(qDesignRatio25<=4,"3.1to4","gt4")))
                    #Start User Risk routine
                    #determine QD Val
                    
                    if qDesignRatio25 <= 2:
                        qdval25 = "1to2"
                    elif qDesignRatio25 > 2 and qDesignRatio25 <= 3:
                        qdval25 = "2.1to3"   
                    elif qDesignRatio25 > 3 and qDesignRatio25 <= 4:
                        qdval25 = "3.1to4"   
                    elif qDesignRatio25 > 4:
                        qdval25 = "gt4"              
                    else:
                        qdval25 = "N/A"

                    if qDesignRatio50 <= 2:
                        qdval50 = "1to2"
                    elif qDesignRatio50 > 2 and qDesignRatio50 <= 3:
                        qdval50 = "2.1to3"   
                    elif qDesignRatio50 > 3 and qDesignRatio50 <= 4:
                        qdval50 = "3.1to4"   
                    elif qDesignRatio50 > 4:
                        qdval50 = "gt4"              
                    else:
                        qdval50 = "N/A"

                    if qDesignRatio100 <= 2:
                        qdval100 = "1to2"
                    elif qDesignRatio100 > 2 and qDesignRatio100 <= 3:
                        qdval100 = "2.1to3"   
                    elif qDesignRatio100 > 3 and qDesignRatio100 <= 4:
                        qdval100 = "3.1to4"   
                    elif qDesignRatio100 > 4:
                        qdval100 = "gt4"           
                    else:
                        qdval100 = "N/A"    

                    #Routine to determine vulnerable value
                    if qdval25 == "1to2" and debrisPotential == "Very Low":
                        vulnerable25 = 0.25
                    elif qdval25 == "1to2" and debrisPotential == "Low":
                        vulnerable25 = 0.30
                    elif qdval25 == "1to2" and debrisPotential == "Moderate":
                        vulnerable25 = 0.42
                    elif qdval25 == "1to2" and debrisPotential == "High":
                        vulnerable25 = 0.64
                    elif qdval25 == "1to2" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable25 = 0.38
                    elif qdval25 == "2.1to3" and debrisPotential == "Low":
                        vulnerable25 = 0.47
                    elif qdval25 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable25 = 0.64
                    elif qdval25 == "2.1to3" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.89
                    elif qdval25 == "3.1to4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "3.1to4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 =="3.1to4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Low":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Moderate":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "High":
                        vulnerable25 = 0.99
                    elif qdval25 == "gt4" and debrisPotential == "Very High":
                        vulnerable25 = 0.99  
                    else:
                        vulnerable25 = "N/A"   

                    if qdval50 == "1to2" and debrisPotential == "Very Low":
                        vulnerable50 = 0.25
                    elif qdval50 == "1to2" and debrisPotential == "Low":
                        vulnerable50 = 0.30
                    elif qdval50 == "1to2" and debrisPotential == "Moderate":
                        vulnerable50 = 0.42
                    elif qdval50 == "1to2" and debrisPotential == "High":
                        vulnerable50 = 0.64
                    elif qdval50 == "1to2" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable50 = 0.38
                    elif qdval50 == "2.1to3" and debrisPotential == "Low":
                        vulnerable50 = 0.47
                    elif qdval50 =="2.1to3" and debrisPotential == "Moderate":
                        vulnerable50 = 0.64
                    elif qdval50 == "2.1to3" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.89
                    elif qdval50 == "3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Very Low":
                        vulnerable50 = 0.99
                    elif qdval50 =="3.1to4" and debrisPotential == "Low":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "Moderate":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential == "High":
                        vulnerable50 = 0.99
                    elif qdval50 == "gt4" and debrisPotential =="Very High":
                        vulnerable50 = 0.99     
                    else:
                        vulnerable50 = "N/A"   

                    if qdval100 == "1to2" and debrisPotential == "Very Low":
                        vulnerable100 = 0.25
                    elif qdval100 == "1to2" and debrisPotential == "Low":
                        vulnerable100 = 0.30
                    elif qdval100 == "1to2" and debrisPotential == "Moderate":
                        vulnerable100 = 0.42
                    elif qdval100 == "1to2" and debrisPotential == "High":
                        vulnerable100 = 0.64
                    elif qdval100 == "1to2" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very Low":
                        vulnerable100 = 0.38
                    elif qdval100 == "2.1to3" and debrisPotential == "Low":
                        vulnerable100 = 0.47
                    elif qdval100 == "2.1to3" and debrisPotential == "Moderate":
                        vulnerable100 = 0.64
                    elif qdval100 == "2.1to3" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "2.1to3" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.89
                    elif qdval100 == "3.1to4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "3.1to4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Low":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Moderate":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "High":
                        vulnerable100 = 0.99
                    elif qdval100 == "gt4" and debrisPotential == "Very High":
                        vulnerable100 = 0.99    
                    else:
                        vulnerable100 = "N/A"     

                    addMsg("25 year vulnerability value is {}".format(vulnerable25))       
                    addMsg("50 year vulnerability value is {}".format(vulnerable50))       
                    addMsg("100 year vulnerability value is {}".format(vulnerable100))      

                    #User consequence calculations 

                    vocFC = ((.59*aadtVehicle)+(.96*aadtTruck))*3*addDist
                    addMsg("vocFC is {}".format(vocFC))     
                    lwFC = (10.62*1.77*aadtVehicle+25.31*aadtTruck)*3*(addTime/60)
                    addMsg("lwFC is {}".format(lwFC))   
                    totalUserConsequence = vocFC+lwFC
                    addMsg("Total User Consequence is {}".format(totalUserConsequence))  

                    #Owner + User Risk Calculations
                    if vulnerable25 == "N/A":
                        OwnerRisk25 = "N/A"
                        UserRisk25 = "N/A"
                    else:    
                        OwnerRisk25 = (unitCost*culvertLength+5000)*vulnerable25*0.04
                        UserRisk25 = totalUserConsequence*vulnerable25*0.04
                    addMsg("25 year vulnerability value is ${}".format(OwnerRisk25))  
                    if vulnerable50 == "N/A":
                        OwnerRisk50 = "N/A"
                        UserRisk50 = "N/A"
                    else:
                        OwnerRisk50 = (unitCost*culvertLength+5000)*vulnerable50*0.02
                        UserRisk50 = totalUserConsequence*vulnerable50*0.02
                    addMsg("50 year vulnerability value is ${}".format(OwnerRisk50))  
                    if vulnerable100 == "N/A":
                        OwnerRisk100 = "N/A"
                        UserRisk100 = "N/A"
                    else:
                        OwnerRisk100 = (unitCost*culvertLength+5000)*vulnerable100*0.01
                        UserRisk100 = totalUserConsequence*vulnerable100*0.01
                        addMsg("100 year vulnerability value is ${}".format(OwnerRisk100))  

                    if vulnerable25 == "N/A" or vulnerable50 == "N/A" or vulnerable100 == "N/A":
                        TotalOwnerRisk = "N/A"
                        TotalUserRisk = "N/A"
                        TotalRisk = "N/A"
                    else:
                        TotalOwnerRisk = OwnerRisk25+OwnerRisk50+OwnerRisk100
                        TotalUserRisk = UserRisk25+UserRisk50+UserRisk100
                        TotalRisk = TotalOwnerRisk+TotalUserRisk

                    addMsg("Total Owner Risk is ${}".format(TotalOwnerRisk))  
                    addMsg("Total User Risk is ${}".format(TotalUserRisk))  
                    addMsg("Total Risk is ${}".format(TotalRisk))  

                    ownerRisk25Col.append(OwnerRisk25)
                    ownerRisk50Col.append(OwnerRisk50)
                    ownerRisk100Col.append(OwnerRisk100)
                    totalOwnerRiskCol.append(TotalOwnerRisk)
                    userRisk25Col.append(UserRisk25)
                    userRisk50Col.append(UserRisk50)
                    userRisk100Col.append(UserRisk100)
                    totalUserRiskCol.append(TotalUserRisk)
                    totalRiskCol.append(TotalRisk)



            count=count+1


        def write_culvertExcel(ingXLS, ingBodyList, printMsgs=True):

                # Creates ingest Excel file from header list and body list data 
                # Start an xls file and populate header rows
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Culvert_Risk_Resiliency')

                #begin formatting excel
            dateMDY = xlwt.easyxf(num_format_str='mm/dd/yyyy')
            colWidthIdxList = []
            colWidthMin = 35
            firstRowFmt = xlwt.easyxf('align: wrap on, vert top;font:bold 1,height 180;')
            secondRowFmt = xlwt.easyxf('align: wrap on, horz center, vert center;')
            bodyFmt = xlwt.easyxf('pattern: pattern solid, fore_colour orange; align: wrap on, horz center, vert center;font:bold 1,height 210;')
            bodyHdrStartRow1 = 1
            bodyHdrStartRow2 = 1
            bodyHdrStartRow3 = 1   
            bodyHdrStartRow4 = 1  
            bodyHdrStartRow5 = 1  
            bodyHdrStartRow6 = 1 
            bodyHdrStartRow7 = 1
            bodyHdrStartRow8 = 1   
            bodyHdrStartRow9 = 1  
            bodyHdrStartRow10 = 1  
            bodyHdrStartRow11 = 1  
            bodyHdrStartRow12 = 1  
            bodyHdrStartRow13 = 1   
            bodyHdrStartRow14 = 1    
            bodyHdrStartRow15 = 1  
            bodyHdrStartRow16 = 1  
            bodyHdrStartRow17= 1   
            bodyHdrStartRow18 = 1   
            bodyHdrStartRow19 = 1    
            bodyHdrStartRow20 = 1  
            bodyHdrStartRow21 = 1  
            bodyHdrStartRow22 = 1  

            #write excel
            for bodyRow in ingBodyList:
                colWidth = colWidthMin
                if len(bodyRow) == 1:
                    ws.write(rowNbr, 0, bodyRow,bodyFmt)
                else:
                    
                    for colNbr, bodyTxt in enumerate(bodyRow):
                        ws.write(rowNbr, colNbr, bodyTxt,bodyFmt)
                for x in range(len(flocCol)):
                    ws.write(bodyHdrStartRow1, 0, flocCol[x],secondRowFmt)  
                    addMsg("Writing excel record for culvert")
                    bodyHdrStartRow1 += 1
                for x in range(len(slopeCol)):
                    ws.write(bodyHdrStartRow2, 1,slopeCol[x],secondRowFmt)  
                    bodyHdrStartRow2 += 1 
                for x in range(len(landcoverCol)):
                    ws.write(bodyHdrStartRow3, 2,landcoverCol[x],secondRowFmt)  
                    bodyHdrStartRow3 += 1    
                for x in range(len(debrisCol)):
                    ws.write(bodyHdrStartRow4, 3, debrisCol[x],secondRowFmt)  
                    bodyHdrStartRow4 += 1
                for x in range(len(equation25Col)): 
                    ws.write(bodyHdrStartRow5, 4, equation25Col[x],secondRowFmt)  
                    bodyHdrStartRow5 += 1
                for x in range(len(qdEvent25Col)): 
                    ws.write(bodyHdrStartRow6, 5, qdEvent25Col[x],secondRowFmt)  
                    bodyHdrStartRow6 += 1
                for x in range(len(qdEvent50Col)):
                    ws.write(bodyHdrStartRow7, 6, qdEvent50Col[x],secondRowFmt)  
                    bodyHdrStartRow7 += 1
                for x in range(len(qdEvent100Col)):
                    ws.write(bodyHdrStartRow8, 7, qdEvent100Col[x],secondRowFmt)  
                    bodyHdrStartRow8 += 1
                for x in range(len(qdDesign25Col)):
                    ws.write(bodyHdrStartRow9, 8, qdDesign25Col[x],secondRowFmt)  
                    bodyHdrStartRow9 += 1
                for x in range(len(qdRatio25Col)):
                    ws.write(bodyHdrStartRow10, 9, qdRatio25Col[x],secondRowFmt)  
                    bodyHdrStartRow10 += 1
                for x in range(len(qdRatio50Col)):
                    ws.write(bodyHdrStartRow11, 10, qdRatio50Col[x],secondRowFmt)  
                    bodyHdrStartRow11 += 1
                for x in range(len(qdRatio100Col)):
                    ws.write(bodyHdrStartRow12, 11, qdRatio100Col[x],secondRowFmt)  
                    bodyHdrStartRow12 += 1
                for x in range(len(drainCol)):
                    ws.write(bodyHdrStartRow13, 12, drainCol[x],secondRowFmt)  
                    bodyHdrStartRow13 += 1
                for x in range(len(ownerRisk25Col)):
                    ws.write(bodyHdrStartRow14, 13, ownerRisk25Col[x],secondRowFmt)  
                    bodyHdrStartRow14 += 1
                for x in range(len(ownerRisk50Col)):
                    ws.write(bodyHdrStartRow15, 14, ownerRisk50Col[x],secondRowFmt)  
                    bodyHdrStartRow15 += 1
                for x in range(len(ownerRisk100Col)):
                    ws.write(bodyHdrStartRow16, 15, ownerRisk100Col[x],secondRowFmt)  
                    bodyHdrStartRow16 += 1
                for x in range(len(totalOwnerRiskCol)):
                    ws.write(bodyHdrStartRow17, 16, totalOwnerRiskCol[x],secondRowFmt)  
                    bodyHdrStartRow17 += 1
                for x in range(len(userRisk25Col)):
                    ws.write(bodyHdrStartRow18, 17, userRisk25Col[x],secondRowFmt)  
                    bodyHdrStartRow18 += 1
                for x in range(len(userRisk50Col)):
                    ws.write(bodyHdrStartRow19, 18, userRisk50Col[x],secondRowFmt)  
                    bodyHdrStartRow19 += 1
                for x in range(len(userRisk100Col)):
                    ws.write(bodyHdrStartRow20, 19, userRisk100Col[x],secondRowFmt)  
                    bodyHdrStartRow20 += 1
                for x in range(len(totalUserRiskCol)):
                    ws.write(bodyHdrStartRow21, 20, totalUserRiskCol[x],secondRowFmt)  
                    bodyHdrStartRow21 += 1
                for x in range(len(totalRiskCol)):
                    ws.write(bodyHdrStartRow22, 21, totalRiskCol[x],secondRowFmt)  
                    bodyHdrStartRow22 += 1
                
                        

                colWidthIdxList.append(colWidth)

            # Save the Excel file 
            wb.save(ingXLS)

        #Creates excel spreadsheet
        def create_culvertExcel(mainPath):

            timeStart = times1[-1]
            outFolder = "\\Out"
            #Output path and file name
            XlsPath = os.path.join(path,  "CulvertRiskResiliency" + "_" + "Batch" + "_" + datetime.now().strftime("%Y%m%d-%H%M%S") +".xls")
            #Column headings
            bodyList = [['FLOC', 'Slope', 'Landcover', 'Debris_Potential', 'Equation', 'qEvent_25Year', 'qEvent_50Year', 'qEvent_100Year', 'Design_Flow', 
            'qEventDesignRatio_25Year', 'qEventDesignRatio_50Year', 'qEventDesignRatio_100Year', 'Drainage_Area', '25-Year Owner Risk', '50-Year Owner Risk', '100-Year Owner Risk', 'Total Owner Risk', '25-Year User Risk', '50-Year User Risk', '100-Year User Risk', 'Total User Risk', 'Total Annual Risk']]
            write_culvertExcel(XlsPath, bodyList)
            addMsg("Excel output created")       


        create_culvertExcel(culverts)
