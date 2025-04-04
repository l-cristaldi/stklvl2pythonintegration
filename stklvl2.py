# %% [markdown]
# Import Required Modules

# %%
import datetime as dt
import numpy as np
import os

from agi.stk12.stkdesktop import STKDesktop
from agi.stk12.stkobjects import *
from agi.stk12.stkutil import *
from agi.stk12.stkx import *
from agi.stk12.utilities.colors import Color, Colors

# %% [markdown]
# Start STK

# %%
uiApp = STKDesktop.StartApplication(visible=True)

# %% [markdown]
# Test

# %%
help(uiApp)

# %%
uiApp.Visible = True
uiApp.UserControl = True

# %%
stkRoot = uiApp.Root

# %%
type(stkRoot)
help(stkRoot)

# %%
stkRoot.NewScenario("IntegrationCertification")
scenario = stkRoot.CurrentScenario

# %%
type(scenario)

# %%
dir(scenario)

# %%
scenario.StartTime = "1 Jun 2022 15:00:00.000"
scenario.StopTime = "2 Jun 2022 15:00:00.000"
stkRoot.Rewind()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 2 ACTIONS REQUIRED ##########

with open("Facilities.txt", "r") as facilityFile:
    for line in facilityFile:
        facilityData = line.strip().split(",")
        
        insertNewFacCmd = f"New / */Facility {facilityData[0]}"
        stkRoot.ExecuteCommand(insertNewFacCmd)
        
########## ACTION 1 : Paste SetPosition Connect command from above on this line ##########
        setPositionCmd = f"SetPosition */Facility/{facilityData[0]} Geodetic {facilityData[2]} {facilityData[1]} Terrain"
        stkRoot.ExecuteCommand(setPositionCmd)
        
########### ACTION 2 : Replace the ? with the Connect command to set the Facility marker and label color to cyan ##########
        setColorCmd = f"Graphics */Facility/{facilityData[0]} SetColor cyan"
        stkRoot.ExecuteCommand(setColorCmd)
        
facilityFile.close()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

satellite = AgSatellite(scenario.Children.New(AgESTKObjectType.eSatellite, "TestSatellite"))

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

#Change some basic display attributes
satelliteBasicGfxAttributes =  satellite.Graphics.Attributes
satelliteBasicGfxAttributes.Color = Color.FromRGB(255,255,0) #Yellow
satelliteBasicGfxAttributes.Line.Width = AgELineWidth.e2 

# # ########## ACTION 1 : Set inheritance of certain 2D graphics settings from the scenario level to false ##########
satelliteBasicGfxAttributes.Inherit = False

satelliteBasicGfxAttributes.IsGroundTrackVisible = False

# %%
#Select Propagator
satellite.SetPropagatorType(AgEVePropagatorType.ePropagatorTwoBody)

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

twoBodyPropagator = satellite.Propagator

########## ACTION 1 : Convert to  the classical orbit state ##########
keplerian = twoBodyPropagator.InitialState.Representation.ConvertTo(AgEOrbitStateType.eOrbitStateClassical)

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 2 ACTIONS REQUIRED ##########
keplerian.SizeShapeType = AgEClassicalSizeShape.eSizeShapeSemimajorAxis

########## ACTION 1 : Set SemiMajorAxis to 7159 km ##########
keplerian.SizeShape.SemiMajorAxis = 7159

########## ACTION 2 : Set Eccentricity to 0 ##########
keplerian.SizeShape.Eccentricity = 0

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

keplerian.Orientation.Inclination = 86.4

########## ACTION 1 : Set argument of perigee to 0 ##########
keplerian.Orientation.ArgOfPerigee = 0

keplerian.Orientation.AscNode.Value = 45

# %%
keplerian.LocationType = AgEClassicalLocation.eLocationTrueAnomaly
keplerian.Location.Value = 45

# %%
twoBodyPropagator.InitialState.Representation.Assign(keplerian)
twoBodyPropagator.Propagate()

# %%
# Remove the test Satellite
satellite.Unload()

#Insert the constellation of Satellites
numOrbitPlanes = 4
numSatsPerPlane = 8
#
for orbitPlaneNum, RAAN in enumerate(range(0,180,180//numOrbitPlanes),1): #RAAN in degrees

    for satNum, trueAnomaly in enumerate(range(0,360,360//numSatsPerPlane), 1): #trueAnomaly in degrees
        
        #Insert satellite
        satellite = AgSatellite(scenario.Children.New(AgESTKObjectType.eSatellite, f"Sat{orbitPlaneNum}{satNum}"))
        
        #Change some basic display attributes
        satelliteBasicGfxAttributes =  satellite.Graphics.Attributes
        satelliteBasicGfxAttributes.Color = Color.FromRGB(255,255,0) #Yellow
        satelliteBasicGfxAttributes.Line.Width = 2     
        satelliteBasicGfxAttributes.Inherit = False
        satelliteBasicGfxAttributes.IsGroundTrackVisible = False
                
        #Select Propagator
        satellite.SetPropagatorType(AgEVePropagatorType.ePropagatorTwoBody)
        
        #Set initial state
        twoBodyPropagator = satellite.Propagator
        keplerian = twoBodyPropagator.InitialState.Representation.ConvertTo(AgEOrbitStateType.eOrbitStateClassical)

        keplerian.SizeShapeType = AgEClassicalSizeShape.eSizeShapeSemimajorAxis
        keplerian.SizeShape.SemiMajorAxis = 7159 #km
        keplerian.SizeShape.Eccentricity = 0

        keplerian.Orientation.Inclination = 86.4 #degrees
        keplerian.Orientation.ArgOfPerigee = 0 #degrees
        keplerian.Orientation.AscNodeType = AgEOrientationAscNode.eAscNodeRAAN
        keplerian.Orientation.AscNode.Value = RAAN  #degrees
        
        keplerian.LocationType = AgEClassicalLocation.eLocationTrueAnomaly
        keplerian.Location.Value = trueAnomaly + (360//numSatsPerPlane/2)*(orbitPlaneNum%2)  #Stagger true anomalies (degrees) for every other orbital plane       
        
        #Propagate
        satellite.Propagator.InitialState.Representation.Assign(keplerian)
        satellite.Propagator.Propagate()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTIONS REQUIRED ##########

########## ACTION 1 : Create a new Constellation Object ##########
sensorConstellation = AgConstellation(scenario.Children.New(AgESTKObjectType.eConstellation,"SensorConstellation"))

#Loop over all satellites
for satellite in scenario.Children.GetElements(AgESTKObjectType.eSatellite):
        
    #Attach sensors to the satellite
    sensor = AgSensor(satellite.Children.New(AgESTKObjectType.eSensor,f"Sensor{satellite.InstanceName[3:]}"))
    
    #Adjust Half Cone Angle
    sensor.CommonTasks.SetPatternSimpleConic(62.5, 2)

    #Adjust the translucency of the sensor projections and the line style
    sensor.VO.PercentTranslucency = 75
    sensor.Graphics.LineStyle = AgELineWidth.e2 
    sensor.Graphics.LineStyle = AgELineStyle.eDotted 
    
    #Add the sensor to the SensorConstellation
    sensorConstellation.Objects.Add(sensor.Path)


# %%
#Create Facility Constellation
facilityConstellation = AgConstellation(scenario.Children.New(AgESTKObjectType.eConstellation, "FacilityConstellation"))

#Loop over each facility
for facility in scenario.Children.GetElements(AgESTKObjectType.eFacility):
    facilityConstellation.Objects.Add(facility.Path)

#Create chain
chain = AgChain(scenario.Children.New(AgESTKObjectType.eChain, "FacsToSensors"))

#Edit some chain graphics properties
chain.Graphics.Animation.Color = Color.FromRGB(0,255,0) #Green
chain.Graphics.Animation.LineWidth = AgELineWidth.e3
chain.Graphics.Animation.IsHighlightVisible = False

#Add objects to chain and compute access
chain.Objects.Add(facilityConstellation.Path)
chain.Objects.Add(sensorConstellation.Path)
chain.ComputeAccess()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

########## ACTION 1 : Replace ? with the scenario stop time ##########
facilityAccess = chain.DataProviders.Item('Object Access').Exec(scenario.StartTime, scenario.StopTime)

# %%
help(facilityAccess)

# %%
print(facilityAccess.Intervals.Count)
print(facilityAccess.Intervals.Item(0).DataSets.GetRow(0))
print(facilityAccess.Intervals.Item(1).DataSets.GetRow(0))

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########
facilityCount = scenario.Children.GetElements(AgESTKObjectType.eFacility).Count

facilityNum = 0
for accessNum in range(facilityAccess.Intervals.Count):
    if 'Fac' in facilityAccess.Intervals.Item(accessNum).DataSets.Item(0).GetValues()[0]:
        facilityNum+=1
        facilityDataSet = facilityAccess.Intervals.Item(accessNum).DataSets
        
        el = facilityDataSet.ElementNames
        
    ########## ACTION 1 : Replace ? with the IAgDrDataSetCollection interface's property that returns the number of rows in the dataset collection ########## 
        numRows = facilityDataSet.RowCount
       
        with open(f"Fac{facilityNum:02}Access.txt", "w") as dataFile:
            dataFile.write(f"{el[0]},{el[2]},{el[3]},{el[4]}\n")
            for row in range(numRows):
                rowData = facilityDataSet.GetRow(row)
                dataFile.write(f"{rowData[0]},{rowData[2]},{rowData[3]},{rowData[4]}\n")        
        dataFile.close()
        
        if facilityNum == 1:
            if os.path.exists("MaxOutageData.txt"):
                open('MaxOutageData.txt', 'w').close()
        
        maxOutage=None
        with open("MaxOutageData.txt", "a") as outageFile:
            
            #If only one row of data, coverage is continuous
            if numRows == 1:
                outageFile.write(f"Fac{facilityNum:02},NA,NA,NA\n")
                print(f"Fac{facilityNum:02}: No Outage")
            
            else:
                #Get StartTimes and StopTimes as lists
                startTimes = list(facilityDataSet.GetDataSetByName("Start Time").GetValues())
                stopTimes = list(facilityDataSet.GetDataSetByName("Stop Time").GetValues())
                
                #convert from strings to datetimes, and create np arrays
                startDatetimes = np.array([dt.datetime.strptime(startTime[:-3], "%d %b %Y %H:%M:%S.%f") for startTime in startTimes])
                stopDatetimes = np.array([dt.datetime.strptime(stopTime[:-3], "%d %b %Y %H:%M:%S.%f") for stopTime in stopTimes])
                
                #Compute outage times
                outages = startDatetimes[1:] - stopDatetimes[:-1]
                
                #Locate max outage and associated start and stop time
                maxOutage = np.amax(outages).total_seconds()
                start = stopTimes[np.argmax(outages)]
                stop = startTimes[np.argmax(outages)+1]
                
                #Write out maxoutage data
                outageFile.write(f"Fac{facilityNum:02},{maxOutage},{start},{stop}\n")
                print(f"Fac{facilityNum:02}: {maxOutage} seconds from {start} until {stop}")
        
        outageFile.close()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

#Get FacTwo object
facTwo = scenario.Children.Item("Fac02")

#Add and configure constraint
facTwoConstraints = facTwo.AccessConstraints
facTwoAzConstraint = facTwoConstraints.AddConstraint(AgEAccessConstraints.eCstrAzimuthAngle)

facTwoAzConstraint.EnableMin = True

########## ACTION 1 : Replace ? with the property that will enable the Max property ##########
facTwoAzConstraint.EnableMax = True

facTwoAzConstraint.Min = 45 #degrees
facTwoAzConstraint.Max = 315 #degrees

# %%
#Compute access
access = facTwo.GetAccess("Satellite/Sat11")
access.ComputeAccess()

# %%
#Get the access data provider
accessDataPrv = access.DataProviders.Item("Access Data").Exec(scenario.StartTime, scenario.StopTime)

#Get Start Time data and print the first access start time
accessStartTimes = accessDataPrv.DataSets.GetDataSetByName("Start Time").GetValues()
print(accessStartTimes[0])

# %%
#Insert aircraft
aircraft = AgAircraft(scenario.Children.New(AgESTKObjectType.eAircraft, "TestAircraft"))

# %%
dir(aircraft)

# %%
#Change AC attitude to Coordinated turn
attitude = aircraft.Attitude
attitude.Basic.SetProfileType(AgEVeProfile.eCoordinatedTurn)

# %%
#Compute aircraft start time
convertUtil = stkRoot.ConversionUtility
aircraftStartTime = convertUtil.NewDate("UTCG",accessStartTimes[0])
aircraftStartTime = aircraftStartTime.Add("min", 30)
print(aircraftStartTime.Format("UTCG"))

# %%
#Load waypoint file
waypoints = np.genfromtxt("FlightPlan.txt", skip_header=1, delimiter=",")
print(waypoints)

# %%
#Set propagtor to GreatArc
aircraft.SetRouteType(AgEVePropagatorType.ePropagatorGreatArc)
route = aircraft.Route

#Set route start time
startEp = aircraft.Route.EphemerisInterval.GetStartEpoch()
startEp.SetExplicitTime(aircraftStartTime.Format("UTCG"))
aircraft.Route.EphemerisInterval.SetStartEpoch(startEp)

#Set the calculation method
aircraft.Route.Method = AgEVeWayPtCompMethod.eDetermineTimeAccFromVel

#Set the altitude reference to MSL
aircraft.Route.SetAltitudeRefType(AgEVeAltitudeRef.eWayPtAltRefMSL)

# %%
#Set unit prefs
stkRoot.UnitPreferences.SetCurrentUnit("DistanceUnit","nm")
stkRoot.UnitPreferences.SetCurrentUnit("TimeUnit","hr")

#Add aircraft waypoints to route
for waypoint in waypoints:
    newWaypoint = route.Waypoints.Add()
    newWaypoint.Latitude = float(waypoint[0]) #degree
    newWaypoint.Longitude = float(waypoint[1]) #degree
    newWaypoint.Altitude = convertUtil.ConvertQuantity("DistanceUnit","ft","nm", waypoint[2]) #ft->nm
    newWaypoint.Speed = waypoint[3] #knots
    newWaypoint.TurnRadius = 1.8 #nautical Miles

#Propagate and reset unit prefs
route.Propagate()
stkRoot.UnitPreferences.ResetUnits()

# %%
#Set graphics properties of the aircraft
aircraftBasicGfxAttributes = aircraft.Graphics.Attributes
aircraftBasicGfxAttributes.Color = Color.FromRGB(255,14,246) #Magenta
aircraftBasicGfxAttributes.Line.Width = AgELineWidth.e3

#Switch to C-130 Model
modelFile = aircraft.VO.Model.ModelData
modelFile.Filename = os.path.abspath("C:/Program Files/AGI/STK 12/STKData/VO/Models/Air/c-130_hercules.glb")

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

#Add aircraft constraint
aircraftConstraints = aircraft.AccessConstraints

########## ACTION 1 : Replace ? with the elevation angle constraint enumeration ##########
elConstraint = aircraftConstraints.AddConstraint(AgEAccessConstraints.eCstrElevationAngle)
elConstraint.EnableMin = True
elConstraint.Min = 10

# %%
#Insert and configure the degraded sensor constellation
degradeSensorConstellation = sensorConstellation.CopyObject("DegradedSensorConstellation")

degradeSensorConstellation.Objects.RemoveName("Satellite/Sat11/Sensor/Sensor11")

# %%
#Insert New Chain
aircraftChain = scenario.Children.New(AgESTKObjectType.eChain, "AcftToSensors")

#Configure chain graphics
aircraftChain.Graphics.Animation.Color = Color.FromRGB(0,255,0) #Green
aircraftChain.Graphics.Animation.LineWidth = AgELineWidth.e3
aircraftChain.Graphics.Animation.IsHighlightVisible = False

#Add objects to chain
aircraftChain.Objects.Add(aircraft.Path)
aircraftChain.Objects.Add(degradeSensorConstellation.Path)
aircraftChain.ComputeAccess()

# %%
########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

########## ACTION 1 : Replace ? with the scenario start time ##########
aircraftAccess = aircraftChain.DataProviders.Item("Complete Access").Exec(scenario.StartTime,scenario.StopTime)


el = aircraftAccess.DataSets.ElementNames
numRows = aircraftAccess.DataSets.RowCount

with open("AircraftAccess.txt", "w") as dataFile:
    dataFile.write(f"{el[0]},{el[1]},{el[2]},{el[3]}\n")
    print(f"{el[0]},{el[1]},{el[2]},{el[3]}")
    
    for row in range(numRows):
        rowData = aircraftAccess.DataSets.GetRow(row)
        dataFile.write(f"{rowData[0]},{rowData[1]},{rowData[2]},{rowData[3]}\n")
        print(f"{rowData[0]},{rowData[1]},{rowData[2]},{rowData[3]}")
        
if numRows == 1:
    print(f"No Outage")

else:
    #Get StartTimes and StopTimes as lists
    startTimes = list(aircraftAccess.DataSets.GetDataSetByName("Start Time").GetValues())
    stopTimes = list(aircraftAccess.DataSets.GetDataSetByName("Stop Time").GetValues())
    
    #convert from strings to datetimes, and create np arrays
    startDatetimes = np.array([dt.datetime.strptime(startTime[:-3], "%d %b %Y %H:%M:%S.%f") for startTime in startTimes])
    stopDatetimes = np.array([dt.datetime.strptime(stopTime[:-3], "%d %b %Y %H:%M:%S.%f") for stopTime in stopTimes])
    
    #Compute outage times
    outages = startDatetimes[1:] - stopDatetimes[:-1]
    
    #Locate max outage and associated start and stop time
    maxOutage = np.amax(outages).total_seconds()
    start = stopTimes[np.argmax(outages)]
    stop = startTimes[np.argmax(outages)+1]
    
    #Write out maxoutage data
    print(f"\nAC Max Outage: {maxOutage} seconds from {start} until {stop}")

# %%
# Get the aircraft LLA State Data Provider
aircraftLLA = aircraft.DataProviders.Item("LLA State")

# %%
#Specify the Fixed Group of the data provider
aircraftLLAFixed = aircraftLLA.Group.Item("Fixed").Exec(scenario.StartTime, scenario.StopTime, 600)

# %%
#Set unit prefs
stkRoot.UnitPreferences.SetCurrentUnit("DistanceUnit","ft")

#Extract desired aircraft LLA data
el = aircraftLLAFixed.DataSets.ElementNames
aircraftLLAFixedRes = np.array(aircraftLLAFixed.DataSets.ToArray())

print(f"{el[0]:30} {el[1]:20} {el[2]:28} {el[11]:15}")
for lla in aircraftLLAFixedRes:
    print(f"{lla[0]:30} {lla[1]:20} {lla[2]:20} {round(float(lla[11])):15}")

#Reset unit prefs
stkRoot.UnitPreferences.ResetUnits()

# %%
#Get aircraft All Postion data provider and print the associated data
facTwoPosData = facTwo.DataProviders.Item("All Position").Exec()
els = facTwoPosData.DataSets.ElementNames
data = facTwoPosData.DataSets.ToArray()[0]
for idx, el in enumerate(els):
    print(f"{el}: {data[idx]}")


