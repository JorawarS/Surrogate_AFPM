Sub square(gap, mat, pm_thic, pm_angle)


Dim oMN As Object, oCst As Object
Set oMN = CreateObject("Magnet.application")
Set oCst = oMN.getConstants 'Get the Constants object

oMN.Visible = False 'Make the MagNet window visible

'Setup Defaults
Call oMN.newDocument
Call oMN.getDocument().beginUndoGroup("Set Default Units", True)
Call oMN.getDocument().setDefaultLengthUnit("Millimeters")
Call oMN.getDocument().endUndoGroup

'OUTER AIRBOX
Call oMN.getDocument().getView().newLine(0, 0, 720, 0)
Call oMN.getDocument().getView().newLine(720, 0, 0, 0)
Call oMN.getDocument().getView().newLine(0, 0, 0, 720)
Call oMN.getDocument().getView().newArc(0, 0, 720, 0, 0, 720)
Call oMN.getDocument().getView().selectAt(183.219131469727, 380.946594238281, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "outer_airbox"
Call oMN.getDocument().getView().makeComponentInALine(412.5, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("outer_airbox", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set outer_airbox Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("outer_airbox", 50)
Call oMN.getDocument().setCurvatureRefinementAngle("outer_airbox", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'AIRBOX 4
Call oMN.getDocument().getView().newArc(0, 0, 247.5, 0, 0, 247.5)
Call oMN.getDocument().getView().newArc(0, 0, 72.5, 0, 0, 72.5)
Call oMN.getDocument().getView().selectAt(102.023254394531, 99.9560165405273, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "outer_airbox4"
Call oMN.getDocument().getView().makeComponentInALine(7.5 + gap * 0.75, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("outer_airbox4", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set outer_airbox4 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("outer_airbox4", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("outer_airbox4", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'AIRBOX 3
Call oMN.getDocument().getView().newArc(0, 0, 245, 0, 0, 245)
Call oMN.getDocument().getView().newArc(0, 0, 75, 0, 0, 75)
Call oMN.getDocument().getView().selectAt(102.023254394531, 89.7381744384766, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "airbox_3"
Call oMN.getDocument().getView().makeComponentInALine(7.5 + gap * 0.5, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("airbox_3", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set airbox_3 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("airbox_3", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("airbox_3", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'AIRBOX 2
Call oMN.getDocument().getView().newArc(0, 0, 242.5, 0, 0, 242.5)
Call oMN.getDocument().getView().newArc(0, 0, 77.5, 0, 0, 77.5)
Call oMN.getDocument().getView().selectAt(123.606964111328, 91.7817459106445, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "airbox_2"
Call oMN.getDocument().getView().makeComponentInALine(7.5 + gap * 0.25, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("airbox_2", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set airbox_2 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("airbox_2", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("airbox_2", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'AIRBOX 1
Call oMN.getDocument().getView().newArc(0, 0, 240, 0, 0, 240)
Call oMN.getDocument().getView().newArc(0, 0, 80, 0, 0, 80)
Call oMN.getDocument().getView().selectAt(90.7174987792969, 105.064933776855, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "airbox_1"
Call oMN.getDocument().getView().makeComponentInALine(7.5, ArrayOfValues, "Name=Virtual Air", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("airbox_1", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set airbox_1 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("airbox_1", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("airbox_1", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'AIRBOX between Stator and shaft
Call oMN.getDocument().getView().newArc(0, 0, 70, 0, 0, 70)
Call oMN.getDocument().getView().selectAt(8.79068374633789, 70.1438751220703, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "airbox_b/w_Stator_rotor"
Call oMN.getDocument().getView().makeComponentInALine(15, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("airbox_b/w_Stator_rotor", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set airbox_b/w_Stator_rotor Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("airbox_b/w_Stator_rotor", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("airbox_b/w_Stator_rotor", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'SHAFT
Call oMN.getDocument().getView().selectAt(26.7676753997803, 31.5569496154785, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "shaft"
Call oMN.getDocument().getView().makeComponentInALine(137.5, ArrayOfValues, "Name=AIR", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().getView().selectObject("shaft", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set shaft Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("shaft", 20)
Call oMN.getDocument().setCurvatureRefinementAngle("shaft", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup

'Stator
Call oMN.getDocument().getView().newArc(0, 0, 112, 0, 105.245573528022, 38.3062560524749)
Call oMN.getDocument().getView().newLine(105.245573528022, 38.3062560524749, 225.526228988618, 82.0848343981605)
Call oMN.getDocument().getView().newArc(11.9999999999999, 7.16708811960747E-13, 123.355287256601, 12, 119.482556413537, 31.4880940485596)
Call oMN.getDocument().getView().newArc(-12, 0, 227.699812265258, 12, 218.323096891429, 67.4631087212962)
Call oMN.getDocument().getView().newLine(119.482556413537, 31.4880940485596, 218.323096891429, 67.4631087212962)
Call oMN.getDocument().getView().newLine(123.355287256601, 12, 227.699812265258, 12)
Call oMN.getDocument().getView().selectAt(139.532440185547, 45.3670082092285, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "Stator_coil1"
Call oMN.getDocument().getView().makeComponentInALine(7.5, ArrayOfValues, "Name=Copper: 5.77e7 Siemens/meter", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().beginUndoGroup("Set Stator_coil1 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("Stator_coil1", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("Stator_coil1", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectAt(153.832321166992, 30.3384170532227, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "Stator_core1"
Call oMN.getDocument().getView().makeComponentInALine(7.5, ArrayOfValues, "Name=S416: 416 Grade stainless steel", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().beginUndoGroup("Set Stator_core1 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("Stator_core1", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("Stator_core1", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("Stator_coil1", oCst.infoSetSelection)
Call oMN.getDocument().getView().selectObject("Stator_core1", infoToggleInSelection)
Call oMN.getDocument().beginUndoGroup("Transform Component")
Call oMN.getDocument().rotateComponent(Array("Stator_coil1", "Stator_core1"), 0, 0, 0, 0, 0, 1, 5, 1)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Transform Component")
Call oMN.getDocument().rotateComponent(oMN.getDocument().copyComponent(Array("Stator_coil1", "Stator_core1"), 1), 0, 0, 0, 0, 0, 1, 30, 1)
Call oMN.getDocument().rotateComponent(oMN.getDocument().copyComponent(Array("Stator_coil1", "Stator_core1"), 1), 0, 0, 0, 0, 0, 1, 60, 1)
Call oMN.getDocument().endUndoGroup

Call oMN.getDocument().getView().selectObject("Stator_coil1 Copy#1", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set Stator_coil1 Copy#1 Properties", True)
Call oMN.getDocument().renameObject("Stator_coil1 Copy#1", "Stator_coil2")
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("Stator_coil1 Copy#2", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set Stator_coil1 Copy#2 Properties", True)
Call oMN.getDocument().renameObject("Stator_coil1 Copy#2", "Stator_coil3")
Call oMN.getDocument().endUndoGroup

Call oMN.getDocument().getView().selectObject("Stator_core1 Copy#1", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set Stator_core1 Copy#1 Properties", True)
Call oMN.getDocument().renameObject("Stator_core1 Copy#1", "Stator_core2")
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("Stator_core1 Copy#2", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Set Stator_core1 Copy#2 Properties", True)
Call oMN.getDocument().renameObject("Stator_core1 Copy#2", "Stator_core3")
Call oMN.getDocument().endUndoGroup


'Create current flow surface
Call oMN.getDocument().makeCurrentFlowSurfaceCoil(1, "Stator_coil1", Array(159.552261352539, 68.9253387451172, 3.75), Array(0.913545457642601, 0.4067366430758, 0))
Call oMN.getDocument().makeCurrentFlowSurfaceCoil(1, "Stator_coil2", Array(102.352752685547, 136.350921630859, 3.75), Array(0.587785252292473, 0.809016994374947, 0))
Call oMN.getDocument().makeCurrentFlowSurfaceCoil(1, "Stator_coil3", Array(19.8220195770264, 168.845169067383, 3.75), Array(0.104528463267654, 0.994521895368273, 0))

Call oMN.getDocument().beginUndoGroup("Set Coil#1 Properties", True)
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "Stator_coil1"
Call oMN.getDocument().addReferencePaths("Coil#1", ArrayOfValues)
Call oMN.getDocument().getView().selectObject("Stator_coil1", oCst.infoSetSelection)
Call oMN.getDocument().setParameter("Coil#1", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("Coil#1", "Current", "0", oCst.infoNumberParameter)
Call oMN.getDocument().setCoilType("Coil#1", oCst.infoStrandedCoil)
Call oMN.getDocument().setCoilNumberOfTurns("Coil#1", 25)
Call oMN.getDocument().endUndoGroup

Call oMN.getDocument().beginUndoGroup("Set Coil#2 Properties", True)
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "Stator_coil2"
Call oMN.getDocument().addReferencePaths("Coil#2", ArrayOfValues)
Call oMN.getDocument().getView().selectObject("Stator_coil2", oCst.infoSetSelection)
Call oMN.getDocument().setParameter("Coil#2", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("Coil#2", "Current", "0", oCst.infoNumberParameter)
Call oMN.getDocument().setCoilType("Coil#2", oCst.infoStrandedCoil)
Call oMN.getDocument().setCoilNumberOfTurns("Coil#2", 25)
Call oMN.getDocument().endUndoGroup

Call oMN.getDocument().beginUndoGroup("Set Coil#3 Properties", True)
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "Stator_coil3"
Call oMN.getDocument().addReferencePaths("Coil#3", ArrayOfValues)
Call oMN.getDocument().getView().selectObject("Stator_coil3", oCst.infoSetSelection)
Call oMN.getDocument().setParameter("Coil#3", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("Coil#3", "Current", "0", oCst.infoNumberParameter)
Call oMN.getDocument().setCoilType("Coil#3", oCst.infoStrandedCoil)
Call oMN.getDocument().setCoilNumberOfTurns("Coil#3", 25)
Call oMN.getDocument().endUndoGroup


'Rotor
Call oMN.getDocument().getView().selectAt(44.6130790710449, 108.553886413574, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(25.9767894744873, 65.4348907470703, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(29.9525318145752, 68.1452255249023, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(30.9464683532715, 69.377197265625, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(32.9343376159668, 71.3483505249023, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(136.303634643555, 24.2870597839355, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(117.667343139648, 32.4180717468262, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(6)
ArrayOfValues(0) = "Rotor"
ArrayOfValues(1) = "Component#2"
ArrayOfValues(2) = "Component#3"
ArrayOfValues(3) = "Component#4"
ArrayOfValues(4) = "Component#5"
ArrayOfValues(5) = "Component#6"
ArrayOfValues(6) = "Component#7"
Call oMN.getDocument().getView().makeComponentInALine(9, ArrayOfValues, "Name=S416: 416 Grade stainless steel", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices)
Call oMN.getDocument().beginUndoGroup("Set Rotor Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("Rotor", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("Rotor", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("Rotor", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Transform Component")
Call oMN.getDocument().shiftComponent(Array("Rotor"), 0, 0, 7.5 + gap + pm_thic, 1)
Call oMN.getDocument().endUndoGroup


'Permanent Magnets
Call oMN.getDocument().getView().newArc(0, 0, 80, 0, 80 * Cos(pm_angle * WorksheetFunction.Pi() / 180), 80 * Sin(pm_angle * WorksheetFunction.Pi() / 180)) ' Construct PM according to  the PM angle parameter
Call oMN.getDocument().getView().newLine(80 * Cos(pm_angle * WorksheetFunction.Pi() / 180), 80 * Sin(pm_angle * WorksheetFunction.Pi() / 180), 240 * Cos(pm_angle * WorksheetFunction.Pi() / 180), 240 * Sin(pm_angle * WorksheetFunction.Pi() / 180))
Call oMN.getDocument().getView().selectAt(89.2830581665039, 7.72557830810547, oCst.infoSetSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(117.771591186523, 10.7677736282349, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
Call oMN.getDocument().getView().selectAt(153.272689819336, 18.1559619903564, oCst.infoAddToSelection, Array(oCst.infoSliceSurface))
ReDim ArrayOfValues(2)
ArrayOfValues(0) = "PM 1"
ArrayOfValues(1) = "Component#8"
ArrayOfValues(2) = "Component#9"
Call oMN.getDocument().getView().makeComponentInALine(pm_thic, ArrayOfValues, "Name=" & mat & ";Type=Uniform;Direction=[0,0,1]", oCst.infoMakeComponentUnionSurfaces Or oCst.infoMakeComponentRemoveVertices) ' Construct PM according to  the PM thickness and material parameter
Call oMN.getDocument().beginUndoGroup("Set PM 1 Properties", True)
Call oMN.getDocument().setBatchUpdatesInGUI(True)
Call oMN.getDocument().setMaxElementSize("PM 1", 10)
Call oMN.getDocument().setCurvatureRefinementAngle("PM 1", 5)
Call oMN.getDocument().setBatchUpdatesInGUI(False)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Transform Component")
Call oMN.getDocument().shiftComponent(Array("PM 1"), 0, 0, 7.5 + gap, 1)
Call oMN.getDocument().rotateComponent(Array("PM 1"), 0, 0, 0, 0, 0, 1, 45 - pm_angle, 1)
Call oMN.getDocument().rotateComponent(oMN.getDocument().copyComponent(Array("PM 1"), 1), 0, 0, 0, 0, 0, 1, 45, 1)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Set PM 1 Copy#1 Properties", True)
Call oMN.getDocument().renameObject("PM 1 Copy#1", "PM 2")
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Set PM 2 Properties", True)
Call oMN.getDocument().assignMaterial("PM 2", "Name=" & mat & " ;Type=Uniform;Direction=[0,0,-1];ReverseMagnetizationDirection=No")
Call oMN.getDocument().endUndoGroup


'Impose Boundary Conditions
Call oMN.getDocument().getView().selectObject("outer_airbox,Face#1", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Assign Boundary Condition")
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "outer_airbox,Face#1"
Call oMN.getDocument().createBoundaryCondition(ArrayOfValues, "BoundaryCondition#2")
Call oMN.getDocument().setMagneticFieldNormal("BoundaryCondition#2")
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("outer_airbox,Face#2", oCst.infoSetSelection)
Call oMN.getDocument().getView().selectObject("outer_airbox,Face#4", oCst.infoToggleInSelection)
Call oMN.getDocument().beginUndoGroup("Assign Boundary Condition")
ReDim ArrayOfValues(1)
ArrayOfValues(0) = "outer_airbox,Face#2"
ArrayOfValues(1) = "outer_airbox,Face#4"
Call oMN.getDocument().createBoundaryCondition(ArrayOfValues, "BoundaryCondition#3")
Call oMN.getDocument().setMagneticFluxTangential("BoundaryCondition#3")
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().getView().selectObject("outer_airbox,Face#5", oCst.infoSetSelection)
Call oMN.getDocument().beginUndoGroup("Assign Boundary Condition")
ReDim ArrayOfValues(0)
ArrayOfValues(0) = "outer_airbox,Face#5"
Call oMN.getDocument().createBoundaryCondition(ArrayOfValues, "BoundaryCondition#4")
ReDim RotationAxis(2)
RotationAxis(0) = 0
RotationAxis(1) = 0
RotationAxis(2) = 1
ReDim Center(2)
Center(0) = 0
Center(1) = 0
Center(2) = 0
Call oMN.getDocument().setEvenPeriodic("BoundaryCondition#4", Null, -90, RotationAxis, Null, Null, Center)
Call oMN.getDocument().endUndoGroup

'For Static 3D solver
Call oMN.getDocument().beginUndoGroup("Set Properties", True)
Call oMN.getDocument().setParameter("", "angles", "0%deg, 5%deg, 10%deg, 15%deg", oCst.infoNumberParameter)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Set Properties", True)
Call oMN.getDocument().setParameter("Rotor", "RotationAngle", "%angles", oCst.infoNumberParameter)
Call oMN.getDocument().setParameter("PM 1", "RotationAngle", "%angles", oCst.infoNumberParameter)
Call oMN.getDocument().setParameter("PM 2", "RotationAngle", "%angles", oCst.infoNumberParameter)
Call oMN.getDocument().endUndoGroup

'Setup Circuit
Call oMN.getDocument().getCircuit().insertCoil("Coil#1", 204, 108)
Call oMN.getDocument().getCircuit().insertCoil("Coil#2", 204, 276)
Call oMN.getDocument().getCircuit().insertCoil("Coil#3", 204, 468)
Call oMN.getDocument().getCircuit().insertCurrentSource(228, 168)
Call oMN.getDocument().getCircuit().insertCurrentSource(228, 372)
Call oMN.getDocument().getCircuit().insertCurrentSource(228, 540)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I1,T1", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#1,T1", TX2, TY2)
ReDim XArrayOfValues(2)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 204
XArrayOfValues(2) = TX2
ReDim YArrayOfValues(2)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 168
YArrayOfValues(2) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#1,T2", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I1,T2", TX2, TY2)
ReDim XArrayOfValues(2)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 273
XArrayOfValues(2) = TX2
ReDim YArrayOfValues(2)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 108
YArrayOfValues(2) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I2,T1", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#2,T2", TX2, TY2)
ReDim XArrayOfValues(3)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 228
XArrayOfValues(2) = 249
XArrayOfValues(3) = TX2
ReDim YArrayOfValues(3)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 312
YArrayOfValues(2) = 312
YArrayOfValues(3) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I2,T2", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#2,T1", TX2, TY2)
ReDim XArrayOfValues(4)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 336
XArrayOfValues(2) = 336
XArrayOfValues(3) = 204
XArrayOfValues(4) = TX2
ReDim YArrayOfValues(4)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 372
YArrayOfValues(2) = 228
YArrayOfValues(3) = 228
YArrayOfValues(4) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I3,T1", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#3,T1", TX2, TY2)
ReDim XArrayOfValues(2)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 204
XArrayOfValues(2) = TX2
ReDim YArrayOfValues(2)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 540
YArrayOfValues(2) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("Coil#3,T2", TX1, TY1)
Call oMN.getDocument().getCircuit().getPositionOfTerminal("I3,T2", TX2, TY2)
ReDim XArrayOfValues(2)
XArrayOfValues(0) = TX1
XArrayOfValues(1) = 273
XArrayOfValues(2) = TX2
ReDim YArrayOfValues(2)
YArrayOfValues(0) = TY1
YArrayOfValues(1) = 468
YArrayOfValues(2) = TY2
Call oMN.getDocument().getCircuit().insertConnection(XArrayOfValues, YArrayOfValues)
Call oMN.getDocument().beginUndoGroup("Set I1 Properties", True)
Call oMN.getDocument().setParameter("I1", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("I1", "Current", "5", oCst.infoNumberParameter)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Set I2 Properties", True)
Call oMN.getDocument().setParameter("I2", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("I2", "Current", "10", oCst.infoNumberParameter)
Call oMN.getDocument().endUndoGroup
Call oMN.getDocument().beginUndoGroup("Set I3 Properties", True)
Call oMN.getDocument().setParameter("I3", "WaveFormType", "DC", oCst.infoStringParameter)
Call oMN.getDocument().setParameter("I3", "Current", "5", oCst.infoNumberParameter)
Call oMN.getDocument().endUndoGroup


'Run Solver
'Call oMN.getDocument().solveTransient3dWithMotion()
Call oMN.getDocument().solveStatic3d

'Save Results
Call oMN.getGlobalResultsView().exportData(oCst.infoDataForce, "Z:\body force.csv", oCst.infoDataFormatLocaleListSeparatorDelimitedLocaleDecimal)
'Close Document
Call oMN.Close(False)
'Call oMN.close(FALSE)


End Sub
