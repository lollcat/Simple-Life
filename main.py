import comtypes.client
import comtypes.gen
import numpy as np
from comtypes import COMError


# tell comtypes to load type libs
cofeTlb = ('{0D1006C7-6086-4838-89FC-FBDCC0E98780}', 1, 0)  # COFE type lib
cofeTypes = comtypes.client.GetModule(cofeTlb)
coTlb = ('{4A5E2E81-C093-11D4-9F1B-0010A4D198C2}', 1, 1)  # CAPE-OPEN v1.1 type lib
coTypes = comtypes.client.GetModule(coTlb)
doc = comtypes.client.CreateObject('COCO_COFE.Document', interface=cofeTypes.ICOFEDocument)


def Solve(doc):
    """
    This function simulates the flowsheet in COCO simulator
    """
    try:
        doc.Solve()
        return True  # solve was sucessful
    except COMError as err:
        return False  # solve failed


doc.Import('Simulator.fsd') # load the file
reactor = doc.GetUnitNames()[0] # Get the reactors name

SimulationInfo = []
for i in range(10):  # lets run this loop a few times so the program takes a vague amount of time to complete
    # Change the value of a unit parameter (i.e change the reactor's temperature to something random)
    new_temperature = np.random.uniform(low=400, high=500)
    doc.GetUnit(reactor).QueryInterface(coTypes.ICapeUtilities).Parameters.QueryInterface(coTypes.ICapeCollection).Item(
                "Temperature").QueryInterface(coTypes.ICapeParameter).value = float(new_temperature)
    sucessful_solve = Solve(doc) # solve the flowsheet with the new reactor temperature
    # let's say we are interested in the info contained in stream 2
    stream_info = doc.GetStream('2').QueryInterface(coTypes.ICapeThermoMaterial).GetOverallProp("flow", "mole")
    SimulationInfo.append((stream_info, sucessful_solve))
print(SimulationInfo)
