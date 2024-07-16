######################################################################################################################################
#
#	SCRIPT CODE: FM
#
#	SCRIPT TITLE: Evaluate Clinical Goals in Raystation
#
#   VERSION: 1.0
#
#   ORIGINAL SCRIPT WRITTEN BY: Fatima Mahmood
#
#   DESCRIPTION & VERSION HISTORY: N/a
#
#	v1.0: (FM) <Use CPython Interpreter> 
#
#                   _____________________________________________________________________________
#                           
#                           SCRIPT VALIDATION DATE IN RAYSTATION SHOULD MATCH FILE DATE
#                   _____________________________________________________________________________
#
######################################################################################################################################

from connect import *
import csv, datetime, sys
import os
import tkinter as tk
from tkinter import filedialog

##############################################################################

# Retrieve the path to RayStation.exe (This will only work if the script is run from within RayStation)
script_path = System.IO.Path.GetDirectoryName(sys.argv[0])
path = script_path.rsplit('\\', 1)[0]
sys.path.append(path)

# Import windows forms
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application, Form, Label, ComboBox, TextBox, Button, MessageBox
from System.Drawing import Point, Size, Font, FontStyle, Color

# Utility function to create 2-dimensional array
def create_array(m, n):
    return System.Array.CreateInstance(System.Object, m, n)

# Get the current patient and case
try:
    patient = get_current('Patient')
    case = get_current('Case')
except Exception as e:
    print('Patient and case have to be loaded. Exiting script.')
    print(f'Error: {e}')
    sys.exit()

# Retrieve patient ID 
patient_id = patient.PatientID

# Define a filename with a timestamp to ensure it's unique and includes an extension
filename = f"S:\\Cancer Services - Radiation Physics\\Auto-planning\\MVision\\Scripts\\clinical_goals_export_{patient_id}_{datetime.datetime.now().strftime('%d%m%y')}.csv"

try:
    with open(filename, "w") as f:
        # Write header to the CSV file
        f.write("Plan Name, ROI Name, Goal Criteria, Acceptance Level, Parameter Value, Clinical Goal Value, Goal Achieved\n")
        
        # Iterate through all plans in the current case
        for plan in case.TreatmentPlans:
            print(f"Processing plan: {plan.Name}")
            
            # Find the clinical goals, there is a separate entry in EvaluationFunctions for each of the clinical goals
            number_of_goals = len(plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions)
            print(f"Number of clinical goals in plan '{plan.Name}': {number_of_goals}")
            
            if number_of_goals == 0:
                print(f"No clinical goals found for plan '{plan.Name}'.")
                continue
            
            results = create_array(number_of_goals, 6)
            
            for idx in range(number_of_goals):
                goal = plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[idx]
                results[idx, 0] = goal.ForRegionOfInterest.Name # roi name
                results[idx, 1] = goal.PlanningGoal.GoalCriteria # criteria (e.g., at most/at least, etc.)
                if goal.PlanningGoal.Type == 'VolumeAtDose':
                    # For VolumeAtDose goals, multiply Level and Value by 100 to give value in %, divide limit by 100 to give value in Gy
                    results[idx, 2] = goal.PlanningGoal.AcceptanceLevel * 100 # level (e.g., At most 50% volume...)
                    results[idx, 3] = goal.PlanningGoal.ParameterValue / 100 # limit (e.g., ...at 40.8 Gy)
                    results[idx, 4] = goal.GetClinicalGoalValue() * 100 # actual value in plan
                else: 
                    # For Dose related goals, multiply limit by 100 to give value in %, divide Level and Value by 100 to give value in Gy        
                    results[idx, 2] = goal.PlanningGoal.AcceptanceLevel / 100 # level (e.g., at least 57Gy...)
                    results[idx, 3] = goal.PlanningGoal.ParameterValue * 100 # limit (e.g., ...at 98% volume)
                    results[idx, 4] = goal.GetClinicalGoalValue() / 100 # actual value in plan
                results[idx, 5] = str(goal.EvaluateClinicalGoal()) # is goal achieved? (TRUE/FALSE)
                
            # Write results to the file
            for i in range(number_of_goals):
                f.write(f"{plan.Name}, {results[i,0]}, {results[i,1]}, {results[i,2]:.2f}, {results[i,3]:.2f}, {results[i,4]:.2f}, {results[i,5]}\r\n")
                
    print(f"Clinical goals have been successfully exported to {filename}")
except Exception as e:
    print(f"Failed to write to file: {e}")
