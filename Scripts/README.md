Scripts for generating the models and storing the results:


  The excel sheet contains the original parameters and the results of simulating all the designs:
  
    Column1 to column 4 : torque on stator at 0,5,10 and 15 degrees respectively , column 5 to column 8: torque on rotor at 0,5,10 and 15 degrees respectively
    
    
  The visual basic files are the codes that run in the of the excel macro. Can also be accessed through the excel sheet. See manuals and tutorials folder for more info:
  
    iteration.vbs: iterates through parameters in the excel sheet and passes them to Script_axial_onefourth.vbs 
    Script_axial_onefourth.vbs: accepts parameters from script in iteration.vbs and creates and simulates the model in MagNet. exports the torque results to another excel file.
    import_results.vbs: opens excel file containing results and copies the required values to original excel sheet in the active row beside the parameter list.
