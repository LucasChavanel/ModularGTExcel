# ModularGTExcel
VBA code for the memoire "Modular GT Cost and performance optimisation"

The code is divided into three categories : 
Module that englobe all the functionning code, with optimisation, creation of cycles in Aspen, extraction of results, modification of datas etc...
UserForm (frm and frx) that are the interfaces used to create different cycles in Excel, to define the global specs and cycles to use in optimisation scenarios.
Class, to regroup our components, their parameters and that allows to calculate performance and costs. Cost formulas are implemented in class.

The Application does not work if an Aspen HYSYS software is installed on the computer.
